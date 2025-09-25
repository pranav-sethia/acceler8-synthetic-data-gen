import json
import itertools
import os
import copy
import random
import re
import openpyxl

# ==============================================================================
# --- CONFIGURATION: EDIT THESE THREE FILENAMES FOR EACH NEW CAPABILITY ---
# ==============================================================================

# 1. The JSON file containing the new capability's questions.
SOURCE_QUESTIONS_FILE = 'response.json'

# 2. The matching, pre-formatted Excel template for this capability.
TEMPLATE_EXCEL_FILE = 'template_SN.xlsx'

# 3. The folder where the generated persona files will be saved.
OUTPUT_DIR = 'personas_SN'

# ==============================================================================
# --- SCRIPT LOGIC (No need to edit below this line) ---
# ==============================================================================

# --- STATIC CONFIGURATION ---
STATE_SCORES = { 'LEARN': 1, 'GROW': 3, 'TEACH': 4 }
RANKING_STATE_SCORES = { 'LEARN': 1, 'GROW': 2, 'TEACH': 4 }
SUBJECTIVE_RESPONSES = {
    'LEARN': "I'm still developing my approach in this area.",
    'GROW': "I have a solid approach and am working on applying it more broadly.",
    'TEACH': "I confidently apply best practices and help mentor others on this topic."
}

# --- HELPER FUNCTIONS ---

def clean_html(raw_html):
    if raw_html:
        no_html = re.sub('<[^<]+?>', '', raw_html)
        no_zwsp = no_html.replace('\u200b', '')
        return no_zwsp.strip()
    return ""

def build_capability_data(source_data):
    if not source_data: return None, {}
    capability_name = source_data[0].get('capabilities', 'Unknown Capability')
    subcapability_map = {}
    for q_source in source_data:
        subcap_name = q_source['sub_capability']
        if subcap_name not in subcapability_map:
            subcapability_map[subcap_name] = []
        new_question = {
            "id": q_source.get('sl_no') or q_source.get('id'),
            "employeeQuestion": clean_html(q_source.get('question')),
            "employeeQuestionType": q_source.get('type'),
            "employeeOptions": [], "employeeResponse": "", "employeeScore": None
        }
        options = q_source.get('options')
        ranking = q_source.get('ranking')
        if options:
            if ranking and len(options) == len(ranking):
                for i, option in enumerate(options):
                    new_question['employeeOptions'].append({"value": option['value'].strip(), "score": ranking[i]})
            else:
                for option in options:
                    new_question['employeeOptions'].append({"value": option['value'].strip()})
        subcapability_map[subcap_name].append(new_question)
    
    capabilityData = {
        "capability": capability_name, "subCapabilities": [],
        "overallAssessment": {"employeeQuestion": "", "employeeResponse": "", "managerQuestion": "", "managerResponse": ""},
        "capabilityScores": {"employeeScore": None, "employeeStage": None, "managerScore": None, "managerStage": None}
    }
    return capabilityData, subcapability_map

def get_option_by_score(question, target_score):
    options = [opt for opt in question.get('employeeOptions', []) if 'score' in opt]
    if not options: return "N/A", None
    closest_option = min(options, key=lambda opt: abs(opt['score'] - target_score))
    return closest_option.get('value'), closest_option.get('score')

def handle_ranking_question(question, target_state):
    target_score = RANKING_STATE_SCORES[target_state]
    scorable_options = [opt for opt in question['employeeOptions'] if 'score' in opt]
    if not scorable_options: return "N/A", 0
    ideal_ranking = sorted(scorable_options, key=lambda x: x['score'])
    if target_score == 1 and len(ideal_ranking) > 1:
        ranked_options = [ideal_ranking[0]] + random.sample(ideal_ranking[1:], len(ideal_ranking) - 1)
    elif target_score == 2 and len(ideal_ranking) >= 2:
        ranked_options = ideal_ranking[:2] + ideal_ranking[2:][::-1]
    else:
        ranked_options = ideal_ranking
    response_text = "\n".join([opt['value'] for opt in ranked_options])
    return response_text, target_score

def get_stage(score):
    if score is None: return None
    if score <= 2.7: return 'LEARN'
    elif score <= 3.5: return 'GROW'
    else: return 'TEACH'

def recalculate_scores(persona_data):
    all_scores = []
    capability_data = persona_data['assessment_capability_results']['capabilityData']
    for subcap in capability_data['subCapabilities']:
        subcap_scores = [q['employeeScore'] for q in subcap['questions'] if q.get('employeeScore') is not None]
        if subcap_scores:
            avg_score = round(sum(subcap_scores) / len(subcap_scores), 2)
            subcap['employeeScore'] = avg_score
            all_scores.extend(subcap_scores)
    if all_scores:
        overall_avg = round(sum(all_scores) / len(all_scores), 2)
        capability_data['capabilityScores']['employeeScore'] = overall_avg
        capability_data['capabilityScores']['employeeStage'] = get_stage(overall_avg)
    return persona_data

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    try:
        with open(SOURCE_QUESTIONS_FILE, 'r') as f:
            source_data = json.load(f)
        template_workbook = openpyxl.load_workbook(TEMPLATE_EXCEL_FILE)
    except FileNotFoundError as e:
        print(f"❌ Error: A required file was not found. Please check the CONFIGURATION section. Details: {e}")
        return

    template_sheet = template_workbook.active
    subcapability_order = []
    for row in template_sheet.iter_rows(min_row=5, max_col=3, values_only=True):
        subcap_name = row[2]
        if subcap_name and subcap_name not in subcapability_order:
            subcapability_order.append(subcap_name)

    base_capability_data, subcapability_map = build_capability_data(source_data)
    if not base_capability_data:
        print("❌ Error: Could not build template from source data.")
        return
        
    states = ['LEARN', 'GROW', 'TEACH']
    combinations = list(itertools.product(states, repeat=len(subcapability_order)))
    print(f"Detected Capability: '{base_capability_data['capability']}'")
    print(f"Using sub-capability order from template: {', '.join(subcapability_order)}")
    print(f"Generating {len(combinations)} persona files...")

    for combo in combinations:
        capability_data_copy = copy.deepcopy(base_capability_data)
        
        for subcap_name in subcapability_order:
            capability_data_copy["subCapabilities"].append({
                "name": subcap_name,
                "questions": subcapability_map.get(subcap_name, []),
                "employeeScore": None
            })

        filename_key = "_".join([s[0] for s in combo])
        filename = f"{OUTPUT_DIR}/persona_{filename_key}.json"

        for i, subcap_name in enumerate(subcapability_order):
            target_state = combo[i]
            target_score = STATE_SCORES[target_state]
            subcap_block = capability_data_copy['subCapabilities'][i]
            
            for question in subcap_block['questions']:
                q_type = question['employeeQuestionType']
                if q_type == 'OBJECTIVE_RANKING':
                    response, score = handle_ranking_question(question, target_state)
                    question['employeeResponse'], question['employeeScore'] = response, score
                elif q_type == 'OBJECTIVE_SINGLE':
                    response, score = get_option_by_score(question, target_score)
                    question['employeeResponse'], question['employeeScore'] = response, score
                elif q_type == 'OBJECTIVE_MULTIPLE' and question['employeeOptions']:
                    question['employeeResponse'] = random.choice(question['employeeOptions'])['value']
                elif q_type == 'SUBJECTIVE':
                    question['employeeResponse'] = SUBJECTIVE_RESPONSES[target_state]

        final_payload = {
            "assessment_capability_results": { "metadata": {
                    "organizationId": "org-stoneX", "organizationName": "StoneX",
                    "employeeName": f"Persona {filename_key}", "employeeTenure": "1 year",
                    "capabilityAssessmentName": f"{base_capability_data['capability']} Assessment",
                    "capabilityAssessmentId": f"SYNTHETIC-{filename_key}", "capabilityAssessmentDate": "2025-09-15"
                }, "capabilityData": capability_data_copy },
            "is_employee_only_assessment": True }
        
        final_payload = recalculate_scores(final_payload)
        with open(filename, 'w') as f:
            json.dump(final_payload, f, indent=2)

    print(f"\n✅ Success! All {len(combinations)} files have been generated in the '{OUTPUT_DIR}' folder.")

if __name__ == '__main__':
    main()