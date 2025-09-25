import os
import json
import random
import requests
import time

API_URL = "http://acceler8-dev-alb-backend-fastapi-836340435.ap-southeast-5.elb.amazonaws.com/v1/capability/summary"
PERSONAS_DIR = "personas_SN"
NUM_TO_TEST = 3
RESULTS_FILE = "api_summary_results.json"

def main():
    try:
        all_persona_files = [f for f in os.listdir(PERSONAS_DIR) if f.endswith('.json')]
        if not all_persona_files:
            print(f"Error: No JSON files found in the '{PERSONAS_DIR}' directory.")
            return
    except FileNotFoundError:
        print(f"Error: The directory '{PERSONAS_DIR}' was not found. Make sure it's in the same location as this script.")
        return

    if len(all_persona_files) < NUM_TO_TEST:
        print(f"Warning: Found fewer files ({len(all_persona_files)}) than the number requested ({NUM_TO_TEST}). Testing all found files.")
        selected_files = all_persona_files
    else:
        selected_files = random.sample(all_persona_files, NUM_TO_TEST)

    print(f"Found {len(all_persona_files)} personas. Randomly selected {len(selected_files)} for testing...")
    
    all_results = []

    for filename in selected_files:
        filepath = os.path.join(PERSONAS_DIR, filename)
        print(f"\nProcessing: {filename}...")

        try:
            with open(filepath, 'r') as f:
                payload = json.load(f)

            response = requests.post(API_URL, json=payload, timeout=60)
            response.raise_for_status() 
            
            response_data = response.json()
            print(f"Success: Received summary for {filename}")

        except requests.exceptions.RequestException as e:
            print(f"Error calling API for {filename}: {e}")
            response_data = {"error": str(e)}
        except json.JSONDecodeError:
            print(f"Error: Could not decode JSON from API response for {filename}.")
            response_data = {"error": "Invalid JSON response from server."}
        except Exception as e:
            print(f"An unexpected error occurred with {filename}: {e}")
            response_data = {"error": str(e)}

        all_results.append({
            "persona_file": filename,
            "request_payload": payload,       
            "api_response": response_data      
        })
        
        time.sleep(1)

    with open(RESULTS_FILE, 'w') as f:
        json.dump(all_results, f, indent=2)

    print(f"\nAll results have been saved to '{RESULTS_FILE}'")

if __name__ == '__main__':
    main()