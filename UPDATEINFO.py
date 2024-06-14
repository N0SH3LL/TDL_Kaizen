import os
import json
import re
import fitz
import SCCCHECK
import FILEGRAB
from datetime import datetime
from difflib import SequenceMatcher

def update_bper_info(bper_dict, base_directories): # grab info for BPERs and write to dict
    for key, value_list in bper_dict.items(): # goes through every item in every bper entry
        for value in value_list:
            if value.get('false_positive', False):
                print(f"Skipping BPER '{key}' marked as false positive.") # skip if marked as false pos
                continue

            if 'manually_linked' in value:
                file_path = value['manually_linked'] # use manually_linked path if assigned
            else:
                source_directory = base_directories['bper']
                file_path = os.path.join(source_directory, f"{key}.pdf")

            if os.path.isfile(file_path):
                valid_to_date, approval_status, tla_present = FILEGRAB.extract_BPER_info(file_path) # ***FIX*** Counterintuitively, the actual pulling of information comes from the FILEGRAB file; just where it started, hasn't been fixed yet. 
                value['Valid to'] = valid_to_date # write these values to the dictionaries
                value['Approval Status'] = approval_status
                value['TLA'] = tla_present
                value['Updated from filename'] = os.path.basename(file_path) # stores file info is from and date it grabbed it
                value['Updated from timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                print(f"Updated BPER '{key}' - 'Valid to': {valid_to_date}, 'Approval Status': {approval_status}, 'TLA': {tla_present}") # print out everything it updated
            else:
                print(f"File not found for BPER: {key}") # if BPER with that name isn't present

def convert_datetime_to_string(obj): # file starts throwing errors unless the datetime objects it gets are converted to strings. Annoying. 
    if isinstance(obj, dict):
        return {key: convert_datetime_to_string(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_datetime_to_string(item) for item in obj]
    elif isinstance(obj, datetime):
        return obj.isoformat()
    return obj

def update_attestation_info(attestation_dict, base_directories): # improved attestation info grabbing 
    for key, value_list in attestation_dict.items():
        for value in value_list: # skip the fals positives
            if value.get('false_positive', False):
                print(f"Skipping Attestation '{key}' marked as false positive.")
                continue

            if 'manually_linked' in value: # use manually linked file path if set
                file_path = value['manually_linked']
            else:
                source_directory = base_directories['attestation'] # otherwise use the attestation directory + attestation name + .pdf
                file_path = os.path.join(source_directory, f"{key}.pdf")

            if os.path.isfile(file_path):
                with fitz.open(file_path) as doc:
                    text = ""
                    for page in doc:
                        text += page.get_text() # if doc exists in directory, look through every page

                approval_status, valid_to_date, review_date, assessment_date, overall_status = FILEGRAB.extract_attest_info(text) # use FILEGRAB extract attestation info 
                if approval_status != "Status: Error": # as long as there is no error, write the values to the dictionary
                    value['Approval Status'] = approval_status
                    value['Valid to'] = valid_to_date
                    value['Review Date'] = review_date
                    value['Assessment Date'] = assessment_date
                    value['Overall Status'] = overall_status
                    value['Updated from filename'] = os.path.basename(file_path)
                    value['Updated from timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    print(f"Updated '{key}' with status: {approval_status}, valid to date: {valid_to_date}, review date: {review_date}, assessment date: {assessment_date}, overall status: {overall_status}")
                else:
                    print(f"Error extracting information for Attestation: {key}")
            else:
                print(f"File not found for Attestation: {key}")

def update_doc_info(doc_dict, base_directories): # grab info for documents and write to dict
    for doc_name, value_list in doc_dict.items(): # goes through every item in every document entry
        for value in value_list:
            if value.get('false_positive', False):
                print(f"Skipping Document '{value['Doc name']}' marked as false positive.") # skip if marked as false pos
                continue

            if 'manually_linked' in value:
                file_path = value['manually_linked'] # use manually_linked path if assigned
            else:
                source_directory = base_directories['doc']
                matching_files = [f for f in os.listdir(source_directory) if (f.endswith('.docx') or f.endswith('.doc') or f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.pdf'))] # has to handle additional file types
                if matching_files:
                    best_match = max(matching_files, key=lambda x: SequenceMatcher(None, doc_name, x).ratio())
                    match_ratio = SequenceMatcher(None, doc_name, best_match).ratio()
                    if match_ratio >= 0.8:
                        file_path = os.path.join(source_directory, best_match) # matching for document names
                    else:
                        print(f"No close match found for Document: {doc_name}") # no matches better than the ratio
                        continue
                else:
                    print(f"No matching file found for Document: {doc_name}") # no matches at all
                    continue

            if os.path.isfile(file_path):
                most_recent_date = FILEGRAB.extract_Doc_info(file_path) # ***FIX*** Counterintuitively, the actual pulling of information comes from the FILEGRAB file; just where it started, hasn't been fixed yet.
                if most_recent_date:
                    for entry in value_list: # write these values to dictionaries
                        entry['Last update'] = most_recent_date
                        entry['Version'] = extract_version(os.path.basename(file_path))
                        entry['Updated from filename'] = os.path.basename(file_path)
                        entry['Updated from timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    print(f"Updated '{doc_name}' with 'Last update': {most_recent_date}, 'Version': {extract_version(os.path.basename(file_path))}")
            else:
                print(f"File not found for Document: {doc_name}")

def extract_version(filename): # grabs the version number, according to WPS naming conventions
    version_match = re.search(r'_(\d{2})(?=\.docx$|\.doc$)', filename)
    return version_match.group(1) if version_match else ''

def update_scc_info(scc_dict, scc_dir): # compares version numbers, then runs scc check if there is a newer version
    for file_path, scc_info in scc_dict.items():
        # Extract the SCC name from the file path
        scc_name = os.path.splitext(os.path.basename(file_path))[0]
        scc_name = re.sub(r'_\d+$', '', scc_name)  # Remove the version number from the SCC name

        # Search for files with a similar name pattern in the specified directory
        matching_files = [f for f in os.listdir(scc_dir) if f.startswith(scc_name) and f.endswith('.xlsx')]

        if matching_files:
            try:
                latest_file = max(matching_files, key=lambda x: int(re.search(r'_(\d+)\.xlsx$', x).group(1)))
            except AttributeError:
                print(f"Skipping SCC file {scc_name} due to version number format mismatch.")
                continue
            
            latest_file_path = os.path.join(scc_dir, latest_file)

            # Extract the version number from the latest file
            version_match = re.search(r'_(\d+)\.xlsx$', latest_file)
            if version_match:
                file_version = int(version_match.group(1))
                # Compare the file version with the stored version in the SCC dictionary
                if 'Version' in scc_info:
                    stored_version = int(scc_info['Version'])
                    if file_version <= stored_version:
                        print(f"Skipping SCC file {latest_file} as it has a version less than or equal to the stored version.")
                        continue

            updated_scc_info = SCCCHECK.process_scc_file(latest_file_path) # run SCC check
            directory_built = scc_info.get('Directory built', False)
            scc_info.update(updated_scc_info)
            scc_info['Directory built'] = directory_built
            scc_info['Version'] = str(file_version)  # Update the stored version in the dictionary
        else:
            print(f"No matching SCC files found for: {scc_name}")

def update_progress_info(progress_file, base_directories=None, scc_dir=None): # writing the dictionaries to progress.json
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    scc_dict = progress_data.get('SCC', {})
    bper_dict = progress_data.get('BPERs', {})
    attestation_dict = progress_data.get('Attestations', {})
    doc_dict = progress_data.get('Documents', {})
    
    if base_directories:
        update_bper_info(bper_dict, base_directories)
        update_attestation_info(attestation_dict, base_directories)
        update_doc_info(doc_dict, base_directories)
    
    if scc_dir:
        update_scc_info(scc_dict, scc_dir)
    
    # write new data to dictionaries
    progress_data['BPERs'] = bper_dict
    progress_data['Attestations'] = attestation_dict
    progress_data['Documents'] = doc_dict
    progress_data['SCC'] = scc_dict

    program_settings = progress_data.get('Program Settings', {})
    program_settings['Pull Info Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    progress_data['Program Settings'] = program_settings
    
    with open(progress_file, 'w') as file:
        json.dump(convert_datetime_to_string(progress_data), file, indent=4) # datetime objects > strings
    
    print("Progress information updated successfully.")
