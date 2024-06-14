import os
import prettytable
from datetime import datetime
import re
import json

def generate_scc_info_docs(progress_file): # create the info docs in this format
    with open(progress_file, 'r') as file:
        progress_data = json.load(file) # open progress.json

    for scc_path, scc_info in progress_data['SCC'].items(): # for each SCC item in the SCC dictionary in progress.json
        scc_name = scc_info['SCC']
        scc_dir = os.path.join(os.path.dirname(progress_file), scc_name)
        doc_path = os.path.join(scc_dir, f"{scc_name}_info.md") # name for info doc

        attestations = {} # create dictionaries for writing to info files
        bpers = {}
        documents = {}

        for attestation_num, attestation_info in progress_data['Attestations'].items(): # grab all attestations that have the SCC we want
            if isinstance(attestation_info, dict) and attestation_info.get('SCC') == scc_name and not attestation_info.get('false_positive', False): # omit those marked as false positives
                attestations[attestation_num] = attestation_info
            elif isinstance(attestation_info, list): # added handling for lists because it was throwing a fit
                for attestation in attestation_info:
                    if isinstance(attestation, dict) and attestation.get('SCC') == scc_name and not attestation.get('false_positive', False):
                        attestations[attestation_num] = attestation

        progress_data['SCC'][scc_path]['Info Doc Path'] = doc_path

        for bper_name, bper_info in progress_data['BPERs'].items(): # grab all BPERs that have the SCC we want
            if isinstance(bper_info, dict) and bper_info.get('SCC') == scc_name and not bper_info.get('false_positive', False):
                bpers[bper_name] = bper_info
            elif isinstance(bper_info, list): # list handling
                for bper in bper_info:
                    if isinstance(bper, dict) and bper.get('SCC') == scc_name and not bper.get('false_positive', False):
                        bpers[bper_name] = bper

        for doc_name, doc_info_list in progress_data['Documents'].items(): # grab all Supporting Dcouments that have the SCC we want
            for doc_info in doc_info_list:
                if isinstance(doc_info, dict) and doc_info.get('SCC') == scc_name and not doc_info.get('false_positive', False):
                    documents[doc_name] = doc_info

        with open(doc_path, 'w') as doc_file: # Actual writing to the text file
            # top section
            doc_file.write(f"# {scc_name}\n\n")
            doc_file.write(f"**SCC Version:**            {scc_info['Version']}\n\n")
            doc_file.write(f"**SCM Name:**               {scc_info['SCM Name']}\n\n")
            doc_file.write(f"**Last Review Date:**       {scc_info['Last Review Date']}\n\n")
            doc_file.write(f"- [{'x' if scc_info['SCC Guidance source presence'] else ' '}] SCC Guidance source\n")
            doc_file.write(f"- [{'x' if scc_info['SCC Policy and Procedure presence'] else ' '}] SCC Policy and Procedure\n")
            doc_file.write(f"- [{'x' if scc_info['SCC System Scope Presence'] else ' '}] SCC System Scope Presence\n")
            doc_file.write(f"- [{'x' if scc_info['Exception column presence'] else ' '}] Exception Column\n")
            doc_file.write(f"- [{'x' if scc_info['Deviation column presence'] else ' '}] Deviation Column\n")
            doc_file.write(f"- [{'x' if scc_info['TLA column presence'] else ' '}] TLA Column\n")
            doc_file.write(f"- [{'x' if scc_info['Compliance method column presence'] else ' '}] Compliance Method Column\n")
            doc_file.write(f"- [{'x' if scc_info['WPS config sup doc presence'] else ' '}] WPS config sup doc\n\n")

            # Attestation section
            doc_file.write("## Attestations\n\n")
            doc_file.write("| Gathered | Attestation Number | Approval Status | Valid To  |\n")
            doc_file.write("| -------- | ------------------ | --------------- | --------- |\n")
            for attestation_num, attestation in attestations.items():
                gathered = 'x' if attestation.get('Gathered', False) else ' '
                approval_status = attestation.get('Approval Status', '').ljust(15)
                valid_to = attestation.get('Valid to', '').ljust(9)
                doc_file.write(f"| [{gathered}]      | {attestation_num[:18].ljust(18)} | {approval_status[:15]} | {valid_to[:9]} |\n")

            # BPER section
            doc_file.write("\n## BPERs\n\n")
            doc_file.write("| Gathered | BPER Name     | Approval Status | Valid To  | TLA |\n")
            doc_file.write("| -------- | ------------- | --------------- | --------- | --- |\n")
            for bper_name, bper in bpers.items():
                gathered = 'x' if bper.get('Gathered', False) else ' '
                approval_status = bper.get('Approval Status', '').ljust(15)
                valid_to = bper.get('Valid to', '').ljust(9)
                tla = 'x' if bper.get('TLA', False) else ' '
                doc_file.write(f"| [{gathered}]      | {bper_name[:13].ljust(13)} | {approval_status[:15]} | {valid_to[:9]} | [{tla}] |\n")

            # Doc Section
            doc_file.write("\n## Documents\n\n")
            doc_file.write("| Gathered | Document Name                                                               | Version | Last Update |\n")
            doc_file.write("| -------- | --------------------------------------------------------------------------- | ------- | ----------- |\n")
            for doc_name, doc in documents.items():
                gathered = 'x' if doc.get('Gathered', False) else ' '
                version = doc.get('Version', '').ljust(7)
                last_update = doc.get('Last update', '').ljust(11)
                doc_file.write(f"| [{gathered}]      | {doc_name[:75].ljust(75)} | {version[:7]} | {last_update[:11]} |\n")

            # Check section
            doc_file.write("\n## Checks\n\n")
            check_methods = {} # dictionary for checks
            for check_id, check_info in progress_data['Checks'].items():
                if check_info.get('SCC') == scc_name:
                    evidence_method = check_info.get('Evidence method', '')
                    if evidence_method not in check_methods:
                        check_methods[evidence_method] = []
                    check_methods[evidence_method].append(check_id) # grab all checks with the SCC we want

            if check_methods: # process all the checks for table formatting
                header = "| " + " | ".join(method.ljust(20) for method in check_methods.keys()) + " |\n"
                separator = "| " + " | ".join("-" * 20 for _ in check_methods.keys()) + " |\n"

                doc_file.write(header)
                doc_file.write(separator)

                max_rows = max(len(checks) for checks in check_methods.values())
                for i in range(max_rows):
                    row = []
                    for method in check_methods.keys():
                        if i < len(check_methods[method]):
                            row.append(check_methods[method][i].ljust(20))
                        else:
                            row.append(" " * 20)
                    doc_file.write("| " + " | ".join(row) + " |\n")
            else:
                doc_file.write("No checks found.\n")

    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)

def format_document_name(name, length=75):
    return name[:length] # supports spacing building the tables; limits the length of doc names

def process_section_with_checkbox(section_name, dict_data, dict_type): # checkboxes
    output = f"\n{section_name}:\n"

    sorted_dict_data = dict(sorted(dict_data.items())) # alphabetizes

    for key, value in sorted_dict_data.items():
        check_mark = '[x]' if value.get('Gathered') else '[ ]' # check box logic
        document_name = format_document_name(key)

        tla_mark = '  [ ] Add TLA Docs' if dict_type == 'bper' and value.get('TLA', False) else '' # TLA box

        if dict_type in ['bper', 'attestation']:
            valid_to = value.get('Valid to', '')[:9]
            total_length = len(document_name) + len(tla_mark)
            spaces_for_alignment = 55 - total_length
            output += f"\t\t{check_mark} {document_name}{tla_mark}{' ' * spaces_for_alignment}Valid to: {valid_to}\n"
        elif dict_type == 'doc':
            last_update = value.get('Last update', '')[:11]
            spaces_for_alignment = 55 - len(document_name)
            output += f"\t\t{check_mark} {document_name}{' ' * spaces_for_alignment}Last update: {last_update}\n"

    return output

def write_checklist(bper_dict, doc_dict, attestation_dict, method_dict, scc_info, master_directory): # pieces together the output for writing
    scc_info_output = process_scc_info(scc_info)
    bper_output = process_section_with_checkbox("BPERs", bper_dict, 'bper')
    doc_output = process_section_with_checkbox("Documents", doc_dict, 'doc')
    attestation_output = process_section_with_checkbox("Attestations", attestation_dict, 'attestation')
    method_output = process_method_section(method_dict)

    scc_name = os.path.splitext(os.path.basename(scc_info['SCC Name']))[0]
    sanitized_scc_name = re.sub(r'(_\d{2})$', '', scc_name).strip()
    checklist_file_path = os.path.join(master_directory, f"{sanitized_scc_name}.txt") # build path for the file

    with open(checklist_file_path, "w") as file: # write each part
        file.write(scc_info_output)
        file.write(attestation_output)
        file.write(bper_output)
        file.write(doc_output)
        file.write(method_output)

    print(f"Made table for {scc_name}")

def process_method_section(method_dict):
    table = prettytable.PrettyTable()
    method_stigs = {} # holder for evidence method stig ids

    for stig_id, details in method_dict.items():
        method = details['Method'].upper().replace('NA', 'N/A') # remove placeholders
        if method not in method_stigs:
            method_stigs[method] = []
        method_stigs[method].append(stig_id) # add unique values from the dict

    sorted_methods = sorted(method_stigs.keys())
    table.field_names = sorted_methods # alphabetize them

    if not method_stigs:
        return "No methods found.\n"
    else:
        max_length = max(len(lst) for lst in method_stigs.values())

    for i in range(max_length): # create the table by adding each stig id
        row = []
        for method in sorted_methods:
            row.append(method_stigs[method][i] if i < len(method_stigs[method]) else '')
        table.add_row(row)

    return f"{table}\n"

def sync_progress_info(progress_file): # pulls information from .md files and updates progress.json
    print("Syncing progress information...")
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)

    for scc_path, scc_info in progress_data['SCC'].items():
        scc_name = scc_info['SCC']
        scc_dir = os.path.join(os.path.dirname(progress_file), scc_name)
        doc_path = os.path.join(scc_dir, f"{scc_name}_info.md")

        print(f"Processing SCC: {scc_name}")

        if os.path.exists(doc_path):
            print(f"  Found info.md file: {doc_path}")
            with open(doc_path, 'r') as doc_file:
                doc_content = doc_file.read()

            print("  Updating Attestations...")
            for attestation_num, attestation_info in progress_data['Attestations'].items():
                if isinstance(attestation_info, dict) and attestation_info.get('SCC') == scc_name:
                    if f"| [x]      | {attestation_num[:18]}" in doc_content:
                        attestation_info['Gathered'] = True
                        print(f"    Marked {attestation_num} as gathered for SCC: {scc_name}")
                    else:
                        attestation_info['Gathered'] = False
                        print(f"    Marked {attestation_num} as not gathered for SCC: {scc_name}")
                elif isinstance(attestation_info, list):
                    for attestation in attestation_info:
                        if isinstance(attestation, dict) and attestation.get('SCC') == scc_name:
                            if f"| [x]      | {attestation_num[:18]}" in doc_content:
                                attestation['Gathered'] = True
                                print(f"    Marked {attestation_num} as gathered for SCC: {scc_name}")
                            else:
                                attestation['Gathered'] = False
                                print(f"    Marked {attestation_num} as not gathered for SCC: {scc_name}")

            print("  Updating BPERs...")
            for bper_name, bper_info_list in progress_data['BPERs'].items():
                for bper_info in bper_info_list:
                    if isinstance(bper_info, dict) and bper_info.get('SCC') == scc_name:
                        if f"| [x]      | {bper_name[:13]}" in doc_content:
                            bper_info['Gathered'] = True
                            print(f"    Marked {bper_name} as gathered for SCC: {scc_name}")
                        else:
                            bper_info['Gathered'] = False
                            print(f"    Marked {bper_name} as not gathered for SCC: {scc_name}")

            print("  Updating Documents...")
            for doc_name, doc_info_list in progress_data['Documents'].items():
                for doc_info in doc_info_list:
                    if isinstance(doc_info, dict) and doc_info.get('SCC') == scc_name:
                        if f"| [x]      | {doc_name[:75]}" in doc_content:
                            doc_info['Gathered'] = True
                            print(f"    Marked {doc_name} as gathered for SCC: {scc_name}")
                        else:
                            doc_info['Gathered'] = False
                            print(f"    Marked {doc_name} as not gathered for SCC: {scc_name}")
        else:
            print(f"  Info.md file not found for SCC: {scc_name}")

    print("Saving updated progress data...")
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)

    print("Sync completed.")

if __name__ == "__main__":
    main()
