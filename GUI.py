import tkinter as tk
from tkinter import filedialog
import os
import json
import JSONTOEXCEL
import KAIZEN
import FILEGRAB
import UPDATEINFO
import SCCREAD
import SCCTABLES
import SPLITBPER
import re
from datetime import datetime
from openpyxl import Workbook

# directory variables
existing_project_dir = None
supporting_docs_dir = None
attestation_dir = None
bpers_dir = None
progress_file = None
scc_dir = None
project_dir = None
template_dir = None

def select_directory(prompt): # pop up for selecting dirs
    directory = filedialog.askdirectory(title=prompt)
    return directory

def show_welcome(): # welcome screen
    clear_frames() # hide other frames
    welcome_screen.pack(fill="both", expand=True) # show welcome screen

def show_options(): 
    clear_frames() # hide other frames
    options_screen.pack(fill="both", expand=True) # show options screen
    dashboard_screen.pack_forget() # hide dashboard screen

def show_gather_docs():
    clear_frames() # hide other frames
    gather_docs_screen.pack(fill="both", expand=True) # show gather docs screen

def show_dashboard():
    clear_frames() # hide other frames
    dashboard_screen.pack(fill="both", expand=True) # show dashboard screen
    refresh_dashboard() # update dashboard data

def refresh_dashboard():
    # Update SCC list
    scc_listbox.delete(0, tk.END) # clear current list
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        for scc_path, scc_data in progress_data['SCC'].items():
            scc_name = scc_data.get('SCC')
            if scc_name:
                scc_listbox.insert(tk.END, scc_name) # add SCC names to listbox

        categories = ['Documents', 'Attestations', 'BPERs']
        gathered_counts = []
        total_counts = []

        for category in categories:
            items = progress_data.get(category, {})
            gathered_count = sum(1 for item_data_list in items.values() for item_data in item_data_list if item_data.get('Gathered', False) and not item_data.get('false_positive', False))
            total_count = sum(1 for item_data_list in items.values() for item_data in item_data_list if not item_data.get('false_positive', False))
            gathered_counts.append(gathered_count)
            total_counts.append(total_count)

        pie_chart_text = ""
        for category, gathered_count, total_count in zip(categories, gathered_counts, total_counts):
            percentage = (gathered_count / total_count) * 100 if total_count > 0 else 0
            pie_chart_text += f"{category}: {gathered_count}/{total_count} ({percentage:.2f}%)\n"

        pie_chart_label.config(text=pie_chart_text)
    else:
        pie_chart_label.config(text="No progress file selected.")
    
    # Update "Items Not Gathered" lists
    not_gathered_attestations = []
    not_gathered_bpers = []
    not_gathered_documents = []
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        for item_type, item_dict in [('Attestations', progress_data.get('Attestations', {})),
                                     ('BPERs', progress_data.get('BPERs', {})),
                                     ('Documents', progress_data.get('Documents', {}))]:
            for item_id, item_data_list in item_dict.items():
                for item_data in item_data_list:
                    if not item_data.get('Gathered', True) and not item_data.get('false_positive', False):
                        item_name = item_data.get('BPER name') or item_data.get('Attestation num') or item_data.get('Doc name')
                        scc = item_data.get('SCC')
                        if item_type == 'Attestations':
                            not_gathered_attestations.append(f"{item_name} - {scc}")
                        elif item_type == 'BPERs':
                            not_gathered_bpers.append(f"{item_name} - {scc}")
                        else:
                            not_gathered_documents.append(f"{item_name} - {scc}")
    
    not_gathered_attestations_listbox.delete(0, tk.END) # clear current list
    for item in not_gathered_attestations:
        not_gathered_attestations_listbox.insert(tk.END, item) # add items to listbox
    
    not_gathered_bpers_listbox.delete(0, tk.END) # clear current list
    for item in not_gathered_bpers:
        not_gathered_bpers_listbox.insert(tk.END, item) # add items to listbox
    
    not_gathered_documents_listbox.delete(0, tk.END) # clear current list
    for item in not_gathered_documents:
        not_gathered_documents_listbox.insert(tk.END, item) # add items to listbox
    
    # Update date labels
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        program_settings = progress_data.get('Program Settings', {})
        last_info_pull_date = program_settings.get('Pull Info Date', 'N/A')
        last_doc_pull_date = program_settings.get('Gather and Sort Date', 'N/A')
        last_checklist_generated_date = program_settings.get('Checklists generated', 'N/A')
    else:
        last_info_pull_date = 'N/A'
        last_doc_pull_date = 'N/A'
        last_checklist_generated_date = 'N/A'
    
    last_info_pull_label.config(text=f"Last Info Pull: {last_info_pull_date}")
    last_doc_pull_label.config(text=f"Last Doc Pull: {last_doc_pull_date}")
    last_checklist_generated_label.config(text=f"Last Checklist Generated: {last_checklist_generated_date}")
    
    # Update pie chart (placeholder)
    # Add code here to update the pie chart based on the latest data

def start_new_project():
    global project_dir
    project_dir = select_directory("Select the project directory") # select project directory
    if project_dir:
        scc_repo = select_directory("Select the SCC Repository") # select SCC repository
        if scc_repo:
            global progress_file
            progress_file = os.path.join(project_dir, "progress.json") # set progress file path
            KAIZEN.build_progress_json(scc_repo, project_dir) # build initial progress.json
            
            with open(progress_file, 'r') as file:
                progress_data = json.load(file)
            
            program_settings = progress_data.get('Program Settings', {})
            program_settings['SCC Directory'] = scc_repo # set SCC directory
            program_settings['Project Directory'] = project_dir # set project directory
            progress_data['Program Settings'] = program_settings
            
            with open(progress_file, 'w') as file:
                json.dump(progress_data, file, indent=4)
            
            load_project_settings() # load project settings
            update_directory_labels() # update directory labels
            show_options() # show options screen
        else:
            error_label.config(text="Please select a valid SCC Repository.")
    else:
        error_label.config(text="Please select a valid project directory.")

def generate_md_files():
    if progress_file and project_dir:
        SCCTABLES.generate_scc_info_docs(progress_file) # generate markdown files
        
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Checklists generated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings
        
        # Store the generated file paths in the respective "SCC" dictionary entry
        for scc_file, scc_data in progress_data.get('SCC', {}).items():
            scc_name = scc_data.get('SCC')
            if scc_name:
                md_file_path = os.path.join(project_dir, f"{scc_name}_info.md")
                if os.path.exists(md_file_path):
                    scc_data['Info Doc Path'] = md_file_path
        
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        generate_md_status.config(text=program_settings['Checklists generated']) # update status label
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")

def update_directory_labels():
    # Update directory labels with the selected paths
    if progress_file:
        progress_file_label.config(text=f"Progress File: {progress_file}")
    else:
        progress_file_label.config(text="Progress File: Not selected")

    if bpers_dir:
        bpers_dir_label.config(text=f"BPERs Directory: {bpers_dir}")
    else:
        bpers_dir_label.config(text="BPERs Directory: Not selected")

    if attestation_dir:
        attestation_dir_label.config(text=f"Attestation Directory: {attestation_dir}")
    else:
        attestation_dir_label.config(text="Attestation Directory: Not selected")

    if supporting_docs_dir:
        supporting_docs_dir_label.config(text=f"Supporting Documents Directory: {supporting_docs_dir}")
    else:
        supporting_docs_dir_label.config(text="Supporting Documents Directory: Not selected")

    if project_dir:
        project_dir_label.config(text=f"Project Directory: {project_dir}")
    else:
        project_dir_label.config(text="Project Directory: Not selected")

    if scc_dir:
        scc_dir_label.config(text=f"SCC Directory: {scc_dir}")
    else:
        scc_dir_label.config(text="SCC Directory: Not selected")

def mark_as_false_positive(item_type):
    # Mark selected items as false positive based on item type
    if item_type == "BPERs":
        selected_items = [not_gathered_bpers_listbox.get(idx) for idx in not_gathered_bpers_listbox.curselection()]
        listbox = not_gathered_bpers_listbox
        dict_key = "BPERs"
    elif item_type == "Attestations":
        selected_items = [not_gathered_attestations_listbox.get(idx) for idx in not_gathered_attestations_listbox.curselection()]
        listbox = not_gathered_attestations_listbox
        dict_key = "Attestations"
    elif item_type == "Documents":
        selected_items = [not_gathered_documents_listbox.get(idx) for idx in not_gathered_documents_listbox.curselection()]
        listbox = not_gathered_documents_listbox
        dict_key = "Documents"
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for selected_item in selected_items:
        item_name, scc = selected_item.split(" - ")
        
        for item_id, item_data_list in progress_data[dict_key].items():
            for item_data in item_data_list:
                if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                    item_data["false_positive"] = True
                    break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)
    
    for idx in reversed(listbox.curselection()): # remove marked items from listbox
        listbox.delete(idx)
    
    item_name, scc = selected_item.split(" - ")
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for item_id, item_data_list in progress_data[dict_key].items():
        for item_data in item_data_list:
            if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                item_data["false_positive"] = True
                break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)
    
    listbox.delete(listbox.curselection())

def manually_link_files(item_type):
    # Manually link selected files to items based on item type
    if item_type == "BPERs":
        selected_items = [not_gathered_bpers_listbox.get(idx) for idx in not_gathered_bpers_listbox.curselection()]
        listbox = not_gathered_bpers_listbox
        dict_key = "BPERs"
        directory = bpers_dir
    elif item_type == "Attestations":
        selected_items = [not_gathered_attestations_listbox.get(idx) for idx in not_gathered_attestations_listbox.curselection()]
        listbox = not_gathered_attestations_listbox
        dict_key = "Attestations"
        directory = attestation_dir
    elif item_type == "Documents":
        selected_items = [not_gathered_documents_listbox.get(idx) for idx in not_gathered_documents_listbox.curselection()]
        listbox = not_gathered_documents_listbox
        dict_key = "Documents"
        directory = supporting_docs_dir
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for selected_item in selected_items:
        item_name, scc = selected_item.split(" - ")
        
        file_path = filedialog.askopenfilename(initialdir=directory, title=f"Select file for {item_name}")
        
        if file_path:
            for item_id, item_data_list in progress_data[dict_key].items():
                if len(item_data_list) > 1:  # Check if there are multiple sub-values
                    for item_data in item_data_list:
                        if item_data.get("Doc name") == item_name:
                            for sub_item_data in item_data_list:
                                sub_item_data["manually_linked"] = file_path
                            break
                else:
                    for item_data in item_data_list:
                        if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                            item_data["manually_linked"] = file_path
                            break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)

def update_existing_project():
    global progress_file, project_dir
    progress_file = filedialog.askopenfilename(title="Select the progress.json file", filetypes=[("JSON Files", "*.json")])
    if progress_file:
        project_dir = os.path.dirname(progress_file) # set project directory
        load_project_settings() # load project settings
        update_directory_labels() # update directory labels
        show_options() # show options screen
    else:
        error_label.config(text="Please select a valid progress.json file.")

def load_project_settings():
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            program_settings = progress_data.get('Program Settings', {})
            global scc_dir, bpers_dir, attestation_dir, supporting_docs_dir, template_dir
            scc_dir = program_settings.get('SCC Directory', '')
            bpers_dir = program_settings.get('BPERs Directory', '')
            attestation_dir = program_settings.get('Attestation Directory', '')
            supporting_docs_dir = program_settings.get('Supporting Documents Directory', '')
            template_dir = program_settings.get('Template Directory', '')
            update_directory_labels() # update directory labels
            update_status_labels(program_settings) # update status labels
            
            # Update SCC list
            scc_list = list(progress_data.get('SCC', {}).keys())
            scc_listbox.delete(0, tk.END) # clear current list
            for scc_path in scc_list:
                scc_data = progress_data.get('SCC', {}).get(scc_path, {})
                scc_name = scc_data.get('SCC')
                if scc_name:
                    scc_listbox.insert(tk.END, scc_name) # add SCC names to listbox
            
            # Update "Items Not Gathered" lists
            not_gathered_attestations = []
            not_gathered_bpers = []
            not_gathered_documents = []
            for item_type, item_dict in [('Attestations', progress_data.get('Attestations', {})),
                                        ('BPERs', progress_data.get('BPERs', {})),
                                        ('Documents', progress_data.get('Documents', {}))]:
                for item_id, item_data_list in item_dict.items():
                    for item_data in item_data_list:
                        if not item_data.get('Gathered', True) and not item_data.get('false_positive', False):
                            item_name = item_data.get('BPER name') or item_data.get('Attestation num') or item_data.get('Doc name')
                            scc = item_data.get('SCC')
                            if item_type == 'Attestations':
                                not_gathered_attestations.append(f"{item_name} - {scc}")
                            elif item_type == 'BPERs':
                                not_gathered_bpers.append(f"{item_name} - {scc}")
                            else:
                                not_gathered_documents.append(f"{item_name} - {scc}")
            
            not_gathered_attestations_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_attestations:
                not_gathered_attestations_listbox.insert(tk.END, item) # add items to listbox
            
            not_gathered_bpers_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_bpers:
                not_gathered_bpers_listbox.insert(tk.END, item) # add items to listbox
            
            not_gathered_documents_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_documents:
                not_gathered_documents_listbox.insert(tk.END, item) # add items to listbox
            
            # Update date labels
            last_info_pull_date = program_settings.get('Pull Info Date', 'N/A')
            last_doc_pull_date = program_settings.get('Gather and Sort Date', 'N/A')
            last_checklist_generated_date = program_settings.get('Checklists generated', 'N/A')
            last_info_pull_label.config(text=f"Last Info Pull: {last_info_pull_date}")
            last_doc_pull_label.config(text=f"Last Doc Pull: {last_doc_pull_date}")
            last_checklist_generated_label.config(text=f"Last Checklist Generated: {last_checklist_generated_date}")

def update_status_labels(program_settings):
    # Update status labels with program settings
    directories_built = program_settings.get('Directories Built', False)
    templates_built = program_settings.get('Templates Built', False)
    gather_sort_date = program_settings.get('Gather and Sort Date', 'Not Done')
    checklist_generated = program_settings.get('Checklists generated', 'Not Done')
    pull_info_date = program_settings.get('Pull Info Date', 'Not Done')

    build_dirs_status.config(text="Done" if directories_built else "Not Done")
    build_templates_status.config(text="Done" if templates_built else "Not Done")
    gather_docs_status.config(text=gather_sort_date)
    generate_md_status.config(text=checklist_generated)
    pull_info_status.config(text=pull_info_date)

def save_project_settings():
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['SCC Directory'] = scc_dir
        program_settings['BPERs Directory'] = bpers_dir
        program_settings['Attestation Directory'] = attestation_dir
        program_settings['Supporting Documents Directory'] = supporting_docs_dir
        program_settings['Template Directory'] = template_dir
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

def select_template_directory():
    global template_dir
    template_dir = select_directory("Please select the Template directory") # select template directory
    if template_dir:
        template_dir_label.config(text=f"Template Directory: {template_dir}") # update label
        save_project_settings() # save settings

def select_scc_directory():
    global scc_dir
    scc_dir = select_directory("Please select the SCC directory") # select SCC directory
    if scc_dir:
        scc_dir_label.config(text=f"SCC Directory: {scc_dir}") # update label
        save_project_settings() # save settings

def select_bpers_directory():
    global bpers_dir
    bpers_dir = select_directory("Please select the BPERs directory") # select BPERs directory
    if bpers_dir:
        bpers_dir_label.config(text=f"BPERs Directory: {bpers_dir}") # update label
        save_project_settings() # save settings

def select_attestation_directory():
    global attestation_dir
    attestation_dir = select_directory("Please select the Attestation directory") # select attestation directory
    if attestation_dir:
        attestation_dir_label.config(text=f"Attestation Directory: {attestation_dir}") # update label
        save_project_settings() # save settings

def select_supporting_docs_directory():
    global supporting_docs_dir
    supporting_docs_dir = select_directory("Please select the Supporting Documents directory") # select supporting docs directory
    if supporting_docs_dir:
        supporting_docs_dir_label.config(text=f"Supporting Documents Directory: {supporting_docs_dir}") # update label
        save_project_settings() # save settings

def select_progress_file():
    global progress_file
    progress_file = filedialog.askopenfilename(title="Select the progress.json file", filetypes=[("JSON Files", "*.json")])
    if progress_file:
        progress_file_label.config(text=f"Progress File: {progress_file}") # update label
        load_project_settings() # load settings
        update_directory_labels() # update directory labels

def select_project_directory():
    global project_dir
    project_dir = select_directory("Select the project directory") # select project directory
    if project_dir:
        project_dir_label.config(text=f"Project Directory: {project_dir}") # update label
        save_project_settings() # save settings

def pull_information():
    if progress_file:
        base_directories = {
            'bper': bpers_dir,
            'attestation': attestation_dir,
            'doc': supporting_docs_dir
        }
        SPLITBPER.process_directory(bpers_dir) # make sure BPERs are split, or split them
        UPDATEINFO.update_progress_info(progress_file, base_directories, scc_dir)
        
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Pull Info Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings
        
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        update_status_labels(program_settings) # update status labels
    else:
        error_label.config(text="Please select a valid progress.json file.")

def sync_button_click():
    if progress_file:
        SCCTABLES.sync_progress_info(progress_file) # sync progress info
    else:
        error_label.config(text="Please select a valid progress.json file.")

def build_dirs():
    if progress_file and project_dir:
        KAIZEN.create_directories(project_dir) # create directories

        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        program_settings = progress_data.get('Program Settings', {})
        program_settings['Directories Built'] = True
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        build_dirs_status.config(text="Done") # update status label
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")

def build_templates():
    if progress_file and project_dir and template_dir:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            method_dict = progress_data.get('Checks', {})
            KAIZEN.build_templates(method_dict, project_dir, template_dir) # build templates

        program_settings = progress_data.get('Program Settings', {})
        program_settings['Templates Built'] = True
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        build_templates_status.config(text="Done") # update status label
    else:
        error_label.config(text="Please select a valid progress.json file, project directory, and template directory.")

def gather_docs():
    if progress_file and bpers_dir and attestation_dir and supporting_docs_dir:
        SPLITBPER.process_directory(bpers_dir) # process BPERs directory

        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            bper_dict = progress_data.get('BPERs', {})
            doc_dict = progress_data.get('Documents', {})
            attestation_dict = progress_data.get('Attestations', {})
            base_directories = {'bper': bpers_dir, 'doc': supporting_docs_dir, 'attestation': attestation_dir}
            
            updated_bper_dict, updated_doc_dict, updated_attestation_dict = FILEGRAB.update_dictionaries_and_copy_files(bper_dict, doc_dict, attestation_dict, base_directories, project_dir)
            
            progress_data['BPERs'] = updated_bper_dict
            progress_data['Documents'] = updated_doc_dict
            progress_data['Attestations'] = updated_attestation_dict

            program_settings = progress_data.get('Program Settings', {})
            pull_info_date = program_settings.get('Pull Info Date')

            if pull_info_date:
                remove_lower_versions(supporting_docs_dir) # remove older versions

            program_settings['Gather and Sort Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            progress_data['Program Settings'] = program_settings
            
            with open(progress_file, 'w') as file:
                json.dump(progress_data, file, indent=4)
            
            gather_docs_status.config(text=program_settings['Gather and Sort Date']) # update status label
    else:
        error_label.config(text="Please select a valid progress.json file and all required directories.")

def remove_lower_versions(directory):
    # Remove lower versions of files in the specified directory
    files = os.listdir(directory)
    file_versions = {}

    for file in files:
        if file.endswith(('.docx', '.doc')):
            file_name, file_extension = os.path.splitext(file)
            version_match = re.search(r'_(\d{2})$', file_name)
            if version_match:
                version = int(version_match.group(1))
                base_name = file_name[:version_match.start()]
                if base_name not in file_versions or version > file_versions[base_name][0]:
                    file_versions[base_name] = (version, file)

    for file in files:
        if file.endswith(('.docx', '.doc')):
            file_name, file_extension = os.path.splitext(file)
            version_match = re.search(r'_(\d{2})$', file_name)
            if version_match:
                base_name = file_name[:version_match.start()]
                if base_name in file_versions and file != file_versions[base_name][1]:
                    file_path = os.path.join(directory, file)
                    os.remove(file_path) # remove lower version file
                    print(f"Removed lower version: {file}")

def remove_scc():
    remove_scc_window = tk.Toplevel(root) # create new window for SCC removal
    remove_scc_window.title("Remove an SCC")
    remove_scc_window.geometry("400x300")

    scc_list_frame = tk.Frame(remove_scc_window)
    scc_list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    scc_list_label = tk.Label(scc_list_frame, text="Select an SCC to remove:")
    scc_list_label.pack()

    scc_list_listbox = tk.Listbox(scc_list_frame, font=("Arial", 10))
    scc_list_listbox.pack(fill="both", expand=True)

    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        for scc_path in progress_data.get('SCC', {}):
            scc_data = progress_data['SCC'][scc_path]
            scc_name = scc_data.get('SCC')
            if scc_name:
                scc_list_listbox.insert(tk.END, scc_name) # add SCC names to listbox

    delete_button = tk.Button(remove_scc_window, text="Delete", command=lambda: delete_scc(scc_list_listbox.get(scc_list_listbox.curselection())))
    delete_button.pack(pady=10)

def delete_scc(scc_name):
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        for key in ['BPERs', 'Attestations', 'Documents']:
            updated_dict = {}
            for item_key, item_list in progress_data[key].items():
                updated_list = [item for item in item_list if item.get('SCC') != scc_name]
                if updated_list:
                    updated_dict[item_key] = updated_list
            progress_data[key] = updated_dict

        progress_data['SCC'] = {k: v for k, v in progress_data['SCC'].items() if v.get('SCC') != scc_name}
        progress_data['Checks'] = {k: v for k, v in progress_data['Checks'].items() if v.get('SCC') != scc_name}

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        error_label.config(text=f"SCC '{scc_name}' removed successfully.")
        load_project_settings()  # Refresh the dashboard after removing an SCC
    else:
        error_label.config(text="Please select a valid progress.json file.")
        error_label.config(text="Please select a valid progress.json file.")

def sort_docs():
    if supporting_docs_dir and attestation_dir and bpers_dir:
        # Perform sorting actions with the selected directories
        pass
    else:
        error_label.config(text="Please select all required directories.")

def update_tracker():
    tracker_file = select_directory("Select the Document Tracker File")
    if tracker_file:
        # Perform actions with the selected Document Tracker File
        pass

def clear_frames():
    for frame in (welcome_screen, options_screen, gather_docs_screen):
        frame.pack_forget() # hide all specified frames

def add_or_redo_scc():
    # Open file explorer dialog to select an Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        # Extract SCC name from the file path and remove extension and trailing "_**"
        scc_name = os.path.splitext(os.path.basename(file_path))[0]
        scc_name = re.sub(r'_\d{2}$', '', scc_name).strip()
        
        # Load progress data from progress.json
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        # Check if the SCC exists in progress data
        if scc_name in progress_data['SCC']:
            # Remove all entries with the matching SCC name
            for key in ['BPERs', 'Attestations', 'Documents', 'SCC', 'Checks']:
                progress_data[key] = {k: v for k, v in progress_data[key].items() if v.get('SCC') != scc_name}
        
        # Process the selected Excel file
        bper_dict, doc_dict, attestation_dict, method_dict = SCCREAD.process_excel_file(file_path)
        
        # Update progress data with the new information
        for key, value in bper_dict.items():
            progress_data['BPERs'][key] = value
        for key, value in doc_dict.items():
            progress_data['Documents'][key] = value
        for key, value in attestation_dict.items():
            progress_data['Attestations'][key] = value
        for stig_id, details in method_dict.items():
            progress_data['Checks'][stig_id] = {
                'SCC': scc_name,
                'Evidence method': details['Evidence Method']
            }
        
        # Save the updated progress data to progress.json
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        error_label.config(text=f"SCC '{scc_name}' added or updated successfully.") # update status message

def open_scc_markdown_file(event):
    selected_scc = scc_listbox.get(scc_listbox.curselection()) # get selected SCC
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            for scc_path, scc_data in progress_data['SCC'].items():
                if scc_data.get('SCC') == selected_scc:
                    md_file_path = scc_data.get('Info Doc Path')
                    if md_file_path and os.path.exists(md_file_path):
                        os.startfile(md_file_path) # open markdown file
                        return
            error_label.config(text=f"Markdown file not found for SCC: {selected_scc}")
    else:
        error_label.config(text="Please select a valid progress.json file.")

def output_progress():
    if progress_file and project_dir:
        try:
            # Load the JSON data from the progress file
            with open(progress_file, 'r') as file:
                data = json.load(file)

            # Create a new workbook
            wb = Workbook()

            # Create a sheet for each high-level dictionary
            for key in data:
                JSONTOEXCEL.create_sheet(wb, key, data[key])

            # Remove the default sheet created by openpyxl
            default_sheet = wb['Sheet']
            wb.remove(default_sheet)

            # Save the workbook in the project directory
            output_file = os.path.join(project_dir, 'progress.xlsx')
            wb.save(output_file)

            error_label.config(text="Progress exported successfully!")
        except Exception as e:
            error_label.config(text=f"Error exporting progress: {str(e)}")
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")

# GUI setup
root = tk.Tk()
root.title("TDL on Easy Mode")
root.geometry("800x600")

# Welcome screen setup
welcome_screen = tk.Frame(root)
welcome_label = tk.Label(welcome_screen, text="Welcome to TDL on Easy Mode!")
welcome_label.pack(pady=20)

new_project_button = tk.Button(welcome_screen, text="Start New Project", command=start_new_project)
new_project_button.pack(pady=10)

existing_project_button = tk.Button(welcome_screen, text="Update Existing Project", command=update_existing_project)
existing_project_button.pack(pady=10)

# Options screen setup
options_screen = tk.Frame(root, bg="#F0F0F0")
options_screen.pack(fill="both", expand=True)
options_label = tk.Label(options_screen, text="What would you like to do?", font=("Arial", 16, "bold"), bg="#F0F0F0")
options_label.pack(pady=20)

button_frame = tk.Frame(options_screen, bg="#F0F0F0")
button_frame.pack(pady=20)

# Pull Information setup
pull_info_frame = tk.LabelFrame(button_frame, text="Pull Information", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
pull_info_frame.pack(side="left", padx=20)
pull_info_button = tk.Button(pull_info_frame, text="Pull", font=("Arial", 10), width=15, command=pull_information)
pull_info_button.pack(pady=5)
pull_info_status = tk.Label(pull_info_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
pull_info_status.pack()

# Build TDL Directories setup
build_dirs_frame = tk.LabelFrame(button_frame, text="Build TDL Directories", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
build_dirs_frame.pack(side="left", padx=20)
build_dirs_button = tk.Button(build_dirs_frame, text="Build", font=("Arial", 10), width=15, command=build_dirs)
build_dirs_button.pack(pady=5)
build_dirs_status = tk.Label(build_dirs_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
build_dirs_status.pack()

# Build Templates setup
build_templates_frame = tk.LabelFrame(button_frame, text="Build Templates", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
build_templates_frame.pack(side="left", padx=20)
build_templates_button = tk.Button(build_templates_frame, text="Build", font=("Arial", 10), width=15, command=build_templates)
build_templates_button.pack(pady=5)
build_templates_status = tk.Label(build_templates_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
build_templates_status.pack()

# Gather and Sort Documents setup
gather_docs_frame = tk.LabelFrame(button_frame, text="Gather and Sort Documents", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
gather_docs_frame.pack(side="left", padx=20)
gather_docs_button = tk.Button(gather_docs_frame, text="Gather", font=("Arial", 10), width=15, command=gather_docs)
gather_docs_button.pack(pady=5)
gather_docs_status = tk.Label(gather_docs_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
gather_docs_status.pack()

# Generate MD Files setup
generate_md_frame = tk.LabelFrame(button_frame, text="Generate MD Files", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
generate_md_frame.pack(side="left", padx=20)
generate_md_button = tk.Button(generate_md_frame, text="Generate", font=("Arial", 10), width=15, command=generate_md_files)
generate_md_button.pack(pady=5)
generate_md_status = tk.Label(generate_md_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
generate_md_status.pack()

# Update Document Tracker setup
update_tracker_frame = tk.LabelFrame(button_frame, text="Update Document Tracker", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
update_tracker_frame.pack(side="left", padx=20)
update_tracker_button = tk.Button(update_tracker_frame, text="Update", font=("Arial", 10), width=15, command=update_tracker)
update_tracker_button.pack(pady=5)
update_tracker_status = tk.Label(update_tracker_frame, text="Not done", font=("Arial", 10), bg="#FFFFFF")
update_tracker_status.pack()

# Additional buttons setup
additional_buttons_frame = tk.Frame(options_screen, bg="#F0F0F0")
additional_buttons_frame.pack(pady=20)

output_progress_button = tk.Button(additional_buttons_frame, text="Output progress", font=("Arial", 12), width=15, command=output_progress)
output_progress_button.pack(side="left", padx=10)

add_redo_scc_button = tk.Button(additional_buttons_frame, text="Add or redo an SCC", font=("Arial", 12), width=20, command=add_or_redo_scc)
add_redo_scc_button.pack(side="left", padx=10)

remove_scc_button = tk.Button(additional_buttons_frame, text="Remove an SCC", font=("Arial", 12), width=20, command=remove_scc)
remove_scc_button.pack(side="left", padx=10)

sync_button = tk.Button(additional_buttons_frame, text="Sync", font=("Arial", 12), width=15, command=sync_button_click)
sync_button.pack(side="left", padx=10)

dashboard_button = tk.Button(additional_buttons_frame, text="Dashboard", font=("Arial", 12), width=15, command=show_dashboard)
dashboard_button.pack(side="right", padx=10)

# Selected Directories setup
directory_labels_frame = tk.LabelFrame(options_screen, text="Selected Directories", font=("Arial", 12), bg="#FFFFFF", padx=10, pady=10)
directory_labels_frame.pack(pady=40)

directory_canvas = tk.Canvas(directory_labels_frame, bg="#FFFFFF", width=800)
directory_canvas.pack(side="left", fill="both", expand=True)

directory_scrollbar = tk.Scrollbar(directory_labels_frame, orient="vertical", command=directory_canvas.yview)
directory_scrollbar.pack(side="right", fill="y")

directory_canvas.configure(yscrollcommand=directory_scrollbar.set)
directory_canvas.bind("<Configure>", lambda e: directory_canvas.configure(scrollregion=directory_canvas.bbox("all")))

directory_frame = tk.Frame(directory_canvas, bg="#FFFFFF", width=800)
directory_canvas.create_window((0, 0), window=directory_frame, anchor="nw")

# BPERs Directory setup
bpers_frame = tk.Frame(directory_frame, bg="#FFFFFF")
bpers_frame.pack(anchor="w", pady=5)
bpers_dir_button = tk.Button(bpers_frame, text="Select", font=("Arial", 10), command=select_bpers_directory)
bpers_dir_button.pack(side="left", padx=10)
bpers_dir_label = tk.Label(bpers_frame, text="BPERs Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
bpers_dir_label.pack(side="left")

# Attestation Directory setup
attestation_frame = tk.Frame(directory_frame, bg="#FFFFFF")
attestation_frame.pack(anchor="w", pady=5)
attestation_dir_button = tk.Button(attestation_frame, text="Select", font=("Arial", 10), command=select_attestation_directory)
attestation_dir_button.pack(side="left", padx=10)
attestation_dir_label = tk.Label(attestation_frame, text="Attestation Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
attestation_dir_label.pack(side="left")

# Supporting Documents Directory setup
supporting_docs_frame = tk.Frame(directory_frame, bg="#FFFFFF")
supporting_docs_frame.pack(anchor="w", pady=5)
supporting_docs_dir_button = tk.Button(supporting_docs_frame, text="Select", font=("Arial", 10), command=select_supporting_docs_directory)
supporting_docs_dir_button.pack(side="left", padx=10)
supporting_docs_dir_label = tk.Label(supporting_docs_frame, text="Supporting Documents Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
supporting_docs_dir_label.pack(side="left")

# SCC Directory setup
scc_frame = tk.Frame(directory_frame, bg="#FFFFFF")
scc_frame.pack(anchor="w", pady=5)
scc_dir_button = tk.Button(scc_frame, text="Select", font=("Arial", 10), command=select_scc_directory)
scc_dir_button.pack(side="left", padx=10)
scc_dir_label = tk.Label(scc_frame, text="SCC Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
scc_dir_label.pack(side="left")

# Progress File setup
progress_file_frame = tk.Frame(directory_frame, bg="#FFFFFF")
progress_file_frame.pack(anchor="w", pady=5)
progress_file_button = tk.Button(progress_file_frame, text="Select", font=("Arial", 10), command=select_progress_file)
progress_file_button.pack(side="left", padx=10)
progress_file_label = tk.Label(progress_file_frame, text="Progress File: Not selected", font=("Arial", 10), bg="#FFFFFF")
progress_file_label.pack(side="left")

# Project Directory setup
project_dir_frame = tk.Frame(directory_frame, bg="#FFFFFF")
project_dir_frame.pack(anchor="w", pady=5)
project_dir_button = tk.Button(project_dir_frame, text="Select", font=("Arial", 10), command=select_project_directory)
project_dir_button.pack(side="left", padx=10)
project_dir_label = tk.Label(project_dir_frame, text="Project Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
project_dir_label.pack(side="left")

# Template Directory setup
template_dir_frame = tk.Frame(directory_frame, bg="#FFFFFF")
template_dir_frame.pack(anchor="w", pady=5)
template_dir_button = tk.Button(template_dir_frame, text="Select", font=("Arial", 10), command=select_template_directory)
template_dir_button.pack(side="left", padx=10)
template_dir_label = tk.Label(template_dir_frame, text="Template Directory: Not selected", font=("Arial", 10), bg="#FFFFFF")
template_dir_label.pack(side="left")

# Gather Docs screen setup
gather_docs_screen = tk.Frame(root)
gather_docs_label = tk.Label(gather_docs_screen, text="Select Document Repositories")
gather_docs_label.pack(pady=20)
gather_docs_buttons_frame = tk.Frame(gather_docs_screen)
gather_docs_buttons_frame.pack(pady=20)

sort_button = tk.Button(gather_docs_screen, text="Sort!", command=sort_docs)
sort_button.pack(pady=20)

# Error label setup
error_label = tk.Label(root, text="", fg="red")
error_label.pack(pady=10)

# Dashboard screen setup
dashboard_screen = tk.Frame(root, bg="#F0F0F0")

back_button = tk.Button(dashboard_screen, text="Back", font=("Arial", 12), width=15, command=show_options)
back_button.pack(side="bottom", padx=10, pady=10)

# Section 1 (SCC List) setup
section1_frame = tk.Frame(dashboard_screen, bg="#FFFFFF", padx=10, pady=10)
section1_frame.pack(side="left", fill="both", expand=True)

section2_frame = tk.Frame(dashboard_screen, bg="#FFFFFF", padx=10, pady=10)
section2_frame.pack(side="left", fill="both", expand=True)

section2_label = tk.Label(section2_frame, text="Items Not Gathered", font=("Arial", 12, "bold"), bg="#FFFFFF")
section2_label.pack(pady=10)

# Items Not Gathered (Attestations) setup
not_gathered_attestations_label = tk.Label(section2_frame, text="Attestations", font=("Arial", 10, "bold"), bg="#FFFFFF")
not_gathered_attestations_label.pack(pady=5)
not_gathered_attestations_listbox = tk.Listbox(section2_frame, font=("Arial", 10), bg="#FFFFFF", selectmode="multiple")
not_gathered_attestations_listbox.pack(fill="both", expand=True)

attestations_buttons_frame = tk.Frame(section2_frame)
attestations_buttons_frame.pack(pady=5)
mark_attestation_false_positive_button = tk.Button(attestations_buttons_frame, text="Mark as False Positive", font=("Arial", 10), command=lambda: mark_as_false_positive("Attestations"))
mark_attestation_false_positive_button.pack(side="left", padx=5)
manually_link_attestations_button = tk.Button(attestations_buttons_frame, text="Assign Match", font=("Arial", 10), command=lambda: manually_link_files("Attestations"))
manually_link_attestations_button.pack(side="left", padx=5)

not_gathered_bpers_label = tk.Label(section2_frame, text="BPERs", font=("Arial", 10, "bold"), bg="#FFFFFF")
not_gathered_bpers_label.pack(pady=5)
not_gathered_bpers_listbox = tk.Listbox(section2_frame, font=("Arial", 10), bg="#FFFFFF", selectmode="multiple")
not_gathered_bpers_listbox.pack(fill="both", expand=True)

bpers_buttons_frame = tk.Frame(section2_frame)
bpers_buttons_frame.pack(pady=5)
mark_bper_false_positive_button = tk.Button(bpers_buttons_frame, text="Mark as False Positive", font=("Arial", 10), command=lambda: mark_as_false_positive("BPERs"))
mark_bper_false_positive_button.pack(side="left", padx=5)
manually_link_bpers_button = tk.Button(bpers_buttons_frame, text="Assign Match", font=("Arial", 10), command=lambda: manually_link_files("BPERs"))
manually_link_bpers_button.pack(side="left", padx=5)

# Items Not Gathered (Documents) setup 
not_gathered_documents_label = tk.Label(section2_frame, text="Documents", font=("Arial", 10, "bold"), bg="#FFFFFF")
not_gathered_documents_label.pack(pady=5)
not_gathered_documents_listbox = tk.Listbox(section2_frame, font=("Arial", 10), bg="#FFFFFF", selectmode="multiple")
not_gathered_documents_listbox.pack(fill="both", expand=True)

documents_buttons_frame = tk.Frame(section2_frame)
documents_buttons_frame.pack(pady=5)
mark_document_false_positive_button = tk.Button(documents_buttons_frame, text="Mark as False Positive", font=("Arial", 10), command=lambda: mark_as_false_positive("Documents"))
mark_document_false_positive_button.pack(side="left", padx=5)
manually_link_documents_button = tk.Button(documents_buttons_frame, text="Assign Match", font=("Arial", 10), command=lambda: manually_link_files("Documents"))
manually_link_documents_button.pack(side="left", padx=5)

# Section 3 (Dates and Chart) setup 
section3_frame = tk.Frame(dashboard_screen, bg="#FFFFFF", padx=10, pady=10)
section3_frame.pack(side="left", fill="both", expand=True)

section1_label = tk.Label(section1_frame, text="SCC List", font=("Arial", 12), bg="#FFFFFF")
section1_label.pack(pady=10)

scc_listbox = tk.Listbox(section1_frame, font=("Arial", 10), bg="#FFFFFF")
scc_listbox.pack(fill="both", expand=True)
scc_listbox.bind("<Double-Button-1>", open_scc_markdown_file)

section3_label = tk.Label(section3_frame, text="Dates and Chart", font=("Arial", 12), bg="#FFFFFF")
section3_label.pack(pady=10)

last_info_pull_label = tk.Label(section3_frame, text="Last Info Pull: N/A", font=("Arial", 10), bg="#FFFFFF")
last_info_pull_label.pack(pady=5)

last_doc_pull_label = tk.Label(section3_frame, text="Last Doc Pull: N/A", font=("Arial", 10), bg="#FFFFFF")
last_doc_pull_label.pack(pady=5)

last_checklist_generated_label = tk.Label(section3_frame, text="Last Checklist Generated: N/A", font=("Arial", 10), bg="#FFFFFF")
last_checklist_generated_label.pack(pady=5)

# Placeholder for the pie chart
pie_chart_label = tk.Label(section3_frame, text="", font=("Arial", 10), bg="#FFFFFF", justify="center")
pie_chart_label.pack(pady=10)

# Show the welcome screen initially
show_welcome()
root.mainloop()
