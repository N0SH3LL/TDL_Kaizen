import tkinter as tk
from tkinter import filedialog
import subprocess

existing_project_dir = None
supporting_docs_dir = None
attestation_dir = None
bpers_dir = None

def select_directory(prompt):
    directory = filedialog.askdirectory(title=prompt)
    return directory

def run_script():
    directory = existing_project_dir
    if directory:
        script_path = "path/to/your/existing/python/script.py"
        subprocess.run(["python", script_path, directory])
    else:
        error_label.config(text="Please select a directory first.")

def show_welcome():
    clear_frames()
    welcome_screen.pack(fill="both", expand=True)

def show_options():
    clear_frames()
    options_screen.pack(fill="both", expand=True)

def show_gather_docs():
    clear_frames()
    gather_docs_screen.pack(fill="both", expand=True)

def start_new_project():
    show_options()

def update_existing_project():
    select_directory()
    if existing_project_dir:
        show_options()
    else:
        error_label.config(text="Please select an existing TDL directory.")

def build_dirs():
    scc_repo = select_directory("Please select the SCC repository")
    if scc_repo:
        # Perform actions with the selected SCC repository directory
        pass

def build_templates():
    template_repo = select_directory("Please select the Template repository")
    if template_repo:
        # Perform actions with the selected Template repository directory
        pass

def gather_docs():
    show_gather_docs()

def select_supporting_docs():
    global supporting_docs_dir
    supporting_docs_dir = select_directory("Please select the Supporting Documents repository")

def select_attestation():
    global attestation_dir
    attestation_dir = select_directory("Please select the Attestation repository")

def select_bpers():
    global bpers_dir
    bpers_dir = select_directory("Please select the BPERs repository")

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
        frame.pack_forget()

root = tk.Tk()
root.title("TDL on Easy Mode")
root.geometry("400x400")

welcome_screen = tk.Frame(root)
welcome_label = tk.Label(welcome_screen, text="Welcome to TDL on Easy Mode!")
welcome_label.pack(pady=20)
new_project_button = tk.Button(welcome_screen, text="Start New Project", command=start_new_project)
new_project_button.pack(pady=10)
existing_project_button = tk.Button(welcome_screen, text="Update Existing Project", command=update_existing_project)
existing_project_button.pack(pady=10)

options_screen = tk.Frame(root)
options_label = tk.Label(options_screen, text="What would you like to do?")
options_label.pack(pady=20)

button_frame = tk.Frame(options_screen)
button_frame.pack(pady=20)

build_dirs_frame = tk.Frame(button_frame)
build_dirs_frame.pack(side="left", padx=20)
build_dirs_button = tk.Button(build_dirs_frame, text="Build the TDL directories", command=build_dirs)
build_dirs_button.pack()
build_dirs_status = tk.Label(build_dirs_frame, text="Not done")
build_dirs_status.pack()

build_templates_frame = tk.Frame(button_frame)
build_templates_frame.pack(side="left", padx=20)
build_templates_button = tk.Button(build_templates_frame, text="Build Templates", command=build_templates)
build_templates_button.pack()
build_templates_status = tk.Label(build_templates_frame, text="Not done")
build_templates_status.pack()

gather_docs_frame = tk.Frame(button_frame)
gather_docs_frame.pack(side="left", padx=20)
gather_docs_button = tk.Button(gather_docs_frame, text="Gather and sort Documents", command=gather_docs)
gather_docs_button.pack()
gather_docs_status = tk.Label(gather_docs_frame, text="Not done")
gather_docs_status.pack()

update_tracker_frame = tk.Frame(button_frame)
update_tracker_frame.pack(side="left", padx=20)
update_tracker_button = tk.Button(update_tracker_frame, text="Update the Document Tracker", command=update_tracker)
update_tracker_button.pack()
update_tracker_status = tk.Label(update_tracker_frame, text="Not done")
update_tracker_status.pack()

missing_docs_frame = tk.Frame(button_frame)
missing_docs_frame.pack(side="left", padx=20)
missing_docs_button = tk.Button(missing_docs_frame, text="See what is missing")
missing_docs_button.pack()
missing_docs_status = tk.Label(missing_docs_frame, text="Not done")
missing_docs_status.pack()

additional_buttons_frame = tk.Frame(options_screen)
additional_buttons_frame.pack(pady=20)

do_all_button = tk.Button(additional_buttons_frame, text="Do it all")
do_all_button.pack(side="left", padx=10)

add_redo_scc_button = tk.Button(additional_buttons_frame, text="Add or redo an SCC")
add_redo_scc_button.pack(side="left", padx=10)

gather_docs_screen = tk.Frame(root)
gather_docs_label = tk.Label(gather_docs_screen, text="Select Document Repositories")
gather_docs_label.pack(pady=20)

gather_docs_buttons_frame = tk.Frame(gather_docs_screen)
gather_docs_buttons_frame.pack(pady=20)

supporting_docs_button = tk.Button(gather_docs_buttons_frame, text="Select Supporting Documents repository", command=select_supporting_docs)
supporting_docs_button.pack(pady=10)

attestation_button = tk.Button(gather_docs_buttons_frame, text="Select Attestation repository", command=select_attestation)
attestation_button.pack(pady=10)

bpers_button = tk.Button(gather_docs_buttons_frame, text="Select BPERs repository", command=select_bpers)
bpers_button.pack(pady=10)

sort_button = tk.Button(gather_docs_screen, text="Sort!", command=sort_docs)
sort_button.pack(pady=20)

error_label = tk.Label(root, text="", fg="red")
error_label.pack(pady=10)

show_welcome()

root.mainloop()
