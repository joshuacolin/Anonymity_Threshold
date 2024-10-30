import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk,OptionMenu

from Anonymity_Threshold_PC import PC
from Anonymity_Threshold_LB import LB
#--------------------------Calling Anonymity Threshold functions -----------------------------------------
def run_mix():
    # get hierarchy
    hierarchy_selection = option.get()

    #Get file paths from user search
    ExportUnits_file_path = file_entry_data1.get()
    Response_file_path = file_entry_data2.get()
    Participants_file_path = file_entry_data3.get()
    
    #Get Employee and Manager IDs
    LevelColumns  = Empl_Man_IDs_Columns_entry.get().split(',')

    # obtaining the anonymity threshold value
    try:
        anonymity_threshold_value = int(anonymity_threshold_entry.get())
    except ValueError:
        messagebox.showerror("ValueError", "Please provide the anonymity threshold value.")
        return

    #error handling
    if not ExportUnits_file_path or not Response_file_path or not Participants_file_path:
        messagebox.showerror("Error", "Please provide all file paths.")
        return
    
    # calling the anonymity threshold script
    if hierarchy_selection == 'Parent-Child':
        PC(ExportUnits_file_path, Response_file_path, Participants_file_path, LevelColumns, anonymity_threshold_value)
    else:
        LB(ExportUnits_file_path, Response_file_path, Participants_file_path, LevelColumns, anonymity_threshold_value)

#--------------------------User Interface -----------------------------------------
def hierarchy_selection_menu(self):
    if option.get() == 'Parent-Child':
        Empl_Man_IDs_Columns_label.configure(text="Provide Employee and Manager ID column names (comma-separated):")
        Empl_Man_IDs_Columns_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W, columnspan=2)
    else:
        Empl_Man_IDs_Columns_label.configure(text="Provide Level column names (comma-separated):")
        Empl_Man_IDs_Columns_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W, columnspan=2)
def browse_files1():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data1.delete(0, tk.END)
    file_entry_data1.insert(0, file_path)

def browse_files2():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data2.delete(0, tk.END)
    file_entry_data2.insert(0, file_path)
    
def browse_files3():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    file_entry_data3.delete(0, tk.END)
    file_entry_data3.insert(0, file_path)
    
# Set up the main application window
root = tk.Tk()
root.title("Manager Anonymity Threshold Checker v1")

# Hierarchy selection
hierarchy_label = tk.Label(root, text="Please select the Hierarchy Type")
hierarchy_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
hierarchy_options = ['Parent-Child', 'Level Based']
option = tk.StringVar()
option.set(hierarchy_options[0])
hierarchy_menu = OptionMenu(root, option, *hierarchy_options, command=hierarchy_selection_menu)
hierarchy_menu.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W, columnspan=2)

# File selection
file_label_data1 = tk.Label(root, text="Org Hierarchy Units File path")
file_label_data1.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
file_entry_data1 = tk.Entry(root, width=30)
file_entry_data1.grid(row=1, column=2, padx=5, pady=5)
browse_button_data1 = tk.Button(root, text="Browse",command=browse_files1)
browse_button_data1.grid(row=1, column=3, padx=5, pady=5)

file_label_data2 = tk.Label(root, text="Response File path")
file_label_data2.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
file_entry_data2 = tk.Entry(root, width=30)
file_entry_data2.grid(row=2, column=2, padx=5, pady=5)
browse_button_data2 = tk.Button(root, text="Browse",command=browse_files2)
browse_button_data2.grid(row=2, column=3, padx=5, pady=5)

file_label_data3 = tk.Label(root, text="Participant File path")
file_label_data3.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
file_entry_data3 = tk.Entry(root, width=30)
file_entry_data3.grid(row=3, column=2, padx=5, pady=5)
browse_button_data3 = tk.Button(root, text="Browse",command=browse_files3)
browse_button_data3.grid(row=3, column=3, padx=5, pady=5)

# Level based column names input
Empl_Man_IDs_Columns_label = tk.Label(root, text="Provide Employee ID and Manager ID (comma-separated):")
Empl_Man_IDs_Columns_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W, columnspan=2)
Empl_Man_IDs_Columns_entry = tk.Entry(root, width=30)
Empl_Man_IDs_Columns_entry.grid(row=4, column=2, padx=5, pady=5, columnspan=2)

#Set anonymity Threshold value
anonymity_threshold_label = tk.Label(root, text="Set anonymity threshold value")
anonymity_threshold_label.grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
anonymity_threshold_entry = tk.Entry(root, width=30)
anonymity_threshold_entry.grid(row=5, column=1, padx=5, pady=5, columnspan=2)


unsupported_chars_list_frame = tk.LabelFrame(root, text = "HOW TO USE:")
unsupported_chars_list_frame.grid(row=6, column=0,columnspan=10, padx=20, pady=20, sticky=tk.W)

unsupported_chars_list_1=tk.Label(unsupported_chars_list_frame, text='1. Select the files to perform the analysis. ')
unsupported_chars_list_1.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
unsupported_chars_list_2=tk.Label(unsupported_chars_list_frame, text='2.1. For Parent-Child Hierarchy: Enter the name of the columns for Employee and Manager ID. ')
unsupported_chars_list_2.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
unsupported_chars_list_2=tk.Label(unsupported_chars_list_frame, text='2.1. For Level-Based Hierarchy: Enter the name of the columns for the different Levels. ')
unsupported_chars_list_2.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
unsupported_chars_list_3=tk.Label(unsupported_chars_list_frame, text='3. Set the Anonymity Threshold value')
unsupported_chars_list_3.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
unsupported_chars_list_4=tk.Label(unsupported_chars_list_frame, text='4.    Click "Get Anonymity Threshold!"')
unsupported_chars_list_4.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)

# Run mix button
run_button = tk.Button(root, text="Get Anonymity Threshold!", command=run_mix)
run_button.grid(row=7, column=0, columnspan=1, padx=20, pady=20)

new=1
url = "https://coda.io/d/QA-Automation_diYJsprOr4k/Managers-Anonymity-Threshold_suB1Q#_luXuW"

def openweb():
    webbrowser.open(url,new=new)

feedback_button= tk.Button(root, text ="Feedback / New ideas", command=openweb)
feedback_button.grid(row=7, column=1, columnspan=1, padx=5, pady=5)

# Start the application
root.mainloop()