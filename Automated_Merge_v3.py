# File Merger for QC SHRU L&D Database
# Purpose: to automatically merge all the training data available
# Created by: Ian Salig U Batangan, Contact Details: isubatangan@gmail.com
# Version Updates by: Airysh Xander M. Espero, Contact Details: derespero@gmail.com

import os 
import venv # For activating virtual environment
import pandas as pd # For data manipulation library, also needs openpyxl sub library of pandas
import re # String manipulation library
import glob # For finding file path library
from pathlib import Path
import msvcrt # For detecting button press
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog
from tkinter import PhotoImage, Label
from tkinter import Tk, Label, Button, PhotoImage, Frame
from PIL import Image, ImageTk

# Script for running necessary commands:
# Creating necessary folders
# Virtual environment creation
# Packages/Librarires Installations
def initial_scripts():
    merged_folder_dir = "MergedFiles"
    if not os.path.exists(merged_folder_dir):
        print("Creating 'MergedFiles' folder...")
        os.mkdir(merged_folder_dir)
        print("'MergedFiles' folder Created!")

    allconcat_file = f"{merged_folder_dir}/AllConcat.csv"
    if not os.path.exists(allconcat_file):
        print(f"Creating 'AllConcat.csv' file...")
        with open(allconcat_file, 'w') as file:
            pass
        print(f"File 'AllConcat.csv' successfully Created!")

    to_merge_folder_dir = "ToMergeFiles"
    if not os.path.exists(to_merge_folder_dir):
        print("Creating 'ToMergeFiles' folder...")
        os.mkdir(to_merge_folder_dir)
        print("'ToMergeFiles' folder Created!")


    default_categories = [
        "PROGRAM DESIGN.+", "TRAINING.+", "LOGISTICS.+",
        "EXPECTATION.+", "ADMINISTRATION.+", "COMMENT.+", "FACILITATOR.+"
    ]

    facilitation_file = "FacilitationColumns.txt"
    if not os.path.exists(facilitation_file):
        print("Creating 'FacilitationColumns.txt' with default categories...")
        with open(facilitation_file, 'w') as file:
            file.write(','.join(default_categories))  # Store as comma-separated values
        print("'FacilitationColumns.txt' successfully created with default values!")

    # venv_dir = "venv"
    # if not os.path.exists(venv_dir):
    #     print("Creating virtual environment...")
    #     venv.create(venv_dir, with_pip=True)
    #     print("venv file created!")

    # print(f"Checking packages...")
    # if not os.path.exists(venv_dir):
    #     os.system('/venv/Scripts/activate.bat && pip install -r requirements.txt')
    
    # else:
    #     os.system('pip install -r requirements.txt')

    # print(f"Virtual environment successfully activated and required packages was installed!")

def merge(variable_name, root_path):
    #determining path
    path_reg = root_path/"ToMergeFiles"/f"{variable_name}_Reg.xlsx/"
    path_post = root_path/"ToMergeFiles"/f"{variable_name}_Post.csv/"

    #reads raw excel and csv
    df_reg = pd.read_excel(path_reg)
    df_post = pd.read_csv(path_post)
    
    df_reg.columns = map(str.upper, df_reg.columns)
    df_post.columns = map(str.upper, df_post.columns)
        
    #cleaning of column names
    #creating a list of the column names
    df_reg_columns = df_reg.columns
    df_post_columns = df_post.columns 
        
        #checking the list if the data set has assessment
    if 'PRE-ASSESSMENT TOTAL' not in df_reg_columns or 'QTOTAL' not in df_post_columns:
        check_pre = 1

    else: 
        check_pre = 0

    #mass removal of extrenous data
    #ISUB
    df_reg_cleaned = [e for e in df_reg_columns if "PRE-ASSESSMENT" not in e and 
                      'EMAIL ADDRESS' not in e and 
                      'NICKNAME' not in e and 
                      'ENDORSEMENT LETTER' not in e and 
                      'CSC UPLOADED' not in e and 
                      'DATE ANSWERED' not in e and 
                      'EXPECTED OUTCOMES'not in e and 
                      'DATA PRIVACY CONSENT' not in e and 
                      'CONTACT NUMBER' not in e]
        
    #using re library to exlude columns with Q in name in specific cases
    df_post_cleaned = [e for e in df_post_columns if not re.match(re.compile('Q.+-' ) , e) and 
                       not re.match(re.compile('Q..' ) , e) and 
                       not re.match(re.compile('Q.' ) , e) and 
                       'EMAIL ADDRESS' not in e and 
                       'NICKNAME' not in e and 
                       'ENDORSEMENT LETTER' not in e and 
                       'CSC UPLOADED' not in e and 
                       'DATE ANSWERED' not in e and 
                       'EXPECTED OUTCOMES'not in e and 
                       'DATA PRIVACY CONSENT'not in e and 
                       'CONTACT NUMBER' not in e]
        
    #adds back the needed column since mass deletion was done previously
    if check_pre == 0:
        df_reg_cleaned = df_reg_cleaned+['PRE-ASSESSMENT TOTAL']
        df_post_cleaned = df_post_cleaned+['QTOTAL']


    #finding maximum score
    Qmax_list = [e for e in df_post_columns if re.match(re.compile('Q.+-' ) , e) ]
    max_score=len(Qmax_list)


    #retrieving specific column names
    df_reg = df_reg.loc[:,df_reg_cleaned]
    df_post = df_post.loc[:,df_post_cleaned]
    df_post= df_post.assign(Maximum_Assesment_Score=max_score,Training_Code=variable_name)

    #debugging
    '''
    print('\n this is registration column \n',df_reg_cleaned,'\n this is post columns\n',df_post_cleaned,'\n this is post data shape \n',df_post_shape, '\n')    
    '''
    #Changes all data in dataframe as string
    for name in df_reg.columns:
        df_reg[f'{name}'] = df_reg[f'{name}'].astype(str)

    for name in df_post.columns:
        df_post[f'{name}'] = df_post[f'{name}'].astype(str)

    #Left Join of post and reg data and drops duplicates
    #checking if data set has full name if not full name is created
    if 'FULL NAME' in df_post_columns and 'FULL NAME' in df_reg_columns:
        df_merge= df_reg.merge(
            df_post,left_on=[
                'FULL NAME',"DESIGNATION/POSITION","DIVISION/ SECTION",
                'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'],
                
                right_on=['FULL NAME',"DESIGNATION/POSITION","DIVISION/ SECTION",
                          'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'
                          ]).drop_duplicates()

    else:
        df_merge= df_reg.merge(
            df_post,left_on=[
                "LAST NAME","FIRST NAME","MIDDLE INITIAL","DESIGNATION/POSITION",
                "DIVISION/ SECTION",'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'
            ],

                right_on=[
                    "LAST NAME","FIRST NAME","MIDDLE INITIAL","DESIGNATION/POSITION","DIVISION/ SECTION",
                    'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE','SEX'
                    ]).drop_duplicates()
        
        #generate Full Name
        df_merge['FULL NAME'] = df_merge["FIRST NAME"]+ " " + df_merge["MIDDLE INITIAL"] + " "+ df_merge["LAST NAME"]

    #saves the merged data to a csv file
    df_merge.to_csv(f"{root_path}/MergedFiles/{variable_name}_merged.csv")
    return df_merge


def getLocalFile():
    # Allow multiple file selection
    filePaths = filedialog.askopenfilenames(title="Select Files to Merge")

    if not filePaths:
        messagebox.showinfo("No Files Selected", "No files were selected.")
        return

    # Define the destination folder
    root_path = Path.cwd()
    to_merge_path = root_path / "ToMergeFiles"

    # Ensure the folder exists
    to_merge_path.mkdir(exist_ok=True)

    # Move selected files to ToMergeFiles
    for file in filePaths:
        file_path = Path(file)
        destination = to_merge_path / file_path.name  # Preserve the original file name

        try:
            destination.write_bytes(file_path.read_bytes())  # Copy & paste the file
            print(f"Moved: {file_path.name} -> {destination}")
        except Exception as e:
            print(f"Error moving {file_path.name}: {e}")

    messagebox.showinfo("Success", "Selected files have been moved to ToMergeFiles!")


# File to store facilitation categories
FILE_NAME = 'FacilitationColumns.txt'

def read_categories():
    # Read facilitation categories from the file.
    if not os.path.exists(FILE_NAME):
        return []
    with open(FILE_NAME, 'r') as file:
        content = file.read().strip()
        return content.split(',')


def write_categories(categories):
    # Write facilitation categories to the file.
    with open(FILE_NAME, 'w') as file:
        file.write(','.join(categories))


def main():
    # Declaring the file path of the root directory
    root_path=Path.cwd()
    print(f'File path is {root_path}')

    # Define the folder to search for files
    to_merge_path = root_path / "ToMergeFiles"
    merged_folder_path = root_path / "MergedFiles"
    allconcat_file_path = merged_folder_path / "AllConcat.csv"

    # Ensure MergedFiles directory exists
    merged_folder_path.mkdir(exist_ok=True)

    # Ensure AllConcat.csv exists before proceeding
    if not allconcat_file_path.exists():
        with open(allconcat_file_path, 'w') as file:
            pass  # Creates an empty AllConcat.csv file
        print(f"Created 'AllConcat.csv' in {merged_folder_path}")

    # Check if the directory is empty
    if not any(to_merge_path.iterdir()):
        messagebox.showerror("Error", "No files found in the ToMergeFiles folder!")
        return
    
    # Define pattern to search for files
    file_pattern = "*_Post.csv"

    # Checks current folder for all the trainings with post assesment and makes a list of the unique trainings
    path_list = to_merge_path.glob(file_pattern)

    training_list=[]
    main_df=pd.DataFrame()
    
    for name in path_list:
        # Use Path to extract the base name without '_Post.csv'
        training_name = name.stem.replace("_Post", "")  # Replace '_Post.csv' suffix with blank
        training_list.append(training_name)
        print(training_name)
        
    total_files=len(training_list)

    #uses list of trainings to look for the file to merge
    for name in training_list:
        temp_df = merge(name, root_path)
        
        # Concatenate with checking duplicates (if necessary)
        main_df = pd.concat([main_df, temp_df])


    # Re-arranging the Data Columns
    main_df_columns= main_df.columns

    #main information
    main_df_columns_tag=["ID",'Training_Code',"FULL NAME","LAST NAME","FIRST NAME","MIDDLE INITIAL",
        "DESIGNATION/POSITION","DIVISION/ SECTION",'DEPARTMENT/ OFFICE/ UNIT/ TASK FORCE','EMPLOYMENT TYPE',
        'SEX','PRE-ASSESSMENT TOTAL', 'QTOTAL', 'Maximum_Assesment_Score']
    
    # Read the file content and split it into a list of patterns
    patterns = []
    with open(FILE_NAME, 'r') as file:
        content = file.read().strip()
        patterns = content.split(',')

    # Track added columns to avoid duplicates
    added_columns = set(main_df_columns_tag)

    main_df_ordered_list = main_df_columns_tag
    for category in patterns:
        matching_columns = [e for e in main_df_columns if re.match(re.compile(category), e)]
    
        # Add only the columns that have not been added yet
        for column in matching_columns:
            if column not in added_columns:
                main_df_ordered_list.append(column)
                added_columns.add(column)

    # Add columns that were not already included in the ordered list
    main_df_columns_others=[e for e in main_df_columns if e not in added_columns]
    main_df_ordered_list.extend(main_df_columns_others) # Add the remaining columns to the ordered list

    # ISUB
    main_df = main_df[main_df_ordered_list]
    main_df.to_csv(f"{root_path}/MergedFiles/AllConcat.csv") 
    messagebox.showinfo("Data Merging", "Registration and Post Evaluation files successfully merged!")



def manage_categories():
    """Manage facilitation categories with pagination."""
    global category_window

    # Ensure only one window is open at a time
    try:
        if category_window.winfo_exists():
            category_window.destroy()
    except NameError:
        pass  

    categories = read_categories()

    category_window = tk.Toplevel(root)
    category_window.title("Manage Categories")
    category_window.attributes('-fullscreen', True)  # Full Screen

    # Center Frame
    center_category_window = tk.Frame(category_window)
    center_category_window.pack(expand=True)

    tk.Label(center_category_window, text="Manage Facilitation Categories", font=("Arial", 16, "bold")).pack(pady=10)

    if not categories:
        tk.Label(center_category_window, text="No categories found.", font=("Arial", 12)).pack(pady=5)
    else:
        # Pagination setup
        categories_per_page = 10
        current_page = [0]

        # Table frame
        table_frame = tk.Frame(center_category_window)
        table_frame.pack(pady=10)

        def display_page():
            """Display only a subset of categories per page."""
            for widget in table_frame.winfo_children():
                widget.destroy()

            start = current_page[0] * categories_per_page
            end = min((current_page[0] + 1) * categories_per_page, len(categories))
            categories_to_display = categories[start:end]

            # Header row
            tk.Label(table_frame, text="No.", font=("Arial", 12, "bold"), width=5, anchor="w").grid(row=0, column=0, padx=5, pady=5)
            tk.Label(table_frame, text="Category", font=("Arial", 12, "bold"), width=40, anchor="w").grid(row=0, column=1, padx=5, pady=5)
            tk.Label(table_frame, text="Action", font=("Arial", 12, "bold"), width=10, anchor="w").grid(row=0, column=2, padx=5, pady=5)

            def delete_category(index):
                """Confirm and delete the selected category."""
                category_name = categories[index][:-2]
                confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{category_name}'?")
                if confirm:
                    del categories[index]
                    write_categories(categories)
                    category_window.destroy()
                    manage_categories()

            for i, cat in enumerate(categories_to_display):
                actual_index = start + i  # Maintain correct index for deletion
                tk.Label(table_frame, text=f"{actual_index + 1}.", font=("Arial", 12), width=5, anchor="w").grid(row=i + 1, column=0, padx=5, pady=5)
                tk.Label(table_frame, text=cat[:-2], font=("Arial", 12), width=40, anchor="w").grid(row=i + 1, column=1, padx=5, pady=5)
                tk.Button(table_frame, text="Delete", command=lambda i=actual_index: delete_category(i), bg="red", fg="white").grid(row=i + 1, column=2, padx=5, pady=5)

        # Navigation buttons
        nav_frame = tk.Frame(center_category_window)
        nav_frame.pack(pady=5)

        def go_prev():
            """Go to the previous page."""
            if current_page[0] > 0:
                current_page[0] -= 1
                display_page()

        def go_next():
            """Go to the next page."""
            if (current_page[0] + 1) * categories_per_page < len(categories):
                current_page[0] += 1
                display_page()

        tk.Button(nav_frame, text="Previous", command=go_prev).pack(side="left", padx=5)
        tk.Button(nav_frame, text="Next", command=go_next).pack(side="left", padx=5)

        # Display the first page
        display_page()

    # Add new category section
    add_frame = tk.Frame(center_category_window)
    add_frame.pack(pady=20)

    tk.Label(add_frame, text="Enter New Category Name:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
    category_entry = tk.Entry(add_frame, width=30)
    category_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(add_frame, text=f"Enter Position (1-{len(categories) + 1}):", font=("Arial", 12)).grid(row=1, column=0, padx=5, pady=5)
    position_entry = tk.Entry(add_frame, width=5)
    position_entry.grid(row=1, column=1, padx=5, pady=5)

    def add_category():
        """Add a new category at the specified position."""
        new_category = category_entry.get().upper()

        if not new_category:
            messagebox.showerror("Error", "Category name cannot be empty.")
            category_window.destroy()
            manage_categories()
            return

        existing_categories = [cat[:-2] for cat in categories]
        if new_category in existing_categories:
            messagebox.showerror("Error", f"Category '{new_category}' already exists.")
            category_window.destroy()
            manage_categories()
            return

        try:
            position = int(position_entry.get())
            if position < 1 or position > len(categories) + 1:
                raise ValueError

            categories.insert(position - 1, f"{new_category}.+")
            write_categories(categories)
            messagebox.showinfo("Success", "New facilitation category successfully added!")
            category_window.destroy()
            manage_categories()

        except ValueError:
            messagebox.showerror("Error", f"Invalid position. Enter a number between 1 and {len(categories) + 1}.")
            category_window.destroy()
            manage_categories()
            return

    tk.Button(add_frame, text="Add Category", command=add_category, bg="green", fg="white").grid(row=2, column=0, columnspan=2, pady=10)

    # Reset to Default Categories
    def reset_to_default():
        """Reset categories to the default list."""
        confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to reset the categories to default?")
        if confirm:
            default_categories = [
                "PROGRAM DESIGN.+", "TRAINING.+", "LOGISTICS.+",
                "EXPECTATION.+", "ADMINISTRATION.+", "COMMENT.+", "FACILITATOR.+"
            ]
            write_categories(default_categories)
            messagebox.showinfo("Success", "Categories have been reset to default.")
            category_window.destroy()
            manage_categories()

    tk.Button(center_category_window, text="Reset to Default", command=reset_to_default, bg="orange", fg="black").pack(pady=10)

    # Back button
    def go_back():
        category_window.destroy()

    tk.Button(center_category_window, text="Back", command=go_back).pack(pady=10)


def viewToMergeFiles():
    """Display a window showing the list of files in the ToMergeFiles folder."""
    global view_merge_window

    # Ensure there's only one window open at a time
    try:
        if view_merge_window.winfo_exists():
            view_merge_window.destroy()
    except NameError:
        pass  

    # Define the folder path
    root_path = Path.cwd()
    to_merge_path = root_path / "ToMergeFiles"

    # Ensure the folder exists
    to_merge_path.mkdir(exist_ok=True)

    # Get the list of files in ToMergeFiles
    files = list(to_merge_path.iterdir())

    # Create the new window
    view_merge_window = tk.Toplevel()
    view_merge_window.title("Files in ToMergeFiles")
    view_merge_window.attributes('-fullscreen', True)  # Full Screen

    # Center Frame to hold all content
    center_view_merge_window = tk.Frame(view_merge_window)
    center_view_merge_window.pack(expand=True)

    # Display the title
    tk.Label(center_view_merge_window, text="Files in ToMergeFiles", font=("Arial", 16)).pack(pady=10)

    # Upload file button
    tk.Button(center_view_merge_window, text="Upload Files", command=getLocalFile, bg="blue", fg="white").pack(pady=10)

    if not files:
        tk.Label(center_view_merge_window, text="No files found in ToMergeFiles.", font=("Arial", 12)).pack(pady=5)
    else:
        # Pagination setup
        files_per_page = 10
        current_page = [0]

        # Table frame
        table_frame = tk.Frame(center_view_merge_window)
        table_frame.pack(pady=10)

        def display_page():
            """Display only a subset of files per page."""
            for widget in table_frame.winfo_children():
                widget.destroy()

            start = current_page[0] * files_per_page
            end = min((current_page[0] + 1) * files_per_page, len(files))
            files_to_display = files[start:end]

            # Header row
            tk.Label(table_frame, text="File Name", font=("Arial", 12, "bold"), width=50, anchor="w").grid(row=0, column=0, padx=5, pady=5)
            tk.Label(table_frame, text="View", font=("Arial", 12, "bold"), width=10).grid(row=0, column=1, padx=5, pady=5)
            tk.Label(table_frame, text="Delete", font=("Arial", 12, "bold"), width=10).grid(row=0, column=2, padx=5, pady=5)

            def open_file(file_path):
                """Open the selected file using the default application."""
                try:
                    os.startfile(file_path)  # Windows
                except AttributeError:
                    os.system(f"open {file_path}")  # MacOS/Linux alternative

            def delete_file(file_path):
                """Confirm and delete the selected file."""
                confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{file_path.name}'?")
                if confirm:
                    try:
                        file_path.unlink()  # Delete the file
                        messagebox.showinfo("Success", f"Deleted '{file_path.name}' successfully!")
                    except Exception as e:
                        messagebox.showerror("Error", f"Could not delete file: {e}")
                    view_merge_window.destroy()  # Close window
                    viewToMergeFiles()  # Reload updated list

            # Display file names
            for i, file in enumerate(files_to_display):
                tk.Label(table_frame, text=file.name, font=("Arial", 12), width=50, anchor="w").grid(row=i + 1, column=0, padx=5, pady=5)
                tk.Button(table_frame, text="View", command=lambda f=file: open_file(f), bg="green", fg="white").grid(row=i + 1, column=1, padx=5, pady=5)
                tk.Button(table_frame, text="Delete", command=lambda f=file: delete_file(f), bg="red", fg="white").grid(row=i + 1, column=2, padx=5, pady=5)

        # Navigation buttons
        nav_frame = tk.Frame(center_view_merge_window)
        nav_frame.pack(pady=5)

        def go_prev():
            """Go to the previous page."""
            if current_page[0] > 0:
                current_page[0] -= 1
                display_page()

        def go_next():
            """Go to the next page."""
            if (current_page[0] + 1) * files_per_page < len(files):
                current_page[0] += 1
                display_page()

        tk.Button(nav_frame, text="Previous", command=go_prev).pack(side="left", padx=5)
        tk.Button(nav_frame, text="Next", command=go_next).pack(side="left", padx=5)

        # Display the first page
        display_page()

    def perform_merge():
        """Perform data merging and close the window afterward."""
        main()  # Call the data merging function
        view_merge_window.destroy()  # Close the viewToMergeFiles window after merging

    # Perform Data Merging Button (Now closes the window after merging)
    tk.Button(center_view_merge_window, text="Perform Data Merging", command=perform_merge, bg="purple", fg="white").pack(pady=10)

    # Back button
    def go_back():
        view_merge_window.destroy()

    tk.Button(center_view_merge_window, text="Back", command=go_back).pack(pady=10)


def viewMergedFiles():
    """Display a window showing the list of files in the MergedFiles folder with options to view and delete."""
    global view_merged_window

    # Ensure only one window is open at a time
    try:
        if view_merged_window.winfo_exists():
            view_merged_window.destroy()
    except NameError:
        pass  # If the window is not defined yet, skip this step

    # Define the folder path
    root_path = Path.cwd()
    merged_path = root_path / "MergedFiles"

    # Ensure the folder exists
    merged_path.mkdir(exist_ok=True)

    # Get the list of files in MergedFiles
    files = list(merged_path.iterdir())
    
    # Pagination variables
    items_per_page = 10
    current_page = 0

    # Create the new window
    view_merged_window = tk.Toplevel(root)
    view_merged_window.title("Manage Merged Files")
    view_merged_window.attributes('-fullscreen', True)  # Full Screen

    # Center Frame to hold all content
    center_view_merged_window = tk.Frame(view_merged_window)
    center_view_merged_window.pack(expand=True)

    tk.Label(center_view_merged_window, text="Files in MergedFiles", font=("Arial", 16, "bold")).pack(pady=10)

    # If no files are present, show a message
    if not files:
        tk.Label(center_view_merged_window, text="No merged files found.", font=("Arial", 12)).pack(pady=5)
    else:
        # Create a frame for table structure
        table_frame = tk.Frame(center_view_merged_window)
        table_frame.pack(pady=10)

        # Header row
        tk.Label(table_frame, text="File Name", font=("Arial", 12, "bold"), width=50, anchor="w").grid(row=0, column=0, padx=5, pady=5)
        tk.Label(table_frame, text="View", font=("Arial", 12, "bold"), width=10).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(table_frame, text="Delete", font=("Arial", 12, "bold"), width=10).grid(row=0, column=2, padx=5, pady=5)

        def open_file(file_path):
            """Open the selected file using the default application."""
            try:
                os.startfile(file_path)  # Windows
            except AttributeError:
                os.system(f"open {file_path}")  # MacOS/Linux alternative

        def delete_file(file_path):
            """Confirm and delete the selected file."""
            confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{file_path.name}'?")
            if confirm:
                try:
                    file_path.unlink()  # Delete the file
                    messagebox.showinfo("Success", f"Deleted '{file_path.name}' successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not delete file: {e}")
                view_merged_window.destroy()  # Close window
                viewMergedFiles()  # Reload updated list

        def display_files():
            """Display files for the current page."""
            # Clear previous entries
            for widget in table_frame.winfo_children():
                widget.destroy()

            # Header row
            tk.Label(table_frame, text="File Name", font=("Arial", 12, "bold"), width=50, anchor="w").grid(row=0, column=0, padx=5, pady=5)
            tk.Label(table_frame, text="View", font=("Arial", 12, "bold"), width=10).grid(row=0, column=1, padx=5, pady=5)
            tk.Label(table_frame, text="Delete", font=("Arial", 12, "bold"), width=10).grid(row=0, column=2, padx=5, pady=5)

            # Calculate start and end indices for pagination
            start_index = current_page * items_per_page 
            end_index = start_index + items_per_page
            
            for i, file in enumerate(files[start_index:end_index]):
                tk.Label(table_frame, text=file.name, font=("Arial", 12), width=50, anchor="w").grid(row=i + 1, column=0, padx=5, pady=5)
                tk.Button(table_frame, text="View", command=lambda f=file: open_file(f), bg="green", fg="white").grid(row=i + 1, column=1, padx=5, pady=5)
                tk.Button(table_frame, text="Delete", command=lambda f=file: delete_file(f), bg="red", fg="white").grid(row=i + 1, column=2, padx=5, pady=5)

        def next_page():
            """Navigate to the next page if available."""
            nonlocal current_page
            if (current_page + 1) * items_per_page < len(files):
                current_page += 1
                display_files()

        def previous_page():
            """Navigate to the previous page if available."""
            nonlocal current_page
            if current_page > 0:
                current_page -= 1
                display_files()

        # Navigation buttons
        nav_frame = tk.Frame(center_view_merged_window)
        nav_frame.pack(pady=10)

        tk.Button(nav_frame, text="Previous", command=previous_page).pack(side=tk.LEFT, padx=5)
        tk.Button(nav_frame, text="Next", command=next_page).pack(side=tk.LEFT, padx=5)

        # Initial display of files
        display_files()

    # Back button
    def go_back():
        view_merged_window.destroy()

    tk.Button(center_view_merged_window, text="Back", command=go_back).pack(pady=10)


def exit_program():
    root.destroy()

# Initialize Tkinter window
root = tk.Tk()
root.title("Facilitation Manager")
root.attributes('-fullscreen', True)  # Full Screen Size
root.configure()  # Set background color for a clean UI

# Run initial scripts
initial_scripts()

# Center Frame to hold all content
center_frame = Frame(root)
center_frame.pack(expand=True)  # Expands to center all content

# Title Frame (for Logo + Title)
title_frame = Frame(center_frame)
title_frame.pack(pady=10)

# Load and display logo
try:
    img = Image.open("./assets/qc_logo.png")  # Update with correct path if needed
    img = img.resize((70, 70))  # Resize for a clean UI
    logo_img = ImageTk.PhotoImage(img)
    logo_label = Label(title_frame, image=logo_img)
    logo_label.pack(side="left", padx=10)  # Adds space between logo and title
except Exception as e:
    print(f"Error loading logo: {e}")

# Title Label
title_label = Label(title_frame, text="WELCOME TO THE HOME PAGE", font=("Arial", 20, "bold"))
title_label.pack(side="left")  # Keep the title next to the logo

# Button Frame (for better spacing)
button_frame = Frame(center_frame)
button_frame.pack(pady=10)

# Buttons
tk.Button(button_frame, text="View Facilitation Categories", command=manage_categories, width=30).pack(pady=5)
tk.Button(button_frame, text="View To Merge File/s", command=viewToMergeFiles, width=30).pack(pady=5)
tk.Button(button_frame, text="View Merged Files", command=viewMergedFiles, width=30).pack(pady=5)
tk.Button(button_frame, text="Exit", command=root.destroy, width=30, bg="red", fg="white").pack(pady=5)

# Run the application
root.mainloop()

# print(f"Running the Python program...")
# main()
# print(f"Finished running the program...")