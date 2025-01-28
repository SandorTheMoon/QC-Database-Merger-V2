# File Merger for QC SHRU L&D Database
# Purpose: to automatically merge all the training data available
# Created by: Ian Salig U Batangan, Contact Details: isubatangan@gmail.com
# Versions Updates by: Airysh Xander M. Espero, Contact Details: derespero@gmail.com

import os 
import venv # For activating virtual environment
import pandas as pd # For data manipulation library, also needs openpyxl sub library of pandas
import re # String manipulation library
import glob # For finding file path library
from pathlib import Path
import msvcrt # For detecting button press
import tkinter as tk
from tkinter import simpledialog, messagebox

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
    print(variable_name)

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

def main():
    # Declaring the file path of the root directory
    root_path=Path.cwd()
    print(f'File path is {root_path}')

    # Define the folder and pattern to search for files
    to_merge_path = root_path / "ToMergeFiles"
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
    with open('FacilitationColumns.txt', 'r') as file:
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

# File to store facilitation categories
FILE_NAME = 'FacilitationColumns.txt'

def read_categories():
    """Read facilitation categories from the file."""
    if not os.path.exists(FILE_NAME):
        return []
    with open(FILE_NAME, 'r') as file:
        content = file.read().strip()
        return content.split(',')

def write_categories(categories):
    """Write facilitation categories to the file."""
    with open(FILE_NAME, 'w') as file:
        file.write(','.join(categories))

def view_categories():
    categories = read_categories()
    if not categories:
        messagebox.showinfo("View Categories", "No facilitation categories found.")
        return

    display = "\n".join(f"{i + 1}. {cat[:-2]}" for i, cat in enumerate(categories))
    messagebox.showinfo("Facilitation Categories", display)

def add_category():
    categories = read_categories()

    # Create a new window for adding a category
    add_window = tk.Toplevel(root)
    add_window.title("Add New Category")

    tk.Label(add_window, text="Current Categories:").pack(pady=5)

    # Display categories
    display = "\n".join(f"{i + 1}. {cat[:-2]}" for i, cat in enumerate(categories))
    tk.Label(add_window, text=display, justify="left").pack(pady=5)

    tk.Label(add_window, text="Enter new facilitation category:").pack(pady=5)
    category_entry = tk.Entry(add_window)
    category_entry.pack(pady=5)

    tk.Label(add_window, text=f"Enter position to insert (1-{len(categories) + 1}):").pack(pady=5)
    position_entry = tk.Entry(add_window)
    position_entry.pack(pady=5)

    def save_category():
        new_category = category_entry.get().upper()
        if not new_category:
            messagebox.showerror("Error", "Category name cannot be empty.")
            return

        if any(new_category in cat for cat in categories):
            messagebox.showerror("Error", f"Category '{new_category}' already exists.")
            return

        try:
            position = int(position_entry.get())
            if position < 1 or position > len(categories) + 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", f"Invalid position. Enter a number between 1 and {len(categories) + 1}.")
            return

        categories.insert(position - 1, f"{new_category}.+")
        write_categories(categories)
        messagebox.showinfo("Success", f"Category '{new_category}' added at position {position}.")
        add_window.destroy()

    tk.Button(add_window, text="Save", command=save_category).pack(pady=10)

def delete_category():
    categories = read_categories()
    if not categories:
        messagebox.showinfo("Delete Category", "No categories to delete.")
        return

    # Create a new window for deleting a category
    delete_window = tk.Toplevel(root)
    delete_window.title("Delete Category")

    tk.Label(delete_window, text="Current Categories:").pack(pady=5)

    # Display categories
    display = "\n".join(f"{i + 1}. {cat[:-2]}" for i, cat in enumerate(categories))
    tk.Label(delete_window, text=display, justify="left").pack(pady=5)

    tk.Label(delete_window, text=f"Enter position to delete (1-{len(categories)}):").pack(pady=5)
    position_entry = tk.Entry(delete_window)
    position_entry.pack(pady=5)

    def confirm_delete():
        try:
            position = int(position_entry.get())
            if position < 1 or position > len(categories):
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", f"Invalid position. Enter a number between 1 and {len(categories)}.")
            return

        deleted_category = categories.pop(position - 1)
        write_categories(categories)
        messagebox.showinfo("Success", f"Category '{deleted_category[:-2]}' deleted successfully.")
        delete_window.destroy()

    tk.Button(delete_window, text="Delete", command=confirm_delete).pack(pady=10)

def perform_data_merging():
    # Placeholder function for merging data
    messagebox.showinfo("Data Merging", "Registration and Post Evaluation files successfully merged!")

def exit_program():
    root.destroy()

# Initialize Tkinter window
root = tk.Tk()
root.title("Facilitation Manager")
root.geometry("400x300")

# Run initial scripts
initial_scripts()

# Title Label
title_label = tk.Label(root, text="WELCOME TO THE HOME PAGE", font=("Arial", 16), pady=10)
title_label.pack()

# Button Frame
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Buttons
tk.Button(button_frame, text="View Facilitation Categories", command=view_categories, width=30).pack(pady=5)
tk.Button(button_frame, text="Add New Facilitation Category", command=add_category, width=30).pack(pady=5)
tk.Button(button_frame, text="Delete Facilitation Category", command=delete_category, width=30).pack(pady=5)
tk.Button(button_frame, text="Perform Data Merging", command=perform_data_merging, width=30).pack(pady=5)
tk.Button(button_frame, text="Exit", command=exit_program, width=30).pack(pady=5)

# Run the application
root.mainloop()

# print(f"Running the Python program...")
# main()
# print(f"Finished running the program...")