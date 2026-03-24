


import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os


# Variables
countries = [
    'Argentina', 'Australia', 'Belgium', 'Brazil', 'Canada',
    'China', 'Denmark', 'Egypt', 'France', 'Germany', 'Greece', 
    'India', 'Indonesia', 'Italy', 'Japan', 'Mexico', 'Netherlands',
    'Nigeria', 'Norway', 'Pakistan', 'Poland', 'Portugal', 'Russia', 
    'Saudi Arabia', 'South Africa', 'South Korea', 'Spain', 'Sweden', 
    'Switzerland', 'Turkey', 'UAE', 'UK', 'USA', 'Ukraine'
    ]

titles_general = [
    "Mr.", "Mrs.", "Miss", "Ms.",
    "Sir", "Madam", "Mx."
]





window = tk.Tk()
window.title("Data Entry Form")

frame = tk.Frame(window)
frame.pack(anchor="nw") 


# Functions

def enter_data():
    accepted = accept_var.get()
    
    if accepted == "Accepted":
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        if firstname and lastname:
            
            
            
            data = {
            "List Number" : "------" ,
            "First Name" : first_name_entry.get(),
            "Last Name" : last_name_entry.get(),
            "Title" : title_combobox.get(),
            "Age" : age_spinbox.get(),
            "Nationality" : nationality_label_combobox.get(),

            "# Courses": numcourses_spinbox.get(),
            "# Semesters" : numsemesters_spinbox.get(),
            "Registration Status" : reg_status_var.get()
            }
            
            filepath = r"C:\Users\Salih\Desktop\Tkinter Data Entry\data.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.append(list(data.keys()))
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append(list(data.values()))
            workbook.save(filepath) 

            return data
        else:
            tk.messagebox.showwarning(title="Error !" , message ="You have to type name and surname")
        

    else:
        tk.messagebox.showwarning(title ="Error !" , message ="You must have accepted terms." )



######################################## Saving User Info ########################################

user_info_frame = tk.LabelFrame( frame , text = "User Information")
user_info_frame.grid(row=0 , column=0 , sticky="news" , padx=20 ,pady=10)


# First and Last name
first_name_label = tk.Label(user_info_frame , text = "First Name")
first_name_label.grid(row=0 , column=0)
last_name_label = tk.Label(user_info_frame , text = "Last Name" )
last_name_label.grid(row=0 , column=1)

first_name_entry = tk.Entry(user_info_frame )
first_name_entry.grid(row=1 , column=0)
last_name_entry = tk.Entry(user_info_frame )
last_name_entry.grid(row=1 , column=1)


# Title
title_label = tk.Label(user_info_frame , text ="Title")
title_label.grid(row=0, column=2)
title_combobox = ttk.Combobox( user_info_frame , values = titles_general)
title_combobox.grid(row=1 , column=2)

# Age
age_label = tk.Label(user_info_frame ,text = "Age")
age_label.grid(row=2 , column=0)
age_spinbox = tk.Spinbox(user_info_frame , from_=18 , to=110)
age_spinbox.grid(row=3 , column=0)

# Nationality
nationality_label = tk.Label(user_info_frame , text = "Nationality")
nationality_label.grid(row=2 , column=1)
nationality_label_combobox = ttk.Combobox(user_info_frame , values = countries)
nationality_label_combobox.grid(row=3 , column= 1)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10 , pady=5)


######################################## Saving Course Info ########################################


course_frame = tk.LabelFrame(frame , text = "Registration Status")
course_frame.grid(row=1 , column=0 , sticky="news" , padx=20 , pady=10)

# Registered Button
reg_status_var = tk.StringVar(value= "Not Registered")
registered_check = tk.Checkbutton(course_frame , text = "Currently Registered" , variable=reg_status_var ,onvalue="Registered" , offvalue="Not registered")
registered_check.grid(row=1 , column=0)

# Number of Courses and Semesters 
numcourses_label = tk.Label(course_frame , text = "# Completed Courses")
numcourses_label.grid(row=1 , column=1)
numcourses_spinbox = tk.Spinbox(course_frame , from_=0 , to="infinity")
numcourses_spinbox.grid(row=2 , column=1) 

numsemesters_label = tk.Label(course_frame , text = "# Semesters")
numsemesters_label.grid(row=1 , column=2)
numsemesters_spinbox = tk.Spinbox(course_frame , from_=0 , to="infinity")
numsemesters_spinbox.grid(row=2 , column=2)

for widget in course_frame.winfo_children():
    widget.grid_configure(padx=10 , pady=5)


######################################## Accept Terms ########################################


terms_frame = tk.LabelFrame(frame ,text = "Terms & Conditions")
terms_frame.grid(row=2 , column=0 , sticky="news" , padx=20 , pady=10)

accept_var = tk.StringVar(value ="Not Accepted")
terms_check = tk.Checkbutton(terms_frame ,text ="I accept the terms and conditions." , variable=accept_var , onvalue="Accepted" , offvalue="Not Accepted")
terms_check.grid(row=0 , column=0)


######################################## Button ########################################


button = tk.Button(frame , text="Enter Data" , command=enter_data)
button.grid(row=3 , column=0 , padx=20 , pady=10)

window.mainloop()

