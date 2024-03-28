from main import sem_paper, ca2_paper
from tkinter import *
from tkinter import filedialog
import os

# Create a Tkinter instance
root = Tk()

# Set the window title
root.title("Question Paper Generator")

# Create a Canvas widget
canvas = Canvas(root)
canvas.pack(side=LEFT, fill="both", expand=True)

# Add a scrollbar
scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.pack(side=RIGHT, fill="y")

# Configure the canvas
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Create a frame inside the canvas to hold the widgets
frame = Frame(canvas)

# Add the frame to the canvas
canvas.create_window((0, 0), window=frame, anchor="nw")

# Global variables to store inputs
subject_name = ""
sub_code = ""
Paper_code = ""
no_of_questions = 0
output_docx_path = ""
output_docx_filled_path = ""
global start, end
start = -1
end = -1

# Function to update global variables with input values
def Unit1():
    global start, end
    if start == -1:
        start = 0
    else:
        end = 0

# Function to update global variables with input values
def Unit2():
    global start, end
    if start == -1:
        start = 2
    else:
        end = 2

# Function to update global variables with input values
def Unit3():
    global start, end
    if start == -1:
        start = 4
    else:
        end = 4

# Function to update global variables with input values
def Unit4():
    global start, end
    if start == -1:
        start = 6
    else:
        end = 6

# Function to update global variables with input values
def Unit5():
    global start, end
    if start == -1:
        start = 8
    else:
        end = 8

# Function to get user inputs
def get_inputs():
    global subject_name, sub_code, Paper_code, weightage, MonthYear, semester
    subject_name = subject_name_entry.get()
    sub_code = sub_code_entry.get()
    Paper_code = Paper_code_entry.get()
    weightage = int(no_of_questions.get())
    MonthYear = Month_Year.get()
    semester = Semester.get()

# Callback function to open file dialog
def openFile():
    global path
    path = filedialog.askopenfilename()
    print(path)

# Callback function to generate question paper for Semester/Model
def sem_button_callback():
    get_inputs()
    sem_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code, MonthYear, semester)
    sem_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code, MonthYear, semester)
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    open_generated_docx()

# Callback function to generate question paper for CA2
def ca2_paper_callback():
    get_inputs()
    ca2_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code, weightage, Month_Year, semester, start, end)
    ca2_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code, weightage, Month_Year, semester, start, end)
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    open_generated_docx()

# Function to open the generated documents
def open_generated_docx():
    if output_docx_path:
        os.system(f'open "{output_docx_path}"')
    if output_docx_filled_path:
        os.system(f'open "{output_docx_filled_path}"')

# Create and pack the widgets
label = Label(frame, text='Select the formatted question bank', font=("Helvetica", 18))
label.pack(pady=10)

button = Button(frame, text='Open File', command=openFile)
button.pack(pady=10)

subject_name_label = Label(frame, text='Enter Subject Name:', font=("Helvetica", 14))
subject_name_label.pack(pady=5)
subject_name_entry = Entry(frame, font=("Helvetica", 14))
subject_name_entry.pack(pady=5)

sub_code_label = Label(frame, text='Enter Subject Code:', font=("Helvetica", 14))
sub_code_label.pack(pady=5)
sub_code_entry = Entry(frame, font=("Helvetica", 14))
sub_code_entry.pack(pady=5)

Paper_code_label = Label(frame, text='Enter Paper Code:', font=("Helvetica", 14))
Paper_code_label.pack(pady=5)
Paper_code_entry = Entry(frame, font=("Helvetica", 14))
Paper_code_entry.pack(pady=5)

no_of_questions_label = Label(frame, text='Enter weightage of Part-A:', font=("Helvetica", 14))
no_of_questions_label.pack(pady=5)
no_of_questions = Entry(frame, font=("Helvetica", 14))
no_of_questions.pack(pady=5)

Month_Year_label = Label(frame, text='Month & Year:', font=("Helvetica", 14))
Month_Year_label.pack(pady=5)
Month_Year = Entry(frame, font=("Helvetica", 14))
Month_Year.pack(pady=5)

Semester_label = Label(frame, text='Semester:', font=("Helvetica", 14))
Semester_label.pack(pady=5)
Semester = Entry(frame, font=("Helvetica", 14))
Semester.pack(pady=5)

Unit_buttons_frame = Frame(frame)
Unit_buttons_frame.pack(pady=20)

label = Label(frame, text='Choose unit :', font=("Helvetica", 18))
label.pack(pady=10)

Unit1_button = Button(Unit_buttons_frame, text='Unit 1', font=("Helvetica", 14), command=Unit1)
Unit1_button.pack(side=LEFT, padx=5)

Unit2_button = Button(Unit_buttons_frame, text='Unit 2', font=("Helvetica", 14), command=Unit2)
Unit2_button.pack(side=LEFT, padx=5)

Unit3_button = Button(Unit_buttons_frame, text='Unit 3', font=("Helvetica", 14), command=Unit3)
Unit3_button.pack(side=LEFT, padx=5)

Unit4_button = Button(Unit_buttons_frame, text='Unit 4', font=("Helvetica", 14), command=Unit4)
Unit4_button.pack(side=LEFT, padx=5)

Unit5_button = Button(Unit_buttons_frame, text='Unit 5', font=("Helvetica", 14), command=Unit5)
Unit5_button.pack(side=LEFT, padx=5)

generate_label = Label(frame, text='Which format of the question paper do you want?', font=("Helvetica", 18))
generate_label.pack(pady=10)

sem_button = Button(frame, text='Semester/Model', font=("Helvetica", 16), command=sem_button_callback)
sem_button.pack(pady=10)

ca2_button = Button(frame, text='Generate CA question', font=("Helvetica", 16), command=ca2_paper_callback)
ca2_button.pack(pady=10)

# Update the scroll region to include the new widgets
frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))

# Run the Tkinter event loop
root.mainloop()
