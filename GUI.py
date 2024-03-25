from main import sem_paper, ca1_paper, ca2_paper
from tkinter import *
from tkinter import filedialog
import os

# Create a Tkinter instance
root = Tk()

# Set the window title and size
root.title("Question Paper Generator")
root.geometry('1080x920')

# Global variables to store inputs
subject_name = ""
sub_code = ""
Paper_code = ""
no_of_questions=0
# Global variable to store the path of the generated documents
output_docx_path = ""
output_docx_filled_path = ""

# Function to update global variables with input values
def get_inputs():
    global subject_name, sub_code, Paper_code, weightage,Month_Year,Semester
    subject_name = subject_name_entry.get()
    sub_code = sub_code_entry.get()
    Paper_code = Paper_code_entry.get()
    weightage=  int(no_of_questions.get())
    Month_Year=Month_Year.get()
    Semester=Semester.get()

# Callback function to open file dialog
def openFile():
    global path
    path = filedialog.askopenfilename()
    print(path)

# Callback function to generate question paper for Semester/Model
def sem_button_callback():
    get_inputs()
    sem_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code)
    sem_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code)
    # Update the global variables with the paths of the generated documents
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    # Open the generated documents
    open_generated_docx()

# Callback function to generate question paper for CA1
def ca1_paper_callback():
    get_inputs()
    ca1_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code,weightage)
    ca1_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code,weightage)
    # Update the global variables with the paths of the generated documents
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    # Open the generated documents
    open_generated_docx()

# Callback function to generate question paper for CA2
def ca2_paper_callback():
    get_inputs()
    ca2_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code,weightage,Month_Year,Semester)
    ca2_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code,weightage,Month_Year,Semester)
    # Update the global variables with the paths of the generated documents
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    # Open the generated documents
    open_generated_docx()

# Function to open the generated documents
def open_generated_docx():
    if output_docx_path:
        os.system(f'open "{output_docx_path}"')
    if output_docx_filled_path:
        os.system(f'open "{output_docx_filled_path}"')

# Create and pack the widgets
label = Label(root, text='Select the formatted question bank', font=("Helvetica", 24))
label.pack(pady=20)

button = Button(root, text='Open File', command=openFile)
button.pack(pady=20)

subject_name_label = Label(root, text='Enter Subject Name:', font=("Helvetica", 16))
subject_name_label.pack(pady=10)
subject_name_entry = Entry(root, font=("Helvetica", 16))
subject_name_entry.pack(pady=5)

sub_code_label = Label(root, text='Enter Subject Code:', font=("Helvetica", 16))
sub_code_label.pack(pady=10)
sub_code_entry = Entry(root, font=("Helvetica", 16))
sub_code_entry.pack(pady=5)

Paper_code_label = Label(root, text='Enter Paper Code:', font=("Helvetica", 16))
Paper_code_label.pack(pady=10)
Paper_code_entry = Entry(root, font=("Helvetica", 16))
Paper_code_entry.pack(pady=5)

no_of_questions = Label(root, text='Enter the weightage of the first unit (Part-A):', font=("Helvetica", 16))
no_of_questions.pack(pady=10)
no_of_questions = Entry(root, font=("Helvetica", 16))
no_of_questions.pack(pady=5)

Month_Year = Label(root, text='Month _ Year :', font=("Helvetica", 16))
Month_Year.pack(pady=10)
Month_Year = Entry(root, font=("Helvetica", 16))
Month_Year.pack(pady=5)

Semester = Label(root, text='Semester :', font=("Helvetica", 16))
Semester.pack(pady=10)
Semester = Entry(root, font=("Helvetica", 16))
Semester.pack(pady=5)

generate_label = Label(root, text='Which format of the question paper do you want?', font=("Helvetica", 24))
generate_label.pack(pady=20)

sem_button = Button(root, text='Semester/Model', font=("Helvetica", 24), command=sem_button_callback)
ca1_button = Button(root, text='CA1', font=("Helvetica", 24), command=ca1_paper_callback)
ca2_button = Button(root, text='CA2', font=("Helvetica", 24), command=ca2_paper_callback)
sem_button.pack(pady=20)
ca1_button.pack(pady=20)
ca2_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
