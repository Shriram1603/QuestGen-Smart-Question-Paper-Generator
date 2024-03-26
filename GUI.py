from main import sem_paper, ca2_paper
from tkinter import *
from tkinter import filedialog
import os

#=====================================UI---------------------------------------------


# from main import sem_paper, ca2_paper

# Create a Tkinter instance
root = Tk()

# Set the window title and size
root.title("Question Paper Generator")
root.geometry('1920x1920')

# Global variables to store inputs
subject_name = ""
sub_code = ""
Paper_code = ""
no_of_questions=0
# Global variable to store the path of the generated documents
output_docx_path = ""
output_docx_filled_path = ""
global start,end
start=-1
end=-1

# Function to update global variables with input values
def Unit1():
    global start,end
    if(start==-1):
        start=0
    else:
        end=0
    
def Unit2():
    global start,end
    if(start==-1):
        start=2
    else:
        end=2
def Unit3():
    global start,end
    if(start==-1):
        start=4
    else:
        end=4
def Unit4():
    global start,end
    if(start==-1):
        start=6
    else:
        end=6
def Unit5():
    global start,end
    if(start==-1):
        start=8
    else:
        end=8
def get_inputs():
    global subject_name, sub_code, Paper_code, weightage,MonthYear,semester
    subject_name = subject_name_entry.get()
    sub_code = sub_code_entry.get()
    Paper_code = Paper_code_entry.get()
    weightage=  int(no_of_questions.get())
    MonthYear=Month_Year.get()
    semester=Semester.get()

# Callback function to open file dialog
def openFile():
    global path
    path = filedialog.askopenfilename()
    print(path)

# Callback function to generate question paper for Semester/Model
def sem_button_callback():
    get_inputs()
    sem_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code,MonthYear,semester)
    sem_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code,MonthYear,semester)
    # Update the global variables with the paths of the generated documents
    global output_docx_path, output_docx_filled_path
    output_docx_path = "output_questions.docx"
    output_docx_filled_path = "output_questions_filled.docx"
    # Open the generated documents
    open_generated_docx()

# Callback function to generate question paper for CA1
# def ca1_paper_callback():
#     get_inputs()
#     ca1_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code,weightage)
#     ca1_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code,weightage)
#     # Update the global variables with the paths of the generated documents
#     global output_docx_path, output_docx_filled_path
#     output_docx_path = "output_questions.docx"
#     output_docx_filled_path = "output_questions_filled.docx"
#     # Open the generated documents
#     open_generated_docx()

# Callback function to generate question paper for CA2
def ca2_paper_callback():
    get_inputs()
    ca2_paper(path, "output_questions.docx", Paper_code, subject_name, sub_code,weightage,Month_Year,semester,start,end)
    ca2_paper(path, "output_questions_filled.docx", Paper_code, subject_name, sub_code,weightage,Month_Year,semester,start,end)
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
label.grid(row=0, column=0, columnspan=2, pady=20)

button = Button(root, text='Open File', command=openFile)
button.grid(row=1, column=0, columnspan=2, pady=20)

subject_name_label = Label(root, text='Enter Subject Name:', font=("Helvetica", 16))
subject_name_label.grid(row=2, column=0, pady=(20, 5), sticky='w')
subject_name_entry = Entry(root, font=("Helvetica", 16))
subject_name_entry.grid(row=2, column=1, pady=(20, 5), sticky='w')

sub_code_label = Label(root, text='Enter Subject Code:', font=("Helvetica", 16))
sub_code_label.grid(row=3, column=0, pady=(20, 5), sticky='w')
sub_code_entry = Entry(root, font=("Helvetica", 16))
sub_code_entry.grid(row=3, column=1, pady=(20, 5), sticky='w')

Paper_code_label = Label(root, text='Enter Paper Code:', font=("Helvetica", 16))
Paper_code_label.grid(row=4, column=0, pady=(20, 5), sticky='w')
Paper_code_entry = Entry(root, font=("Helvetica", 16))
Paper_code_entry.grid(row=4, column=1, pady=(20, 5), sticky='w')

no_of_questions_label = Label(root, text='Enter the weightage of the first unit (Part-A):', font=("Helvetica", 16))
no_of_questions_label.grid(row=5, column=0, pady=(20, 5), sticky='w')
no_of_questions = Entry(root, font=("Helvetica", 16))
no_of_questions.grid(row=5, column=1, pady=(20, 5), sticky='w')

Month_Year_label = Label(root, text='Month _ Year :', font=("Helvetica", 16))
Month_Year_label.grid(row=6, column=0, pady=(20, 5), sticky='w')
Month_Year = Entry(root, font=("Helvetica", 16))
Month_Year.grid(row=6, column=1, pady=(20, 5), sticky='w')

Semester_label = Label(root, text='Semester :', font=("Helvetica", 16))
Semester_label.grid(row=7, column=0, pady=(20, 5), sticky='w')
Semester = Entry(root, font=("Helvetica", 16))
Semester.grid(row=7, column=1, pady=(20, 5), sticky='w')



Unit_buttons_frame = Frame(root)
Unit_buttons_frame.grid(row=9, column=0, columnspan=5, pady=(0, 20))

label = Label(root, text='Choose unit :', font=("Helvetica", 24))
label.grid(row=0, column=0, padx=10,pady=20,)

Unit1_button = Button(Unit_buttons_frame, text='Unit 1', font=("Helvetica", 24), command=Unit1)
Unit1_button.grid(row=0, column=1, padx=10, pady=20)

Unit2_button = Button(Unit_buttons_frame, text='Unit 2', font=("Helvetica", 24), command=Unit2)
Unit2_button.grid(row=0, column=2, padx=10, pady=20)

Unit3_button = Button(Unit_buttons_frame, text='Unit 3', font=("Helvetica", 24), command=Unit3)
Unit3_button.grid(row=0, column=3, padx=10, pady=20)

Unit4_button = Button(Unit_buttons_frame, text='Unit 4', font=("Helvetica", 24), command=Unit4)
Unit4_button.grid(row=0, column=4, padx=10, pady=20)

Unit5_button = Button(Unit_buttons_frame, text='Unit 5', font=("Helvetica", 24), command=Unit5)
Unit5_button.grid(row=0, column=5, padx=10, pady=20)

generate_label = Label(root, text='Which format of the question paper do you want?', font=("Helvetica", 24))
generate_label.grid(row=10, column=0, columnspan=5, pady=(20, 10))

sem_button = Button(root, text='Semester/Model', font=("Helvetica", 24), command=sem_button_callback)
sem_button.grid(row=11, column=0, columnspan=2, padx=10, pady=20)

ca2_button = Button(root, text='Generate CA question', font=("Helvetica", 24), command=ca2_paper_callback)
ca2_button.grid(row=11, column=3, columnspan=2,padx=10, pady=45)

# Run the Tkinter event loop
root.mainloop()
