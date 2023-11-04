import random
import pandas as pd
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
import sys
import os

if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.dirname(__file__))

excel_file_path = os.path.join(base_dir, 'AZ-900 data.xlsx')

root = tk.Tk()
root.title('AZ-900 training')
root.geometry('800x500')
root.minsize(800, 500)

style = ThemedStyle(root)
style.set_theme('clam')

style.configure('My.TFrame', background='#206bc7')

page1 = ttk.Frame(root)
page1.pack(fill='both', expand=True)
page2 = ttk.Frame(root)
page2.pack(fill='both', expand=True)
page3 = ttk.Frame(root)
page3.pack(fill='both', expand=True)

def show_page(page):
    if page == page1:
        page3.pack_forget()
        page2.pack_forget()
        page1.pack(fill='both', expand=True)
    elif page == page2:
        page1.pack_forget()
        page2.pack(fill='both', expand=True)
    elif page == page3:
        page2.pack_forget()
        page3.pack(fill='both', expand=True)

to_page_3_button = ttk.Button(page2, text="Stop", command=lambda: (show_page(page3), results()))
to_page_3_button.pack(anchor='ne',padx=20, pady=20)

to_page_1_button = ttk.Button(page3, text="Back to menu", command=lambda: (show_page(page1), reset()))
to_page_1_button.pack(anchor='ne',padx=20, pady=20)

frame2_in_page2 = ttk.Frame(page2, style='My.TFrame')
frame2_in_page2.pack(side='bottom', fill='x', expand=False)

df = pd.read_excel(excel_file_path)
questions = df.to_dict(orient='records')  

random.shuffle(questions)

question_index = 0

amount_of_questions = len(questions)

azure_label = ttk.Label(page1, text="AZ-900", font=('', 90), foreground="#206bc7")
azure_label.pack(anchor='center', padx=0, pady=30  )
azure_label2 = ttk.Label(page1, text=f"There are {amount_of_questions} questions", font=('', 12), foreground="#206bc7",)
azure_label2.pack(anchor='center', padx=0, pady=0  )

question_label = ttk.Label(page2, text="", justify='center' ,wraplength=500)
question_label.pack(anchor='center', padx=10, pady=10)

answer_label =ttk.Label(page2, text="")
answer_label.pack(anchor='center', padx=10, pady=10)

def remove_choices():
    for widget in page2.winfo_children():
        if isinstance(widget, ttk.Radiobutton):
            widget.destroy()


def new_question():
    global question_index, random_correct_answer, selected_option, random_choices, radio_button

    if question_index < len(questions):
        random_question = questions[question_index]['Question']
        random_correct_answer = str(questions[question_index]['Correct Answer'])
        random_choices = [
            str(questions[question_index]['Option1']).strip(), 
            str(questions[question_index]['Option2']).strip(), 
            str(questions[question_index]['Option3']).strip(), 
            str(questions[question_index]['Option4']).strip(),
            str(questions[question_index]['Option5']).strip()
        ]

        remove_choices()  

        selected_option = tk.StringVar()
        
        for choice in random_choices:
            if choice != 'nan':
                radio_button = ttk.Radiobutton(page2, text=choice, variable=selected_option, value=choice)
                radio_button.pack()

        question_label.configure(text=f"Question:\n{random_question}")
        answer_label.configure(text="")
        
        question_index += 1
        if question_index == len(questions):
            next_question_button.pack_forget()
    

amount_correct = 0
amount_answered = 0

def check_option():
    global amount_correct, amount_answered

    if selected_option.get()[0] in random_correct_answer:
        answer_label.configure(text="Correct")
        amount_correct += 1
        amount_answered += 1
    else:
        answer_label.configure(text=f'Wrong, correct answer is:\n{random_correct_answer}', justify='center')
        amount_answered += 1
    

to_page_2_button = ttk.Button(page1, text="Start", command=lambda: (show_page(page2), new_question()))
to_page_2_button.pack(anchor='center',padx=40, pady=40)

next_question_button = ttk.Button(frame2_in_page2, text="Next Question", command=new_question)
next_question_button.pack(side='right', padx=20, pady=20)

submit_answer_button = ttk.Button(frame2_in_page2, text='Submit', command=check_option)
submit_answer_button.pack(side='right', padx=20, pady=20)

results_label = ttk.Label(page3, text='', font=('Tahome', 25), foreground='#206bc7')
results_label.pack(anchor='center', padx=10, pady=10)



def results():
    global results_label
    precentage_correct = 0.0
    if amount_answered != 0:
        precentage_correct = (100 / amount_answered) * amount_correct
    results_label.configure(text=f'Score: {precentage_correct:.0f}% \n{amount_correct} out of {amount_answered} answers were correct',)


def reset():
    global question_index, amount_correct, amount_answered
    next_question_button.pack(side='right', padx=20, pady=20)
    random.shuffle(questions)
    question_index = 0
    amount_answered = 0
    amount_correct = 0
    remove_choices

    
show_page(page1)

root.mainloop()