import pandas as pd
from bs4 import BeautifulSoup

excel_file = "AZ-900 data.xlsx"
try:
    existing_data = pd.read_excel(excel_file)
except FileNotFoundError:
    existing_data = pd.DataFrame()

with open('data.txt', 'r') as file:
    html_content = file.read()

soup = BeautifulSoup(html_content, 'html.parser')
cards = soup.find_all(class_="card")

data = []

for card in cards:
    question = card.find(class_="question_text").find('p').text.strip()
    if question not in existing_data['Question'].values:
        correct_answer = card.find(class_="answer_block green accent-1").find('p').find('strong').text.strip()
        choices_list = card.find(class_="choices-list list-unstyled").find_all('li')
        choices = [choice.text.strip() for choice in choices_list]

        data.append({    
            'Question': question,
            'Correct Answer': correct_answer,
            'Option1': choices[0] if len(choices) > 0 else '',
            'Option2': choices[1] if len(choices) > 1 else '',
            'Option3': choices[2] if len(choices) > 2 else '',
            'Option4': choices[3] if len(choices) > 3 else '',
            'Option5': choices[4] if len(choices) > 4 else '',
            'Option6': choices[5] if len(choices) > 5 else '',
        })

combined_data = pd.concat([existing_data, pd.DataFrame(data)], ignore_index=True)

combined_data.to_excel(excel_file, index=False)
print(f"Data appended to {excel_file}")
