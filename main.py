from docx import Document
import pandas as pd

# Reading the Word document
doc = Document(R"path\to\docx\file")

# Extracting data from the table
data = []
table = doc.tables[0]

# Function to extract text from a cell, accounting for empty values
def extract_text(cell):
    return cell.text.strip() if cell.text.strip() else None

# Ensure that the number of rows is a multiple of 5
row_count = len(table.rows)
if row_count % 5 != 0:
    print(f"Warning: The number of rows ({row_count}) is not a multiple of 5. Some questions may be skipped.")

for i in range(0, row_count, 5):  # Reading 5 rows (5 rows per question)
    try:
        question_row = table.rows[i]
        question_number = extract_text(question_row.cells[0])
        question = extract_text(question_row.cells[1])

        # If the question number is missing, use the index as the question number
        if not question_number:
            question_number = str(i // 5 + 1)

        answers = []
        for j in range(0, 5):  # Extracting 5 rows with answer options
            answer_row = table.rows[i + j]
            answer_text = extract_text(answer_row.cells[2])
            is_bold = any(run.bold for run in answer_row.cells[2].paragraphs[0].runs)
            correct_indicator = '*' if is_bold else ''
            
            # Removing indicators before the answer text (letter, dot, comma, and space)
            if '.' in answer_text:
                answer_text = answer_text.split('.', 1)[1].strip()
            elif ',' in answer_text:
                answer_text = answer_text.split(',', 1)[1].strip()

            answers.append(f"{correct_indicator}{answer_text}")

        formatted_answers_str = '\n'.join(answers)
        data.append([question_number, question, formatted_answers_str])
    except IndexError:
        # Error handling for incomplete data with question number output
        remaining_rows = len(table.rows) - i
        question_number = extract_text(table.rows[i].cells[0]) if i < len(table.rows) else str(i // 5 + 1)
        print(f"Not enough rows for question number {question_number}, starting from row {i}. Remaining rows: {remaining_rows}. Skipping.")

# Converting data to DataFrame for further processing
df = pd.DataFrame(data, columns=['Question Number', 'Question Content', 'Answer Options'])

# Creating an XML file for Moodle
with open('questions.xml', 'w', encoding='utf-8') as file:
    file.write('<quiz>\n')
    for index, row in df.iterrows():
        file.write('  <question type="multichoice">\n')
        file.write(f'    <name>\n      <text>Question {row["Question Number"]}</text>\n    </name>\n')
        file.write(f'    <questiontext format="html">\n      <text><![CDATA[{row["Question Content"]}]]></text>\n    </questiontext>\n')
        answers = row["Answer Options"].split('\n')
        for answer in answers:
            correct = '100' if answer.startswith('*') else '0'
            answer_text = answer.lstrip('*')
            file.write(f'    <answer fraction="{correct}">\n')
            file.write(f'      <text><![CDATA[{answer_text}]]></text>\n')
            file.write('    </answer>\n')
        file.write('  </question>\n')
    file.write('</quiz>')
