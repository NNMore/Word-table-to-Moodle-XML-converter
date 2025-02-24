from docx import Document
import pandas as pd

# Чтение документа Word
doc = Document(R"path\to\docx\file")

# Извлечение данных из таблицы
data = []
table = doc.tables[0]

# Функция для извлечения текста из ячейки с учетом пустых значений
def extract_text(cell):
    return cell.text.strip() if cell.text.strip() else None

# Убедимся, что количество строк кратно 5
row_count = len(table.rows)
if row_count % 5 != 0:
    print(f"Предупреждение: Количество строк ({row_count}) не кратно 5. Некоторые вопросы могут быть пропущены.")

for i in range(0, row_count, 5):  # Чтение по 5 строк (5 строк на один вопрос)
    try:
        question_row = table.rows[i]
        question_number = extract_text(question_row.cells[0])
        question = extract_text(question_row.cells[1])

        # Если номер вопроса отсутствует, используем индекс как номер вопроса
        if not question_number:
            question_number = str(i // 5 + 1)

        answers = []
        for j in range(0, 5):  # Извлечение 5 строк с вариантами ответов
            answer_row = table.rows[i + j]
            answer_text = extract_text(answer_row.cells[2])
            is_bold = any(run.bold for run in answer_row.cells[2].paragraphs[0].runs)
            correct_indicator = '*' if is_bold else ''
            
            # Удаление обозначений перед текстом ответа (буква, точка, запятая и пробел)
            if '.' in answer_text:
                answer_text = answer_text.split('.', 1)[1].strip()
            elif ',' in answer_text:
                answer_text = answer_text.split(',', 1)[1].strip()

            answers.append(f"{correct_indicator}{answer_text}")

        formatted_answers_str = '\n'.join(answers)
        data.append([question_number, question, formatted_answers_str])
    except IndexError:
        # Обработка ошибок для неполных данных с выводом номера вопроса
        remaining_rows = len(table.rows) - i
        question_number = extract_text(table.rows[i].cells[0]) if i < len(table.rows) else str(i // 5 + 1)
        print(f"Недостаточно строк для вопроса номер {question_number}, начиная с строки {i}. Осталось строк: {remaining_rows}. Пропуск.")

# Преобразование данных в DataFrame для дальнейшей обработки
df = pd.DataFrame(data, columns=['Номер вопроса', 'Содержание вопроса', 'Варианты ответов'])

# Создание XML-файла для Moodle
with open('questions.xml', 'w', encoding='utf-8') as file:
    file.write('<quiz>\n')
    for index, row in df.iterrows():
        file.write('  <question type="multichoice">\n')
        file.write(f'    <name>\n      <text>Вопрос {row["Номер вопроса"]}</text>\n    </name>\n')
        file.write(f'    <questiontext format="html">\n      <text><![CDATA[{row["Содержание вопроса"]}]]></text>\n    </questiontext>\n')
        answers = row["Варианты ответов"].split('\n')
        for answer in answers:
            correct = '100' if answer.startswith('*') else '0'
            answer_text = answer.lstrip('*')
            file.write(f'    <answer fraction="{correct}">\n')
            file.write(f'      <text><![CDATA[{answer_text}]]></text>\n')
            file.write('    </answer>\n')
        file.write('  </question>\n')
    file.write('</quiz>')
