## Description

This script is designed to extract questions and answers from a table in a Word document (.docx) and convert them into an XML format compatible with Moodle. It utilizes the `python-docx` library for handling Word documents and `pandas` for data processing.

## Installation

Before using the script, ensure you have the required libraries installed. You can install them using pip:

- `pip install python-docx pandas`

## Usage

1. **Prepare the Word Document**: Ensure your Word document contains a table where each group of 5 rows corresponds to one question. The first row should contain the question number and text, while the next 4 rows should contain the answer options.

2. **Set the File Path**: In the script, change the line `R"path\to\docx\file"` to the path of your .docx file.

3. **Run the Script**: Execute the script in your Python environment. It will extract data from the table and create a `questions.xml` file in the current directory.

## Table Structure

The table in the Word document should have the following structure:

| Question Number  | Question Content                    | Answer Options    |
|------------------|-------------------------------------|-------------------|
| 1                | What is the sky color?              | 1. Blue           |
|                  |                                     | 2. Green          |
|                  |                                     | 3. Red            |
|                  |                                     | 4. Yellow         |
|                  |                                     | 5. Black          |
| 2                | What is the grass color?            | 1. Blue           |
|                  |                                     | 2. Green          |
|                  |                                     | 3. Red            |
|                  |                                     | 4. Yellow         |
|                  |                                     | 5. Black          |
| 3                | What is the capital of France?      | 1. Berlin         |
|                  |                                     | 2. Madrid         |
|                  |                                     | 3. Paris          |
|                  |                                     | 4. Rome           |
|                  |                                     | 5. Lisbon         |
| 4                | Which is a programming language?    | 1. HTML           |
|                  |                                     | 2. CSS            |
|                  |                                     | 3. Python         |
|                  |                                     | 4. SQL            |
|                  |                                     | 5. Markdown       |
| 5                | What is 2 + 2?                      | 1. 3              |
|                  |                                     | 2. 4              |
|                  |                                     | 3. 5              |
|                  |                                     | 4. 6              |
|                  |                                     | 5. 7              |
|                  |                                     | 6. 8              |
|                  |                                     | 7. 9              |
|                  |                                     | 8. 10             |
| 6                | What is the largest planet?         | 1. Earth          |
|                  |                                     | 2. Mars           |
|                  |                                     | 3. Jupiter        |
|                  |                                     | 4. Saturn         |
|                  |                                     | 5. Neptune        |
|                  |                                     | 6. Uranus         |
|                  |                                     | 7. Venus          |
|                  |                                     | 8. Mercury        |
| 7                | What is the boiling point of water? | 1. 100°C          |
|                  |                                     | 2. 0°C            |
|                  |                                     | 3. 50°C           |
|                  |                                     | 4. 25°C           |
|                  |                                     | 5. 75°C           |
|                  |                                     | 6. 150°C          |
|                  |                                     | 7. 200°C          |
|                  |                                     | 8. 300°C          |
| 8                | Which of the following is correct?  | 1. *Option A      |
|                  |                                     | 2. Option B       |
|                  |                                     | 3. Option C       |
|                  |                                     | 4. Option D       |
|                  |                                     | 5. Option E       |

**Note**: 
- If the question number is missing, the script will automatically assign a number based on the index.
- The correct answer can be indicated by making the text bold (e.g., `*Option A`).
- The script can handle questions with varying numbers of answer options, as long as they are grouped in sets of 5 rows.
