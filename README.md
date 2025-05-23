# 📝 answer_finder

A handy Python-based tool that automatically extracts quiz questions from DOCX lab manuals, generates professional student-like answers using Google Gemini, and inserts the answers back into the DOCX in a clean, structured format.

---

## 🚀 Features

- **Extracts questions** from DOCX lab manuals.
- **Generates detailed answers** using Google Gemini API.
- **Inserts answers** directly under each question in the DOCX file.
- **Maintains formatting** and styles (bullets, bold, italics, tables, diagrams).
- **Professional structured answers** (Intro ➔ Body ➔ Conclusion).

---

## 🗂️ Folder Structure

```bash
pipeline/
    ├── answer_generator.py # Script to generate answers using Gemini API
    ├── answers_data.json # Generated answers saved as JSON
    ├── document_writer.py # Script to insert answers into DOCX
    ├── lab_manual_with_answers.docx # Final output lab manual with inserted answers
    ├── main.py # Main controller script (extract ➔ generate ➔ insert)
    ├── question_extractor.py # Script to extract questions from DOCX
    └── questions_data.json # Extracted questions saved as JSON
```

---

---

## 📦 Requirements

- Python 3.10+
- `docx`, `PyMuPDF`, `google-generativeai`, `dotenv`
- Google `Gemini API key` or any other LLM API key

Install dependencies:

```bash
pip install -r requirements.txt
```

## ⚙ Setup

##### 1. Clone the repo

```bash
git clone https://github.com/abhipatel2005/answer_finder.git
cd pipeline
```

##### 2.Get a Google Gemini API Key:

- Go to Google AI Studio
- Get your Gemini Flash 2.0 API key.

##### 3. Add your API key to .env:

```bash
GEMINI_API_KEY=your_key_here
```

## 🏃 How to Run

##### Basic usage:

```bash
python main.py --input "path/to/your/lab_manual.docx" --output "output_with_answers.docx"
```

##### Output:

- `questions_data.json` with extracted questions.

- `answers_data.json` with generated answers.

- `answers.docx` with answers neatly inserted.

## 🛠️ Troubleshooting

- Inserted 0 answers: Ensure `answers_data.json` matches questions. Try deleting the file and rerunning.

- Gemini errors: Check your API key in `.env` and that you're using gemini-2.0-flash.

- Formatting issues: The tool expects numbered questions like 1. What is .... Ensure your DOCX follows this pattern.

## 🙌 Credits

- Created by Abhi Patel.
- Powered by Google Gemini.
