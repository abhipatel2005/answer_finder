import os
import asyncio
import json
from dotenv import load_dotenv
import google.generativeai as genai

# Load environment variables
load_dotenv()

# Initialize Gemini API
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

async def generate_detailed_answer(question, context):
    """Process quiz questions using Gemini API with detailed answers"""
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    prompt = f"""You are a student writing answers for a lab manual.
Please generate a clean, professional, educational answer to the following question, based on the lab manual content.

Format your answer for direct insertion into a Microsoft Word document:
- Start with a short introduction (2–3 lines).
- After that, create detailed points (bullet points if applicable).
- End with a short conclusion (optional).
- Insert <break> where a paragraph break should occur.
- For bullet points, start lines with '•'
- To format bold text, wrap with <b>text</b>
- To format italic text, wrap with <i>text</i>
- To format both bold and italic, wrap with <bi>text</bi>
- For tables, describe using: <table>row1col1|row1col2;row2col1|row2col2</table>
- For diagrams, describe with: <diagram>description</diagram>

Make sure the answer flows like a mini textbook explanation. Keep spacing and clarity.

Question: {question}

Context: {context[:10000]}
"""
    
    try:
        response = model.generate_content(prompt)
        answer = response.text.strip() if response.text else 'No answer generated.'
        print(f"\nQ: {question}\nA: {answer[:100]}...\n")  # Preview first 100 chars
        return answer
    except Exception as err:
        print(f"Error generating answer for '{question}': {str(err)}")
        return 'Error generating answer.'

async def generate_and_save_answers(questions_json, answers_json):
    """Generate answers for questions and save to JSON file"""
    print(f"Loading questions from {questions_json}")
    
    with open(questions_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    questions = data.get("questions", [])
    context = data.get("context", "")
    
    if questions:
        print(f"Generating answers for {len(questions)} questions...")
        
        answers = []
        for i, question in enumerate(questions):
            print(f"Processing question {i+1}/{len(questions)}")
            answer = await generate_detailed_answer(question, context)
            answers.append(answer)
        
        answers_data = {
            "qa_pairs": [{"question": q, "answer": a} for q, a in zip(questions, answers)],
            "question_indices": data.get("question_indices", []),
            "question_texts": data.get("question_texts", []),
            "question_formats": data.get("question_formats", [])
        }
        
        with open(answers_json, 'w', encoding='utf-8') as f:
            json.dump(answers_data, f, ensure_ascii=False, indent=2)
        
        print(f"Successfully saved {len(answers)} answers to {answers_json}")
        return True
    else:
        print("No questions found.")
        return False

async def update_answers(questions_json, answers_json):
    """Update answers only for new questions"""
    print(f"Loading questions from {questions_json}")
    
    with open(questions_json, 'r', encoding='utf-8') as f:
        questions_data = json.load(f)
    
    questions = questions_data.get("questions", [])
    context = questions_data.get("context", "")
    
    try:
        with open(answers_json, 'r', encoding='utf-8') as f:
            answers_data = json.load(f)
            qa_pairs = answers_data.get("qa_pairs", [])
    except (FileNotFoundError, json.JSONDecodeError):
        answers_data = {
            "qa_pairs": [],
            "question_indices": questions_data.get("question_indices", []),
            "question_texts": questions_data.get("question_texts", []),
            "question_formats": questions_data.get("question_formats", [])
        }
        qa_pairs = []
    
    existing_qa = {pair["question"].strip(): pair["answer"] for pair in qa_pairs}
    new_answers = 0
    
    for i, question in enumerate(questions):
        if question.strip() not in existing_qa:
            print(f"Processing new question {i+1}/{len(questions)}")
            answer = await generate_detailed_answer(question, context)
            qa_pairs.append({"question": question, "answer": answer})
            new_answers += 1
        else:
            print(f"Question {i+1} already answered.")
    
    answers_data["qa_pairs"] = qa_pairs
    answers_data["question_indices"] = questions_data.get("question_indices", [])
    answers_data["question_texts"] = questions_data.get("question_texts", [])
    answers_data["question_formats"] = questions_data.get("question_formats", [])
    
    with open(answers_json, 'w', encoding='utf-8') as f:
        json.dump(answers_data, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {new_answers} new answers. Total: {len(qa_pairs)}")
    return True

if __name__ == "__main__":
    questions_json = "questions_data.json"
    answers_json = "answers_data.json"
    asyncio.run(update_answers(questions_json, answers_json))