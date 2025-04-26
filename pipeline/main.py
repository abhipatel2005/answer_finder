# main.py
import os
import asyncio
import argparse
from question_extractor import extract_and_save_questions
from answer_generator import update_answers
from document_writer import write_answers_to_document

async def process_lab_manual(input_file, output_file, questions_json=None, answers_json=None):
    """Process a lab manual document end-to-end"""
    # Set default filenames if not provided
    if questions_json is None:
        questions_json = "questions_data.json"
    if answers_json is None:
        answers_json = "answers_data.json"
    
    print("=== Lab Manual Processing System ===")
    print(f"Input document: {input_file}")
    print(f"Output document: {output_file}")
    print(f"Questions data: {questions_json}")
    print(f"Answers data: {answers_json}")
    print("==================================")
    
    # Step 1: Extract questions from the document
    print("\nStep 1: Extracting questions...")
    success = extract_and_save_questions(input_file, questions_json)
    if not success:
        print("Failed to extract questions. Aborting.")
        return False
    
    # Step 2: Generate answers (only for new questions)
    print("\nStep 2: Generating answers...")
    success = await update_answers(questions_json, answers_json)
    if not success:
        print("Failed to generate answers. Aborting.")
        return False
    
    # Step 3: Write answers to the document
    print("\nStep 3: Writing answers to document...")
    success = write_answers_to_document(input_file, answers_json, output_file)
    if not success:
        print("Failed to write answers to document.")
        return False
        
    print("\n=== Processing complete! ===")
    print(f"Document with answers saved to: {output_file}")
    return True

def setup_argparse():
    """Set up command line arguments"""
    parser = argparse.ArgumentParser(description='Process lab manual to add answers to questions.')
    parser.add_argument('--input', '-i', required=True, help='Input docx file path')
    parser.add_argument('--output', '-o', required=True, help='Output docx file path')
    parser.add_argument('--questions', '-q', default='questions_data.json', help='JSON file to save/load questions data')
    parser.add_argument('--answers', '-a', default='answers_data.json', help='JSON file to save/load answers data')
    parser.add_argument('--extract-only', action='store_true', help='Only extract questions, don\'t generate answers')
    parser.add_argument('--generate-only', action='store_true', help='Only generate answers, don\'t modify document')
    parser.add_argument('--write-only', action='store_true', help='Only write answers to document, don\'t extract or generate')
    return parser

async def main():
    parser = setup_argparse()
    args = parser.parse_args()
    
    # Validate that input file exists
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        return False
    
    if args.extract_only:
        # Only extract questions
        print("Extracting questions only...")
        return extract_and_save_questions(args.input, args.questions)
        
    elif args.generate_only:
        # Only generate answers
        print("Generating answers only...")
        if not os.path.exists(args.questions):
            print(f"Error: Questions file '{args.questions}' not found.")
            return False
        return await update_answers(args.questions, args.answers)
        
    elif args.write_only:
        # Only write answers to document
        print("Writing answers to document only...")
        if not os.path.exists(args.answers):
            print(f"Error: Answers file '{args.answers}' not found.")
            return False
        return write_answers_to_document(args.input, args.answers, args.output)
        
    else:
        # Process everything
        return await process_lab_manual(args.input, args.output, args.questions, args.answers)

if __name__ == "__main__":
    asyncio.run(main())