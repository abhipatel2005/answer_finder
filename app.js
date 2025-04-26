const fs = require('fs');
const pdfParse = require('pdf-parse');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
// const { HfInference } = require('@huggingface/inference');
const axios = require('axios');
const dotenv = require('dotenv');
dotenv.config();

// Initialize Hugging Face Inference API
// const hf = new HfInference(process.env.HF_TOKEN);

// Read and Parse PDF
async function extractLabManual(path) {
    const dataBuffer = fs.readFileSync(path);
    const data = await pdfParse(dataBuffer);
    return data.text;
}

// Process quiz questions using Hugging Face QA Model with detailed answers
async function generateDetailedAnswer(question, context) {
    const apiKey = process.env.GEMINI_API_KEY;

    const payload = {
        contents: [
            {
                parts: [
                    {
                        text: `You are an expert tutor writing content for a lab manual. 
    Please generate a clear, concise, and well-structured answer in plain text (no bullet points, no line breaks, no markdown) to the following question based on the lab manual content.
    
    Ensure the answer reads like a textbook explanation or model written response.
    
    Question: ${question}
    
    Context: ${context.slice(0, 10000)}`
                    }
                ]
            }
        ]
    };

    try {
        const res = await axios.post(
            'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey,
            payload,
            {
                headers: {
                    'Content-Type': 'application/json'
                }
            }
        );

        const answer = res.data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || 'No answer generated.';
        console.log(`\nQ: ${question}\nA: ${answer}\n`);
        return answer;
    } catch (err) {
        console.error(`Error generating answer for "${question}":`, err.message);
        return 'Error generating answer.';
    }
}

// Create a new PDF with detailed answers inserted below questions
async function generateBlankPDFWithQnA(questions, answers, outputPath) {
    const pdfDoc = await PDFDocument.create();
    let page = pdfDoc.addPage();
    const { width, height } = page.getSize();
    const fontSize = 12;
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const margin = 50;
    const lineHeight = fontSize * 1.5;
    let yPosition = height - margin;

    for (let i = 0; i < questions.length; i++) {
        const qLines = wrapText(`Q${i + 1}: ${questions[i].replace(/\n/g, ' ')}`, font, fontSize, width - margin * 2);
        const aLines = wrapText(`A${i + 1}: ${answers[i].replace(/\n/g, ' ')}`, font, fontSize, width - margin * 2);

        for (const line of qLines) {
            if (yPosition < margin) {
                page = pdfDoc.addPage();
                yPosition = height - margin;
            }
            page.drawText(line, {
                x: margin,
                y: yPosition,
                size: fontSize,
                font,
                color: rgb(0, 0, 0)
            });
            yPosition -= lineHeight;
        }

        yPosition -= lineHeight / 2;

        for (const line of aLines) {
            if (yPosition < margin) {
                page = pdfDoc.addPage();
                yPosition = height - margin;
            }
            page.drawText(line, {
                x: margin,
                y: yPosition,
                size: fontSize,
                font,
                color: rgb(0, 0, 0)
            });
            yPosition -= lineHeight;
        }

        yPosition -= lineHeight * 1.5; // space between QnAs
    }

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
    console.log("Blank PDF with Questions and Answers generated successfully!");
}

// Helper to wrap text
function wrapText(text, font, fontSize, maxWidth) {
    const words = text.split(' ');
    const lines = [];
    let line = '';
    for (let word of words) {
        const testLine = line + word + ' ';
        const testWidth = font.widthOfTextAtSize(testLine, fontSize);
        if (testWidth <= maxWidth) {
            line = testLine;
        } else {
            lines.push(line.trim());
            line = word + ' ';
        }
    }
    if (line.trim()) lines.push(line.trim());
    return lines;
}


// Improved Question Extraction Regex
function extractQuestionsFromQuizSection(quizSection) {
    // Regex to capture numbered questions followed by any text, even if it doesn't end in a '?' or '.'.
    const questionRegex = /\d+\.\s*(.*?)(?=\d+\.\s|$)/gs;
    let questions = [];
    let match;
    while ((match = questionRegex.exec(quizSection)) !== null) {
        questions.push(match[1].trim());
    }
    // console.log(questions.length)
    return questions;
}

// Main Function
(async () => {
    const manualText = await extractLabManual("E:\\ClgStu\\Sem 5\\Computer Networks(3150710)\\3150710-Computer-Network Lab manual.pdf");

    // Extract all occurrences of quiz sections using multiple start and end possibilities
    const quizSections = [...manualText.matchAll(/(Quiz\(.*?\)|Quiz: \(Sufficient space to be provided for the answers\))(.*?)(References used by the students:?|Suggested Reference:?)/gs)];

    if (quizSections.length > 0) {
        let allQuestions = [];

        // Iterate through all quiz sections
        for (const section of quizSections) {
            const quizSection = section[2].trim();  // Extracting the actual content between 'Quiz' and the reference section
            console.log("Extracted Quiz Section: ", quizSection);

            // Extract questions using improved regex
            const questions = extractQuestionsFromQuizSection(quizSection);
            console.log("Extracted Questions: ", questions);

            allQuestions = allQuestions.concat(questions);  // Collect all questions from each section
        }

        if (allQuestions.length === 0) {
            console.error("No questions found in the quiz sections.");
            return;
        }

        // Answer questions using Hugging Face with detailed knowledge
        const answers = [];
        for (const question of allQuestions) {
            const answer = await generateDetailedAnswer(question, manualText);
            answers.push(answer);
        }
        console.log("Generated Detailed Answers: ", answers);

        // Generate a new PDF with detailed answers below the questions
        await generateBlankPDFWithQnA(allQuestions, answers, 'answered.pdf');
        console.log('Executed Successfully...!')
    } else {
        console.error("Quiz sections not found.");
        console.log('Code Executed...!, check the problems');
    }
})();
