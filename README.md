# PaperCraft Pro — Advanced Exam Paper Generator

## Main file
`app.py`

## Deploy to Streamlit Community Cloud
- Branch: `main`
- Main file path: `app.py`

## Features
- Professional iLovePDF-inspired workspace
- Multi-page dashboard
- PDF, DOCX and TXT material upload
- Knowledge-base extraction
- Multiple Choice questions
- True/False questions
- Short-answer questions
- Long-answer questions
- Difficulty controls
- Bloom's Taxonomy selection
- Question editing and marks
- Paper designer
- School/university branding
- PDF export
- Word export
- PDF answer key
- Word answer key

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Important
This version is a strong local/RAG-inspired generator using uploaded material as the source. For true generative AI, semantic vector search, citations, and higher-quality question generation, connect an LLM provider such as OpenAI/Gemini/Claude plus a vector database.
