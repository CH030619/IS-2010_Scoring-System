# IS 2010 Automated Scoring System with AI Feedback

> **Smart, Scalable, and Precise Grading Engine for Microsoft Excel Assignments.**

---

This project is an automated grading solution designed for the IS 2010 course at the University of Utah. 
It goes beyond simple value matching by analyzing Excel formulas, cell properties, and hidden XML structures (for Sparklines), providing students with personalized, AI-generated feedback through OpenAI's GPT-4o-mini.

---

## Core Objectives
* **Automated Precision**: Compare student submissions against a master key with 1:1 formula and value matching.
* **Intelligent Feedback**: Leverage OpenAI API to analyze errors and provide pedagogical guidance.
* **Professional Reporting**: Generate comprehensive PDF diagnostic reports for each student.

---

## Key Technical Features

### Performance Optimization (Read-only & Caching)
* **Read-only Mode**: Utilizes `openpyxl`'s `read_only=True` to minimize memory footprint and maximize loading speed for large batches.


### Decoupled AI Pipeline
The AI feedback generation is decoupled from the main grading loop. This "Generate on Demand" architecture ensures system stability and optimizes API token costs by only processing errors during report generation.

---

## Getting Started
### 1. Installation
```bash
pip install streamlit pandas openpyxl openai fpdf lxml
```
### 2. Configuration (API Credentials)
```bash
* **Create a .streamlit/secrets.toml file in the root directory:
OPENAI_API_KEY = "your_openai_api_key_here"
```
### 3. Run the Application
```bash
streamlit run guided_lab_2.py
```
### Guided Lab Guidelines & Cautions
```bash
To ensure accuracy, students must strictly follow these rules:

* File Naming: Filenames must include a valid UNID (e.g., u1234567).

* Sheet Integrity: Do not rename sheets or insert/delete rows/columns. The system relies on fixed cell coordinates.

* Formula vs. Value: Manual inputs (typing values) where a formula is required will be marked as Incorrect.

* Formatting: Only .xlsx files are supported. No macros (.xlsm) or password protection allowed.
```
### 🛠 Tech Stack
```bash
* Engine: Python 3.x

* Framework: Streamlit

* Excel Processing: Openpyxl, Pandas

* AI Logic: OpenAI GPT-4o-mini

* PDF Engine: FPDF

* Parsing: Lxml (XML structures)

