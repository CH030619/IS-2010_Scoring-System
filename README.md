## 📊 IS 2010 Automated Scoring System [AI-Powered Precision Grading for Guided/Homework Lab]
* This system is a specialized grading engine developed for the IS 2010 course at the University of Utah.
* It ensures academic integrity and provides pedagogical value by analyzing formulas, values, and hidden XML structures (Sparklines) with high-speed automation.

### ✨ Why This System?
* Formula-Level Analysis: Beyond simple values, it verifies the logic behind the numbers (1:1 string comparison).

* AI-Driven Pedagogy: Uses OpenAI GPT-4o-mini to explain why a student was wrong, acting as a 24/7 TA.

* Dynamic Sparkline Support: Deep-parses XML structures to grade visual data trends that standard libraries often miss.

* One-Click Reporting: Automatically generates professional PDF diagnostic reports for the entire class.


### 🛠 Tech Stack & Architecture
```bash
* Framework: Streamlit

* Data Engines: Openpyxl, Pandas, Lxml

* AI Logic: OpenAI GPT-4o-mini

* PDF Engine: FPDF
```

### 📋 Professor's Quick Start Guide
```bash
1. Environment Setup
Ensure you have set your environment

2. Install required libraries:
pip install streamlit openpyxl pandas openai fpdf lxml zipfile math

3. API Configuration
Create a .streamlit/secrets.toml file in your root directory:
Add your API key as follows:
OPENAI_API_KEY = "your api key here"

4. Launching the App
Enter the following in your terminal:
streamlit run "Scoring_System"
```

### ⚠️ Essential Guidelines for Success
To maintain 100% grading accuracy, please ensure students are briefed on these rules:

### 🎨 The "10 Standard Colors" Rule
* The engine utilizes openpyxl to detect answer keys via Cell Shading.

* Requirement: Instructors must use one of the 10 Standard Colors (Standard Palette) for answer keys.

* Note: Custom RGB/Theme colors or font-only changes will not be recognized by the engine.

### 📝 Student Submission Rules
* File Naming: Must include a valid UNID (e.g., u1234567_Lab2.xlsx).

* Structure: Students must not rename sheets or shift cell coordinates.

* Logic: Manual inputs (typing 0.33) will be marked incorrect if the key requires a formula (=1/3).

* Format: Only .xlsx is supported (No Macros/Password protection).

### 🧪 Quick Test (Demo)
1) Download the testing folders (Guided_Lab_2 & Guided_Lab_3)

2) Upload professor's file (xlsx) in the Professor's File slot.

3) Upload student's file (u0000000_lab_x).xlsx in the Student's File slot.

4) Select the Fill Color you used for the answer key.

5) Click Start Grading and review the AI-generated PDF reports.
