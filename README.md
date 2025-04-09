# AttendanceClerk
Thou art welcome, to yet another one of my creations, born partly out of my refusal to do repetitive, unintuitive tasks, and partly from my love of bringing ideas to life through code.  
![welcome gif](welcome.gif)
## The Program
AttendanceClerk is a program that automates the attendance marking process. It was built for an institution that trains chartered accountants, there is an excel workbook that contains a list of all registered students for each course (e.g Advanced Financial Management), after holding a class on Teams the list of meeting participants is downloaded and then someone has to manually check and tick the name of each person that attended that class in the main workbook. Imagine having to do this for 20+ classes a week, week after week after week, the process is not only painstaking but also not very effective.

## How It Works
Going into the details is redundant as the code is already properly documented so i will instead outline high level steps of how it does what it does:
- Accepts two files as input: the main attendance workbook and the participant list from Teams
- Loads the workbook, preserving all existing data, taking note of everything exactly as is (This is important to preserve formulas and untouched sheets)
- Extracts relevant details from the Teams file
- Determines the appropriate worksheet to update and locates the date column based on the course and class date extracted from the Teams file
- Matches participants from the Teams list with names in the worksheet and marks their attendance
- Adds new entries for attendees not already listed in the workbook
- Applies updates to formulas, fonts, merged cells, and text alignment as needed
- Finalizes the process by saving and presenting the updated file
### Pros:
1. Significantly reduces manual effort, freeing up valuable time and giving you some part of your life back.
2. It's an effective solution to issues
3. It reduces risk of human errors
4. It requires neither wages nor salaries — a key benefit of automation. (This is the ultimate benefit of automation for money-focused companies)
5. Preserves Excel Integrity: The tool keeps formulas, merged cells, and existing formats intact — a crucial feature when working with templates already wired for downstream calculations.
6. Smart Matching: The name extraction and cleaning logic is resilient against common inconsistencies in user-submitted names — uppercase, prefixes, or different naming conventions.
7. Visual and Professional Output: Visual elements are retained or improved, making the final sheet more presentable — little need for post-run formatting touchups.
8. Lightweight Deployment: Since it runs via Streamlit or raw Python, there's no complicated setup required.
### Cons:
1. Performance Bottlenecks: These can occur due to openpyxl’s speed when reading/writing large Excel files (an operation that would typically take 0.3-0.7 seconds can take up to 20+ seconds as a result of this.)
2. This code is made particularly for files that are formatted and structured in a precise way, anything outside that might produce errors (That’s acceptable, however, since both the institution and Microsoft Teams generate files in a consistent format.)
3. No Built-in Preview or Undo: Once the file is processed, changes are saved — there’s no preview step or rollback.
4. Hardcoded Logic Assumptions: Some logic is tailored for a very specific use case.
5. No User Input Validation in UI: If someone uploads a wrong file (like a malformed .csv), there’s no inline validation.
6. Session Persistence Limits: If a user refreshes or navigates away, the session state may reset — temporary file storage or caching.
## 📦 Prerequisites
Make sure you have the following installed:
- [Python 3.10+](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/)
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Streamlit](https://streamlit.io/)
  
You can install them using:
```bash
pip install -r requirements.txt
```
## ▶️ How to Run
You can run the program in two different modes:

1. Streamlit Interface (Recommended)
    ```bash
    streamlit run main.py
    ```
    
2. Pure Python (No UI)
   ```bash
   python pure_code.py
   ```
   Ensure that both attendance.xlsx and class_participants.csv are placed in the working directory before running.
## ⚠️ Error Handling
If an error occurs during execution, the program will display a custom error page (4xx or 5xx) with a description of the issue and an option to contact the developer directly via email.

This ensures that issues can be reported and resolved promptly without interrupting your workflow.
## 👨‍💻 Author
Created with care by [Emmanuel](https://www.linkedin.com/in/ebi-emmanuel/) — fueled by frustration with manual work and a love for automation.

Feel free to reach out if you encounter issues or have suggestions for improvement.