📊 Python Automation for MIS Reporting
📌 Project Overview

This project demonstrates how Python can be used to automate a repetitive MIS reporting task.
The script reads a raw Excel workbook containing multiple worksheets and automatically separates each worksheet into individual Excel files.
This automation helps reduce manual effort and speeds up the reporting process.

🚩 Problem

In my daily workflow, I download raw data from an internal reporting tool.
The downloaded Excel file contains multiple worksheets with different report names.
For reporting purposes, each worksheet must be saved as a separate Excel file before analysis.
Previously, this process involved:
Opening the Excel workbook
Copying each worksheet manually
Saving each sheet as a new Excel file
Renaming the files correctly
This manual task took around 2 hours every day.

⚙️ Solution

I developed a Python script that automates the entire process.
The script performs the following steps:
Reads the Excel workbook.
Identifies all worksheets in the file.
Extracts each worksheet automatically.
Saves each worksheet as an individual Excel file.
Stores the files in the required folder with the correct report name.

⏱ Impact

After implementing this automation:
Manual work reduced from ~2 hours to ~10 minutes
Improved efficiency in daily reporting workflow
Reduced chances of manual errors
Saved approximately 1 hour 50 minutes per day

🛠 Tools Used

Python
pandas
openpyxl

📂 Output

The script generates:
Separate Excel files for each worksheet
Files saved automatically in the specified folder path

📈 Learning Outcome

This project helped me understand:
Python automation for reporting workflows
Excel data processing using pandas
Improving operational efficiency using code
