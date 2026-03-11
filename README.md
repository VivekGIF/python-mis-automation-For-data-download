As part of improving my data analytics skills, I recently worked on automating one of my daily MIS reporting tasks using Python.

In my workflow, I download raw data from an internal tool. The file contains multiple worksheets with different report names. As per the reporting process, each sheet needs to be separated and saved as an individual Excel file before starting the analysis.

Earlier, this was a completely manual process. I had to open the workbook, copy each worksheet, and save them separately with the correct report name. Since the file contains many sheets, this task alone used to take nearly 2 hours every day.

To solve this, I wrote a Python script that:
• Reads the Excel workbook
• Extracts all worksheets automatically
• Saves each worksheet as a separate Excel file in the required folder
• Keeps the correct report name for each file

With this automation, the entire process now takes around 10 minutes instead of 2 hours.

My goal was to reduce repetitive manual work and make the reporting process faster and more efficient.
