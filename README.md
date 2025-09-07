# MoodleAssignmentUploaderAgent
A Python GUI application that automates uploading student assignments to Moodle.


Features
Simple GUI interface

Batch upload PDF assignments to Moodle

Auto student name matching from filenames

Progress tracking and detailed reports

Save configuration for future use

Quick Install
For End Users:
Download the latest release for your OS

Extract the ZIP file

Run MoodleUploader.exe (Windows) or MoodleUploader (Linux/Mac)

No Python installation required

For Developers:
bash
git clone https://github.com/Nyakudziguma/MoodleAssignmentUploaderAgent
pip install -r requirements.txt
python moodle_uploader.py
Usage
Enter your Moodle URL, username, and password

Provide Course ID and Assignment ID (from Moodle URLs)

Select folder containing student PDF files

Click "Start Upload"

Monitor progress in the status window

File Naming
Name PDF files as: Firstname Lastname.pdf

Example: John_Doe.pdf, Mary_Smith.pdf

Requirements
Google Chrome browser

Moodle account with grading permissions

PDF files named with student names

Troubleshooting
Ensure Chrome is installed

Check student names match Moodle records

Verify login credentials and course permissions

Building from Source
bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "MoodleUploader" --add-data "config.json;." moodle_uploader.py