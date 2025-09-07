# 📚 MoodleAssignmentUploaderAgent

A Python GUI application that automates uploading student assignments to **Moodle**.

## ✨ Features
- 🖥️ Simple, user-friendly GUI interface
- 📂 Batch upload **PDF assignments** directly to Moodle
- 🔍 Automatic student name matching from filenames
- 📊 Real-time progress tracking and detailed reports
- 💾 Save & reuse configuration for future uploads

---

## 🚀 Quick Install

### 🔹 For End Users
1. Download the latest release for your OS (Windows, macOS, or Linux).
2. Extract the **ZIP** file.
3. Run the appropriate executable:
   - `MoodleUploader.exe` (Windows)
   - `MoodleUploader` (Linux/macOS)

> No Python installation required for users who download the packaged release.


### 🔹 For Developers
```bash
git clone https://github.com/Nyakudziguma/MoodleAssignmentUploaderAgent
cd MoodleAssignmentUploaderAgent
pip install -r requirements.txt
python moodle_uploader.py
```

---

## 🛠️ Usage
1. Start the application.
2. Enter your **Moodle URL**, **username**, and **password**.
3. Provide the **Course ID** and **Assignment ID** (these appear in Moodle assignment URLs).
4. Select the folder containing the student PDF files.
5. Click **Start Upload**.
6. Monitor upload progress and view the final report in the status window.

---

## 📑 File Naming Convention

PDF files must be named using the student's name so the app can match students automatically. Use one of these formats:

```
Firstname_Lastname.pdf
Firstname Lastname.pdf
```

**Examples:** `John_Doe.pdf`, `Mary Smith.pdf`

> Filenames should match the names registered in Moodle to ensure correct student-to-file mapping.

---

## 📋 Requirements
- Google Chrome browser installed (used by the automation).
- Moodle account with grading/upload permissions for the target course and assignment.
- PDF files named with student names (see File Naming Convention).

---

## 🛠️ Troubleshooting
- Make sure **Google Chrome** is installed and up to date.
- Verify your Moodle **username** and **password** are correct and have permission to upload files/grade.
- Confirm student names in PDFs match Moodle rosters exactly (or adjust the matching rules in code).
- If uploads fail, check logs and the generated report for which files failed and why.

---

## 🏗️ Building from Source (Standalone Executable)

To build a single-file executable using PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed \
  --name "MoodleUploader" \
  --add-data "config.json;." moodle_uploader.py
```

**Notes:**
- On Windows, use `--add-data "config.json;."` (semicolon separator). On macOS/Linux, replace with `--add-data "config.json:."` (colon separator) if necessary.

---

## 🔐 Security & Privacy
- Credentials are used to log in to Moodle and should be protected.
- If you enable "save configuration", credentials are stored in `config.json` by default — consider encrypting or securing that file.
- Do not commit sensitive credentials to version control.

---

## ✅ Reporting & Logs
- After a batch run, the application generates a report listing:
  - Successfully uploaded files
  - Failed uploads and error details
  - Any filename-to-student mismatches
- Check the `logs/` folder (or the configured log destination) for detailed error traces.

---


