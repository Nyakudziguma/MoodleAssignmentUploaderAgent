import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import time
from datetime import datetime
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class MoodleUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Moodle Assignment Uploader")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Configuration variables
        self.moodle_url = tk.StringVar(value="https://accs.credspace.co.zw")
        self.username = tk.StringVar(value="")
        self.password = tk.StringVar(value="")
        self.course_id = tk.StringVar(value="")
        self.assignment_id = tk.StringVar(value="")
        self.folder_path = tk.StringVar(value="")
        
        # Status variables
        self.is_running = False
        self.driver = None
        
        self.setup_ui()
        self.load_config()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Moodle Assignment Uploader", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Configuration section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # Moodle URL
        ttk.Label(config_frame, text="Moodle URL:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.moodle_url, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="we")
        
        # Username
        ttk.Label(config_frame, text="Username:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.username, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="we")
        
        # Password
        ttk.Label(config_frame, text="Password:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.password, show="*", width=50).grid(row=2, column=1, padx=5, pady=5, sticky="we")
        
        # Course ID
        ttk.Label(config_frame, text="Course ID:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.course_id, width=50).grid(row=3, column=1, padx=5, pady=5, sticky="we")
        
        # Assignment ID
        ttk.Label(config_frame, text="Assignment ID:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.assignment_id, width=50).grid(row=4, column=1, padx=5, pady=5, sticky="we")
        
        # Folder path
        ttk.Label(config_frame, text="Files Folder:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(config_frame, textvariable=self.folder_path, width=40).grid(row=5, column=1, padx=5, pady=5, sticky="we")
        ttk.Button(config_frame, text="Browse", command=self.browse_folder).grid(row=5, column=2, padx=5, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Save Configuration", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Start Upload", command=self.start_upload).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Stop", command=self.stop_upload).pack(side=tk.LEFT, padx=5)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Status text
        self.status_text = scrolledtext.ScrolledText(progress_frame, height=15, width=80)
        self.status_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.status_text.config(state=tk.DISABLED)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)
        
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
    
    def save_config(self):
        config = {
            "moodle_url": self.moodle_url.get(),
            "username": self.username.get(),
            "course_id": self.course_id.get(),
            "assignment_id": self.assignment_id.get(),
            "folder_path": self.folder_path.get()
        }
        
        try:
            with open("config.json", "w") as f:
                json.dump(config, f)
            self.log_message("Configuration saved successfully!")
            messagebox.showinfo("Success", "Configuration saved!")
        except Exception as e:
            self.log_message(f"Error saving configuration: {str(e)}")
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    def load_config(self):
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                self.moodle_url.set(config.get("moodle_url", ""))
                self.username.set(config.get("username", ""))
                self.course_id.set(config.get("course_id", ""))
                self.assignment_id.set(config.get("assignment_id", ""))
                self.folder_path.set(config.get("folder_path", ""))
            self.log_message("Configuration loaded successfully!")
        except FileNotFoundError:
            self.log_message("No saved configuration found.")
        except Exception as e:
            self.log_message(f"Error loading configuration: {str(e)}")
    
    def log_message(self, message):
        self.status_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def start_upload(self):
        if self.is_running:
            self.log_message("Upload already in progress!")
            return
            
        # Validate inputs
        if not all([self.moodle_url.get(), self.username.get(), self.password.get(), 
                   self.course_id.get(), self.assignment_id.get(), self.folder_path.get()]):
            messagebox.showerror("Error", "Please fill all fields!")
            return
        
        if not os.path.exists(self.folder_path.get()):
            messagebox.showerror("Error", "Selected folder does not exist!")
            return
        
        # Start the upload process in a separate thread
        self.is_running = True
        self.progress["value"] = 0
        threading.Thread(target=self.run_upload, daemon=True).start()
    
    def stop_upload(self):
        self.is_running = False
        self.log_message("Stopping upload process...")
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
    
    def safe_click(self, element):
        """Click an element using JavaScript to avoid click interception"""
        self.driver.execute_script("arguments[0].click();", element)

    def find_student_row(self, student_name):
        """Case-insensitive search for student name in the table"""
        search_name = student_name.lower()
        
        # Find all student name links
        student_links = self.driver.find_elements(By.CSS_SELECTOR, "td.username a")
        
        for link in student_links:
            if search_name in link.text.lower():
                return link.find_element(By.XPATH, "./ancestor::tr")
        
        raise Exception(f"Student '{student_name}' not found in grading table")

    def wait_for_grading_table(self):
        """Wait for the grading table to be fully loaded and ready"""
        self.log_message("  Waiting for grading table to load...")
        WebDriverWait(self.driver, 25).until(EC.presence_of_element_located((By.ID, "submissions")))
        WebDriverWait(self.driver, 25).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#submissions tbody tr")))
        self.log_message("  Grading table ready")

    def upload_file_to_filepicker(self, full_file_path, filename):
        """Handle the file upload process with better error handling"""
        self.log_message("  On edit page, looking for file input...")
        
        # Wait for the page to fully load
        WebDriverWait(self.driver, 25).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".fp-btn-add")))
        time.sleep(2)  # Longer wait for page stability
        
        # Click the "Add" button to open the file picker dialog
        add_button = WebDriverWait(self.driver, 25).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".fp-btn-add"))
        )
        self.safe_click(add_button)
        self.log_message("  Opened file picker dialog")
        
        # Wait for the file picker dialog to be fully visible
        WebDriverWait(self.driver, 25).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".file-picker")))
        time.sleep(1)  # Wait for dialog to stabilize
        
        # Click the "Upload a file" tab
        upload_tab = WebDriverWait(self.driver, 25).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Upload a file')]"))
        )
        self.safe_click(upload_tab)
        self.log_message("  Switched to 'Upload a file' tab")
        
        # Wait for the tab to activate
        time.sleep(1)
        
        # FIND FILE INPUT WITH MULTIPLE STRATEGIES
        file_input = None
        file_input_selectors = [
            "input[type='file'][name='repo_upload_file']",
            "input[type='file']",
            "#repo_upload_file",
            ".file-picker input[type='file']"
        ]
        
        for selector in file_input_selectors:
            try:
                file_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                self.log_message(f"  Found file input using: {selector}")
                
                # Ensure the input is visible and enabled
                if file_input.is_displayed() and file_input.is_enabled():
                    break
                else:
                    file_input = None
            except:
                continue
        
        if not file_input:
            raise Exception("Could not find usable file input element")
        
        # UPLOAD FILE WITH BETTER HANDLING
        self.log_message(f"  Uploading file: {filename}")
        
        # Clear any existing value first (just in case)
        try:
            file_input.clear()
        except:
            pass
        
        # Send the file path to the input
        absolute_path = os.path.abspath(full_file_path)
        self.log_message(f"  Using absolute path: {absolute_path}")
        
        file_input.send_keys(absolute_path)
        self.log_message("  File path sent to input")
        
        # WAIT FOR FILE TO BE PROCESSED BEFORE CLICKING UPLOAD
        self.log_message("  Waiting for file to be processed by the dialog...")
        time.sleep(2)  # Crucial wait for the file to be recognized
        
        # CLICK THE "UPLOAD THIS FILE" BUTTON
        self.log_message("  Looking for upload button...")
        
        upload_button = None
        upload_button_selectors = [
            "button.fp-upload-btn",
            ".fp-upload-btn",
            "button[type='submit']",
            "input[type='submit'][value*='Upload']"
        ]
        
        for selector in upload_button_selectors:
            try:
                upload_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                self.log_message(f"  Found upload button using: {selector}")
                break
            except:
                continue
        
        if not upload_button:
            raise Exception("Could not find upload button")
        
        # Scroll the button into view and click it
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", upload_button)
        time.sleep(0.5)
        
        self.log_message("  Clicking 'Upload this file'...")
        self.safe_click(upload_button)
        
        # WAIT FOR FILE TO BE UPLOADED
        self.log_message("  Waiting for file upload to complete...")
        
        # Wait for the file picker dialog to close
        try:
            WebDriverWait(self.driver, 10).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, ".file-picker"))
            )
            self.log_message("  File picker dialog closed")
        except:
            self.log_message("  ⚠ File picker dialog did not close as expected")
        
        # Additional wait for file processing
        time.sleep(3)
        
        return True

    def run_upload(self):
        try:
            # Setup WebDriver with webdriver-manager
            self.log_message("Setting up Chrome WebDriver...")
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(service=service, options=options)
            
            # 1. LOGIN
            self.log_message("Logging in...")
            self.driver.get(f"{self.moodle_url.get()}/login/index.php")
            username_field = WebDriverWait(self.driver, 25).until(
                EC.element_to_be_clickable((By.ID, "username"))
            )
            username_field.send_keys(self.username.get())
            password_field = self.driver.find_element(By.ID, "password")
            password_field.send_keys(self.password.get())
            login_button = self.driver.find_element(By.ID, "loginbtn")
            self.safe_click(login_button)
            WebDriverWait(self.driver, 25).until(EC.url_contains("my/"))
            self.log_message("Login successful.")

            # 2. NAVIGATE TO THE ASSIGNMENT GRADING PAGE
            grading_url = f"{self.moodle_url.get()}/mod/assign/view.php?id={self.assignment_id.get()}&action=grading"
            self.log_message(f"Navigating to grading page: {grading_url}")
            self.driver.get(grading_url)
            self.wait_for_grading_table()

            # 3. GET LIST OF FILES
            files = [f for f in os.listdir(self.folder_path.get()) if f.lower().endswith('.pdf')]
            total_files = len(files)
            
            if total_files == 0:
                self.log_message("No PDF files found in the specified folder!")
                return
            
            successful_uploads = []
            failed_uploads = []
            
            # 4. PROCESS EACH FILE
            for i, filename in enumerate(files):
                if not self.is_running:
                    self.log_message("Upload process stopped by user")
                    break
                    
                # Update progress
                progress = (i / total_files) * 100
                self.progress["value"] = progress
                self.root.update_idletasks()
                
                # Extract student name from filename
                student_name_from_file = os.path.splitext(filename)[0].replace('_', ' ')
                full_file_path = os.path.join(self.folder_path.get(), filename)
                self.log_message(f"\n=== Processing student: {student_name_from_file} ({i+1}/{total_files}) ===")

                try:
                    # 5. FIND THE STUDENT'S ROW (CASE-INSENSITIVE)
                    self.log_message("  Finding student row...")
                    student_row = self.find_student_row(student_name_from_file)
                    self.log_message(f"  Found row for {student_name_from_file}")

                    # 6. FIND AND CLICK THE ACTION MENU IN THE STATUS COLUMN
                    status_cell = student_row.find_element(By.CLASS_NAME, "status")
                    action_menu_button = status_cell.find_element(
                        By.CSS_SELECTOR, "a[data-bs-toggle='dropdown'][aria-haspopup='true']"
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", action_menu_button)
                    time.sleep(1)
                    self.log_message("  Clicking action menu button...")
                    self.safe_click(action_menu_button)
                    
                    # 7. WAIT FOR DROPDOWN MENU TO BE VISIBLE AND CLICK "EDIT SUBMISSION"
                    self.log_message("  Waiting for dropdown menu to appear...")
                    
                    dropdown_menu = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.show"))
                    )
                    
                    edit_submission_link = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//a[.//span[contains(text(), 'Edit submission')]]"))
                    )
                    
                    self.log_message("  Clicking 'Edit submission'...")
                    self.safe_click(edit_submission_link)
                    self.log_message("  Navigated to edit submission page")

                    # 8. UPLOAD THE FILE WITH IMPROVED HANDLING
                    upload_success = self.upload_file_to_filepicker(full_file_path, filename)
                    
                    if not upload_success:
                        raise Exception("File upload failed or file not found in manager after upload")
                    
                    # 9. SAVE CHANGES
                    self.log_message("  Saving changes...")
                    
                    save_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value*='Save'], button[type='submit']"))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                    time.sleep(1)
                    
                    self.log_message("  Clicking save button...")
                    save_button.click()

                    # 10. WAIT FOR CONFIRMATION
                    self.log_message("  Waiting for save confirmation...")
                    time.sleep(3)
                    
                    success_indicators = [".alert-success", ".notifysuccess", ".success"]
                    success_found = any(self.driver.find_elements(By.CSS_SELECTOR, indicator) for indicator in success_indicators)
                    
                    if success_found or "view.php" in self.driver.current_url:
                        self.log_message(f"  ✅ Successfully uploaded for {student_name_from_file}")
                        successful_uploads.append(filename)
                    else:
                        raise Exception("Save confirmation not found")

                    # 11. GO BACK TO GRADING LIST
                    self.log_message("  Returning to grading list...")
                    self.driver.get(grading_url)
                    self.wait_for_grading_table()
                    self.log_message("  Returned to grading list")

                except Exception as e:
                    error_msg = f"ERROR: {filename} - {student_name_from_file} - {str(e)}"
                    self.log_message(f"  ❌ {error_msg}")
                    failed_uploads.append((filename, student_name_from_file, str(e)))
                    
                    try:
                        self.driver.get(grading_url)
                        self.wait_for_grading_table()
                    except:
                        try:
                            self.driver.refresh()
                            self.wait_for_grading_table()
                        except:
                            pass

                time.sleep(3)

            # FINAL REPORT
            self.progress["value"] = 100
            self.log_message("\n" + "="*60)
            self.log_message("BATCH UPLOAD COMPLETE - SUMMARY REPORT")
            self.log_message("="*60)
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.log_message(f"\nTimestamp: {timestamp}")
            self.log_message(f"Total files processed: {total_files}")
            self.log_message(f"Successful uploads: {len(successful_uploads)}")
            self.log_message(f"Failed uploads: {len(failed_uploads)}")
            
            if len(successful_uploads) > 0:
                self.log_message(f"Success rate: {len(successful_uploads)/total_files*100:.1f}%")
            
            if successful_uploads:
                self.log_message(f"\n✅ SUCCESSFUL UPLOADS ({len(successful_uploads)}):")
                self.log_message("-" * 50)
                for i, filename in enumerate(successful_uploads, 1):
                    student_name = os.path.splitext(filename)[0].replace('_', ' ')
                    self.log_message(f"{i:2d}. {student_name} ({filename})")
            
            if failed_uploads:
                self.log_message(f"\n❌ FAILED UPLOADS ({len(failed_uploads)}):")
                self.log_message("-" * 50)
                for i, (filename, student_name, error) in enumerate(failed_uploads, 1):
                    self.log_message(f"{i:2d}. {student_name} ({filename})")
                    self.log_message(f"    Error: {error[:80]}{'...' if len(error) > 80 else ''}")

        except Exception as e:
            self.log_message(f"Fatal error: {str(e)}")
            messagebox.showerror("Error", f"A fatal error occurred: {str(e)}")
        finally:
            self.is_running = False
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

def main():
    root = tk.Tk()
    app = MoodleUploaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()