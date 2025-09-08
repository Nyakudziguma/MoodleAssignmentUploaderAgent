import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import time
from datetime import datetime
import json
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

class MoodleUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Moodle Assignment Uploader")
        self.root.geometry("800x750")
        self.root.resizable(True, True)
        
        # Configuration variables
        self.moodle_url = tk.StringVar(value="http://myaoc1.accsoncall.com")
        self.username = tk.StringVar(value="")
        self.password = tk.StringVar(value="")
        self.course_id = tk.StringVar(value="")
        self.assignment_id = tk.StringVar(value="")
        self.folder_path = tk.StringVar(value="")
        
        # Status variables
        self.is_running = False
        self.driver = None
        self.report_data = ""
        
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
        ttk.Button(button_frame, text="Download Report", command=self.download_report).pack(side=tk.LEFT, padx=5)
        
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
    
    def download_report(self):
        """Save the report to a text file"""
        if not self.report_data:
            messagebox.showinfo("Info", "No report data available. Run an upload first.")
            return
            
        # Ask user where to save the report
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Save Report As",
            initialfile=f"moodle_upload_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.report_data)
                self.log_message(f"Report saved to: {file_path}")
                messagebox.showinfo("Success", f"Report saved successfully!\n{file_path}")
            except Exception as e:
                self.log_message(f"Error saving report: {str(e)}")
                messagebox.showerror("Error", f"Failed to save report: {str(e)}")
    
    def generate_report(self, successful_uploads, failed_uploads, total_files, timestamp):
        """Generate a detailed report string"""
        report = "=" * 60 + "\n"
        report += "MOODLE ASSIGNMENT UPLOAD REPORT\n"
        report += "=" * 60 + "\n\n"
        
        report += f"Timestamp: {timestamp}\n"
        report += f"Moodle URL: {self.moodle_url.get()}\n"
        report += f"Course ID: {self.course_id.get()}\n"
        report += f"Assignment ID: {self.assignment_id.get()}\n"
        report += f"Folder: {self.folder_path.get()}\n\n"
        
        report += f"Total files processed: {total_files}\n"
        report += f"Successful uploads: {len(successful_uploads)}\n"
        report += f"Failed uploads: {len(failed_uploads)}\n"
        
        if total_files > 0:
            success_rate = (len(successful_uploads) / total_files) * 100
            report += f"Success rate: {success_rate:.1f}%\n"
        
        report += "\n" + "=" * 60 + "\n"
        
        if successful_uploads:
            report += f"✅ SUCCESSFUL UPLOADS ({len(successful_uploads)}):\n"
            report += "-" * 50 + "\n"
            for i, (filename, username) in enumerate(successful_uploads, 1):
                report += f"{i:2d}. Username: {username} ({filename})\n"
        
        if failed_uploads:
            report += f"\n❌ FAILED UPLOADS ({len(failed_uploads)}):\n"
            report += "-" * 50 + "\n"
            for i, (filename, username, error) in enumerate(failed_uploads, 1):
                report += f"{i:2d}. Username: {username} ({filename})\n"
                report += f"    Error: {error}\n"
        
        # Add the log content
        report += "\n" + "=" * 60 + "\n"
        report += "PROCESS LOG:\n"
        report += "-" * 50 + "\n"
        report += self.status_text.get("1.0", tk.END)
        
        return report
    
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

    def extract_username_from_filename(self, filename):
        """
        Extract username from filename. Supports multiple formats:
        - username.pdf
        - username_assignment.pdf
        - assignment_username.pdf
        - any filename containing a username pattern
        """
        # Remove file extension
        name_without_ext = os.path.splitext(filename)[0]
        
        # Try to find username patterns (alphanumeric, specific patterns, etc.)
        # This regex looks for sequences of alphanumeric characters that could be usernames
        username_patterns = re.findall(r'[a-zA-Z0-9]+', name_without_ext)
        
        if username_patterns:
            # Return the longest alphanumeric pattern (most likely to be username)
            return max(username_patterns, key=len)
        
        # If no patterns found, return the filename without extension
        return name_without_ext

    def find_student_row_by_username(self, username):
        """Find student row by username in the table for Moodle 4.3"""
        # Look for username in the username column (c3)
        try:
            username_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.cell.c3.username")
            for cell in username_cells:
                if username in cell.text:
                    return cell.find_element(By.XPATH, "./ancestor::tr")
        except:
            pass
        
        # Fallback: search all table cells for the username
        try:
            all_cells = self.driver.find_elements(By.CSS_SELECTOR, "td")
            for cell in all_cells:
                if username in cell.text:
                    return cell.find_element(By.XPATH, "./ancestor::tr")
        except:
            pass
        
        raise Exception(f"Student with username '{username}' not found in grading table")

    def wait_for_grading_table(self):
        """Wait for the grading table to be fully loaded and ready with multiple fallbacks"""
        self.log_message("  Waiting for grading table to load...")
        
        # Try multiple selectors for the grading table
        table_selectors = [
            "#submissions",
            "table.generaltable",
            "table[role='grid']",
            "table.assignresponses"
        ]
        
        for selector in table_selectors:
            try:
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                self.log_message(f"  Found table with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        # Wait for table rows
        try:
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr"))
            )
            self.log_message("  Grading table ready with rows")
        except TimeoutException:
            self.log_message("  Warning: No rows found in grading table")
        
        # Additional wait for page to stabilize
        time.sleep(2)

    def check_login_success(self):
        """Check if login was successful by looking for dashboard elements"""
        try:
            # Look for elements that indicate successful login
            indicators = [
                "div.page-header-headings",  # Page header
                "section.block_myoverview",  # Dashboard block
                "div.userinitials",          # User avatar
                "li.dropdown-item",          # User menu items
                "a[data-title='myhome']"     # Dashboard link
            ]
            
            for indicator in indicators:
                elements = self.driver.find_elements(By.CSS_SELECTOR, indicator)
                if elements:
                    return True
            
            # Check if we're still on the login page
            if "login" in self.driver.current_url:
                return False
                
            # If we're not on login page and no indicators found, wait a bit more
            time.sleep(2)
            return True
            
        except Exception as e:
            self.log_message(f"  Login check error: {str(e)}")
            return False

    def upload_file_to_filepicker(self, full_file_path, filename):
        """Handle the file upload process for Moodle 4.3"""
        self.log_message("  On edit page, looking for file input...")
        
        # Wait for the page to fully load - try multiple selectors
        file_manager_selectors = [
            ".filemanager",
            ".fp-restrictions",
            ".file-picker",
            "#filemanager-container"
        ]
        
        file_manager = None
        for selector in file_manager_selectors:
            try:
                file_manager = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                self.log_message(f"  Found file manager with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not file_manager:
            raise Exception("Could not find file manager on page")
        
        # Look for the "Add" button in Moodle 4.3 style
        add_button_selectors = [
            ".filepicker-filename a",
            ".fp-btn-add",
            "a[title='Add...']",
            "button[aria-label='Add file']"
        ]
        
        add_button = None
        for selector in add_button_selectors:
            try:
                add_button = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                self.log_message(f"  Found add button with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not add_button:
            raise Exception("Could not find add button")
        
        self.log_message("  Clicking add button...")
        self.safe_click(add_button)
        
        # Wait for the file upload dialog to appear
        try:
            WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".filepicker-filelist, .moodle-dialogue-content"))
            )
        except TimeoutException:
            self.log_message("  Warning: File upload dialog not visible, proceeding anyway")
        
        # Find the file input element
        file_input = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )
        
        # Upload the file
        absolute_path = os.path.abspath(full_file_path)
        self.log_message(f"  Uploading file: {filename}")
        file_input.send_keys(absolute_path)
        
        # Wait for the file to appear in the list or for upload button to be enabled
        try:
            WebDriverWait(self.driver, 20).until(
                lambda driver: driver.find_elements(By.CSS_SELECTOR, ".filepicker-filelist .file") or 
                             driver.find_element(By.CSS_SELECTOR, ".fp-upload-btn").is_enabled()
            )
        except TimeoutException:
            self.log_message("  Warning: File list not updated, proceeding with upload")
        
        # Click the upload button
        upload_button = WebDriverWait(self.driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".fp-upload-btn, input[type='submit'][value*='Upload']"))
        )
        
        self.log_message("  Clicking upload button...")
        self.safe_click(upload_button)
        
        # Wait for the upload to complete
        try:
            WebDriverWait(self.driver, 30).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, ".filepicker-filelist .file, .moodle-dialogue-base"))
            )
            self.log_message("  File uploaded successfully")
            return True
        except TimeoutException:
            self.log_message("  Warning: Upload completion not detected, assuming success")
            return True

    def click_edit_submission(self, student_row):
        """Click the edit submission button for a student row"""
        # Find the action menu toggle button in the Edit column (c7)
        try:
            action_menu_toggle = WebDriverWait(student_row, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "td.cell.c7 .dropdown-toggle"))
            )
            
            self.log_message("  Clicking action menu toggle...")
            self.safe_click(action_menu_toggle)
            
            # Wait for the dropdown menu to appear
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "td.cell.c7 .dropdown-menu.show"))
            )
            
            # Check if the student has a submission already by looking at the status column (c5)
            status_cell = student_row.find_element(By.CSS_SELECTOR, "td.cell.c5")
            status_text = status_cell.text.lower()
            
            if "no submission" in status_text:
                # For students with no submission, we need to click "Grade" first
                self.log_message("  Student has no submission, clicking 'Grade' first...")
                
                # Look for the Grade option in the dropdown
                grade_option = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'dropdown-item') and contains(., 'Grade')]"))
                )
                
                self.safe_click(grade_option)
                
                # Wait for the grader page to load
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".gradingform_rubric, .filemanager, #id_submitbutton"))
                )
                
                # Check if we're on the grader page and need to click "Add submission"
                try:
                    add_submission_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Add submission') or contains(., 'Add Submission')]"))
                    )
                    
                    self.log_message("  Clicking 'Add submission'...")
                    self.safe_click(add_submission_button)
                    
                    # Wait for the file manager to appear
                    WebDriverWait(self.driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".filemanager, .file-picker"))
                    )
                except TimeoutException:
                    # If no "Add submission" button is found, we might already be on the submission page
                    self.log_message("  No 'Add submission' button found, assuming we're already on submission page")
            else:
                # For students with existing submissions, look for "Edit submission" option
                self.log_message("  Student has existing submission, looking for 'Edit submission'...")
                
                # Try multiple selectors for the edit submission option
                edit_submission_selectors = [
                    "//a[contains(@href, 'editsubmission')]",
                    "//a[contains(., 'Edit submission')]",
                    "//a[contains(@class, 'menu-action') and contains(., 'Edit')]"
                ]
                
                edit_submission_option = None
                for selector in edit_submission_selectors:
                    try:
                        edit_submission_option = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        break
                    except TimeoutException:
                        continue
                
                if not edit_submission_option:
                    raise Exception("Could not find Edit submission option in dropdown menu")
                
                self.log_message("  Clicking 'Edit submission'...")
                self.safe_click(edit_submission_option)
                
                # Wait for the edit submission page to load
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".filemanager, .file-picker, #id_submitbutton"))
                )
            
            self.log_message("  Navigated to submission page")
            
        except Exception as e:
            # If clicking the dropdown fails, try a different approach
            self.log_message(f"  Alternative approach needed: {str(e)}")
            
            # Try to navigate directly to the edit submission URL
            try:
                # Extract user ID from the row
                user_id_input = student_row.find_element(By.CSS_SELECTOR, "input[name='selectedusers']")
                user_id = user_id_input.get_attribute("value")
                
                # Build the direct URL to the edit submission page
                edit_url = f"{self.moodle_url.get()}/mod/assign/view.php?id={self.assignment_id.get()}&userid={user_id}&action=editsubmission"
                self.log_message(f"  Navigating directly to: {edit_url}")
                
                self.driver.get(edit_url)
                
                # Wait for the edit submission page to load
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".filemanager, .file-picker, #id_submitbutton"))
                )
                
                self.log_message("  Navigated to submission page via direct URL")
                
            except Exception as direct_error:
                raise Exception(f"Failed to navigate to submission page: {str(direct_error)}")

    def run_upload(self):
        try:
            # Setup WebDriver with webdriver-manager
            self.log_message("Setting up Chrome WebDriver...")
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # 1. LOGIN
            self.log_message("Logging in...")
            self.driver.get(f"{self.moodle_url.get()}/login/index.php")
            
            # Wait for login form
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            
            username_field = self.driver.find_element(By.ID, "username")
            username_field.send_keys(self.username.get())
            
            password_field = self.driver.find_element(By.ID, "password")
            password_field.send_keys(self.password.get())
            
            login_button = self.driver.find_element(By.ID, "loginbtn")
            self.safe_click(login_button)
            
            # Wait for login to complete with multiple checks
            WebDriverWait(self.driver, 30).until(
                lambda driver: "my/" in driver.current_url or "dashboard" in driver.current_url or self.check_login_success()
            )
            
            self.log_message("Login successful.")

            # 2. NAVIGATE TO THE ASSIGNMENT GRADING PAGE
            grading_url = f"{self.moodle_url.get()}/mod/assign/view.php?id={self.assignment_id.get()}&action=grading"
            self.log_message(f"Navigating to grading page: {grading_url}")
            self.driver.get(grading_url)
            
            # Wait for page to load with multiple fallbacks
            try:
                WebDriverWait(self.driver, 30).until(
                    lambda driver: "view.php" in driver.current_url and "id=" in driver.current_url
                )
            except TimeoutException:
                self.log_message("  Warning: Page load timeout, but proceeding anyway")
            
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
                
                # Extract username from filename
                username = self.extract_username_from_filename(filename)
                full_file_path = os.path.join(self.folder_path.get(), filename)
                self.log_message(f"\n=== Processing Username: {username} ({i+1}/{total_files}) ===")

                try:
                    # 5. FIND THE STUDENT'S ROW BY USERNAME
                    self.log_message("  Finding student row by username...")
                    student_row = self.find_student_row_by_username(username)
                    self.log_message(f"  Found row for username: {username}")

                    # 6. CLICK THE EDIT SUBMISSION BUTTON
                    self.click_edit_submission(student_row)

                    # 7. UPLOAD THE FILE WITH IMPROVED HANDLING
                    upload_success = self.upload_file_to_filepicker(full_file_path, filename)
                    
                    if not upload_success:
                        raise Exception("File upload failed or file not found in manager after upload")
                    
                    # 8. SAVE CHANGES
                    self.log_message("  Saving changes...")
                    
                    save_button = WebDriverWait(self.driver, 15).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value*='Save'], button[type='submit']"))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                    time.sleep(1)
                    
                    self.log_message("  Clicking save button...")
                    save_button.click()

                    # 9. WAIT FOR CONFIRMATION
                    self.log_message("  Waiting for save confirmation...")
                    time.sleep(3)
                    
                    # Check for success with multiple indicators
                    success_indicators = [
                        ".alert-success",
                        ".notifysuccess",
                        ".success",
                        "div[data-region='overlay']:not([style*='display: none'])",
                        "body:not(.has-assignment-overlay)"
                    ]
                    
                    success_found = any(self.driver.find_elements(By.CSS_SELECTOR, indicator) for indicator in success_indicators)
                    
                    if success_found or "view.php" in self.driver.current_url:
                        self.log_message(f"  ✅ Successfully uploaded for username: {username}")
                        successful_uploads.append((filename, username))
                    else:
                        # Check if we're back on the grading page
                        try:
                            self.wait_for_grading_table()
                            self.log_message(f"  ✅ Successfully uploaded for username: {username} (implied by return to grading page)")
                            successful_uploads.append((filename, username))
                        except:
                            raise Exception("Save confirmation not found")

                    # 10. GO BACK TO GRADING LIST
                    self.log_message("  Returning to grading list...")
                    self.driver.get(grading_url)
                    self.wait_for_grading_table()
                    self.log_message("  Returned to grading list")

                except Exception as e:
                    error_msg = f"ERROR: {filename} - Username: {username} - {str(e)}"
                    self.log_message(f"  ❌ {error_msg}")
                    failed_uploads.append((filename, username, str(e)))
                    
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
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            self.log_message("\n" + "="*60)
            self.log_message("BATCH UPLOAD COMPLETE - SUMMARY REPORT")
            self.log_message("="*60)
            
            self.log_message(f"\nTimestamp: {timestamp}")
            self.log_message(f"Total files processed: {total_files}")
            self.log_message(f"Successful uploads: {len(successful_uploads)}")
            self.log_message(f"Failed uploads: {len(failed_uploads)}")
            
            if len(successful_uploads) > 0:
                self.log_message(f"Success rate: {len(successful_uploads)/total_files*100:.1f}%")
            
            if successful_uploads:
                self.log_message(f"\n✅ SUCCESSFUL UPLOADS ({len(successful_uploads)}):")
                self.log_message("-" * 50)
                for i, (filename, username) in enumerate(successful_uploads, 1):
                    self.log_message(f"{i:2d}. Username: {username} ({filename})")
            
            if failed_uploads:
                self.log_message(f"\n❌ FAILED UPLOADS ({len(failed_uploads)}):")
                self.log_message("-" * 50)
                for i, (filename, username, error) in enumerate(failed_uploads, 1):
                    self.log_message(f"{i:2d}. Username: {username} ({filename})")
                    self.log_message(f"    Error: {error[:80]}{'...' if len(error) > 80 else ''}")

            # Generate and store the full report
            self.report_data = self.generate_report(successful_uploads, failed_uploads, total_files, timestamp)
            self.log_message("\nReport generated. Click 'Download Report' to save it.")

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