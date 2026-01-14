#!/usr/bin/env pythonw
"""
Task Manager GUI - A graphical application to manage daily tasks.
"""
import os
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
from typing import Dict, List, Optional, Any, Tuple
import subprocess  # Add this import for the subprocess module
import webbrowser
import logging
import threading
import pymssql  # Changed: Import pymssql instead of pyodbc
import getpass # Import the 'getpass' module to get username
import tkinter.messagebox as messagebox # Import messagebox
import platform
import sys # <<< Add sys
import appdirs # <<< Add appdirs (requires pip install appdirs)
import keyring # Added for secure password storage
import smartdb_login # Added for scraping functionality
from tkinter import filedialog # Added for askdirectory
import time # Added for timer
import shutil # Added for moving files

# --- Determine Log File Path ---
def get_log_file_path():
    """Determines the appropriate absolute path for the log file, prioritizing the script/exe directory."""
    app_name = "TaskManager"
    # author_name = "TaskManagerApp" # Not needed if appdirs isn't used

    # <<< REMOVED appdirs logic >>>
    # try:
    #     # Preferred: User's AppData directory (log subfolder)
    #     log_dir = appdirs.user_log_dir(app_name, author_name)
    #     # ... (rest of appdirs block removed) ...
    # except Exception as e_appdata:
    #     print(f"[WARNING] Could not use AppData directory ({log_dir}): {e_appdata}. Trying executable directory.")

    # <<< MODIFIED: Make executable/script directory the primary attempt >>>
    try:
        if getattr(sys, 'frozen', False):
            # Running as bundled executable
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as a script
            base_path = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(base_path, 'task_manager.log')
        # Basic write test
        try:
            # <<< Use 'a' append mode for testing to avoid clearing existing file if test fails later >>>
            with open(log_path, 'a') as f: 
                f.write("") # Test write permission
        except Exception as e:
            print(f"[WARNING] Could not write to log file in executable/script directory: {e}")
            raise # Re-raise to trigger fallback
        # <<< MODIFIED: Log message indicates this is the intended location >>>
        print(f"[INFO] Attempting to log to executable/script directory: {log_path}") 
        return log_path
    except Exception as e_exe_dir:
        print(f"[ERROR] Could not use executable/script directory: {e_exe_dir}")
        # Final fallback: Current working directory (less reliable)
        log_path = 'task_manager.log'
        print(f"[CRITICAL] Falling back to current working directory for log: {os.path.abspath(log_path)}")
        return log_path

# Configure logging *using the absolute path*
log_file_path = get_log_file_path() # <<< Call the function (no change here)

# <<< Get the root logger and remove existing handlers >>>
# This prevents adding handlers multiple times if the setup code is run again
root_logger = logging.getLogger()
# Remove all handlers associated with the root logger object.
for handler in root_logger.handlers[:]:
    root_logger.removeHandler(handler)
    handler.close() # Ensure handlers are closed before removing

# <<< Create handlers explicitly >>>
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s')

# File Handler
try:
    file_handler = logging.FileHandler(log_file_path, mode='w') # Overwrite log each run
    file_handler.setFormatter(log_formatter)
    root_logger.addHandler(file_handler)
except Exception as e:
     print(f"[ERROR] Failed to create file handler for log file '{log_file_path}': {e}")
     file_handler = None # Ensure it's None if creation failed

# Stream Handler (Console)
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(log_formatter)
root_logger.addHandler(stream_handler)

# Set level on the root logger
root_logger.setLevel(logging.INFO)

# Log the chosen path after basicConfig is set up
logging.info(f"--- Log Start ---")
logging.info(f"Log file path determined: {log_file_path}") # (no change here)
# <<< Add an immediate flush >>>
if file_handler:
    file_handler.flush()
# <<< Add an immediate test message >>>
logging.info("--- Initial logging setup complete. This message should appear in the log. ---")
if file_handler:
    file_handler.flush()

# List of users that can be assigned tasks
USERS = ["All", "Jay", "Jude", "Jorgen", "Earl", "Philip", "Sam", "Glenn"]

class TaskManagerApp:
    """GUI application for managing tasks."""
    
    def __init__(self, root, task_manager):
        """Initialize the GUI application."""
        self.root = root
        self.root.title("Task Manager")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=('Helvetica', 10))
        self.style.configure("TLabel", font=('Helvetica', 10))
        self.style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        self.style.configure("Title.TLabel", font=('Helvetica', 16, 'bold'))
        
        # Try to set a modern theme if available
        try:
            self.root.tk.call("source", "azure.tcl")
            self.style.theme_use("azure")
        except tk.TclError:
            # If azure theme is not available, try other themes
            available_themes = self.style.theme_names()
            for theme in ["clam", "alt", "vista"]:
                if theme in available_themes:
                    self.style.theme_use(theme)
                    break
        
        # Get current user
        self.current_user = self._get_current_user()
        


        # Connection string - Not used directly with pymssql's connect()
        # self.CONNECTION_STRING = f'DRIVER={{SQL Server}};SERVER={self.SERVER};DATABASE={self.DATABASE};Trusted_Connection=yes;'
        
        # Initialize connection_error flag and task manager
        self.connection_error = False
        try:
            self.task_manager = task_manager
        except Exception as e:
            self.connection_error = True
            messagebox.showerror("Connection Error", 
                                f"Could not connect to the SQL Server database.\n\n"
                                f"Error: {str(e)}\n\n"
                                f"The application will run without database connection.")
            logging.error("Connection error at startup", exc_info=True) # Log the exception
            self.task_manager = None
        
        # Initialize deleted tasks stack
        self.deleted_tasks_stack = []
        self.max_undo_stack_size = 10  # Maximum number of undo operations to keep
        
        # Initialize refresh lock to prevent multiple simultaneous refresh operations
        self.refresh_lock = threading.Lock()
        self.refresh_in_progress = False
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        self.create_header()
        
        # Create task list frame
        self.create_task_list()
        
        # Create control panel
        self.create_control_panel()
        
        # Create status bar
        self.create_status_bar()
        
        # Set initial status
        self.set_status("Loading tasks...")
        
        # Schedule task loading after UI is displayed
        self.root.after(100, self.refresh_task_list)
        
        # Bind the window close event to cleanup method
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def _get_current_user(self):
        """Get the current user's username."""
        try:
            username = os.getenv('USERNAME') or os.getenv('USER') or 'Unknown User'
            logging.info(f"Current user: {username}") # Log the username
            return username
        except:
            return 'Unknown User'
    
    def on_close(self):
        """Handle application close event - clean up resources and close database connections."""
        try:
            logging.info("Application shutdown sequence started.") # Log shutdown start
            # Close database connection if it exists
            if hasattr(self, 'task_manager') and self.task_manager:
                if hasattr(self.task_manager, 'conn') and self.task_manager.conn:
                    try:
                        # Commit any pending changes
                        self.task_manager.conn.commit()
                        # Close the connection
                        self.task_manager.conn.close()
                        logging.info("Database connection closed properly")
                    except Exception as e:
                        logging.error(f"Error closing database connection: {str(e)}", exc_info=True)
        except Exception as e:
            logging.error(f"Error during application shutdown: {str(e)}", exc_info=True)
        finally:
            logging.info("Flushing log handlers before exit...")
            # <<< Explicitly flush and close handlers on exit >>>
            for handler in logging.getLogger().handlers:
                try:
                    handler.flush()
                    handler.close()
                except Exception as e:
                    print(f"Error closing handler {handler}: {e}") # Print error during close
            # Close the application
            self.root.destroy()
    
    def create_header(self):
        """Create the application header."""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create a left frame for title and storage indicator
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, padx=5)
        
        title_label = ttk.Label(title_frame, text="Task Manager", style="Title.TLabel")
        title_label.pack(side=tk.TOP, anchor=tk.W)
        
        # Show storage location indicator under the title
        if self.connection_error:
            self.storage_indicator = ttk.Label(title_frame, text="(No Database)", foreground="red")
        else:
            self.storage_indicator = ttk.Label(title_frame, text="(SQL Server)", foreground="green")
        self.storage_indicator.pack(side=tk.TOP, anchor=tk.W)
        
        # Filter controls
        filter_frame = ttk.Frame(header_frame)
        filter_frame.pack(side=tk.RIGHT, padx=5)
        
        # Category filter
        ttk.Label(filter_frame, text="Category:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.category_var = tk.StringVar(value="All")
        self.category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var, width=15, state="readonly")
        self.category_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.update_category_filter()
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # Main Staff filter
        ttk.Label(filter_frame, text="Main Staff:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.main_staff_var = tk.StringVar(value="All")
        self.main_staff_combo = ttk.Combobox(filter_frame, textvariable=self.main_staff_var, width=15, values=USERS, state="readonly")
        self.main_staff_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.main_staff_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # User filter
        ttk.Label(filter_frame, text="User:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.user_var = tk.StringVar(value="All")
        self.user_combo = ttk.Combobox(filter_frame, textvariable=self.user_var, width=15, values=USERS, state="readonly")
        self.user_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.user_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_task_list())
        
        # Show completed checkbox - with fixed width to ensure full text is visible
        ttk.Label(filter_frame, text="Status:").pack(side=tk.LEFT, padx=(0, 5))
        
        status_frame = ttk.Frame(filter_frame)
        status_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        self.show_completed_var = tk.BooleanVar(value=False)
        show_completed_check = ttk.Checkbutton(
            status_frame, 
            text="Show Completed Tasks", 
            variable=self.show_completed_var,
            command=self.refresh_task_list,
            width=22  # Increased width to ensure full text is visible including the "s"
        )
        show_completed_check.pack(side=tk.LEFT)
    
    def create_task_list(self):
        """Create the task list with a treeview."""
        # Create frame with scrollbar
        task_frame = ttk.Frame(self.main_frame)
        task_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create treeview with both vertical and horizontal scrollbars
        columns = (
            "status", "title", "rev", "applied_vessel",
            "priority", "main_staff", "assigned_to",
            "qtd_mhr", "actual_mhr"  # Add new columns here
        )
        
        # Create a frame to hold the treeview and scrollbars
        tree_frame = ttk.Frame(task_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create the treeview
        self.task_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # Define standard font for all treeview elements - make it bold
        standard_font = ('TkDefaultFont', 9, 'bold')
        
        # Configure the treeview to use the standard font
        self.style.configure("Treeview", font=standard_font)
        self.style.configure("Treeview.Heading", font=('TkDefaultFont', 9, 'bold'))
        
        # Define headings with proper command for toggling sort direction
        self.task_tree.heading("status", text="Status", command=lambda: self.sort_treeview("status", True))
        self.task_tree.heading("title", text="Equipment Name", command=lambda: self.sort_treeview("title", True))
        self.task_tree.heading("rev", text="Rev", command=lambda: self.sort_treeview("rev", True))
        self.task_tree.heading("applied_vessel", text="Applied Vessel", command=lambda: self.sort_treeview("applied_vessel", True))
        self.task_tree.heading("priority", text="Priority", command=lambda: self.sort_treeview("priority", True))
        self.task_tree.heading("main_staff", text="Main Staff", command=lambda: self.sort_treeview("main_staff", True))
        self.task_tree.heading("assigned_to", text="Assigned To", command=lambda: self.sort_treeview("assigned_to", True))
        self.task_tree.heading("qtd_mhr", text="Qtd Mhr", command=lambda: self.sort_treeview("qtd_mhr", True))  # Heading for Qtd Mhr
        self.task_tree.heading("actual_mhr", text="Actual Mhr", command=lambda: self.sort_treeview("actual_mhr", True))  # Heading for Actual Mhr
        
        # Define optimized column widths - REDUCED WIDTHS FOR BETTER VISIBILITY
        self.task_tree.column("status", width=120, minwidth=100, anchor=tk.W)  # Reduced width
        self.task_tree.column("title", width=250, minwidth=200) # Reduced width
        self.task_tree.column("rev", width=40, minwidth=40, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("applied_vessel", width=100, minwidth=100) # Reduced width
        self.task_tree.column("priority", width=70, minwidth=70, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("main_staff", width=80, minwidth=80, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("assigned_to", width=80, minwidth=80, anchor=tk.CENTER) # Reduced width
        self.task_tree.column("qtd_mhr", width=70, minwidth=70, anchor=tk.CENTER)  # Width for Qtd Mhr - slightly reduced
        self.task_tree.column("actual_mhr", width=70, minwidth=70, anchor=tk.CENTER)  # Width for Actual Mhr - slightly reduced
        
        # Initialize sorting variables
        self.sort_column = None
        self.sort_reverse = False
        
        # Add vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Add horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.task_tree.xview)
        self.task_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Grid layout for treeview and scrollbars to ensure they're always visible
        self.task_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Configure the grid to expand properly
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Create a loading indicator overlay
        self.loading_frame = ttk.Frame(tree_frame, style="TFrame")
        self.loading_label = ttk.Label(
            self.loading_frame, 
            text="Loading tasks...", 
            font=('Helvetica', 12),
            padding=20
        )
        self.loading_label.pack(expand=True, fill=tk.BOTH)
        
        # Bind double-click to view task details
        self.task_tree.bind("<Double-1>", self.view_task_details)
        
        # Bind right-click to show context menu
        self.task_tree.bind("<Button-3>", self.show_context_menu)
        
        # Bind single-click to check for link clicks
        self.task_tree.bind("<ButtonRelease-1>", self.check_link_click)
        
        # Bind mouse motion to change cursor over links
        self.task_tree.bind("<Motion>", self.update_cursor)
        
        # Configure tag for links - use bold font for consistency
        self.task_tree.tag_configure("link", foreground="#0066cc", font=standard_font)
        
        # Configure tag for link hover effect
        self.task_tree.tag_configure("link_hover", foreground="#0099ff", font=standard_font)
        
        # Configure tag for link click effect
        self.task_tree.tag_configure("link_click", foreground="#003366", font=standard_font)
        
        # Configure tag for completed tasks with gray text and background
        self.task_tree.tag_configure("completed", foreground="#9e9e9e", background="#e0e0e0", font=standard_font)  # Gray text and light gray background
    
    def show_loading_indicator(self):
        """Show the loading indicator over the treeview."""
        if hasattr(self, 'loading_frame'):
            self.loading_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            self.root.update_idletasks()
    
    def hide_loading_indicator(self):
        """Hide the loading indicator."""
        if hasattr(self, 'loading_frame'):
            self.loading_frame.place_forget()
            self.root.update_idletasks()

    def _get_due_date_color(self, due_date_str):
        """Calculate background color based on due date proximity."""
        if not due_date_str:
            return "#ffffff"  # White for no due date

        try:
            due_date = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
            today = datetime.date.today()
            days_until_due = (due_date - today).days

            if days_until_due < 0:
                return "#ffcccc"  # More intense light red for overdue
            elif days_until_due == 0:
                return "#ffd9d9"  # More intense red for due today
            elif days_until_due <= 3:
                return "#ffe6e6"  # Light red for near due
            elif days_until_due <= 7:
                return "#fff2cc"  # More intense light yellow for upcoming
            elif days_until_due <= 14:
                return "#e6ffcc"  # More intense light yellow-green
            else:
                return "#ccffcc"  # More intense light green for far future
        except ValueError:
            return "#ffffff"  # White for invalid date format

    def create_control_panel(self):
        """Create the control panel with buttons."""
        # Create a container with a subtle border
        control_frame = ttk.Frame(self.main_frame, style="TFrame")
        control_frame.pack(fill=tk.X, pady=10)
        
        # Main button frame
        button_frame = ttk.Frame(control_frame, style="TFrame")
        button_frame.pack(side=tk.LEFT)
        
        # Add task button
        add_btn = ttk.Button(
            button_frame, 
            text=" Add Task", 
            command=self.add_task,
            style="TButton"
        )
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Undo button
        self.undo_btn = ttk.Button(
            button_frame,
            text=" Undo",
            command=self.undo_delete,
            state=tk.DISABLED,  # Initially disabled
            style="TButton"
        )
        self.undo_btn.pack(side=tk.LEFT, padx=5)
        
        # Refresh button
        refresh_btn = ttk.Button(
            button_frame, 
            text=" Refresh", 
            command=self.refresh_task_list,
            style="TButton"
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # Search bar frame
        search_frame = ttk.Frame(control_frame, style="TFrame")
        search_frame.pack(side=tk.RIGHT, padx=(20, 0))
        
        # Search label
        search_label = ttk.Label(search_frame, text="Search:", style="TLabel")
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Search entry
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            search_frame, 
            textvariable=self.search_var, 
            width=20,
            font=('Helvetica', 10)
        )
        self.search_entry.pack(side=tk.LEFT)
        
        # Bind search events
        self.search_var.trace('w', self.on_search_change)
        self.search_entry.bind('<Return>', lambda e: self.refresh_task_list())
        self.search_entry.bind('<Escape>', self.clear_search)
        
        # Clear search button - smaller size to match entry height
        clear_search_btn = ttk.Button(
            search_frame,
            text="✕",
            command=self.clear_search,
            width=2,  # Reduced width
            style="TButton"
        )
        clear_search_btn.pack(side=tk.LEFT, padx=(2, 0))
    
    def create_status_bar(self):
        """Create the status bar at the bottom of the main window."""
        status_frame = ttk.Frame(self.root, relief=tk.FLAT, padding="2 2 2 2")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        separator = ttk.Separator(status_frame, orient=tk.HORIZONTAL)
        separator.pack(fill=tk.X, pady=(0, 2))
        
        # Status bar with task statistics
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(
            status_frame, 
            textvariable=self.status_var, 
            anchor=tk.W,
            style="Status.TLabel"
        )
        status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Add user info to the right side of the status bar
        username = self.task_manager.get_sql_username()
        if username == "TaskUser1":
            display_name = "Jay"
        elif username == "TaskUser2":
            display_name = "Jude"
        elif username == "TaskUser3":
            display_name = "Jorgen"
        elif username == "TaskUser4":
            display_name = "Earl"
        elif username == "TaskUser5":
            display_name = "Philip"
        elif username == "TaskUser6":
            display_name = "Samuel"
        elif username == "TaskUser7":
            display_name = "Glenn"
        else:
            display_name = username  # Default to username if not in the list

        user_label = ttk.Label(
            status_frame, 
            text=f"User: {display_name}",
            anchor=tk.E,
            style="Status.TLabel"
        )
        user_label.pack(side=tk.RIGHT, padx=10)
        
        # Initialize status
        self.update_status()
    
    def update_status(self):
        """Update the status bar with task statistics."""
        # Get filtered non-deleted tasks based on active filters
        category = None if self.category_var.get() == "All" else self.category_var.get()
        main_staff = None if self.main_staff_var.get() == "All" else self.main_staff_var.get()
        assigned_to = None if self.user_var.get() == "All" else self.user_var.get()
        
        # Get non-deleted tasks with current filters applied
        filtered_tasks = self.task_manager.get_filtered_tasks(
            show_completed=True,
            category=category,
            main_staff=main_staff,
            assigned_to=assigned_to,
            show_deleted=False  # Explicitly exclude deleted tasks
        )
        
        total = len(filtered_tasks)
        completed = sum(1 for task in filtered_tasks if task["completed"])
        pending = total - completed
        
        # Create a more informative status message (total now excludes completed tasks)
        status_text = f"Total Tasks: {pending} | Completed: {completed} | Pending: {pending}"
        
        # Add filter information if filters are active
        filters_active = []
        
        if self.category_var.get() != "All":
            filters_active.append(f"Category: {self.category_var.get()}")
            
        if self.main_staff_var.get() != "All":
            filters_active.append(f"Main Staff: {self.main_staff_var.get()}")
            
        if self.user_var.get() != "All":
            filters_active.append(f"Assigned To: {self.user_var.get()}")
        
        if filters_active:
            status_text += " | Filters: " + ", ".join(filters_active)
        
        # Add search information if search is active
        if hasattr(self, 'search_var') and self.search_var.get().strip():
            search_term = self.search_var.get().strip()
            status_text += f" | Search: '{search_term}'"
        
        self.status_var.set(status_text)
    
    def refresh_task_list(self):
        """Refresh the task list in the treeview."""
        # Check if refresh is already in progress
        with self.refresh_lock:
            if self.refresh_in_progress:
                return  # Exit early if refresh is already running
            self.refresh_in_progress = True
        
        # Show loading indicator
        self.set_status("Loading tasks...")
        self.show_loading_indicator()

        # Use threading to perform task refresh in the background
        threading.Thread(target=self._perform_task_refresh_threaded, daemon=True).start()

    def _perform_task_refresh_threaded(self):
        """Threaded method to perform task refresh."""
        try:
            self._perform_task_refresh() # Call the original refresh method
        except Exception as e:
            logging.error("Error during task refresh in thread", exc_info=True)
            # Handle errors that occur during task refresh in the background thread
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error refreshing task list: {str(e)}", parent=self.root)) # Use lambda for deferred execution
        finally:
            # Reset the refresh flag and ensure loading indicator is hidden even if there's an error
            with self.refresh_lock:
                self.refresh_in_progress = False
            self.root.after(0, self.hide_loading_indicator) # Use root.after to run in main thread
            self.root.after(0, self.update_status) # Update status bar in main thread

    def _perform_task_refresh(self):
        """Perform the actual task refresh after UI has updated."""
        try:
            # First, explicitly reload all tasks from database to get the latest changes
            try:
                self.task_manager.reload_tasks() # Force reload tasks from database
            except Exception as e:
                logging.error(f"Error reloading tasks: {str(e)}", exc_info=True)
                raise Exception(f"Database error: {str(e)}")

            # Clear existing items in the treeview
            for item in self.task_tree.get_children():
                self.task_tree.delete(item)

            # Get filtered tasks
            show_completed = self.show_completed_var.get()
            category = None if self.category_var.get() == "All" else self.category_var.get()
            main_staff = None if self.main_staff_var.get() == "All" else self.main_staff_var.get()
            assigned_to = None if self.user_var.get() == "All" else self.user_var.get()

            # Get search term if available
            search_term = ""
            if hasattr(self, 'search_var'):
                search_term = self.search_var.get().lower().strip()

            # Get filtered tasks
            filtered_tasks = self.task_manager.get_filtered_tasks(
                show_completed=show_completed,
                category=category,
                main_staff=main_staff,
                assigned_to=assigned_to
            )

            # Apply search filter if needed
            if search_term:
                filtered_tasks = [
                    task for task in filtered_tasks if self._task_matches_search(task, search_term)
                ]

            # Apply sorting if a sort column is set
            if self.sort_column:
                self.apply_sorting(filtered_tasks)

            # Track the maximum status text length to adjust column width
            max_status_length = 0

            # Define a standard font for all task items - make it bold
            standard_font = ('TkDefaultFont', 9, 'bold')

            # Add tasks to treeview with improved visual indicators
            for task in filtered_tasks:
                # Create status with time remaining - using standardized format for better alignment
                if task["completed"]:
                    status = "✓ Completed"
                    status_tags = ["completed"]
                else:
                    # Add due date information for pending tasks
                    if task["due_date"]:
                        try:
                            due_date = datetime.datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                            today = datetime.date.today()
                            days_until_due = (due_date - today).days
                            
                            if days_until_due < 0:
                                # Use consistent format for overdue
                                status = f"! {abs(days_until_due)} Days Overdue"
                                status_tags = ["overdue"]
                            elif days_until_due == 0:
                                status = "! Due Today"
                                status_tags = ["due_today"]
                            elif days_until_due == 1:
                                status = "! Due Tomorrow"
                                status_tags = ["pending_soon"]
                            else:
                                # Use consistent format for days left
                                status = f"• {days_until_due} Days Left"
                                status_tags = ["pending"]
                                if days_until_due <= 3:
                                    status_tags = ["pending_soon"]
                        except ValueError:
                            status = "• No Due Date"
                            status_tags = ["pending"]
                    else:
                        status = "• No Due Date"
                        status_tags = ["pending"]
                
                # Update max status length for column width adjustment
                max_status_length = max(max_status_length, len(status))
                
                # Format due date
                due_date = task["due_date"] if task["due_date"] else "No due date"
                
                # Format staff assignments
                main_staff = task.get("main_staff", "Unassigned")
                assigned_to = task.get("assigned_to", "Unassigned")
                
                # Get values for new columns (with defaults if not present)
                applied_vessel = task.get("applied_vessel", "")
                rev = task.get("rev", "")
                
                # Set row tags for styling based on priority and completion
                if task["completed"]:
                    # For completed tasks, only use the completed tag to ensure it overrides all other styles
                    tags = ["completed"]
                else:
                    # For pending tasks, use priority and status tags
                    tags = [task["priority"]] + status_tags
                
                # Format priority with emoji (without link indicator)
                priority_display = {
                    "high": "🔴 High",
                    "medium": "🔵 Medium",
                    "low": "🟢 Low"
                }.get(task["priority"].lower(), task["priority"].capitalize())
                
                # Create link display
                link_display = "🔗" if task.get("link") else ""
                
                # Create a unique tag for this task's due date color
                due_date_tag = f"due_date_{task['id']}"
                tags.append(due_date_tag)
                
                # Configure the tag with the appropriate background color
                bg_color = self._get_due_date_color(task["due_date"])
                self.task_tree.tag_configure(due_date_tag, background=bg_color)
                
                # For completed tasks, make sure "completed" is the first tag to ensure it has priority
                if task["completed"]:
                    # Remove any status tags that might override the completed style
                    tags = [tag for tag in tags if tag not in ["overdue", "due_today", "pending_soon", "pending"]]
                    # Put completed tag first to ensure it has priority
                    tags.insert(0, "completed")
                
                # Add link tag if there's a link
                if task.get("link"):
                    tags.append("link")
                
                self.task_tree.insert(
                    "", tk.END, 
                    values=(
                        status,
                        task["title"],
                        task.get("rev", ""),
                        task.get("applied_vessel", ""),
                        priority_display,
                        task.get("main_staff", ""),
                        task.get("assigned_to", ""),
                        task.get("qtd_mhr", ""),  # Qtd Mhr value
                        task.get("actual_mhr", "") # Actual Mhr value
                    ),
                    tags=tags,
                    iid=str(task["id"])  # Use task ID as the item ID for direct lookup
                )
            
            # Adjust the status column width based on content
            # Calculate pixel width: approximate 8 pixels per character plus some padding
            if max_status_length > 0:
                # Use a multiplier for average character width (depends on font)
                char_width = 8  # Approximate width in pixels per character
                padding = 20    # Extra padding for the column
                new_width = max(150, min(300, max_status_length * char_width + padding))
                self.task_tree.column("status", width=new_width, minwidth=150)
            
            # Configure tag colors for status indicators with consistent bold font
            self.task_tree.tag_configure("completed", foreground="#9e9e9e", font=standard_font)  # Gray text
            self.task_tree.tag_configure("overdue", foreground="#d32f2f", font=standard_font)  # Red text
            self.task_tree.tag_configure("due_today", foreground="#f57c00", font=standard_font)  # Orange text
            self.task_tree.tag_configure("pending_soon", foreground="#7b1fa2", font=standard_font)  # Purple text
            self.task_tree.tag_configure("pending", foreground="#1976d2", font=standard_font)  # Blue text
            
            # Configure priority tags with consistent bold font
            self.task_tree.tag_configure("high", font=standard_font)
            self.task_tree.tag_configure("medium", font=standard_font)
            self.task_tree.tag_configure("low", font=standard_font)
            
            # Configure link tag with consistent bold font but still make it stand out
            self.task_tree.tag_configure("link", foreground="#0066cc", font=standard_font)
            self.task_tree.tag_configure("link_hover", foreground="#0099ff", font=standard_font)
            self.task_tree.tag_configure("link_click", foreground="#003366", font=standard_font)
            
            # Show "no results" message if no tasks match the filters
            if not filtered_tasks:
                self.task_tree.insert("", tk.END, values=("No tasks match your filters", "", "", "", "", "", "", "", ""), tags=["no_results"])
                self.task_tree.tag_configure("no_results", foreground="#9e9e9e", font=standard_font)
            
        except Exception as e:
            logging.error("Error during _perform_task_refresh", exc_info=True)
            # Re-raise to be caught in _perform_task_refresh_threaded - No, handle here and show error
            self.root.after(0, lambda: messagebox.showerror("Error", "Error refreshing task list. See log for details.", parent=self.root)) # Show error in main thread
        finally:
            # Hide loading indicator and update status are now handled in _perform_task_refresh_threaded
            pass # No changes needed here
    
    def update_category_filter(self):
        """Update the category filter dropdown with available categories."""
        categories = set(task["category"] for task in self.task_manager.tasks)
        categories = sorted(list(categories))
        
        # Add "All" option
        categories = ["All"] + categories
        
        # Update combobox values
        current = self.category_var.get()
        self.category_combo["values"] = categories
        
        # If current category is not in the list, reset to "All"
        if current not in categories:
            self.category_var.set("All")
    
    def get_selected_task_id(self):
        """Get the ID of the selected task (single selection)."""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showinfo("Information", "Please select a task first.")
            return None
        
        # Get the task ID directly from the item ID
        try:
            task_id = int(selection[0])
            return task_id
        except (ValueError, TypeError):
            # Fallback to the old method if the item ID is not a valid task ID
            item = selection[0]
            values = self.task_tree.item(item, "values")
            
            # Find the task with matching values
            for task in self.task_manager.tasks:
                # Match by title, rev, and applied_vessel
                if (task["title"] == values[1] and
                    task.get("rev", "") == values[2] and
                    task.get("applied_vessel", "") == values[3]):
                    return task["id"]
            
            messagebox.showerror("Error", "Could not identify the selected task.")
            return None
    
    def get_selected_task_ids(self):
        """Get the IDs of the selected tasks."""
        selection = self.task_tree.selection()
        if not selection:
            messagebox.showinfo("Information", "Please select at least one task.")
            return None
        
        # Get the IDs directly from the item IDs
        task_ids = []
        for item in selection:
            try:
                task_id = int(item)
                task_ids.append(task_id)
            except (ValueError, TypeError):
                # Fallback to the old method if the item ID is not a valid task ID
                values = self.task_tree.item(item, "values")
                
                # Find the task with matching values
                for task in self.task_manager.tasks:
                    # Match by title, rev, and applied_vessel
                    if (task["title"] == values[1] and
                        task.get("rev", "") == values[2] and
                        task.get("applied_vessel", "") == values[3]):
                        task_ids.append(task["id"])
                        break
        
        if not task_ids:
            messagebox.showerror("Error", "Could not identify the selected tasks.")
            return None
            
        return task_ids
    
    def add_task(self):
        """Open dialog to add a new task."""
        dialog = TaskDialog(self.root, "Add Task")
        if dialog.result:
            # Attempt to add the task using the TaskManager
            # TaskManager.add_task now returns the new ID on success, or None on failure
            new_task_id = self.task_manager.add_task(**dialog.result)

            if new_task_id is not None:
                # Success: Task was added to DB and local list
                self.refresh_task_list() # Refresh to show the new task
                self.set_status(f"Task '{dialog.result['title']}' (ID: {new_task_id}) added successfully.")
                self.update_category_filter() # Update filters if a new category was added
            else:
                # Failure: TaskManager.add_task already logged the error
                messagebox.showerror("Error Adding Task",
                                     "Failed to save the task to the database.\n"
                                     "This might be due to a temporary network issue or database contention.\n\n"
                                     "Please try adding the task again.",
                                     parent=self.root)
                # No refresh needed as the task wasn't added successfully
    
    def edit_task(self):
        """Edit the selected task."""

        task_id = self.get_selected_task_id()
        if not task_id:
            return
        
        # Get the task
        task = self.task_manager._find_task_by_id(task_id)
        if not task:
            messagebox.showerror("Error", f"Task with ID {task_id} not found.")
            return
        
        # Open dialog with task data
        dialog = TaskDialog(self.root, "Edit Task", task)
        if dialog.result:
            # Update the task
            self.task_manager.update_task(task_id, **dialog.result)
            
            # Refresh the task list to show the updated task
            self.refresh_task_list()
            
            # Show confirmation message
            self.set_status(f"Task {task_id} updated successfully.")
    
    def complete_task(self):
        """Mark the selected task(s) as completed or pending."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return
        
        # Check if all selected tasks have the same completion status
        tasks = [self.task_manager._find_task_by_id(task_id) for task_id in task_ids]
        all_completed = all(task["completed"] for task in tasks if task)
        all_pending = all(not task["completed"] for task in tasks if task)
        
        count = len(task_ids)
        
        # If all tasks have the same status, offer to toggle them
        if all_completed:
            message = f"Mark {count} selected task{'s' if count > 1 else ''} as pending?"
            if messagebox.askyesno("Confirm", message):
                self.batch_update_completion_status(task_ids, completed=False)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as pending.")
        elif all_pending:
            message = f"Mark {count} selected task{'s' if count > 1 else ''} as completed?"
            if messagebox.askyesno("Confirm", message):
                self.batch_update_completion_status(task_ids, completed=True)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as completed.")
        else:
            # Mixed status - ask what to do
            choice = messagebox.askyesnocancel("Mixed Status", 
                                             f"The {count} selected tasks have mixed completion statuses.\n\n"
                                             "What would you like to do?\n\n"
                                             "• Yes = Mark All as Completed\n"
                                             "• No = Mark All as Pending\n"
                                             "• Cancel = Do Nothing",
                                             icon=messagebox.QUESTION)
            
            if choice is True:  # Yes corresponds to "Mark All as Completed"
                self.batch_update_completion_status(task_ids, completed=True)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as completed.")
            elif choice is False:  # No corresponds to "Mark All as Pending"
                self.batch_update_completion_status(task_ids, completed=False)
                self.set_status(f"{count} task{'s' if count > 1 else ''} marked as pending.")
        
        self.refresh_task_list()
    
    def batch_update_completion_status(self, task_ids, completed):
        """Update completion status for multiple tasks at once."""
        for task_id in task_ids:
            self.task_manager.update_task(task_id, completed=completed)
        
        # No need to call refresh_task_list() here as it will be called by the calling method
    
    def delete_task(self):
        """Delete the selected task(s) using soft delete."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return

        count = len(task_ids)
        message = (f"Are you sure you want to delete {count} task{'s' if count > 1 else ''}?\n\n"
                   "Deleted tasks can be recovered later from the 'View Deleted Tasks' option.")

        if not messagebox.askyesno("Confirm Deletion", message, parent=self.root):
            return

        # For single task deletion:
        if count == 1:
            task_id_to_delete = task_ids[0]
            # Call the newly added delete_task method in TaskManager
            if self.task_manager.delete_task(task_id_to_delete):
                # Successfully soft-deleted
                self.refresh_task_list()
                self.set_status(f"Task {task_id_to_delete} marked as deleted.")
                # Enable the Undo button after deletion
                if hasattr(self, 'undo_btn'):
                    self.undo_btn.config(state=tk.NORMAL)
            else:
                # Deletion failed (error message likely shown by TaskManager)
                 self.set_status(f"Failed to delete task {task_id_to_delete}.")

        else:
            # For batch deletion:
            # Call the newly added batch_delete_tasks method
            succeeded, failed = self.task_manager.batch_delete_tasks(task_ids)
            if succeeded:
                self.refresh_task_list()
                self.set_status(f"Marked {len(succeeded)} task{'s' if len(succeeded) > 1 else ''} as deleted.")
                # Enable the Undo button after batch deletion
                if hasattr(self, 'undo_btn'):
                    self.undo_btn.config(state=tk.NORMAL)
            if failed:
                messagebox.showwarning("Partial Deletion",
                                     f"{len(failed)} task{'s' if len(failed) > 1 else ''} could not be deleted. Check logs for details.",
                                     parent=self.root)
            if not succeeded and not failed:
                 # This case might happen if all task_ids were invalid or a major DB error occurred
                 self.set_status(f"Failed to delete selected tasks.")
    
    def view_task_details(self, event):
        """View details of the selected task."""
        # If called from an event, get the task ID from the clicked item
        if event:
            item = self.task_tree.identify_row(event.y)
            if not item:
                return
            self.task_tree.selection_set(item)
        
        task_id = self.get_selected_task_id()
        if not task_id:
            return
        
        # Get the task - reload tasks first to ensure we have the latest data
        self.task_manager.reload_tasks()
        task = self.task_manager._find_task_by_id(task_id)
        if not task:
            messagebox.showerror("Error", f"Task with ID {task_id} not found.")
            return
        
        # Format dates for display
        start_date = task.get('date_started', 'Not specified')
        request_date = task.get('requested_date', 'Not specified')
        due_date = task.get('due_date', 'No due date')
        created_date = task.get('created_date', 'Unknown')
        last_modified = task.get('last_modified', 'Not modified')
        modified_by = task.get('modified_by', 'Not modified')
        
        # Show task details
        details = (
            f"Equipment Name: {task['title']}\n\n"
            f"Rev: {task.get('rev', 'Not specified')}\n\n"
            f"Description: {task['description'] or 'No description'}\n\n"
            f"Start: {start_date}\n\n"
            f"Request: {request_date}\n\n"
            f"Due: {due_date}\n\n"
            f"Category: {task['category']}\n\n"
            f"Created date: {created_date}\n\n"
            f"Last Modified: {last_modified}\n\n"
            f"Modified by: {modified_by}"
        )
        
        messagebox.showinfo(f"Task Details", details)
    
    def show_context_menu(self, event):
        """Show the context menu for tasks."""
        # Get the item under cursor
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        
        # Select the item if it's not already selected
        if item not in self.task_tree.selection():
            self.task_tree.selection_set(item)
        
        # Get the number of selected items
        selection_count = len(self.task_tree.selection())
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        
        # Task actions submenu
        task_menu = tk.Menu(context_menu, tearoff=0)
        if selection_count == 1:
            task_menu.add_command(label="View Details", command=lambda: self.view_task_details(None))
            task_menu.add_command(label="Edit Task", command=self.edit_task)
            task_menu.add_command(label="Toggle Complete/Pending", command=self.complete_task)
        else:
            task_menu.add_command(label=f"Toggle {selection_count} Tasks", command=self.complete_task)
        
        context_menu.add_cascade(label="Task Actions", menu=task_menu)
        
        # Add Assign To submenu
        assign_menu = tk.Menu(context_menu, tearoff=0)
        for staff in USERS[1:]:  # Skip "All" from the USERS list
            assign_menu.add_command(
                label=staff,
                command=lambda s=staff: self.assign_task_to(s)
            )
        context_menu.add_cascade(label="Assign To", menu=assign_menu)
        
        # Copy options submenu (only for single selection)
        if selection_count == 1:
            try:
                task_id = int(item)
                task = self.task_manager._find_task_by_id(task_id)
                if task:
                    # Create Copy submenu
                    copy_menu = tk.Menu(context_menu, tearoff=0)
                    has_copy_items = False
                    
                    # Add copy options for hidden fields
                    if task.get("drawing_no"):
                        copy_menu.add_command(
                            label="Copy Drawing No.", 
                            command=lambda: self.copy_to_clipboard(task.get("drawing_no", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("request_no"):
                        copy_menu.add_command(
                            label="Copy Request No.", 
                            command=lambda: self.copy_to_clipboard(task.get("request_no", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("link"):
                        copy_menu.add_command(
                            label="Copy Folder Link", 
                            command=lambda: self.copy_to_clipboard(task.get("link", ""))
                        )
                        has_copy_items = True
                    
                    if task.get("sdb_link"):
                        copy_menu.add_command(
                            label="Copy SDB Link", 
                            command=lambda: self.copy_to_clipboard(task.get("sdb_link", ""))
                        )
                        has_copy_items = True
                    
                    # Add dates submenu if any dates are present
                    dates = {
                        "Requested Date": task.get("requested_date"),
                        "Date Started": task.get("date_started"),
                        "Due Date": task.get("due_date")
                    }
                    
                    if any(dates.values()):
                        dates_menu = tk.Menu(copy_menu, tearoff=0)
                        for label, date in dates.items():
                            if date:
                                dates_menu.add_command(
                                    label=f"Copy {label}", 
                                    command=lambda d=date: self.copy_to_clipboard(d)
                                )
                        copy_menu.add_cascade(label="Copy Dates", menu=dates_menu)
                        has_copy_items = True
                    
                    # Only add the Copy menu if there are items in it
                    if has_copy_items:
                        context_menu.add_cascade(label="Copy", menu=copy_menu)
                    
                    # Create Open submenu
                    open_menu = tk.Menu(context_menu, tearoff=0)
                    has_open_items = False
                    
                    if task.get("link"):
                        open_menu.add_command(
                            label="Open Folder Link", 
                            command=lambda: self.open_link(task.get("link", ""))
                        )
                        has_open_items = True
                    
                    if task.get("sdb_link"):
                        open_menu.add_command(
                            label="Open SDB Link", 
                            command=lambda: self.open_link(task.get("sdb_link", ""))
                        )
                        has_open_items = True
                    
                    # Only add the Open menu if there are items in it
                    if has_open_items:
                        context_menu.add_cascade(label="Open", menu=open_menu)
                    
            except (ValueError, TypeError):
                # If we can't get the task ID directly, try to find the task by values
                values = self.task_tree.item(item, "values")
                if len(values) >= 4:  # Make sure we have enough values
                    for task in self.task_manager.tasks:
                        # Match by title, rev, and applied_vessel
                        if (task["title"] == values[1] and
                            task.get("rev", "") == values[2] and
                            task.get("applied_vessel", "") == values[3]):
                            
                            # Create Copy submenu
                            copy_menu = tk.Menu(context_menu, tearoff=0)
                            has_copy_items = False
                            
                            # Add copy options for hidden fields
                            if task.get("drawing_no"):
                                copy_menu.add_command(
                                    label="Copy Drawing No.", 
                                    command=lambda: self.copy_to_clipboard(task.get("drawing_no", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("request_no"):
                                copy_menu.add_command(
                                    label="Copy Request No.", 
                                    command=lambda: self.copy_to_clipboard(task.get("request_no", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("link"):
                                copy_menu.add_command(
                                    label="Copy Folder Link", 
                                    command=lambda: self.copy_to_clipboard(task.get("link", ""))
                                )
                                has_copy_items = True
                            
                            if task.get("sdb_link"):
                                copy_menu.add_command(
                                    label="Copy SDB Link", 
                                    command=lambda: self.copy_to_clipboard(task.get("sdb_link", ""))
                                )
                                has_copy_items = True
                            
                            # Add dates submenu if any dates are present
                            dates = {
                                "Requested Date": task.get("requested_date"),
                                "Date Started": task.get("date_started"),
                                "Due Date": task.get("due_date")
                            }
                            
                            if any(dates.values()):
                                dates_menu = tk.Menu(copy_menu, tearoff=0)
                                for label, date in dates.items():
                                    if date:
                                        dates_menu.add_command(
                                            label=f"Copy {label}", 
                                            command=lambda d=date: self.copy_to_clipboard(d)
                                        )
                                copy_menu.add_cascade(label="Copy Dates", menu=dates_menu)
                                has_copy_items = True
                            
                            # Only add the Copy menu if there are items in it
                            if has_copy_items:
                                context_menu.add_cascade(label="Copy", menu=copy_menu)
                            
                            # Create Open submenu
                            open_menu = tk.Menu(context_menu, tearoff=0)
                            has_open_items = False
                            
                            if task.get("link"):
                                open_menu.add_command(
                                    label="Open Folder Link", 
                                    command=lambda: self.open_link(task.get("link", ""))
                                )
                                has_open_items = True
                            
                            if task.get("sdb_link"):
                                open_menu.add_command(
                                    label="Open SDB Link", 
                                    command=lambda: self.open_link(task.get("sdb_link", ""))
                                )
                                has_open_items = True
                            
                            # Only add the Open menu if there are items in it
                            if has_open_items:
                                context_menu.add_cascade(label="Open", menu=open_menu)
                            
                            break
        
        context_menu.add_separator()
        
        # Delete option shows count of items to be deleted
        if selection_count == 1:
            context_menu.add_command(label="Delete Task", command=self.delete_task)
        else:
            context_menu.add_command(label=f"Delete {selection_count} Tasks", command=self.delete_task)
        
        # Add a 'View Deleted Tasks' option
        context_menu.add_separator()
        context_menu.add_command(label="View Deleted Tasks", command=self.view_deleted_tasks)
        
        # Display the context menu
        context_menu.tk_popup(event.x_root, event.y_root)
    
    def check_link_click(self, event):
        """Check if a link cell was clicked and open the link if so."""
        region = self.task_tree.identify_region(event.x, event.y)
        if region == "cell":
            # Get the item and column that was clicked
            item = self.task_tree.identify_row(event.y)
            column = self.task_tree.identify_column(event.x)
            
            # Check if the link column was clicked (column #8)
            if column == "#8":  # Link column
                # Get the values from the display
                values = self.task_tree.item(item, "values")
                link_display = values[7] if len(values) > 7 else ""
                
                # Only proceed if it contains a link indicator
                if link_display == "🔗":
                    # Get the actual task to find the real link
                    try:
                        task_id = int(item)
                        task = self.task_manager._find_task_by_id(task_id)
                        if task and task.get("link"):
                            # Apply click effect
                            tags = list(self.task_tree.item(item, "tags"))
                            if "link" in tags:
                                tags.remove("link")
                            if "link_hover" in tags:
                                tags.remove("link_hover")
                            tags.append("link_click")
                            self.task_tree.item(item, tags=tags)
                            
                            # Schedule to revert the effect after a short delay
                            self.root.after(150, lambda i=item: self.revert_link_click_effect(i))
                            
                            # Open the link
                            self.open_link(task["link"])
                    except (ValueError, TypeError):
                        # If we can't get the task ID directly, try to find the task
                        for task in self.task_manager.tasks:
                            # Check if this is the right task by matching other values
                            if (task["title"] == values[1] and
                                task.get("rev", "") == values[2] and
                                task.get("applied_vessel", "") == values[3]):
                                if task.get("link"):
                                    # Open the link
                                    self.open_link(task["link"])
                                break
    
    def revert_link_click_effect(self, item):
        """Revert the link click effect after a short delay."""
        if item in self.task_tree.get_children():
            tags = list(self.task_tree.item(item, "tags"))
            if "link_click" in tags:
                tags.remove("link_click")
            if "link" not in tags:
                tags.append("link")
            self.task_tree.item(item, tags=tags)
    
    def open_link(self, link):
        """Open a link which can be a URL, file path, or network path."""
        
        if not link:
            messagebox.showerror("Error", "No link provided.")
            return
            
        try:
            # Check if it's a network path or local file path
            if link.startswith("\\\\") or os.path.exists(link) or link.startswith(("C:", "D:", "E:", "F:", "\\")):
                # It's a file path - use the appropriate command based on OS
                if os.name == 'nt':  # Windows
                    # For network paths, ensure they're properly formatted
                    if link.startswith("//"):
                        # Convert forward slashes to backslashes for Windows
                        link = link.replace("/", "\\")
                    
                    # Check if it's a directory
                    if os.path.isdir(link) or (link.startswith("\\\\") and not os.path.splitext(link)[1]):
                        # Open directory in Explorer
                        subprocess.run(['explorer', link], shell=True)
                    else:
                        # Open file with default application
                        os.startfile(link)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['xdg-open', link] if os.name == 'posix' else ['open', link])
            else:
                # It's a web URL - add http:// if needed
                if not link.startswith(("http://", "https://", "www.")):
                    link = f"http://{link}"
                webbrowser.open(link)
                
            self.set_status(f"Opening: {link}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not open link: {link}\n\nError: {str(e)}")
            # Log more detailed error information
            print(f"Error opening link: {link}")
            print(f"Error details: {str(e)}")
            print(f"Error type: {type(e)}")
    
    def set_status(self, message):
        """Set a temporary status message."""
        self.status_var.set(message)
        # Schedule to reset after 5 seconds
        self.root.after(5000, self.update_status)
    
    def on_search_change(self, *args):
        """Handle search text changes with debouncing."""
        # Cancel any existing search timer
        if hasattr(self, 'search_timer'):
            self.root.after_cancel(self.search_timer)
        
        # Set a new timer to refresh after 300ms of no typing
        self.search_timer = self.root.after(300, self.refresh_task_list)
    
    def clear_search(self):
        """Clear the search field and refresh the task list."""
        self.search_var.set("")
        self.refresh_task_list()
    
    def _task_matches_search(self, task, search_term):
        """Check if a task matches the search term across all relevant fields."""
        search_term = search_term.lower().strip()
        
        # Fields to search in (convert to lowercase for case-insensitive search)
        searchable_fields = [
            task.get("title", ""),
            task.get("description", ""),
            task.get("applied_vessel", ""),
            task.get("category", ""),
            task.get("main_staff", ""),
            task.get("assigned_to", ""),
            task.get("rev", ""),
            task.get("drawing_no", ""),
            task.get("request_no", ""),
            task.get("link", ""),
            task.get("sdb_link", ""),
            str(task.get("id", "")),  # Convert ID to string for searching
            str(task.get("qtd_mhr", "")),  # Convert to string for searching
            str(task.get("actual_mhr", "")),  # Convert to string for searching
        ]
        
        # Check if search term matches any field
        for field in searchable_fields:
            if field and search_term in str(field).lower():
                return True
        
        # Also check date fields if they exist
        date_fields = ["due_date", "requested_date", "date_started"]
        for date_field in date_fields:
            date_value = task.get(date_field)
            if date_value and search_term in str(date_value).lower():
                return True
        
        return False
    
    def update_cursor(self, event):
        """Update cursor based on whether it's over a link."""
        region = self.task_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.task_tree.identify_column(event.x)
            item = self.task_tree.identify_row(event.y)
            
            # Reset all link items to normal link style
            for all_item in self.task_tree.get_children():
                tags = list(self.task_tree.item(all_item, "tags"))
                if "link_hover" in tags:
                    tags.remove("link_hover")
                    if "link" not in tags:
                        tags.append("link")
                    self.task_tree.item(all_item, tags=tags)
            
            # Check if over the link column (column #8)
            if column == "#8" and item:
                # Get the values from the display
                values = self.task_tree.item(item, "values")
                link_display = values[7] if len(values) > 7 else ""
                
                # Only change cursor if it contains a link indicator
                if link_display == "🔗":
                    # Get the actual task to confirm it has a link
                    try:
                        task_id = int(item)
                        task = self.task_manager._find_task_by_id(task_id)
                        if task and task.get("link"):
                            self.task_tree.config(cursor="hand2")
                            
                            # Apply hover effect to this link
                            tags = list(self.task_tree.item(item, "tags"))
                            if "link" in tags:
                                tags.remove("link")
                            if "link_hover" not in tags:
                                tags.append("link_hover")
                            self.task_tree.item(item, tags=tags)
                            return
                    except (ValueError, TypeError):
                        # Try to find the task by values
                        for task in self.task_manager.tasks:
                            if (task["title"] == values[1] and
                                task.get("rev", "") == values[2] and
                                task.get("applied_vessel", "") == values[3]):
                                if task.get("link"):
                                    self.task_tree.config(cursor="hand2")
                                    
                                    # Apply hover effect to this link
                                    tags = list(self.task_tree.item(item, "tags"))
                                    if "link" in tags:
                                        tags.remove("link")
                                    if "link_hover" not in tags:
                                        tags.append("link_hover")
                                    self.task_tree.item(item, tags=tags)
                                    return
                                break
            
            # Reset cursor if not over a link
            self.task_tree.config(cursor="")
    
    def copy_to_clipboard(self, text):
        """Copy text to clipboard."""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.set_status(f"Copied to clipboard: {text}")
    
    def sort_treeview(self, column, reset=True):
        """
        Sort treeview by a column.
        
        Args:
            column: Column to sort by
            reset: Whether to reset the sort direction if clicking the same column
        """
        # Update sort indicators in column headings
        for col in self.task_tree["columns"]:
            # Remove any existing sort indicators
            heading_text = self.task_tree.heading(col)["text"]
            heading_text = heading_text.replace(" ▲", "").replace(" ▼", "")
            self.task_tree.heading(col, text=heading_text)
        
        # Determine sort direction
        if self.sort_column == column and reset:
            # Toggle sort direction if clicking the same column
            self.sort_reverse = not self.sort_reverse
        else:
            # Default to ascending for a new column
            self.sort_column = column
            self.sort_reverse = False
        
        # Add sort indicator to column heading
        heading_text = self.task_tree.heading(column)["text"]
        indicator = " ▲" if not self.sort_reverse else " ▼"
        self.task_tree.heading(column, text=heading_text + indicator)
        
        # Refresh the task list to apply the sorting
        self.refresh_task_list()
    
    def apply_sorting(self, tasks):
        """
        Apply sorting to the task list.
        
        Args:
            tasks: List of tasks to sort
        """
        # Define custom key functions for special columns
        def priority_key(task):
            # Sort by priority level (high, medium, low)
            priority_order = {"high": 0, "medium": 1, "low": 2}
            return priority_order.get(task.get("priority", "").lower(), 3)
        
        def status_key(task):
            # Sort by completion status and due date
            if task.get("completed", False):
                return (2, "")  # Completed tasks at the bottom
            
            if not task.get("due_date"):
                return (1, "")  # Tasks without due date in the middle
            
            # Tasks with due date sorted by days until due
            try:
                due_date = datetime.datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                today = datetime.date.today()
                days_until_due = (due_date - today).days
                
                # Return a tuple for proper sorting
                if days_until_due < 0:
                    return (0, f"{abs(days_until_due):04d}")  # Overdue tasks first
                else:
                    return (0, f"{days_until_due+10000:04d}")  # Then by days until due
            except ValueError:
                return (1, "")  # Invalid date format in the middle
        
        # Sort the tasks based on the selected column
        if self.sort_column == "status":
            tasks.sort(key=status_key, reverse=self.sort_reverse)
        elif self.sort_column == "priority":
            tasks.sort(key=priority_key, reverse=self.sort_reverse)
        elif self.sort_column == "rev":
            # Special handling for revision numbers (numeric sorting when possible)
            def rev_key(task):
                rev = task.get("rev", "")
                # Try to extract numeric part for numeric sorting
                try:
                    # If it's a pure number, convert to int for proper sorting
                    if rev.isdigit():
                        return int(rev)
                    # If it has a format like "Rev 1" or "R2", extract the number
                    import re
                    match = re.search(r'(\d+)', rev)
                    if match:
                        return int(match.group(1))
                except:
                    pass
                # Fall back to string sorting
                return rev.lower()
            
            tasks.sort(key=rev_key, reverse=self.sort_reverse)
        else:
            # For other columns, sort by the column value
            # Use a case-insensitive sort for text columns
            tasks.sort(
                key=lambda task: str(task.get(self.sort_column, "")).lower(),
                reverse=self.sort_reverse
            )

    def assign_task_to(self, staff_name):
        """Assign the selected task(s) to a staff member."""
        task_ids = self.get_selected_task_ids()
        if not task_ids:
            return
            
        count = len(task_ids)
        # Update each selected task
        for task_id in task_ids:
            self.task_manager.update_task(task_id, assigned_to=staff_name)
        
        # Refresh the task list and show confirmation
        self.refresh_task_list()
        self.set_status(f"{count} task{'s' if count > 1 else ''} assigned to {staff_name}")
        
        # Disable the Undo button after any other action
        if hasattr(self, 'undo_btn'):
            self.undo_btn.config(state=tk.DISABLED)

    def undo_delete(self):
        """Undo the most recent deletion (which might be multiple tasks)"""
        if not self.task_manager.deleted_tasks_stack:
            self.set_status("Nothing to undo")
            return
        
        # Get the most recent batch of deleted tasks
        recent_deletion = self.task_manager.deleted_tasks_stack.pop()
        
        # Recover all tasks in this batch
        recovered_ids = self.task_manager.recover_batch(recent_deletion)
        
        # Update the UI
        if recovered_ids:
            self.refresh_task_list()
            if len(recovered_ids) == 1:
                self.set_status(f"Restored task {recovered_ids[0]}")
            else:
                self.set_status(f"Restored {len(recovered_ids)} tasks")
        else:
            self.set_status("Failed to restore tasks")
        
        # Disable the Undo button if no more items in the stack
        if hasattr(self, 'undo_btn'):
            if not self.task_manager.deleted_tasks_stack:
                self.undo_btn.config(state=tk.DISABLED)
            else:
                self.undo_btn.config(state=tk.NORMAL)
    
    def position_window_center(self):
        """Positions the application window in the center of the screen."""
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.root.geometry(f"+{x}+{y}")

    def bind_mousewheel_to_treeview(self):
        """Binds mousewheel scrolling to the task treeview for different platforms."""
        if platform.system() == 'Windows':
            self.task_tree.bind("<MouseWheel>", self.on_mousewheel_windows)
        elif platform.system() == 'Darwin':  # macOS
            self.task_tree.bind("<MouseWheel>", self.on_mousewheel_macos)
        elif platform.system() == 'Linux':
            self.task_tree.bind("<Button-4>", self.on_mousewheel_linux)
            self.task_tree.bind("<Button-5>", self.on_mousewheel_linux)

    def on_mousewheel_windows(self, event):
        """Handles mousewheel scrolling on Windows."""
        self.task_tree.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_mousewheel_macos(self, event):
        """Handles mousewheel scrolling on macOS."""
        self.task_tree.yview_scroll(event.delta, "units")

    def on_mousewheel_linux(self, event):
        """Handles mousewheel scrolling on Linux."""
        if event.num == 4:
            self.task_tree.yview_scroll(-1, "units")
        elif event.num == 5:
            self.task_tree.yview_scroll(1, "units")

    def run(self): # ADD THIS run METHOD
        """Starts the Tkinter main event loop."""
        self.root.mainloop()

    def view_deleted_tasks(self):
        """Show a dialog with all deleted tasks for recovery or permanent deletion."""
        # Get all deleted tasks
        deleted_tasks = self.task_manager.get_filtered_tasks(show_completed=True, show_deleted=True)
        deleted_tasks = [t for t in deleted_tasks if t.get("deleted", False)]
        
        if not deleted_tasks:
            messagebox.showinfo("Deleted Tasks", "No deleted tasks found")
            return
            
        # Create a new toplevel window
        deleted_window = tk.Toplevel(self.root)
        deleted_window.title("Deleted Tasks")
        deleted_window.geometry("800x500")
        deleted_window.transient(self.root)
        deleted_window.grab_set()
        
        # Create main frame
        main_frame = ttk.Frame(deleted_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add info label
        ttk.Label(main_frame, 
                  text="These tasks have been deleted but can be recovered. Select tasks and use the buttons below.",
                  wraplength=780).pack(pady=(0, 10))
        
        # Create treeview for deleted tasks
        columns = ("id", "title", "due_date", "priority", "category", "assigned_to")
        tree = ttk.Treeview(main_frame, columns=columns, show='headings')
        
        # Set column headings
        tree.heading('id', text='ID')
        tree.heading('title', text='Title')
        tree.heading('due_date', text='Due Date')
        tree.heading('priority', text='Priority')
        tree.heading('category', text='Category')
        tree.heading('assigned_to', text='Assigned To')
        
        # Set column widths
        tree.column('id', width=50, anchor=tk.CENTER)
        tree.column('title', width=250)
        tree.column('due_date', width=100, anchor=tk.CENTER)
        tree.column('priority', width=80, anchor=tk.CENTER)
        tree.column('category', width=100)
        tree.column('assigned_to', width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add button frame
        button_frame = ttk.Frame(deleted_window, padding=10)
        button_frame.pack(fill=tk.X)
        
        # Populate the tree
        for task in deleted_tasks:
            tree.insert('', tk.END, values=(
                task['id'],
                task['title'],
                task.get('due_date', ''),
                task.get('priority', ''),
                task.get('category', ''),
                task.get('assigned_to', '')
            ))
        
        # Function to recover selected tasks
        def recover_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showinfo("Recover Task", "Please select at least one task to recover")
                return
                
            recovered_count = 0
            for item in selected_items:
                task_id = int(tree.item(item, 'values')[0])
                if self.task_manager.recover_task(task_id):
                    recovered_count += 1
                    
            if recovered_count > 0:
                self.refresh_task_list()
                self.set_status(f"Recovered {recovered_count} task(s)")
                # Refresh the deleted tasks window
                for item in tree.get_children():
                    tree.delete(item)
                    
                # Re-populate with remaining deleted tasks
                remaining_deleted = [t for t in self.task_manager.get_filtered_tasks(
                    show_completed=True, show_deleted=True) if t.get("deleted", False)]
                    
                for task in remaining_deleted:
                    tree.insert('', tk.END, values=(
                        task['id'],
                        task['title'],
                        task.get('due_date', ''),
                        task.get('priority', ''),
                        task.get('category', ''),
                        task.get('assigned_to', '')
                    ))
                    
                if not remaining_deleted:
                    deleted_window.destroy()
        
        # Function to permanently delete selected tasks
        def permanently_delete_selected():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showinfo("Delete Task", "Please select at least one task to permanently delete")
                return
                
            task_count = len(selected_items)
            confirm = messagebox.askyesno(
                "Confirm Permanent Deletion",
                f"Are you sure you want to permanently delete {task_count} task(s)?\n\n"
                "This action cannot be undone.",
                icon=messagebox.WARNING
            )
            
            if not confirm:
                return
                
            deleted_count = 0
            for item in selected_items:
                task_id = int(tree.item(item, 'values')[0])
                if self.task_manager.permanently_delete_task(task_id):
                    deleted_count += 1
                    
            if deleted_count > 0:
                self.set_status(f"Permanently deleted {deleted_count} task(s)")
                # Refresh the deleted tasks window
                for item in tree.get_children():
                    tree.delete(item)
                    
                # Re-populate with remaining deleted tasks
                remaining_deleted = [t for t in self.task_manager.get_filtered_tasks(
                    show_completed=True, show_deleted=True) if t.get("deleted", False)]
                    
                for task in remaining_deleted:
                    tree.insert('', tk.END, values=(
                        task['id'],
                        task['title'],
                        task.get('due_date', ''),
                        task.get('priority', ''),
                        task.get('category', ''),
                        task.get('assigned_to', '')
                    ))
                    
                if not remaining_deleted:
                    deleted_window.destroy()
        
        # Function to recover all tasks
        def recover_all():
            confirm = messagebox.askyesno(
                "Confirm Recovery",
                "Are you sure you want to recover all deleted tasks?",
                icon=messagebox.QUESTION
            )
            
            if not confirm:
                return
                
            recovered = self.task_manager.recover_all_deleted_tasks()
            if recovered > 0:
                self.refresh_task_list()
                self.set_status(f"Recovered all {recovered} deleted task(s)")
                deleted_window.destroy()
        
        # Function to permanently delete all tasks
        def permanently_delete_all():
            confirm = messagebox.askyesno(
                "Confirm Permanent Deletion",
                "Are you sure you want to PERMANENTLY DELETE ALL deleted tasks?\n\n"
                "This action cannot be undone.",
                icon=messagebox.WARNING
            )
            
            if not confirm:
                return
                
            deleted = self.task_manager.permanently_delete_all_deleted_tasks()
            if deleted > 0:
                self.set_status(f"Permanently deleted all {deleted} task(s)")
                deleted_window.destroy()
        
        # Add buttons
        ttk.Button(button_frame, text="Recover Selected", command=recover_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Permanently Delete Selected", command=permanently_delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Recover All", command=recover_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Permanently Delete All", command=permanently_delete_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=deleted_window.destroy).pack(side=tk.RIGHT, padx=5)


class TaskDialog:
    """Dialog for adding or editing a task."""
    
    def __init__(self, parent, title, task=None):
        """Initialize the dialog."""
        self.result = None
        self.parent = parent
        
        # Create dialog window but keep it hidden until fully built
        self.dialog = tk.Toplevel(parent)
        self.dialog.withdraw()  # Hide the window during setup
        self.dialog.title(title)
        
        # Set fixed dialog size - ADJUSTED HEIGHT FOR SCRAPE FIELD
        self.dialog_width = 630
        self.dialog_height = 670 # Increased height
        self.dialog.geometry(f"{self.dialog_width}x{self.dialog_height}")
        self.dialog.minsize(self.dialog_width, self.dialog_height)
        self.dialog.maxsize(self.dialog_width, self.dialog_height)  # Prevent resizing
        self.dialog.transient(parent)
        
        # Create main container with fixed padding
        main_container = ttk.Frame(self.dialog)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create canvas with scrollbar - use exact width calculation
        canvas_width = self.dialog_width - 40  # Account for dialog padding (20px on each side)
        self.canvas = tk.Canvas(main_container, width=canvas_width, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.canvas.yview) # Define scrollbar here
        
        # --- ADD FIELD MAPPING ---
        # Copy the field mapping from smartdb_login.py for easy access
        self.FIELDS_TO_EXTRACT = {
            "Request No": "Application_number",
            "Requested Date": "id_shinseibi",
            "Deadline": "id_yoteibi",
            "Revision No": "REV_No",
            "Applied Vessel": "Ship_No",
            "Drawing No": "DWG_No",
            "Description": "Drawing_Name_Common_1", # Often used for 'Equipment Name'
            "Note": "id_bikouran",
            "Modeling Staff": "acModeler", # Could map to Main Staff or Assigned To
            # We also need the file path if downloaded
            "Downloaded File Path": "Downloaded File Path" # Key added by login_to_smartdb
        }
        # Define which dialog field variable corresponds to which scraped field name
        # --- MOVED THIS BLOCK AFTER create_form CALL ---
        # self.FIELD_VAR_MAP = {
        #     "Request No": self.request_no_var,
        #     "Equipment Name": self.title_var, # Mapping 'Description' from scrape to 'Equipment Name' dialog field
        #     "Applied Vessel": self.applied_vessel_var,
        #     "Revision No": self.rev_var,
        #     "Drawing No": self.drawing_no_var,
        #     # Dates need special handling (parsing and setting DateEntry)
        #     "Requested Date": self.requested_date_picker,
        #     "Due Date": self.due_date_picker, # Mapping 'Deadline' from scrape
        #     # Links
        #     "Link": self.link_var, # Mapping 'Downloaded File Path' to 'Link' dialog field
        #     # Text Area
        #     "Description": self.description_text, # Mapping 'Note' from scrape to 'Description' dialog field
        #     # Comboboxes
        #     "Main Staff": self.main_staff_var, # Mapping 'Modeling Staff' from scrape
        #     "Assigned To": self.assigned_to_var # Can also map 'Modeling Staff' here if desired
        # }
        # --- END FIELD MAPPING ---
        
        # Create a frame inside the canvas for the form - use exact width
        form_width = canvas_width - 20  # Account for scrollbar width and some padding
        self.form_frame = ttk.Frame(self.canvas, width=form_width)
        
        # Pack scrollbar and canvas AFTER they are both defined
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure canvas - use exact coordinates to prevent shifting
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.form_frame, anchor=tk.NW, width=form_width)
        
        # Create form elements
        self.create_form(self.form_frame, task)
        
        # --- DEFINE FIELD MAPPING AFTER WIDGETS ARE CREATED ---
        self.FIELD_VAR_MAP = {
            "Request No": self.request_no_var,
            "Equipment Name": self.title_var, # Mapping 'Description' from scrape to 'Equipment Name' dialog field
            "Applied Vessel": self.applied_vessel_var,
            "Revision No": self.rev_var,
            "Drawing No": self.drawing_no_var,
            # Dates need special handling (parsing and setting DateEntry)
            "Requested Date": self.requested_date_picker,
            "Due Date": self.due_date_picker, # Mapping 'Deadline' from scrape
            # Links
            "Link": self.link_var, # Mapping 'Downloaded File Path' to 'Link' dialog field
            # Text Area
            "Description": self.description_text, # Mapping 'Note' from scrape to 'Description' dialog field
            # Comboboxes
            "Main Staff": self.main_staff_var, # Mapping 'Modeling Staff' from scrape
            "Assigned To": self.assigned_to_var # Can also map 'Modeling Staff' here if desired
        }
        # --- END FIELD MAPPING DEFINITION ---
        
        # Create button frame at the bottom (outside the scrollable area)
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save).pack(side=tk.RIGHT, padx=5)
        
        # Bind canvas configuration and mouse wheel
        self.form_frame.bind('<Configure>', self.on_frame_configure)
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Use a more reliable way to bind the mousewheel
        self.bind_mousewheel()
        
        # Bind dialog close event to ensure proper cleanup
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # Position the dialog
        self.position_dialog()
        
        # Pre-configure the canvas scroll region before showing the dialog
        self.dialog.update_idletasks()
        self.on_frame_configure()
        
        # Now show the dialog
        self.dialog.deiconify()
        
        # Grab focus and wait
        self.dialog.grab_set()
        self.dialog.focus_force()
        self.dialog.wait_window()
    
    def bind_mousewheel(self):
        """Bind mousewheel events to the canvas."""
        # Windows and macOS use different events and deltas
        if sys.platform.startswith('win'):
            self.canvas.bind_all("<MouseWheel>", self.on_mousewheel_windows)
        elif sys.platform.startswith('darwin'):
            self.canvas.bind_all("<MouseWheel>", self.on_mousewheel_macos)
        else:
            # Linux
            self.canvas.bind_all("<Button-4>", self.on_mousewheel_linux)
            self.canvas.bind_all("<Button-5>", self.on_mousewheel_linux)
    
    def unbind_mousewheel(self):
        """Unbind mousewheel events when dialog is closed."""
        try:
            # Unbind all mousewheel events
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        except:
            pass  # Ignore errors if the canvas is already destroyed
    
    def on_mousewheel_windows(self, event):
        """Handle mousewheel scrolling on Windows."""
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_mousewheel_macos(self, event):
        """Handle mousewheel scrolling on macOS."""
        try:
            self.canvas.yview_scroll(int(-1 * event.delta), "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_mousewheel_linux(self, event):
        """Handle mousewheel scrolling on Linux."""
        try:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_frame_configure(self, event=None):
        """Reset the scroll region to encompass the inner frame"""
        try:
            # Set the scroll region to the entire canvas
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            # Ensure the form frame stays at the left edge
            self.canvas.coords(self.canvas_frame, 0, 0)
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def on_canvas_configure(self, event):
        """When the canvas is resized, resize the inner frame to match"""
        try:
            # Maintain a fixed width for the inner frame
            width = event.width
            self.canvas.itemconfig(self.canvas_frame, width=width)
            
            # Ensure the form frame stays at the left edge
            self.canvas.coords(self.canvas_frame, 0, 0)
        except:
            pass  # Ignore errors if the canvas is destroyed
    
    def cancel(self):
        """Cancel the dialog and clean up resources."""
        # Unbind mousewheel events before destroying the dialog
        self.unbind_mousewheel()
        self.dialog.destroy()
    
    def position_dialog(self):
        """Position the dialog on the same monitor as the parent window."""
        # Update the dialog's geometry
        self.dialog.update_idletasks()
        
        # Get parent window geometry
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # Calculate position
        x = parent_x + (parent_width - self.dialog_width) // 2
        y = parent_y + (parent_height - self.dialog_height) // 2
        
        # Set the position
        self.dialog.geometry(f"+{x}+{y}")
    
    def create_form(self, parent, task=None):
        """Create the form fields in a single frame."""
        # Configure grid
        parent.columnconfigure(0, weight=0, minsize=120)  # Label column - fixed width
        parent.columnconfigure(1, weight=1)  # Entry column - expandable
        parent.columnconfigure(2, weight=0, minsize=20)  # Button column - fixed width
        
        # Request No. and Equipment Name
        row = 0
        # --- ADD SCRAPE ROW ---
        ttk.Label(parent, text="Scrape URL:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        scrape_frame = ttk.Frame(parent)
        scrape_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)
        scrape_frame.columnconfigure(0, weight=1) # Make entry expand

        self.scrape_url_var = tk.StringVar() # Variable for scrape URL
        scrape_entry = ttk.Entry(scrape_frame, textvariable=self.scrape_url_var)
        scrape_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Add Scrape button
        scrape_btn = ttk.Button(scrape_frame, text="Scrape", command=self.scrape_data, style="BrowseButton.TButton") # Reuse browse button style for size consistency
        scrape_btn.pack(side=tk.RIGHT, fill=tk.Y)
        
        # --- END SCRAPE ROW ---
        
        row += 1 # Increment row index for next element
        ttk.Label(parent, text="Request No.:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.request_no_var = tk.StringVar(value=task.get("request_no", "") if task else "")
        ttk.Entry(parent, textvariable=self.request_no_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Equipment Name:*").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.title_var = tk.StringVar(value=task["title"] if task else "")
        ttk.Entry(parent, textvariable=self.title_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Vessel Information
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Applied Vessel:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.applied_vessel_var = tk.StringVar(value=task.get("applied_vessel", "") if task else "")
        ttk.Entry(parent, textvariable=self.applied_vessel_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Rev:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.rev_var = tk.StringVar(value=task.get("rev", "") if task else "")
        ttk.Entry(parent, textvariable=self.rev_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        row += 1
        ttk.Label(parent, text="Drawing No.:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        self.drawing_no_var = tk.StringVar(value=task.get("drawing_no", "") if task else "")
        ttk.Entry(parent, textvariable=self.drawing_no_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Dates
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        # Create a frame for dates that spans both columns and uses grid
        date_frame = ttk.Frame(parent)
        date_frame.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=5)
        date_frame.columnconfigure(1, weight=1)  # Make the date entries expand
        
        # Configure date frame columns for proper alignment
        date_frame.columnconfigure(0, minsize=120)  # Same as label column
        date_frame.columnconfigure(1, weight=1)     # First date picker
        date_frame.columnconfigure(2, minsize=80)   # Second label
        date_frame.columnconfigure(3, weight=1)     # Second date picker
        date_frame.columnconfigure(4, minsize=80)   # Third label
        date_frame.columnconfigure(5, weight=1)     # Third date picker
        
        # Add date pickers with proper alignment
        ttk.Label(date_frame, text="Requested Date:").grid(row=0, column=0, sticky=tk.W)
        self.requested_date_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.requested_date_picker.grid(row=0, column=1, sticky=tk.EW, padx=(0,10))
        
        ttk.Label(date_frame, text="Started date:").grid(row=0, column=2, sticky=tk.W)
        self.date_started_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.date_started_picker.grid(row=0, column=3, sticky=tk.EW, padx=(0,10))
        
        ttk.Label(date_frame, text="Due Date:").grid(row=0, column=4, sticky=tk.W)
        self.due_date_picker = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.due_date_picker.grid(row=0, column=5, sticky=tk.EW)
        
        # Set existing dates if available
        if task:
            for picker, date_key in [(self.requested_date_picker, "requested_date"),
                                   (self.date_started_picker, "date_started"),
                                   (self.due_date_picker, "due_date")]:
                if task.get(date_key):
                    try:
                        date_value = datetime.datetime.strptime(task[date_key], "%Y-%m-%d").date()
                        picker.set_date(date_value)
                    except ValueError:
                        pass
        
        # Links
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Link:").grid(row=row, column=0, sticky=tk.W, pady=5)
        
        # Create a frame for the link field and browse button
        link_frame = ttk.Frame(parent)
        link_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)
        link_frame.columnconfigure(0, weight=1)  # Make the entry expand
        
        self.link_var = tk.StringVar(value=task.get("link", "") if task else "")
        
        # Create a custom style for the browse button to match entry height
        # FIX: Use the style from the parent window instead of the dialog
        style = ttk.Style()
        style.configure("BrowseButton.TButton", padding=0)
        
        # Create the entry
        link_entry = ttk.Entry(link_frame, textvariable=self.link_var)
        link_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Add browse button - using pack instead of grid for better height matching
        browse_btn = ttk.Button(link_frame, text="Browse", command=self.browse_folder, style="BrowseButton.TButton")
        browse_btn.pack(side=tk.RIGHT, fill=tk.Y)
        
        row += 1
        ttk.Label(parent, text="SDB Link:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.sdb_link_var = tk.StringVar(value=task.get("sdb_link", "") if task else "")
        sdb_link_entry = ttk.Entry(parent, textvariable=self.sdb_link_var)
        sdb_link_entry.grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Description
        row += 1
        ttk.Separator(parent, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        row += 1
        ttk.Label(parent, text="Description:").grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0,5))
        
        # Create a frame to contain both text widget and scrollbar in the same row as the label
        desc_frame = ttk.Frame(parent)
        desc_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)
        desc_frame.columnconfigure(0, weight=1)  # Make text widget expand
        
        self.description_text = tk.Text(desc_frame, height=4, wrap=tk.WORD)
        self.description_text.grid(row=0, column=0, sticky=tk.EW)
        if task and task.get("description"):
            self.description_text.insert("1.0", task["description"])
        
        # Add scrollbar to description inside the frame
        desc_scrollbar = ttk.Scrollbar(desc_frame, orient="vertical", command=self.description_text.yview)
        desc_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.description_text.configure(yscrollcommand=desc_scrollbar.set)
        
        # Priority
        row += 1
        ttk.Label(parent, text="Priority:").grid(row=row, column=0, sticky=tk.W, pady=5)
        priority_frame = ttk.Frame(parent)
        priority_frame.grid(row=row, column=1, sticky=tk.W, pady=5)
        
        self.priority_var = tk.StringVar(value=task["priority"] if task else "medium")
        ttk.Radiobutton(priority_frame, text="Low", variable=self.priority_var, value="low").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(priority_frame, text="Medium", variable=self.priority_var, value="medium").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(priority_frame, text="High", variable=self.priority_var, value="high").pack(side=tk.LEFT, padx=5)
        
        # Category
        row += 1
        ttk.Label(parent, text="Category:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.category_var = tk.StringVar(value=task["category"] if task else "general")
        ttk.Entry(parent, textvariable=self.category_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        
        # Main Staff
        row += 1
        ttk.Label(parent, text="Main Staff:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.main_staff_var = tk.StringVar(value=task.get("main_staff", "") if task else "")
        self.main_staff_combo = ttk.Combobox(parent, textvariable=self.main_staff_var, 
                                           values=USERS[1:], state="readonly")
        self.main_staff_combo.grid(row=row, column=1, sticky=tk.EW, pady=5)
        self.disable_combobox_mousewheel(self.main_staff_combo)  # Disable mousewheel scrolling
        row += 1
        
        # Assigned To
        ttk.Label(parent, text="Assigned To:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.assigned_to_var = tk.StringVar(value=task.get("assigned_to", "") if task else "")
        self.assigned_to_combo = ttk.Combobox(parent, textvariable=self.assigned_to_var, 
                                            values=USERS[1:], state="readonly")
        self.assigned_to_combo.grid(row=row, column=1, sticky=tk.EW, pady=5)
        self.disable_combobox_mousewheel(self.assigned_to_combo)  # Disable mousewheel scrolling
        row += 1
        
        # Qtd Mhr
        ttk.Label(parent, text="Qtd Mhr:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.qtd_mhr_var = tk.StringVar(value=str(task.get("qtd_mhr", "")) if task and task.get("qtd_mhr") is not None else "") # Ensure default is empty string
        ttk.Entry(parent, textvariable=self.qtd_mhr_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        row += 1

        # Actual Mhr
        ttk.Label(parent, text="Actual Mhr:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.actual_mhr_var = tk.StringVar(value=str(task.get("actual_mhr", "")) if task and task.get("actual_mhr") is not None else "") # Ensure default is empty string
        ttk.Entry(parent, textvariable=self.actual_mhr_var).grid(row=row, column=1, sticky=tk.EW, pady=5)
        row += 1

        # Status (only for editing)
        if task is not None:
            row += 1
            ttk.Label(parent, text="Status:").grid(row=row, column=0, sticky=tk.W, pady=5)
            self.completed_var = tk.BooleanVar(value=task["completed"])
            ttk.Checkbutton(parent, text="Completed", variable=self.completed_var).grid(row=row, column=1, sticky=tk.W, pady=5)
    
    def disable_combobox_mousewheel(self, combobox):
        """Disable mousewheel scrolling for a combobox to prevent accidental changes."""
        def block_mousewheel(event):
            return "break"
        
        # Bind mousewheel events for different platforms
        combobox.bind("<MouseWheel>", block_mousewheel)  # Windows
        combobox.bind("<Button-4>", block_mousewheel)    # Linux
        combobox.bind("<Button-5>", block_mousewheel)    # Linux
        # Remove the incorrect macOS binding that was causing the error
    
    def browse_folder(self):
        """Browse for a folder and set it as the link."""
        # Get the applied vessel value
        applied_vessel = self.applied_vessel_var.get().strip()
        
        # Check if applied vessel is provided
        if not applied_vessel:
            messagebox.showwarning("Missing Information", 
                                 "Please enter a value for Applied Vessel before browsing.",
                                 parent=self.dialog)
            return
        
        # Define the base project path
        base_path = r"\\srb096154\01_CESSD_SCG_CAD\01_Projects"
        vessel_path = os.path.join(base_path, applied_vessel)
        
        # Check if the vessel folder exists
        if not os.path.exists(vessel_path):
            # Ask if the user wants to create the folder
            create_folder = messagebox.askyesno(
                "Create Folder",
                f"Folder for vessel '{applied_vessel}' does not exist.\n\nDo you want to create it?",
                parent=self.dialog
            )
            
            if create_folder:
                try:
                    # Create the folder
                    os.makedirs(vessel_path, exist_ok=True)
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"Could not create folder: {vessel_path}\n\nError: {str(e)}",
                        parent=self.dialog
                    )
                    return
            else:
                return
        
        # Open folder dialog
        try:
            from tkinter import filedialog
            
            # Use the vessel path as the initial directory if it exists
            initial_dir = vessel_path if os.path.exists(vessel_path) else os.path.dirname(vessel_path)
            
            folder_path = filedialog.askdirectory(
                title="Select Folder",
                initialdir=initial_dir,
                parent=self.dialog
            )
            
            # Update the link field if a folder was selected
            if folder_path:
                self.link_var.set(folder_path)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"An error occurred while browsing for folders: {str(e)}",
                parent=self.dialog
            )
    
    def save(self):
        """Save the form data."""
        # Validate required fields
        validation_errors = []
        
        # Request No validation
        request_no = self.request_no_var.get().strip()
        if not request_no:
            validation_errors.append("Request No. is required")
        
        # Equipment Name validation
        title = self.title_var.get().strip()
        if not title:
            validation_errors.append("Equipment Name is required")
        
        # Applied Vessel validation
        applied_vessel = self.applied_vessel_var.get().strip()
        if not applied_vessel:
            validation_errors.append("Applied Vessel is required")
        
        # Rev validation
        rev = self.rev_var.get().strip()
        if not rev:
            validation_errors.append("Rev is required")
        else:
            if not rev.isdigit(): # Check if rev is numeric
                validation_errors.append("Rev must be a valid integer")
            else:
                try:
                    rev_num = int(rev)
                    if rev_num < 0: # Check if rev is non-negative
                        validation_errors.append("Rev must be a non-negative integer")
                    else:
                        self.rev_var.set(str(rev_num)) # Store the converted integer back
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Rev must be a valid integer")
        
        # Link validation
        link = self.link_var.get().strip()
        if not link:
            validation_errors.append("Link is required")
        
        # SDB Link validation
        sdb_link = self.sdb_link_var.get().strip()
        if not sdb_link:
            validation_errors.append("SDB Link is required")
        
        # Main Staff validation
        main_staff = self.main_staff_var.get().strip()
        if not main_staff:
            validation_errors.append("Main Staff is required")
        
        # Priority validation
        priority = self.priority_var.get().lower()
        if priority not in ["low", "medium", "high"]:
            validation_errors.append("Priority must be one of: Low, Medium, High")
        
        # Category validation (example: alphanumeric and spaces only)
        category = self.category_var.get().strip()
        if not category:
            validation_errors.append("Category is required")
        elif not category.replace(" ", "").isalnum(): # Allow alphanumeric and spaces
            validation_errors.append("Category must be alphanumeric and spaces only")
        
        # Main Staff and Assigned To validation (check against USERS list)
        main_staff = self.main_staff_var.get()
        if main_staff and main_staff not in USERS[1:]: # USERS[1:] to exclude "All"
            validation_errors.append(f"Main Staff must be one of: {', '.join(USERS[1:])}")

        assigned_to = self.assigned_to_var.get()
        if assigned_to and assigned_to not in USERS[1:]: # USERS[1:] to exclude "All"
            validation_errors.append(f"Assigned To must be one of: {', '.join(USERS[1:])}")
        
        # Qtd Mhr validation
        qtd_mhr_str = self.qtd_mhr_var.get().strip()
        if qtd_mhr_str: # Only validate if not empty
            if not qtd_mhr_str.isdigit():
                validation_errors.append("Qtd Mhr must be an integer")
            else:
                try:
                    qtd_mhr = int(qtd_mhr_str)
                    if qtd_mhr < 0:
                        validation_errors.append("Qtd Mhr must be a non-negative integer")
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Qtd Mhr must be an integer")
        else:
            qtd_mhr = 0 # Default to 0 if empty

        # Actual Mhr validation
        actual_mhr_str = self.actual_mhr_var.get().strip()
        if actual_mhr_str: # Only validate if not empty
            if not actual_mhr_str.isdigit():
                validation_errors.append("Actual Mhr must be an integer")
            else:
                try:
                    actual_mhr = int(actual_mhr_str)
                    if actual_mhr < 0:
                        validation_errors.append("Actual Mhr must be a non-negative integer")
                except ValueError: # Redundant, but kept for clarity
                    validation_errors.append("Actual Mhr must be an integer")
        else:
            actual_mhr = 0 # Default to 0 if empty

        # Show all validation errors if any
        if validation_errors:
            error_message = "Please correct the following errors:\n\n" + "\n".join(f"• {error}" for error in validation_errors)
            messagebox.showerror("Validation Error", error_message, parent=self.dialog)
            return
        
        # Get dates - ensure we get the raw string value from the picker
        try:
            requested_date = self.requested_date_picker.get()
            if requested_date:
                # Store as string in YYYY-MM-DD format
                requested_date = requested_date
        except Exception as e:
            print(f"Error getting requested date: {str(e)}")
            requested_date = None
            
        try:
            date_started = self.date_started_picker.get()
            if date_started:
                # Store as string in YYYY-MM-DD format
                date_started = date_started
        except Exception as e:
            print(f"Error getting start date: {str(e)}")
            date_started = None
            
        try:
            due_date = self.due_date_picker.get()
            if due_date:
                # Store as string in YYYY-MM-DD format
                due_date = due_date
        except Exception as e:
            print(f"Error getting due date: {str(e)}")
            due_date = None
        
        # Create result with all fields
        self.result = {
            "request_no": request_no,
            "title": title,
            "description": self.description_text.get("1.0", tk.END).strip() or None,  # Make None if empty
            "requested_date": requested_date,
            "date_started": date_started,
            "due_date": due_date,
            "priority": self.priority_var.get(),
            "category": self.category_var.get().strip() or "general",
            "main_staff": main_staff,
            "assigned_to": self.assigned_to_var.get().strip() or None,
            "applied_vessel": applied_vessel,
            "rev": rev,
            "drawing_no": self.drawing_no_var.get().strip() or None,  # Make None if empty
            "link": link,
            "sdb_link": sdb_link,
            "qtd_mhr": qtd_mhr if qtd_mhr_str else 0, # Use validated integer value, default to 0 if empty
            "actual_mhr": actual_mhr if actual_mhr_str else 0, # Use validated integer value, default to 0 if empty
            "request_no": request_no,
            "requested_date": requested_date,
            "date_started": date_started
        }
        
        # Add completed status if editing
        if hasattr(self, "completed_var"):
            self.result["completed"] = self.completed_var.get()
        
        # Unbind mousewheel events before destroying the dialog
        self.unbind_mousewheel()
        
        # Close dialog
        self.dialog.destroy()

    def scrape_data(self):
        """Handles scraping data from SmartDB using the provided URL."""
        logging.info("Starting scrape_data method.") # Log start
        url = self.scrape_url_var.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter a SmartDB URL to scrape.", parent=self.dialog)
            logging.warning("Scrape cancelled: No URL provided.")
            return

        # --- Set SDB Link field with the Scrape URL ---
        self.sdb_link_var.set(url)
        logging.info(f"Set SDB Link field to: {url}")
        # ----------------------------------------------

        # --- Check Applied Vessel and prompt if needed ---
        applied_vessel = self.applied_vessel_var.get().strip()
        if not applied_vessel:
            logging.info("Applied Vessel is missing, prompting user.")
            # Prompt user for Applied Vessel
            applied_vessel = simpledialog.askstring("Applied Vessel Required",
                                                  "Please enter the Applied Vessel number:",
                                                  parent=self.dialog)
            if not applied_vessel:  # User cancelled or entered empty string
                messagebox.showwarning("Missing Information",
                                     "Applied Vessel is required to proceed.",
                                     parent=self.dialog)
                logging.warning("Scrape cancelled: User did not provide Applied Vessel.")
                return
            # Set the Applied Vessel field
            self.applied_vessel_var.set(applied_vessel)
            logging.info(f"Applied Vessel set by user: {applied_vessel}")

        # --- Define standard location and construct base path ---
        standard_location = r"\\srb096154\01_CESSD_SCG_CAD\01_Projects"
        base_vessel_path = os.path.join(standard_location, applied_vessel)
        logging.info(f"Base vessel path defined: {base_vessel_path}")

        # Check if the vessel folder exists
        if not os.path.exists(base_vessel_path):
            logging.warning(f"Base vessel folder does not exist: {base_vessel_path}")
            create_folder = messagebox.askyesno(
                "Create Folder",
                f"Folder for vessel '{applied_vessel}' does not exist at:\n{base_vessel_path}\n\nDo you want to create it?",
                parent=self.dialog
            )

            if create_folder:
                try:
                    os.makedirs(base_vessel_path, exist_ok=True)
                    logging.info(f"Created base vessel folder: {base_vessel_path}")
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"Could not create folder: {base_vessel_path}\n\nError: {str(e)}",
                        parent=self.dialog
                    )
                    logging.error(f"Failed to create base vessel folder: {e}", exc_info=True)
                    return
            else:
                logging.warning("Scrape cancelled: User chose not to create base vessel folder.")
                return

        # Show info message about folder selection
        logging.info("Prompting user to select final folder structure.")
        messagebox.showinfo("Select Final Folder",
                          f"You will now be prompted to select or create the final folder structure.\n\n"
                          f"Please navigate to or create the desired subfolder structure\n"
                          f"(e.g., SmartDB/Equipment/MAIN ENGINE)\n\n"
                          f"Starting from: {base_vessel_path}",
                          parent=self.dialog)

        final_target_folder = filedialog.askdirectory(
            title="Select Final Folder",
            initialdir=base_vessel_path,  # Start from the vessel folder
            parent=self.dialog
        )

        if not final_target_folder:
            messagebox.showwarning("Folder Selection Cancelled",
                                 "No final folder selected. Downloads will be skipped, but scraping will proceed.", # Adjusted message
                                 parent=self.dialog)
            logging.warning("No final folder selected by user. Downloads will be skipped.")
            # Allow proceeding without download if user cancels folder selection
            # final_target_folder will be None, which is handled later
        else:
            logging.info(f"Final target folder selected by user: {final_target_folder}")
            # Set the Link field to the final selected path immediately
            self.link_var.set(final_target_folder)
            logging.info(f"Set Link field to final path: {final_target_folder}")


        # --- Create a LOCAL Temporary Directory for Browser Download ---
        temp_download_dir = None
        if final_target_folder: # Only create if needed for download
            try:
                import tempfile
                temp_download_dir = tempfile.mkdtemp(prefix="smartdb_dl_")
                logging.info(f"Created temporary local download directory: {temp_download_dir}")
            except Exception as temp_err:
                logging.error(f"ERROR: Could not create temporary download directory: {temp_err}")
                messagebox.showerror("Error",
                                   f"Could not create a temporary local folder for download:\n{temp_err}\n\nDownloads cannot proceed.",
                                   parent=self.dialog)
                # If temp dir fails, we cannot download, nullify the temp dir variable
                temp_download_dir = None
                # Do not return here, allow text scraping to continue if possible


        # --- Get Credentials ---
        email = None
        password = None
        try:
            logging.info("Getting email via smartdb_login.get_email_address...")
            email = smartdb_login.get_email_address(smartdb_login.USER_EMAIL)
            if not email:
                messagebox.showerror("Authentication Error", "Email address is required for scraping.", parent=self.dialog)
                logging.error("Scrape failed: Email address not obtained.")
                if temp_download_dir: shutil.rmtree(temp_download_dir, ignore_errors=True)
                return
            logging.info(f"Using email: {email}")

            logging.info("Getting password via smartdb_login.store_or_get_password...")
            password = smartdb_login.store_or_get_password(smartdb_login.KEYRING_SERVICE_NAME, email)
            if not password:
                messagebox.showerror("Authentication Error", "Password is required for scraping.", parent=self.dialog)
                logging.error("Scrape failed: Password not obtained.")
                if temp_download_dir: shutil.rmtree(temp_download_dir, ignore_errors=True)
                return
            logging.info("Password obtained successfully.")

        except Exception as auth_e:
            messagebox.showerror("Authentication Error", f"Failed to get credentials: {auth_e}", parent=self.dialog)
            logging.error(f"Error getting credentials for scraping: {auth_e}", exc_info=True)
            if temp_download_dir: shutil.rmtree(temp_download_dir, ignore_errors=True)
            return

        # --- Start Timer ---
        start_time = time.monotonic()
        logging.info("Starting scraping process timer.")

        # --- Perform Scraping ---
        messagebox.showinfo("Scraping Started",
                            "Starting web scraping process...\nBrowser may run in the background.\nFiles will download to a temporary folder first.\nThis may take a moment.",
                            parent=self.dialog)
        self.dialog.update_idletasks() # Ensure message box is shown

        extracted_data = None
        success = False
        driver = None # Initialize driver variable
        temp_user_data_dir_path = None # <<< Initialize variable for user data dir path
        moved_files_summary = []
        download_info = {}

        try:
            logging.info(f"Calling smartdb_login.login_to_smartdb with URL: {url}")
            # Pass the LOCAL TEMP directory for downloads if it was created
            # Capture the returned temp_user_data_dir_path
            success, extracted_data, driver, temp_user_data_dir_path = smartdb_login.login_to_smartdb( # <<< Capture path
                url, email, password, target_download_folder=temp_download_dir
            )
            logging.info(f"Login/Scrape process finished. Success: {success}")
            if temp_user_data_dir_path:
                logging.info(f"WebDriver used temporary user data directory: {temp_user_data_dir_path}")
            if extracted_data:
                 logging.info(f"Extracted data keys: {list(extracted_data.keys())}")
                 download_info = extracted_data.get("Downloaded Files", {})
                 logging.info(f"Download info found: {len(download_info)} item(s)")

            # --- Populate Text Fields (if successful) ---
            if success and extracted_data:
                logging.info("Populating dialog text fields from extracted data...")
                self.populate_fields_from_scrape(extracted_data) # Populate text fields first
            elif success and not extracted_data:
                 logging.warning("Login successful, but no text data found to populate fields.")
                 messagebox.showwarning("No Text Data", "Login successful, but no text data found to populate fields.", parent=self.dialog)
            elif not success:
                 logging.error("Login or text data scraping failed.")
                 messagebox.showerror("Scraping Failed", "Login or text data scraping failed. Check logs for details.", parent=self.dialog)
                 # Don't proceed to file move if scraping failed

            # --- Move Downloaded Files from Temp to Final Destination (only if scraping succeeded and folders exist) ---
            if success and final_target_folder and temp_download_dir and download_info:
                logging.info(f"\nMoving downloaded files from '{temp_download_dir}' to '{final_target_folder}'...")

                for itemkey, info in download_info.items():
                    # Check status from smartdb_login (should be "Success" if downloaded to temp)
                    if info.get("status") == "Success" and "path" in info:
                        temp_file_path = info["path"]
                        # Double check the temp file exists before moving
                        if not os.path.exists(temp_file_path):
                             logging.warning(f"  Skipping move: Temp file '{temp_file_path}' not found for itemkey '{itemkey}'.")
                             info["error"] = "Temporary file missing before move"
                             info["status"] = "Failed (Move Error)"
                             continue

                        original_filename = os.path.basename(temp_file_path)
                        final_destination_path = os.path.join(final_target_folder, original_filename)

                        try:
                            logging.info(f"  Moving '{original_filename}' to '{final_destination_path}'...")
                            # Ensure final directory exists (should have been checked, but double-check)
                            os.makedirs(final_target_folder, exist_ok=True)

                            # Handle potential overwrite? For now, move will overwrite by default.
                            # Add check if needed:
                            # if os.path.exists(final_destination_path):
                            #     logging.warning(f"    Destination file '{final_destination_path}' already exists. Overwriting.")

                            shutil.move(temp_file_path, final_destination_path)
                            logging.info(f"  Successfully moved: {original_filename}")
                            moved_files_summary.append(original_filename)

                            # Update the path in extracted_data to the final location
                            info["path"] = final_destination_path
                            info["status"] = "Success (Moved)" # Update status

                        except Exception as move_err:
                            logging.error(f"  ERROR moving file '{original_filename}': {move_err}", exc_info=True)
                            info["error"] = f"Move Failed: {move_err}"
                            info["status"] = "Failed (Move Error)" # Update status if move fails
                    elif info.get("status") == "Failed":
                         logging.warning(f"  Skipping move for itemkey '{itemkey}': Initial download failed.")
                    # Add case for if status is already "Success (Moved)" - shouldn't happen here but good practice
                    elif info.get("status") == "Success (Moved)":
                         logging.info(f"  File for itemkey '{itemkey}' already marked as moved.")


                # --- Update the Link field to the FINAL TARGET FOLDER if files were moved ---
                # This is slightly redundant as we set it earlier, but confirms it reflects the download/move outcome
                if moved_files_summary and final_target_folder: # Check if any files were successfully moved
                    self.link_var.set(final_target_folder) # Set Link to the folder path
                    logging.info(f"Confirmed Link field remains set to target directory after move: {final_target_folder}")
                    # Update the key in extracted_data as well for consistency, pointing to the folder
                    if extracted_data: extracted_data["Primary Downloaded File Path"] = final_target_folder
                elif final_target_folder: # If a folder was selected but no files were moved (or found)
                     self.link_var.set(final_target_folder) # Ensure it's still set
                     logging.info(f"Link field remains set to selected target folder (no files moved): {final_target_folder}")


            # Prepare download result message based on final status (including move results)
            download_result_message = ""
            if final_target_folder: # Check if a target folder was selected initially
                 # Recalculate counts based on final status in download_info
                 success_moved_count = sum(1 for info in download_info.values() if info.get("status") == "Success (Moved)")
                 fail_move_count = sum(1 for info in download_info.values() if info.get("status") == "Failed (Move Error)")
                 fail_dl_count = sum(1 for info in download_info.values() if info.get("status") == "Failed") # Original download failures

                 if success_moved_count > 0:
                     download_result_message += f"\n\nSuccessfully downloaded and moved {success_moved_count} file(s) to:\n{final_target_folder}"
                 if fail_move_count > 0:
                     download_result_message += f"\n\nEncountered errors moving {fail_move_count} downloaded file(s)."
                 if fail_dl_count > 0:
                     download_result_message += f"\n\nFailed to download {fail_dl_count} file(s) initially."
                 # Check if scraping was successful but no downloadable files were found/processed
                 if success and not download_info and success_moved_count==0 and fail_move_count==0 and fail_dl_count==0:
                      download_result_message = "\n\nNo files found to download."
                 # If scraping failed, don't report download status
                 elif not success:
                      download_result_message = ""

            elif not final_target_folder: # User cancelled selection or temp dir failed
                 download_result_message = "\n\nAttachment download skipped (no final folder specified/accessible)."


        except Exception as scrape_e:
            success = False # Ensure success is False
            messagebox.showerror("Scraping Process Error", f"An unexpected error occurred during scraping/moving: {scrape_e}", parent=self.dialog)
            # Log the error including traceback
            logging.error(f"Error during scrape_data execution: {scrape_e}", exc_info=True) # <<< Ensure traceback is logged
        finally:
            # --- Stop Timer ---
            elapsed_time = time.monotonic() - start_time
            elapsed_time_str = f"{elapsed_time:.2f} seconds"
            logging.info(f"Scraping process finished. Total time: {elapsed_time_str}")

            # --- Ensure driver is closed ---
            if driver:
                logging.info("Closing WebDriver...")
                try:
                    driver.quit()
                    logging.info("WebDriver closed.")
                except Exception as quit_e:
                    logging.error(f"Error closing WebDriver: {quit_e}", exc_info=True)

            # --- Clean up Temporary DOWNLOAD Directory ---
            if temp_download_dir and os.path.exists(temp_download_dir): # Check existence
                try:
                    logging.info(f"Cleaning up temporary DOWNLOAD directory: {temp_download_dir}")
                    shutil.rmtree(temp_download_dir, ignore_errors=True)
                    logging.info("Temporary download directory removal attempted.")
                except Exception as clean_err:
                    logging.warning(f"Could not remove temporary download directory '{temp_download_dir}': {clean_err}")

            # --- Clean up Temporary USER DATA Directory ---
            if temp_user_data_dir_path and os.path.exists(temp_user_data_dir_path): # <<< Add cleanup here
                try:
                    logging.info(f"Cleaning up temporary USER DATA directory: {temp_user_data_dir_path}")
                    # Use ignore_errors=True for robustness, as sometimes files might be locked briefly after browser close
                    shutil.rmtree(temp_user_data_dir_path, ignore_errors=True)
                    logging.info("Temporary user data directory removal attempted.")
                except Exception as clean_user_err:
                    # Log warning, but don't stop execution
                    logging.warning(f"Could not remove temporary user data directory '{temp_user_data_dir_path}': {clean_user_err}")
            elif temp_user_data_dir_path:
                 logging.warning(f"Temporary user data directory path was set ('{temp_user_data_dir_path}') but directory not found for cleanup.")


            # --- Show Final Status Message ---
            final_message = ""
            if success:
                final_message = "Scraping process finished."
                # Add download/move info message
                final_message += download_result_message
            else:
                # Check if it was just a login failure or a broader exception
                if extracted_data is None: # Indicates likely exception before or during login call
                     final_message = "Scraping process failed due to an unexpected error before completion."
                else: # Login call returned success=False
                     final_message = "Scraping process failed during login or data extraction."


            final_message += f"\n\nTotal time: {elapsed_time_str}"
            messagebox.showinfo("Scraping Complete", final_message, parent=self.dialog)
            logging.info(f"Final status message shown to user: {final_message}")

    def populate_fields_from_scrape(self, extracted_data):
        """Populates the dialog fields based on the scraped data dictionary."""
        # This function should remain largely the same for populating text fields.
        # The crucial change is HOW the 'Link' field gets its final value, which is now handled
        # in the scrape_data method *after* the file move is attempted.
        # We'll remove the explicit setting of the 'Link' field from here based on "Primary Downloaded File Path"
        # as that key might point to the temp location before the move, or the final location after.
        # scrape_data now handles setting self.link_var directly after the move.

        if not extracted_data:
            return

        populated_fields = []
        print("Populating fields with:", extracted_data)
        try:
            # Handle Request No
            req_no_val = extracted_data.get("Request No", "")
            if req_no_val != "Not Found" and req_no_val:
                self.FIELD_VAR_MAP["Request No"].set(req_no_val)
                populated_fields.append("Request No")

            # Handle Equipment Name (from Scraped "Description")
            equip_name_val = extracted_data.get("Description", "") # Key is "Description" from FIELDS_TO_EXTRACT
            if equip_name_val != "Not Found" and equip_name_val:
                 self.FIELD_VAR_MAP["Equipment Name"].set(equip_name_val)
                 populated_fields.append("Equipment Name (from Scraped Description)")

            # Handle Applied Vessel
            vessel_val = extracted_data.get("Applied Vessel", "")
            if vessel_val != "Not Found" and vessel_val:
                 self.FIELD_VAR_MAP["Applied Vessel"].set(vessel_val)
                 populated_fields.append("Applied Vessel")

            # Handle Revision No
            rev_val = extracted_data.get("Revision No", "")
            if rev_val != "Not Found" and rev_val:
                 # Check for 'O' and replace with '0'
                 if rev_val.strip().upper() == 'O':
                     rev_val = '0'
                     print("  Interpreted Revision 'O' as '0'.")
                 self.FIELD_VAR_MAP["Revision No"].set(rev_val)
                 populated_fields.append("Revision No")

            # Handle Drawing No
            dwg_no_val = extracted_data.get("Drawing No", "")
            if dwg_no_val != "Not Found" and dwg_no_val:
                 self.FIELD_VAR_MAP["Drawing No"].set(dwg_no_val)
                 populated_fields.append("Drawing No")

            # --- Link field is now handled AFTER potential move in scrape_data ---
            # Remove this block:
            # link_path_val = extracted_data.get("Primary Downloaded File Path", "")
            # if link_path_val:
            #      self.FIELD_VAR_MAP["Link"].set(link_path_val)
            #      populated_fields.append("Link (Downloaded File Path)")

            # Handle Description (from Scraped "Note")
            note_val = extracted_data.get("Note", "") # Key is "Note"
            desc_widget = self.FIELD_VAR_MAP["Description"]
            if note_val != "Not Found" and note_val and isinstance(desc_widget, tk.Text):
                 desc_widget.delete('1.0', tk.END)
                 desc_widget.insert('1.0', note_val)
                 populated_fields.append("Description (from Scraped Note)")

            # Handle Modeling Staff - Find first name match in known USERS
            staff_val = extracted_data.get("Modeling Staff", "")
            matched_staff = None
            if staff_val and staff_val != "Not Found":
                staff_val_lower = staff_val.lower()
                for known_staff_name in USERS[1:]: # Iterate through valid staff first names
                    if known_staff_name.lower() in staff_val_lower:
                        matched_staff = known_staff_name
                        print(f"  Matched Modeling Staff '{staff_val}' to known staff '{matched_staff}'")
                        break # Found the first match

            if matched_staff:
                 self.FIELD_VAR_MAP["Main Staff"].set(matched_staff)
                 populated_fields.append("Main Staff (from Modeling Staff)")
                 # Optionally set Assigned To as well (if desired):
                 # self.FIELD_VAR_MAP["Assigned To"].set(matched_staff)
                 # if "Assigned To" not in populated_fields: populated_fields.append("Assigned To")
            elif staff_val and staff_val != "Not Found": # If staff name was scraped but no match found
                 print(f"Warning: Scraped Modeling Staff '{staff_val}' did not match any known staff in USERS list.")

            # --- Handle Dates ---
            date_format = "%Y/%m/%d" # Format used in SmartDB example

            # Handle Requested Date
            req_date_str = extracted_data.get("Requested Date", "") # Key is "Requested Date"
            if req_date_str and req_date_str != "Not Found":
                try:
                    req_date_obj = datetime.datetime.strptime(req_date_str, date_format).date()
                    target_widget = self.FIELD_VAR_MAP["Requested Date"]
                    if isinstance(target_widget, DateEntry):
                         target_widget.set_date(req_date_obj)
                         populated_fields.append("Requested Date")
                    else:
                         print(f"Warning: Target widget for Requested Date is not a DateEntry ({type(target_widget)}).")
                except ValueError:
                    print(f"Warning: Could not parse Requested Date '{req_date_str}' with format {date_format}")
                except Exception as date_err:
                     print(f"Error setting Requested Date: {date_err}")

            # Handle Due Date (from Scraped "Deadline")
            due_date_str = extracted_data.get("Deadline", "") # Key is "Deadline"
            if due_date_str and due_date_str != "Not Found":
                 try:
                     # Clean the date string (remove day in parentheses)
                     cleaned_due_date_str = due_date_str.split('(')[0].strip()
                     print(f"  Parsing cleaned Deadline: '{cleaned_due_date_str}'")
                     due_date_obj = datetime.datetime.strptime(cleaned_due_date_str, date_format).date()
                     target_widget = self.FIELD_VAR_MAP["Due Date"]
                     if isinstance(target_widget, DateEntry):
                         target_widget.set_date(due_date_obj)
                         populated_fields.append("Due Date (from Scraped Deadline)")
                     else:
                          print(f"Warning: Target widget for Due Date is not a DateEntry ({type(target_widget)}).")
                 except ValueError:
                     print(f"Warning: Could not parse Due Date (Deadline) '{cleaned_due_date_str}' with format {date_format}")
                 except Exception as date_err:
                     print(f"Error setting Due Date: {date_err}")

            # --- Final Feedback ---
            if populated_fields:
                print(f"Successfully populated fields: {', '.join(populated_fields)}")
                # Message box is now shown after download attempt/move in the main scrape_data flow
            else:
                print("No matching text data found to populate fields.")
                # Message box is now shown after download attempt/move in the main scrape_data flow

        except Exception as populate_e:
             logging.error(f"Error populating fields: {populate_e}", exc_info=True)
             raise # Re-raise to be caught by the main try block


class TaskManager:
    """Class to manage daily tasks."""
    
    def __init__(self):
        """Initialize the task manager."""
        # Database connection info
        self.server = "10.195.103.198"  # Server address (IP address)
        self.database = "TaskManagerDB"
        self.user = None  # Will be set by _get_user_credentials
        self.password = None  # Will be set by _get_user_credentials
        
        # Initialize internal state
        self._tasks = []  # List of tasks loaded from the database
        self._deleted_tasks = []  # Track recently deleted tasks for undo (used by older undo logic)
        # self._transaction_active = False # REMOVED - Not needed with autocommit=True
        
        # Create standardized connection references
        self.conn = None  # Main connection object
        self.cursor = None  # Main cursor object
        
        # Initialize everything
        self._get_user_credentials()  # Set up credentials
        self._connect()  # Connect to database
        if self.conn: # Only load if connection succeeded
            self._tasks = self._load_tasks()  # Load tasks from database
            self._ensure_deleted_field_exists()
        else:
            # Handle case where initial connection failed
             self._tasks = []
             logging.error("Initial database connection failed. Task list will be empty.")
             # Consider raising an error or exiting if connection is critical

        self.deleted_tasks_stack = []  # Stack of deletion transactions (each a list of task IDs)
    
    def _ensure_deleted_field_exists(self):
        """Ensure the 'deleted' field exists in the Tasks table."""
        if not self.conn or not self.cursor: return # Cannot proceed without connection
        try:
            # Check if the deleted column exists
            self.cursor.execute("""
                SELECT COUNT(*)
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'Tasks'
                AND COLUMN_NAME = 'deleted'
            """)
            column_exists = self.cursor.fetchone()[0] > 0

            if not column_exists:
                # Add the column if it doesn't exist
                logging.info("Attempting to add 'deleted' column...")
                self.cursor.execute("""
                    ALTER TABLE Tasks
                    ADD deleted BIT NOT NULL DEFAULT 0
                """)
                # No explicit commit needed due to autocommit=True
                logging.info("Added 'deleted' column to Tasks table (autocommitted).")
        except pymssql.Error as e:
            logging.error("Error checking/adding deleted column", exc_info=True)
            # Rollback might not be strictly necessary with autocommit but good practice if error occurs
            try: self.conn.rollback()
            except Exception: pass
            messagebox.showerror("Schema Update Error", f"Could not ensure 'deleted' column exists: {e}")
    
    def _get_user_credentials(self):
        """Retrieves SQL Server credentials based on the current user."""
        username = getpass.getuser()  # Get the current logged-in username

        # Set credentials based on username
        if username == "a0011071":
            self.user = "TaskUser1"
            self.password = "pass1"
        elif username == "a0010756":
            self.user = "TaskUser2"
            self.password = "pass1"
        elif username == "a0012923":
            self.user = "TaskUser3"
            self.password = "pass1"
        elif username == "a0010751":
            self.user = "TaskUser4"
            self.password = "pass1"
        elif username == "a0012501":
            self.user = "TaskUser5"
            self.password = "pass1"
        elif username == "a0008432":
            self.user = "TaskUser6"
            self.password = "pass1"
        elif username == "a0003878":
            self.user = "TaskUser7"
            self.password = "pass1"
        else:
            # Use default credentials for other users
            print(f"Warning: Username '{username}' is not recognized as a specific user. Using default login.")
            self.user = "DefaultFallbackUser"
            self.password = "pass1"

        # Store sql_username as instance attribute - not needed with this approach
        # self.sql_username = sql_username
    
    def _connect(self):
        """Connect to the SQL Server database with autocommit=True."""
        try:
            # Close existing connection if any
            if self.conn:
                try:
                    self.conn.close()
                except Exception as close_e:
                     logging.warning(f"Ignoring error while closing previous connection: {close_e}")
                self.conn = None
                self.cursor = None

            # Connect to the database with autocommit enabled
            self.conn = pymssql.connect(
                server=self.server,
                database=self.database,
                user=self.user,
                password=self.password,
                autocommit=True # <<< SET AUTOCOMMIT TO TRUE
            )

            # Create a cursor
            self.cursor = self.conn.cursor()

            logging.info("Successfully connected to the SQL Server database (autocommit=True).")
            # Test connection (optional but good)
            # self.cursor.execute("SELECT @@SERVERNAME")
            # logging.info(f"Connected to server: {self.cursor.fetchone()[0]}")

        except pymssql.Error as e:
            logging.error(f"Error connecting to database: {str(e)}", exc_info=True)
            self.conn = None # Ensure conn is None on failure
            self.cursor = None
            # Raise the error so the caller (__init__) knows connection failed
            # Avoid messagebox here, let __init__ or main handle startup failure
            raise ConnectionError(f"Database connection failed: {e}") from e
        except Exception as e:
             logging.error(f"Unexpected error during connection: {e}", exc_info=True)
             self.conn = None
             self.cursor = None
             raise ConnectionError(f"Unexpected connection error: {e}") from e
    
    def _load_tasks(self) -> List[Dict[str, Any]]:
        """Load tasks from the SQL Server database."""
        if not self.conn or not self.cursor:
             logging.error("Cannot load tasks: No database connection.")
             return []
        try:
            # Ensure we select all columns, including the 'deleted' column
            self.cursor.execute("SELECT * FROM [dbo].[Tasks]")
            rows = self.cursor.fetchall()
            column_names = [desc[0] for desc in self.cursor.description]

            # Convert rows to a list of dictionaries using column names
            tasks = []
            for row in rows:
                task = dict(zip(column_names, row))

                # --- Data Type Conversions and Formatting ---
                # Convert dates to string YYYY-MM-DD or None
                for date_key in ['created_date', 'due_date', 'requested_date', 'date_started']:
                    if date_key in task and isinstance(task[date_key], datetime.datetime):
                        task[date_key] = task[date_key].strftime("%Y-%m-%d")
                    elif date_key in task and isinstance(task[date_key], datetime.date):
                         task[date_key] = task[date_key].strftime("%Y-%m-%d")
                    elif date_key in task and not isinstance(task[date_key], str):
                         task[date_key] = None # Handle potential non-date, non-string values

                # Format last_modified
                if 'last_modified' in task and isinstance(task['last_modified'], datetime.datetime):
                    task['last_modified'] = task['last_modified'].strftime("%Y-%m-%d %H:%M:%S")
                elif 'last_modified' in task:
                     task['last_modified'] = None # Or keep as is if needed

                # Ensure boolean fields are bool
                if 'completed' in task:
                    task['completed'] = bool(task['completed'])
                if 'deleted' in task:
                    task['deleted'] = bool(task['deleted'])
                else: # Handle case where 'deleted' column might be missing initially
                    task['deleted'] = False

                # Ensure numeric fields are numbers (or None/0 if appropriate)
                for num_key in ['qtd_mhr', 'actual_mhr']:
                     if num_key in task and task[num_key] is not None:
                         try:
                             task[num_key] = int(task[num_key])
                         except (ValueError, TypeError):
                              task[num_key] = 0 # Or None, depending on desired default
                     else:
                          task[num_key] = 0 # Default if null

                tasks.append(task)

            logging.info(f"Loaded {len(tasks)} tasks from the database.")
            return tasks
        except pymssql.Error as e:
            logging.error("Error loading tasks from database", exc_info=True)
            messagebox.showerror("Database Error", f"Error loading tasks: {str(e)}")
            return []
        except Exception as e:
            logging.error(f"Unexpected error loading tasks: {str(e)}", exc_info=True)
            messagebox.showerror("Application Error", f"Unexpected error loading tasks: {str(e)}")
            return []
    
    def _get_current_user(self):
        """Get the current user's username."""
        try:
            return os.getenv('USERNAME') or os.getenv('USER') or 'Unknown User'
        except:
            return 'Unknown User'
    
    def add_task(self, title: str, description: str = "", due_date: Optional[str] = None, 
                priority: str = "medium", category: str = "general", 
                main_staff: Optional[str] = None, assigned_to: Optional[str] = None,
                applied_vessel: str = "", rev: str = "", drawing_no: str = "",
                link: str = "", sdb_link: str = "", request_no: str = "",
                requested_date: Optional[str] = None, date_started: Optional[str] = None,
                qtd_mhr: int = 0, actual_mhr: int = 0) -> Optional[int]: # Return new ID or None on failure
        """
        Add a new task to the SQL Server database (autocommitted).
        Returns the new task ID on success, None on failure.
        """
        if not self.conn or not self.cursor:
             logging.error("Cannot add task: No database connection.")
             messagebox.showerror("Database Error", "No database connection available.")
             return None

        try:
            # Get current date and time for metadata columns
            # Ensure datetime objects are used for timestamp columns if DB expects them
            current_datetime = datetime.datetime.now()
            # Use datetime object directly if column type is DATETIME/DATETIME2
            created_date_dt = current_datetime
            last_modified_dt = current_datetime
            # Or format if column type is VARCHAR/NVARCHAR
            # created_date_str = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
            # last_modified_str = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

            created_by = self._get_current_user()
            modified_by = created_by

            # --- Input Validation and Formatting ---
            # Validate/Format dates (ensure they are 'YYYY-MM-DD' strings or None)
            def format_date(date_input):
                if isinstance(date_input, datetime.date):
                    return date_input.strftime("%Y-%m-%d")
                elif isinstance(date_input, str) and date_input:
                    try:
                        # Validate format YYYY-MM-DD
                        datetime.datetime.strptime(date_input, '%Y-%m-%d')
                        return date_input
                    except ValueError:
                        logging.warning(f"Invalid date format provided: {date_input}. Setting to None.")
                        return None
                return None # Return None for empty strings, None, or other invalid types

            due_date = format_date(due_date)
            requested_date = format_date(requested_date)
            date_started = format_date(date_started)

            # Ensure MHR values are integers
            try:
                qtd_mhr = int(qtd_mhr) if qtd_mhr is not None else 0
            except (ValueError, TypeError):
                 logging.warning(f"Invalid qtd_mhr value: {qtd_mhr}. Setting to 0.")
                 qtd_mhr = 0
            try:
                actual_mhr = int(actual_mhr) if actual_mhr is not None else 0
            except (ValueError, TypeError):
                 logging.warning(f"Invalid actual_mhr value: {actual_mhr}. Setting to 0.")
                 actual_mhr = 0


            # --- Database Interaction (autocommitted) ---
            sql = """
                INSERT INTO Tasks (
                    title, description, created_date, created_by, due_date, priority, category,
                    main_staff, assigned_to, completed, applied_vessel, rev, drawing_no,
                    link, sdb_link, request_no, requested_date, date_started,
                    qtd_mhr, actual_mhr, last_modified, modified_by, deleted
                ) OUTPUT INSERTED.id
                VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                );
            """
            # Prepare parameters tuple - ensure order matches INSERT list
            params = (
                title, description, created_date_dt, created_by, due_date, priority.lower(), category,
                main_staff, assigned_to, False, applied_vessel, rev, drawing_no, # completed=False
                link, sdb_link, request_no, requested_date, date_started,
                qtd_mhr, actual_mhr, last_modified_dt, modified_by, False # deleted=False
            )

            logging.debug(f"Executing INSERT with params: {params}")
            self.cursor.execute(sql, params)
            result = self.cursor.fetchone()

            if result is None:
                logging.error("INSERT statement did not return an ID via OUTPUT clause (autocommit might have failed implicitly).")
                # No explicit rollback needed for autocommit errors usually, DB handles it
                return None

            new_id = result[0]
            # No explicit commit needed here

            # --- Update Local Cache ---
            created_date_str = created_date_dt.strftime("%Y-%m-%d")
            last_modified_str = last_modified_dt.strftime("%Y-%m-%d %H:%M:%S")
            new_task = {
                "id": new_id,
                "title": title, "description": description, "created_date": created_date_str,
                "created_by": created_by, "due_date": due_date, "priority": priority, "category": category,
                "main_staff": main_staff, "assigned_to": assigned_to, "completed": False,
                "applied_vessel": applied_vessel, "rev": rev, "drawing_no": drawing_no,
                "link": link, "sdb_link": sdb_link, "request_no": request_no,
                "requested_date": requested_date, "date_started": date_started,
                "qtd_mhr": qtd_mhr, "actual_mhr": actual_mhr,
                "last_modified": last_modified_str, "modified_by": modified_by, "deleted": False
            }
            self._tasks.append(new_task) # Add to local list

            logging.info(f"Task {new_id} added successfully by {created_by} (autocommitted).")
            return new_id # Return the new ID on success

        except pymssql.Error as e:
            logging.error(f"Database error adding task: {str(e)}", exc_info=True)
            # No explicit rollback needed
            return None # Indicate failure

        except Exception as e:
            logging.error(f"Unexpected error adding task: {str(e)}", exc_info=True)
            # No explicit rollback needed
            return None # Indicate failure
    
    def get_filtered_tasks(self, show_completed: bool = False, category: Optional[str] = None, 
                          main_staff: Optional[str] = None, assigned_to: Optional[str] = None,
                          show_deleted: bool = False) -> List[Dict[str, Any]]:
        """Get tasks filtered by various criteria."""
        filtered_tasks = []
        
        for task in self._tasks:
            # Skip deleted tasks unless explicitly requested
            if task.get("deleted", False) and not show_deleted:
                continue
                
            # Apply other filters as before
            if not show_completed and task.get("completed", False):
                continue
            
            task_category = task.get("category", "")
            if category and category != "All":
                if not task_category or task_category.lower() != category.lower():
                    continue
                
            task_main_staff = task.get("main_staff", "")
            if main_staff and main_staff != "All":
                if not task_main_staff or task_main_staff.lower() != main_staff.lower():
                    continue
                
            task_assigned_to = task.get("assigned_to", "")
            if assigned_to and assigned_to != "All":
                if not task_assigned_to or task_assigned_to.lower() != assigned_to.lower():
                    continue
                
            filtered_tasks.append(task)
        
        return filtered_tasks
    
    def update_task(self, task_id: int, **kwargs) -> bool: # Return True/False
        """Update a task in the SQL Server database (autocommitted). Returns True on success, False on failure."""
        if not self.conn or not self.cursor:
             logging.error(f"Cannot update task {task_id}: No database connection.")
             return False
        try:
            # --- Input Validation and Formatting ---
            def validate_date(date_str):
                if date_str:
                    try:
                        # Ensure the date is in the correct format YYYY-MM-DD
                        datetime.datetime.strptime(date_str, "%Y-%m-%d")
                        return date_str
                    except ValueError:
                         logging.warning(f"Invalid date format during update: {date_str}. Setting to None.")
                         return None
                return None

            # Process date fields if present in kwargs
            if 'due_date' in kwargs:
                kwargs['due_date'] = validate_date(kwargs['due_date'])
            if 'requested_date' in kwargs:
                kwargs['requested_date'] = validate_date(kwargs['requested_date'])
            if 'date_started' in kwargs:
                kwargs['date_started'] = validate_date(kwargs['date_started'])

            # Convert any remaining date objects to strings (shouldn't happen if TaskDialog passes strings)
            for key, value in kwargs.items():
                if isinstance(value, datetime.date):
                    kwargs[key] = value.strftime("%Y-%m-%d")

            # Ensure MHR values are integers if present
            if 'qtd_mhr' in kwargs:
                 try:
                     kwargs['qtd_mhr'] = int(kwargs['qtd_mhr']) if kwargs['qtd_mhr'] is not None else None
                 except (ValueError, TypeError):
                     logging.warning(f"Invalid qtd_mhr value during update: {kwargs['qtd_mhr']}. Skipping update for this field.")
                     del kwargs['qtd_mhr'] # Remove invalid value
            if 'actual_mhr' in kwargs:
                 try:
                     kwargs['actual_mhr'] = int(kwargs['actual_mhr']) if kwargs['actual_mhr'] is not None else None
                 except (ValueError, TypeError):
                     logging.warning(f"Invalid actual_mhr value during update: {kwargs['actual_mhr']}. Skipping update for this field.")
                     del kwargs['actual_mhr'] # Remove invalid value


            # --- Metadata and Query Building ---
            now_dt = datetime.datetime.now() # Use datetime object for DB if appropriate
            # now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
            current_user = self._get_current_user()

            # Add last_modified and modified_by to kwargs for update
            kwargs['last_modified'] = now_dt
            kwargs['modified_by'] = current_user

            # Build the SQL SET clause dynamically
            set_clauses = []
            params = []
            for key, value in kwargs.items():
                # Ensure we don't try to update the ID itself
                if key != "id":
                    # Use correct placeholder syntax for pymssql
                    set_clauses.append(f"[{key}] = %s") # Bracket column names for safety
                    params.append(value)

            # Check if there's anything to update
            if not set_clauses:
                logging.warning(f"No valid fields provided for updating task {task_id}.")
                return True # No changes needed, consider it a success

            # Finalize the SQL query
            sql = f"UPDATE [dbo].[Tasks] SET {', '.join(set_clauses)} WHERE id = %s"
            params.append(task_id) # Add the task_id for the WHERE clause

            # --- Database Execution (autocommitted) ---
            self.cursor.execute(sql, params)
            rows_affected = self.cursor.rowcount # Check if the row was actually updated

            if rows_affected == 0:
                # This might mean the task_id doesn't exist or the values were the same
                logging.warning(f"Update task {task_id}: Task not found or no changes made in DB (Rows affected: 0).")
                # Check if task exists locally to differentiate
                if not self._find_task_by_id_local(task_id):
                     messagebox.showerror("Error", f"Task with ID {task_id} not found for update.")
                     return False
                 # If task exists but no rows affected, maybe values were identical? Treat as success.


            logging.info(f"Task {task_id} updated successfully by {current_user} (autocommitted).")
            return True # Indicate success

        except pymssql.Error as e:
            logging.error(f"Database error updating task {task_id}: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error updating task: {str(e)}")
            return False # Indicate failure

        except Exception as e:
            logging.error(f"Unexpected error updating task {task_id}: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Application Error", f"An unexpected error occurred during update: {str(e)}")
            return False # Indicate failure


    def _find_task_by_id(self, task_id: int) -> Optional[Dict[str, Any]]:
        """Find a task by ID directly from the SQL Server database."""
        if not self.conn or not self.cursor:
             logging.error(f"Cannot find task {task_id}: No database connection.")
             return None
        try:
            # SQL query to find a task by ID
            sql = "SELECT * FROM Tasks WHERE id = %s"
            self.cursor.execute(sql, (task_id,))
            row = self.cursor.fetchone()

            if row:
                # Convert row to a dictionary using column descriptions
                column_names = [desc[0] for desc in self.cursor.description]
                task = dict(zip(column_names, row))

                # Apply the same data type conversions as in _load_tasks
                for date_key in ['created_date', 'due_date', 'requested_date', 'date_started']:
                    if date_key in task and isinstance(task[date_key], (datetime.datetime, datetime.date)):
                        task[date_key] = task[date_key].strftime("%Y-%m-%d")
                    elif date_key in task and not isinstance(task[date_key], str):
                        task[date_key] = None
                if 'last_modified' in task and isinstance(task['last_modified'], datetime.datetime):
                    task['last_modified'] = task['last_modified'].strftime("%Y-%m-%d %H:%M:%S")
                elif 'last_modified' in task:
                    task['last_modified'] = None
                if 'completed' in task: task['completed'] = bool(task['completed'])
                if 'deleted' in task: task['deleted'] = bool(task['deleted'])
                else: task['deleted'] = False
                for num_key in ['qtd_mhr', 'actual_mhr']:
                    if num_key in task and task[num_key] is not None:
                        try: task[num_key] = int(task[num_key])
                        except (ValueError, TypeError): task[num_key] = 0
                    else: task[num_key] = 0

                return task
            else:
                logging.warning(f"Task with ID {task_id} not found in database.")
                return None
        except pymssql.Error as e:
            logging.error(f"Database error finding task by ID {task_id}: {str(e)}", exc_info=True)
            # Avoid showing messagebox here, return None
            return None
        except Exception as e:
            logging.error(f"Unexpected error finding task {task_id}: {str(e)}", exc_info=True)
            return None

    def _find_task_by_id_local(self, task_id: int) -> Optional[Dict[str, Any]]:
        """Find a task by ID in the local cache (_tasks list)."""
        for task in self._tasks:
            if task.get('id') == task_id:
                return task
        return None


    def reload_tasks(self):
        """Reload tasks from the database into the local cache."""
        logging.info("Reloading tasks from database...")
        # Re-establish connection if needed (optional, depends on connection stability)
        # self._connect()
        if self.conn:
            self._tasks = self._load_tasks()
            logging.info(f"Task reload complete. {len(self._tasks)} tasks loaded.")
            return self._tasks
        else:
            logging.error("Cannot reload tasks, no database connection.")
            self._tasks = []
            # Show error to user?
            messagebox.showerror("Connection Error", "Cannot reload tasks, database connection lost.")
            return []

    def begin_transaction(self):
        """Begin an explicit database transaction."""
        try:
            # Check if connection is valid before executing
            if self.conn and self.cursor:
                 # pymssql doesn't have an explicit BEGIN TRANSACTION command like standard SQL sometimes.
                 # Turning autocommit off implicitly starts a transaction.
                 # However, since we connect with autocommit=True, we need to manage this carefully.
                 # For pymssql, explicit transaction control might involve SET IMPLICIT_TRANSACTIONS OFF/ON
                 # or simply relying on the commit/rollback methods to define boundaries.
                 # Let's assume calling commit/rollback manages the scope when autocommit is True globally.
                 # A safer approach might be to temporarily disable autocommit if pymssql supports it per-session.
                 # For simplicity here, we'll just log the intention. The commit/rollback call IS the boundary.
                 logging.debug("Starting explicit transaction block (commit/rollback will define scope).")
            else:
                 logging.error("Cannot begin transaction: No database connection.")
                 raise pymssql.Error("No database connection to begin transaction.")
        except pymssql.Error as e:
            logging.error(f"Error attempting to manage transaction state: {e}", exc_info=True)
            raise

    def commit_transaction(self):
        """Commit the current explicit database transaction."""
        try:
             if self.conn:
                 self.conn.commit()
                 logging.debug("Explicit transaction committed.")
             else:
                 logging.error("Cannot commit transaction: No database connection.")
                 # Raise error? Or just log? Let's raise.
                 raise pymssql.Error("No database connection to commit transaction.")
        except pymssql.Error as e:
            logging.error(f"Error committing explicit transaction: {e}", exc_info=True)
            raise # Re-raise so the caller knows commit failed

    def rollback_transaction(self):
        """Rollback the current explicit database transaction."""
        try:
             if self.conn:
                 self.conn.rollback()
                 logging.debug("Explicit transaction rolled back.")
             else:
                 logging.error("Cannot rollback transaction: No database connection.")
                 # Don't raise here, as rollback is usually called during error handling
        except pymssql.Error as e:
            logging.error(f"Error rolling back explicit transaction: {e}", exc_info=True)
            # Don't re-raise during error handling rollback

    def get_sql_username(self):
        """Get the SQL username used for the database connection."""
        return self.user

    def recover_task(self, task_id: int) -> bool:
        """Recover a soft-deleted task (autocommitted)."""
        if not self.conn or not self.cursor:
             logging.error(f"Cannot recover task {task_id}: No database connection.")
             messagebox.showerror("Database Error", "No database connection available.")
             return False

        # Find task locally or in DB (logic remains same as previous version)
        task = self._find_task_by_id_local(task_id)
        task_from_db = None
        if not task or not task.get("deleted", False):
             task_from_db = self._find_task_by_id(task_id)
             if not task_from_db: logging.warning(f"Recover task {task_id}: Task not found."); return False
             if not task_from_db.get("deleted", False): logging.warning(f"Recover task {task_id}: Task is not deleted."); return False
             task = task_from_db # Use DB version

        try:
            now_dt = datetime.datetime.now()
            current_user = self._get_current_user()
            sql = "UPDATE [dbo].[Tasks] SET deleted = 0, last_modified = %s, modified_by = %s WHERE id = %s AND deleted = 1"
            params = (now_dt, current_user, task_id)

            self.cursor.execute(sql, params)
            rows_affected = self.cursor.rowcount
            # No explicit commit needed

            if rows_affected > 0:
                 logging.info(f"Task {task_id} recovered in DB by {current_user} (autocommitted).")
                 # Update local cache (logic remains same)
                 task["deleted"] = False
                 task["last_modified"] = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                 task["modified_by"] = current_user
                 if not self._find_task_by_id_local(task_id): self._tasks.append(task) # Add back if missing locally
                 return True
            else:
                 logging.warning(f"Recover task {task_id} affected 0 rows (autocommit). Task might not exist or was already recovered.")
                 # (Sync logic remains same)
                 db_task = self._find_task_by_id(task_id)
                 if db_task and not db_task.get('deleted'):
                      if task: # Update local if it exists
                         task['deleted'] = False
                         task['last_modified'] = db_task.get('last_modified')
                         task['modified_by'] = db_task.get('modified_by')
                      return True # State matches request
                 return False

        except pymssql.Error as e:
            logging.error(f"Database error recovering task {task_id}: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error recovering task: {str(e)}")
            return False
        except Exception as e:
             logging.error(f"Unexpected error recovering task {task_id}: {str(e)}", exc_info=True)
             # No explicit rollback needed
             messagebox.showerror("Application Error", f"An unexpected error occurred during recovery: {str(e)}")
             return False

    def permanently_delete_task(self, task_id: int) -> bool:
        """Permanently remove a task from the system (autocommitted)."""
        if not self.conn or not self.cursor:
             logging.error(f"Cannot permanently delete task {task_id}: No database connection.")
             messagebox.showerror("Database Error", "No database connection available.")
             return False
        try:
            # Delete from the database
            sql = "DELETE FROM [dbo].[Tasks] WHERE id = %s"
            self.cursor.execute(sql, (task_id,))
            rows_affected = self.cursor.rowcount
            # No explicit commit needed

            if rows_affected > 0:
                 logging.info(f"Task {task_id} permanently deleted from DB (autocommitted).")
                 # Remove from the local list
                 original_len = len(self._tasks)
                 self._tasks = [task for task in self._tasks if task["id"] != task_id]
                 if len(self._tasks) == original_len:
                      logging.warning(f"Task {task_id} permanently deleted from DB but was not found in local cache.")

                 # Remove from older deleted_tasks list if present
                 self._deleted_tasks = [t for t in self._deleted_tasks if t["id"] != task_id]
                 # Also consider removing from the main deleted_tasks_stack if it was there?
                 # This might interfere with undo if the permanent delete is meant to be undoable itself.
                 # Let's assume permanent delete is final and doesn't affect undo stack.

                 return True
            else:
                 logging.warning(f"Permanent delete for task {task_id} affected 0 rows. Task may not have existed.")
                 # Check if it existed locally
                 if self._find_task_by_id_local(task_id):
                      # It existed locally but not in DB? Sync issue. Remove locally.
                      self._tasks = [task for task in self._tasks if task["id"] != task_id]
                 return False # Didn't delete anything from DB

        except pymssql.Error as e:
            logging.error(f"Database error permanently deleting task {task_id}: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error permanently deleting task: {str(e)}")
            return False
        except Exception as e:
             logging.error(f"Unexpected error permanently deleting task {task_id}: {str(e)}", exc_info=True)
             # No explicit rollback needed
             messagebox.showerror("Application Error", f"An unexpected error occurred during permanent deletion: {str(e)}")
             return False


    def permanently_delete_all_deleted_tasks(self) -> int:
        """Permanently remove all soft-deleted tasks (autocommitted)."""
        if not self.conn or not self.cursor:
             logging.error("Cannot permanently delete all tasks: No database connection.")
             messagebox.showerror("Database Error", "No database connection available.")
             return 0
        try:
            # Delete all deleted tasks from the database
            sql = "DELETE FROM [dbo].[Tasks] WHERE deleted = 1"
            self.cursor.execute(sql)
            deleted_count = self.cursor.rowcount
            # No explicit commit needed

            if deleted_count > 0:
                 logging.info(f"Permanently deleted {deleted_count} tasks marked as deleted from DB (autocommitted).")
            else:
                 logging.info("No tasks marked as deleted found to permanently delete.")

            # Update local list regardless of DB count, to ensure consistency
            original_count = len(self._tasks)
            self._tasks = [task for task in self._tasks if not task.get("deleted", False)]
            local_removed_count = original_count - len(self._tasks)
            if local_removed_count != deleted_count:
                 logging.warning(f"DB deleted count ({deleted_count}) differs from local cache removed count ({local_removed_count}). Cache might have been out of sync.")

            # Clear the older deleted_tasks list
            self._deleted_tasks = []
            # Clear the main undo stack? This is debatable. If permanent delete is final, maybe clear it.
            # self.deleted_tasks_stack = [] # Uncomment if permanent delete should clear undo history

            return deleted_count
        except pymssql.Error as e:
            logging.error(f"Database error permanently deleting all tasks: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error permanently deleting tasks: {str(e)}")
            return 0
        except Exception as e:
             logging.error(f"Unexpected error permanently deleting all tasks: {str(e)}", exc_info=True)
             # No explicit rollback needed
             messagebox.showerror("Application Error", f"An unexpected error occurred during permanent deletion: {str(e)}")
             return 0

    @property
    def tasks(self):
        """Get the list of tasks. Use property for consistent access."""
        return self._tasks

    def recover_batch(self, task_ids: List[int]) -> List[int]:
        """Recover multiple tasks at once (each recovery is autocommitted)."""
        # This method calls self.recover_task which now uses autocommit.
        # Therefore, the overall batch recovery is NOT atomic.
        # If atomicity is required for batch recovery, this needs explicit transaction handling like batch_delete.
        # For now, we keep it simple: recover one by one.
        recovered_ids = []
        for task_id in task_ids:
            if self.recover_task(task_id): # recover_task now uses autocommit
                recovered_ids.append(task_id)
        return recovered_ids

    def delete_task(self, task_id: int) -> bool:
        """
        Soft delete a task by marking it as deleted (autocommitted).
        Updates local cache and undo stack. Returns True/False.
        """
        if not self.conn or not self.cursor:
            logging.error(f"Cannot delete task {task_id}: No database connection.")
            messagebox.showerror("Database Error", "No database connection available.")
            return False

        task_to_delete = self._find_task_by_id_local(task_id)
        task_copy = None
        if task_to_delete and not task_to_delete.get("deleted", False):
             task_copy = task_to_delete.copy()
             # Keep older _deleted_tasks list update if needed for compatibility
             self._deleted_tasks.append(task_copy)

        try:
            now_dt = datetime.datetime.now()
            current_user = self._get_current_user()

            sql = "UPDATE [dbo].[Tasks] SET deleted = 1, last_modified = %s, modified_by = %s WHERE id = %s AND deleted = 0"
            params = (now_dt, current_user, task_id)

            self.cursor.execute(sql, params)
            rows_affected = self.cursor.rowcount
            # No explicit commit needed

            if rows_affected > 0:
                logging.info(f"Task {task_id} marked as deleted in DB by {current_user} (autocommitted).")
                if task_to_delete: # Update local cache if it exists
                    task_to_delete["deleted"] = True
                    task_to_delete["last_modified"] = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                    task_to_delete["modified_by"] = current_user
                self.deleted_tasks_stack.append([task_id]) # Add to primary undo stack
                return True
            elif task_to_delete and task_to_delete.get("deleted"):
                 logging.warning(f"Task {task_id} was already marked as deleted.")
                 return True # State matches request
            else:
                 logging.warning(f"Soft delete for task {task_id} affected 0 rows (autocommit). Task might not exist or was already deleted.")
                 if task_to_delete and not task_to_delete.get("deleted"):
                      messagebox.showwarning("Sync Issue?", f"Could not mark task {task_id} as deleted in the database. It might have been deleted by another user. Refreshing data is recommended.")
                 return False # DB state not changed as expected

        except pymssql.Error as e:
            logging.error(f"Database error soft-deleting task {task_id}: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error deleting task: {str(e)}")
            if task_copy and task_copy in self._deleted_tasks: self._deleted_tasks.remove(task_copy)
            return False
        except Exception as e:
             logging.error(f"Unexpected error deleting task {task_id}: {str(e)}", exc_info=True)
             # No explicit rollback needed
             messagebox.showerror("Application Error", f"An unexpected error occurred during deletion: {str(e)}")
             if task_copy and task_copy in self._deleted_tasks: self._deleted_tasks.remove(task_copy)
             return False

    def batch_delete_tasks(self, task_ids: List[int]) -> Tuple[List[int], List[int]]:
        """
        Soft delete multiple tasks using an explicit transaction.
        Updates local cache and undo stack. Returns (succeeded_ids, failed_ids).
        """
        if not task_ids: return [], []
        if not self.conn or not self.cursor:
            logging.error("Cannot batch delete tasks: No database connection.")
            messagebox.showerror("Database Error", "No database connection available.")
            return [], list(task_ids) # Return all as failed

        success_ids = []
        failed_ids = [] # Start with empty failed
        tasks_to_undo = []

        # Prepare data for undo stack before starting transaction
        eligible_task_ids_for_delete = []
        for task_id in task_ids:
             task = self._find_task_by_id_local(task_id)
             if task and not task.get("deleted", False):
                  tasks_to_undo.append(task.copy())
                  eligible_task_ids_for_delete.append(task_id)
             else:
                 # Task doesn't exist locally or is already deleted
                 failed_ids.append(task_id)

        if not eligible_task_ids_for_delete:
             logging.warning("Batch delete: No eligible tasks found to delete.")
             return [], list(task_ids) # All originally passed IDs failed/skipped

        try:
            # Begin explicit transaction (overrides autocommit for this block)
            self.begin_transaction()

            now_dt = datetime.datetime.now()
            current_user = self._get_current_user()
            placeholders = ', '.join(['%s'] * len(eligible_task_ids_for_delete))

            sql = f"""
                UPDATE [dbo].[Tasks]
                SET deleted = 1, last_modified = %s, modified_by = %s
                WHERE id IN ({placeholders}) AND deleted = 0
            """
            params = [now_dt, current_user] + eligible_task_ids_for_delete

            self.cursor.execute(sql, params)
            rows_affected = self.cursor.rowcount

            # Commit the explicit transaction
            self.commit_transaction()

            logging.info(f"Batch delete transaction committed for {len(eligible_task_ids_for_delete)} eligible tasks. Rows affected in DB: {rows_affected}.")

            # If rows_affected doesn't match len(eligible_task_ids_for_delete), some might have been deleted between local check and DB update.
            # We'll consider the operation successful for those IDs passed to the DB if no error occurred.

            # Update local cache for successfully processed IDs
            now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
            for task_id in eligible_task_ids_for_delete:
                 task = self._find_task_by_id_local(task_id)
                 if task: # Should exist as we checked before
                     task['deleted'] = True
                     task['last_modified'] = now_str
                     task['modified_by'] = current_user
                 success_ids.append(task_id) # Add to success list

            # Add successful deletions to the undo stack
            if success_ids:
                self.deleted_tasks_stack.append(success_ids)
                self._deleted_tasks.extend(tasks_to_undo)

            logging.info(f"Batch delete result: {len(success_ids)} succeeded, {len(failed_ids)} failed/skipped.")

        except pymssql.Error as e:
            logging.error(f"Database error during batch delete transaction: {str(e)}", exc_info=True)
            try:
                self.rollback_transaction() # Rollback explicit transaction
            except pymssql.Error as rb_e:
                logging.error(f"Error during rollback after batch delete error: {rb_e}", exc_info=True)
            messagebox.showerror("Database Error", f"Error deleting tasks: {str(e)}")
            success_ids = [] # Reset success on error
            failed_ids = list(task_ids) # Mark all originally passed IDs as failed

        except Exception as e:
            logging.error(f"Unexpected error during batch delete transaction: {str(e)}", exc_info=True)
            try:
                 if self.conn: self.rollback_transaction()
            except Exception as rb_e:
                 logging.error(f"Error during rollback after unexpected batch delete error: {rb_e}", exc_info=True)
            messagebox.showerror("Application Error", f"An unexpected error occurred during batch deletion: {str(e)}")
            success_ids = []
            failed_ids = list(task_ids)

        return list(set(success_ids)), list(set(failed_ids))

    def recover_all_deleted_tasks(self) -> int:
        """
        Recover all soft-deleted tasks by setting their 'deleted' flag to 0 (autocommitted).
        Returns the number of tasks recovered in the database.
        """
        if not self.conn or not self.cursor:
             logging.error("Cannot recover all deleted tasks: No database connection.")
             messagebox.showerror("Database Error", "No database connection available.")
             return 0

        recovered_count_db = 0
        try:
            now_dt = datetime.datetime.now()
            current_user = self._get_current_user()

            # Update all deleted tasks in the database in one go
            sql = "UPDATE [dbo].[Tasks] SET deleted = 0, last_modified = %s, modified_by = %s WHERE deleted = 1"
            params = (now_dt, current_user)

            self.cursor.execute(sql, params)
            recovered_count_db = self.cursor.rowcount
            # No explicit commit needed due to autocommit=True

            if recovered_count_db > 0:
                 logging.info(f"Recovered {recovered_count_db} tasks marked as deleted in DB (autocommitted).")

                 # Update local cache
                 now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                 recovered_in_cache = 0
                 for task in self._tasks:
                     if task.get("deleted", False):
                         task["deleted"] = False
                         task["last_modified"] = now_str
                         task["modified_by"] = current_user
                         recovered_in_cache += 1
                 logging.info(f"Updated {recovered_in_cache} tasks in local cache during recover all.")
                 # Note: DB count and cache count might differ if cache was out of sync

            else:
                 logging.info("No tasks marked as deleted found to recover.")

            # Clear the older _deleted_tasks list (if still used)
            self._deleted_tasks = []
            # Clear the main undo stack? Recover all might be considered a bulk action that clears undo.
            # self.deleted_tasks_stack = [] # Uncomment if Recover All should clear undo history

            return recovered_count_db

        except pymssql.Error as e:
            logging.error(f"Database error recovering all deleted tasks: {str(e)}", exc_info=True)
            # No explicit rollback needed
            messagebox.showerror("Database Error", f"Error recovering all tasks: {str(e)}")
            return 0
        except Exception as e:
             logging.error(f"Unexpected error recovering all deleted tasks: {str(e)}", exc_info=True)
             # No explicit rollback needed
             messagebox.showerror("Application Error", f"An unexpected error occurred during recover all: {str(e)}")
             return 0


def main():
    """Main function to run the Task Manager application."""
    # Logging is now configured above, before main() is called

    print(f"--- Task Manager Application ---")
    print(f"Attempting to write logs to: {log_file_path}")
    print(f"Log file should be located in the same directory as the executable/script.")
    print(f"If issues occur, please check this log file.")
    print(f"------------------------------")

    logging.info("Application starting in main()...") # First log message in main

    root = tk.Tk()
    try:
        # Initialize TaskManager - connection happens here
        task_manager = TaskManager()

        app = TaskManagerApp(root, task_manager=task_manager)
        app.run() # This blocks until the window is closed
        # Code here runs *after* the window is closed, 'on_close' handles shutdown logging/flushing

    except ConnectionError as ce:
         logging.error(f"Fatal Error: Database connection failed on startup. {ce}", exc_info=False) 
         messagebox.showerror("Fatal Error", f"Database connection failed on startup:\n{ce}\n\nApplication cannot continue.")
         try: root.destroy()
         except: pass 
         # <<< Add explicit flush/close here for error exit path >>>
         logging.info("Flushing log handlers after connection error exit...")
         for handler in logging.getLogger().handlers:
             try:
                 handler.flush()
                 handler.close()
             except Exception as e:
                 print(f"Error closing handler {handler}: {e}")

    except Exception as e:
         logging.error("Unhandled exception in main", exc_info=True)
         messagebox.showerror("Fatal Error", f"An unexpected error occurred: {str(e)}\nCheck task_manager.log for details.")
         try: root.destroy()
         except: pass
         # <<< Add explicit flush/close here for general error exit path >>>
         logging.info("Flushing log handlers after general error exit...")
         for handler in logging.getLogger().handlers:
             try:
                 handler.flush()
                 handler.close()
             except Exception as e:
                 print(f"Error closing handler {handler}: {e}")

if __name__ == "__main__":
    main()
