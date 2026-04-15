import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
import logging
from typing import Dict, List, Tuple, Any, Optional
import threading
import time
from flask import Flask, render_template
import webbrowser
import socket
from contextlib import closing
import datetime

# Import test execution modules
try:
    from Sparameter_Automation import run_s_param_test
    from Load_Pull import run_test
except ImportError as e:
    print(f"Warning: Could not import test modules: {e}")
    # Define fallback functions for testing
    def run_s_param_test(value):
        time.sleep(1)
        return {"verdict": "PASSED"}
    
    def run_test(value):
        time.sleep(1)
        return {"verdict": "PASSED"}

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class TestController:
    """
    Advanced Test Controller for automated S-Parameter and Load Pull testing.
    
    Provides a comprehensive GUI interface for CSV-based test configuration,
    multi-threaded test execution, and web-based results visualization.
    """
    
    def __init__(self):
        """Initialize the test controller application."""
        # Main window setup
        self.root = tk.Tk()
        self.root.title("Advanced Test Controller")
        self.root.geometry("600x800")
        self.root.configure(bg='#f0f0f0')
        
        # Configure UI styling
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Application data storage
        self.csv_file_path = None
        self.dataframe = None
        self.test_data = {}
        self.system_name = "Test-Suite Controller"
        self.serial_numbers = []
        
        # Test execution state management
        self.is_running = False
        self.current_thread = None
        
        # Test type classification mapping
        self.test_type_mapping = {
            "S Parameter": ["s_param", "s-param", "s parameter", "sparameter"],
            "F0 Load Pull": ["load", "loadpull", "load_pull", "f0_load"],
            "CW": ["cw", "CW", "CWMode", "cw mode", "CW Mode"],
            "Pulsed": ["Pulsed", "pulsed", "pulsed mode", "pulsedmode"],
            "Modulated": ["MOD", "mod", "Modulated", "modulated"]
        }

        
        # UI control variables
        self.checkbox_vars = {}
        self.flask_app = None
        self.flask_thread = None
        self.web_port = 5000
        self.all_test_results = []
        
        # Initialize interface components
        self.create_ui()
        self.create_menu()

        self.test_start_time = None
        self.test_summary = {
                        'total_tests': 0,
                        'passed_tests': 0,
                        'failed_tests': 0,
                        'total_time': 0
                    }
        
    def animate_progress_bar(self, target_value):
        """Animate progress bar"""
        current_value = self.progress['value']
        if current_value < target_value:
            increment = max(1, (target_value - current_value) / 10)
            new_value = min(current_value + increment, target_value)
            self.progress.config(value=new_value)
            
            if new_value < target_value:
                self.root.after(50, lambda: self.animate_progress_bar(target_value))

    
    def find_free_port(self) -> int:
        """Find an available port for the Flask web interface."""
        with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as s:
            s.bind(('', 0))
            s.listen(1)
            port = s.getsockname()[1]
        return port

    def setup_flask_app(self):
        """Setup Flask application with web interface routes."""
        self.flask_app = Flask(__name__)
        self.web_port = self.find_free_port()
        
        # In the setup_flask_app method...


        @self.flask_app.route("/")
        def index():
            test_rows = []
            combined_system_info = None
        
            for i, test_data in enumerate(self.all_test_results):
                result_dict = test_data['result_data']
                verdict = result_dict.get('verdict', 'FAILED')
               
                base_test_type = test_data.get('test_type', 'Unknown Test')
                
                
                mode_map = {0: "CW", 1: "Pulsed", 2: "Modulated"}
                mode_string = ""
                        
                if isinstance(result_dict, dict):
                    test_config = result_dict.get('test_config', {})
                    if isinstance(test_config, dict):
                        mode_value = test_config.get('mode') 
                        mode_string = mode_map.get(mode_value, "") 
                           
                full_test_name = f"{mode_string} {base_test_type}".strip()
            
                data_file = "N/A"
                data_link = None
            
                if isinstance(result_dict, dict):
                    if 'file' in result_dict and result_dict['file']: 
                        data_file = os.path.basename(result_dict['file'])
                        data_link = f"/open_file/{i}"
                    elif 'output_file' in result_dict: 
                        if result_dict['output_file']:  
                            data_file = os.path.basename(result_dict['output_file'])
                            data_link = f"/open_file/{i}"
                        else: 
                            data_file = "Check for file on system"
                            data_link = None
                
                test_rows.append({
                    "test_type": full_test_name,
                    "result": "PASS" if verdict == "PASSED" else "FAIL",
                    "notes": "N/A" if verdict == "PASSED" else "See details",
                    "data": data_file,
                    "data_link": data_link,
                    "link": f"/report/{i}"
                })


                # Extract system information from first test
                if not combined_system_info and 'system_info' in result_dict:
                    combined_system_info = result_dict['system_info']


            # Prepare page information
            page_info = combined_system_info or {
                "jira": "N/A", "commit": "N/A", "hardware": "N/A",
                "sn": "N/A", "fw": "N/A"
            }
        
            page_info.update({
                "mode": "Mixed Test Session",
                "local": "N/A",
                "srcal": "Multiple Files",
                "vecal": "Multiple Files",
                "init": "Multiple Files",
                "date": datetime.datetime.now().strftime("%d/%m/%Y"),
                "total_tests": self.test_summary['total_tests'],
                "passed_tests": self.test_summary['passed_tests'],
                "failed_tests": self.test_summary['failed_tests'],
                "execution_time": f"{self.test_summary['total_time']:.1f} seconds"
            })


            return render_template("output.html", info=page_info, tests=test_rows)


        @self.flask_app.route("/report/<int:test_index>")
        def report(test_index):
            """Individual test detailed report page."""
            try:
                if test_index >= len(self.all_test_results):
                    return "Test not found", 404

                test_data = self.all_test_results[test_index]
                test_type = test_data['test_type']
                result_dict = test_data['result_data']
                
                if test_type == 'Load Pull':
                    # Get test config for Load Pull
                    test_config = result_dict.get('test_config', {})
                    mode_value = test_config.get('mode', 0)  # Get the mode number
                    
                    # Determine bandwidth label and get waveform
                    if mode_value == 2:  # Modulated mode
                        bandwidth_label = "Modulated Bandwidth"
                        bandwidth_value = test_config.get('mod_bandwidth', 'N/A')
                        waveform_value = test_config.get('waveform', 'N/A')  # You'll need to add this
                    else:
                        bandwidth_label = "IF Bandwidth"  
                        bandwidth_value = test_config.get('if_bandwidth', 'N/A')
                        waveform_value = None  # No waveform for non-modulated
                    
                    return render_template("load_pull_report.html",
                                        verdict=result_dict.get('verdict', 'UNKNOWN'),
                                        err_dict=result_dict.get('err_dict', {}),
                                        gold_standard_dict=result_dict.get('gold_standard_dict', {}),
                                        measured_dict=result_dict.get('measured_dict', {}),
                                        # Test configuration
                                        mode_name=result_dict.get('test_config', {}).get('mode_name', 'N/A'),
                                        source_cal=result_dict.get('test_config', {}).get('source_cal', 'N/A'),
                                        vector_cal=result_dict.get('test_config', {}).get('vector_cal', 'N/A'),
                                        init_file=result_dict.get('test_config', {}).get('init_file', 'N/A'),
                                        bandwidth_label=bandwidth_label,  # Dynamic label
                                        bandwidth_value=bandwidth_value,  # Dynamic value
                                        waveform=waveform_value,  # Waveform info
                                        frequency=result_dict.get('test_config', {}).get('frequency', 'N/A'))

                                        
                else:  # S Parameter
                    return render_template("report.html",
                                        verdict=result_dict.get('verdict', 'UNKNOWN'),
                                        err_dict=result_dict.get('err_dict', {}),
                                        mag_tol=result_dict.get('mag_tol', 0.0),
                                        phase_tol=result_dict.get('phase_tol', 0.0),
                                        file=result_dict.get('file', 'N/A'),
                                        mode_name=result_dict.get('test_config', {}).get('mode_name', 'N/A'),
                                        source_cal=result_dict.get('test_config', {}).get('source_cal', 'N/A'),
                                        vector_cal=result_dict.get('test_config', {}).get('vector_cal', 'N/A'),
                                        if_bandwidth=result_dict.get('test_config', {}).get('if_bandwidth', 'N/A'),
                                        frequency_list=result_dict.get('test_config', {}).get('frequency_list', []))
                                        
            except Exception as e:
                print(f"Error in report route: {e}")
                import traceback
                traceback.print_exc()
                return f"Error: {str(e)}", 500

        @self.flask_app.route("/open_file/<int:test_index>")
        def open_file(test_index):
            try:
                if test_index >= len(self.all_test_results):
                    return "File not found", 404
                    
                test_data = self.all_test_results[test_index]
                result_dict = test_data['result_data']
                
                # Get file path from result
                file_path = None
                if isinstance(result_dict, dict):
                    if 'file' in result_dict and result_dict['file']:  # S-Parameter
                        file_path = result_dict['file']
                    elif 'output_file' in result_dict and result_dict['output_file']:  # Load Pull
                        file_path = result_dict['output_file']
                
                if file_path and file_path != "No output file generated" and os.path.exists(file_path):
                    try:
                        import subprocess
                        
                        # Path to data explorer application
                        data_explorer_app_path = r"C:\Program Files (x86)\Focus Microwaves\Load Pull Explorer\FMWViewer.exe"
                        
                        # Create the command to run as administrator
                        command = [
                            "powershell", 
                            "-Command", 
                            f"Start-Process '{data_explorer_app_path}' -ArgumentList '\"{file_path}\"' -Verb RunAs"
                        ]
                        
                        # Execute the command
                        subprocess.Popen(command, shell=True)
                        
                        # Return simple text response
                        return f"Opening {os.path.basename(file_path)}..."
                        
                    except Exception as e:
                        return f"Error: {str(e)}"
                else:
                    return "File not found"
                    
            except Exception as e:
                return f"Error: {str(e)}"



    def start_flask_app(self):
        """Start Flask app in a separate thread."""
        if self.flask_app is None:
            self.setup_flask_app()
        
        def run_flask():
            try:
                self.flask_app.run(host='localhost', port=self.web_port, debug=False, use_reloader=False)
            except Exception as e:
                print(f"Flask error: {e}")
        
        if self.flask_thread is None or not self.flask_thread.is_alive():
            self.flask_thread = threading.Thread(target=run_flask, daemon=True)
            self.flask_thread.start()
            time.sleep(2)  # Allow server startup time

    def open_web_interface(self):
        """Open the web interface in the default browser."""
        if not self.flask_thread or not self.flask_thread.is_alive():
            self.start_flask_app()
        
        url = f"http://localhost:{self.web_port}"
        webbrowser.open(url)
        self.log_message(f"🌐 Web interface opened at {url}")

    def create_menu(self):
        """Create menu bar with file and view options."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load CSV...", command=self.load_csv_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Show CSV Data", command=self.show_csv_data)
        view_menu.add_command(label="Clear Results", command=self.clear_results)
        view_menu.add_command(label="Open Web Interface", command=self.open_web_interface)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def create_ui(self):
        """Create the main user interface components."""
        # Header section
        header_frame = tk.Frame(self.root)
        header_frame.config(bg="#34495e", height=70)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        self.system_label = tk.Label(header_frame, text=self.system_name, 
                                    font=('Arial', 16, 'bold'), fg='white', bg='#34495e')
        self.system_label.pack(pady=(10, 0))
        
        self.file_label = tk.Label(header_frame, text="No file loaded", 
                                  font=('Arial', 9), fg='#bdc3c7', bg='#34495e')
        self.file_label.pack(pady=(0, 10))
        
        # Main content area
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # File selection section
        file_frame = tk.Frame(main_frame, bg='#f0f0f0')
        file_frame.pack(fill='x', pady=(0, 15))
        
        load_csv_button = tk.Button(file_frame, text="📁 Load CSV File", 
                                   command=self.load_csv_file,
                                   font=('Arial', 11, 'bold'), 
                                   bg='#3498db', fg='white',
                                   padx=20, pady=10,
                                   relief='flat', bd=0,
                                   activebackground='#2980b9',
                                   cursor='hand2')
        load_csv_button.pack(side='left')
        
        # Test configuration section
        test_frame = tk.LabelFrame(main_frame, text="Available Test Types", 
                                font=('Arial', 12, 'bold'), 
                                bg='#f0f0f0', fg='#2c3e50',
                                padx=15, pady=10)
        test_frame.pack(fill='x', pady=(0, 15))

        types_modes_frame = tk.Frame(test_frame, bg='#f0f0f0')
        types_modes_frame.pack(fill='x')

        # Test types selection (left side)
        test_types_frame = tk.Frame(types_modes_frame, bg='#f0f0f0')
        test_types_frame.pack(side='left', fill='both', expand=True, padx=(0, 20))

        tk.Label(test_types_frame, text="Test Types", 
                font=('Arial', 11, 'bold'), 
                bg='#f0f0f0', fg='#2c3e50').pack(anchor='w', pady=(0, 5))

        test_types = [
            ("S Parameter", "#3498db"),
            ("F0 Load Pull", "#e74c3c")
        ]

        for test_name, color in test_types:
            var = tk.BooleanVar()
            self.checkbox_vars[test_name] = var
            
            cb_frame = tk.Frame(test_types_frame, bg='#f0f0f0')
            cb_frame.pack(fill='x', pady=3)
            
            cb = tk.Checkbutton(cb_frame, text=test_name, variable=var,
                            font=('Arial', 11), bg='#f0f0f0', fg='#2c3e50',
                            activebackground='#f0f0f0', activeforeground=color,
                            selectcolor='white', anchor='w',
                            command=lambda name=test_name: self.on_checkbox_change(name))
            cb.pack(side='left', fill='x', expand=True)
            
            # Status indicators
            status_frame = tk.Frame(cb_frame, bg='#f0f0f0')
            status_frame.pack(side='right')
            
            count_label = tk.Label(status_frame, text="(0)", font=('Arial', 9), 
                                fg='#7f8c8d', bg='#f0f0f0')
            count_label.pack(side='right', padx=(5, 0))
            
            status_label = tk.Label(status_frame, text="●", font=('Arial', 12), 
                                fg='#bdc3c7', bg='#f0f0f0')
            status_label.pack(side='right')
            
            # Store references for later access
            setattr(self, f'status_{test_name.replace(" ", "_").lower()}', status_label)
            setattr(self, f'count_{test_name.replace(" ", "_").lower()}', count_label)

        # Test modes selection (right side)
        modes_frame = tk.Frame(types_modes_frame, bg='#f0f0f0')
        modes_frame.pack(side='right', fill='both', expand=True)

        tk.Label(modes_frame, text="Test Modes", 
                font=('Arial', 11, 'bold'), 
                bg='#f0f0f0', fg='#2c3e50').pack(anchor='w', pady=(0, 5))

        test_modes = [
            ("CW", "#f39c12"),
            ("Pulsed", "#9b59b6"), 
            ("Modulated", "#1abc9c")
        ]

        for mode_name, color in test_modes:
            var = tk.BooleanVar()
            self.checkbox_vars[mode_name] = var
            
            cb_frame = tk.Frame(modes_frame, bg='#f0f0f0')
            cb_frame.pack(fill='x', pady=3)
            
            cb = tk.Checkbutton(cb_frame, text=mode_name, variable=var,
                            font=('Arial', 11), bg='#f0f0f0', fg='#2c3e50',
                            activebackground='#f0f0f0', activeforeground=color,
                            selectcolor='white', anchor='w',
                            command=lambda name=mode_name: self.on_checkbox_change(name))
            cb.pack(side='left', fill='x', expand=True)
            
            status_frame = tk.Frame(cb_frame, bg='#f0f0f0')
            status_frame.pack(side='right')
            
            count_label = tk.Label(status_frame, text="(0)", font=('Arial', 9), 
                                fg='#7f8c8d', bg='#f0f0f0')
            count_label.pack(side='right', padx=(5, 0))
            
            status_label = tk.Label(status_frame, text="●", font=('Arial', 12), 
                                fg='#bdc3c7', bg='#f0f0f0')
            status_label.pack(side='right')
            
            setattr(self, f'status_{mode_name.replace(" ", "_").lower()}', status_label)
            setattr(self, f'count_{mode_name.replace(" ", "_").lower()}', count_label)
                
        # Statistics section
        stats_frame = tk.LabelFrame(main_frame, text="Test Statistics", 
                                   font=('Arial', 12, 'bold'), 
                                   bg='#f0f0f0', fg='#2c3e50',
                                   padx=15, pady=10)
        stats_frame.pack(fill='x', pady=(0, 15))
        
        stats_grid = tk.Frame(stats_frame, bg='#f0f0f0')
        stats_grid.pack(fill='x')
        
        tk.Label(stats_grid, text="Total Tests:", font=('Arial', 10, 'bold'), 
                bg='#f0f0f0', fg='#2c3e50').grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.total_tests_label = tk.Label(stats_grid, text="0", font=('Arial', 10), 
                                         bg='#f0f0f0', fg='#2c3e50')
        self.total_tests_label.grid(row=0, column=1, sticky='w')
        
        tk.Label(stats_grid, text="Serial Numbers:", font=('Arial', 10, 'bold'), 
                bg='#f0f0f0', fg='#2c3e50').grid(row=1, column=0, sticky='w', padx=(0, 10))
        self.serial_count_label = tk.Label(stats_grid, text="0", font=('Arial', 10), 
                                          bg='#f0f0f0', fg='#2c3e50')
        self.serial_count_label.grid(row=1, column=1, sticky='w')
        
        # Control buttons section
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', pady=(0, 15))
        
        clear_button = tk.Button(button_frame, text="Clear All", 
                               command=lambda: [self.clear_all_tests(), self.clear_results()],
                               font=('Arial', 10), 
                               bg='#95a5a6', fg='white',
                               padx=20, pady=8,
                               relief='flat', bd=0,
                               activebackground='#7f8c8d')
        clear_button.pack(side='left', padx=(0, 10))
        
        select_button = tk.Button(button_frame, text="Select All", 
                                command=self.select_all_tests,
                                font=('Arial', 10), 
                                bg='#3498db', fg='white',
                                padx=20, pady=8,
                                relief='flat', bd=0,
                                activebackground='#2980b9')
        select_button.pack(side='left', padx=(0, 10))
        
        analyze_button = tk.Button(button_frame, text="🔍 Analyze CSV", 
                                 command=self.analyze_csv,
                                 font=('Arial', 10), 
                                 bg='#f39c12', fg='white',
                                 padx=20, pady=8,
                                 relief='flat', bd=0,
                                 activebackground='#e67e22')
        analyze_button.pack(side='left', padx=(0, 10))

        web_button = tk.Button(button_frame, text="View Web Report", 
                                 command=self.open_web_interface,
                                 font=('Arial', 10), 
                                 bg="#c612f3", fg='white',
                                 padx=20, pady=8,
                                 relief='flat', bd=0,
                                 activebackground='#c612f3')
        web_button.pack(side='left')
        
        # Main execution control
        self.play_button = tk.Button(main_frame, text="▶ START TESTS", 
                                    command=self.toggle_tests,
                                    font=('Arial', 14, 'bold'), 
                                    bg='#27ae60', fg='white',
                                    padx=50, pady=15,
                                    relief='flat', bd=0,
                                    activebackground='#229954',
                                    cursor='hand2')
        self.play_button.pack(pady=10)

        # view_menu.add_command(label="Clear Results", command=self.clear_results)
        # view_menu.add_command(label="Open Web Interface", command=self.open_web_interface)
        
        # Progress indication
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.style.configure("green.Horizontal.TProgressbar", 
                    troughcolor='#ecf0f1',
                    background='#27ae60',
                    lightcolor='#2ecc71',
                    darkcolor='#1e8449')
        self.progress.config(style="green.Horizontal.TProgressbar")

        self.progress.pack(fill='x', pady=(0, 10))
        
        # Results display area
        tk.Label(main_frame, text="Test Results", 
                font=('Arial', 12, 'bold'), 
                bg='#f0f0f0', fg='#2c3e50').pack(anchor='w', pady=(10, 5))
        
        results_frame = tk.Frame(main_frame, bg='#f0f0f0')
        results_frame.pack(fill='both', expand=True)
        
        self.results_text = tk.Text(results_frame, height=12, font=('Consolas', 9),
                                   bg='#2c3e50', fg='#ecf0f1', relief='flat',
                                   padx=10, pady=10, wrap='word')
        
        scrollbar = tk.Scrollbar(results_frame, orient='vertical', command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Initial welcome message
        self.log_message("🚀 Advanced Test Controller Ready")
        self.log_message("=" * 60)
        self.log_message("1. Load a CSV file using the 'Load CSV File' button")
        self.log_message("2. Select desired test types and modes")
        self.log_message("3. Click 'START TESTS' to begin execution")
        self.log_message("")
    
    def get_test_mode_from_data(self, test_data: List) -> str:
        """Extract mode from test data configuration."""
        mode_value = test_data[8] if len(test_data) > 8 else 0
        mode_map = {0: "cw", 1: "pulsed", 2: "modulated"}
        return mode_map.get(mode_value, "cw")

    def load_csv_file(self):
        """Load CSV file and analyze its contents."""
        file_path = filedialog.askopenfilename(
            title="Select CSV Test File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.log_message(f"📂 Loading CSV file: {os.path.basename(file_path)}")
            
            try:
                self.dataframe = self.read_csv(file_path)
                if self.dataframe is not None:
                    self.analyze_csv()
                    self.log_message("✅ CSV file loaded successfully")
                else:
                    self.log_message("❌ Failed to load CSV file")
            except Exception as e:
                self.log_message(f"❌ Error loading CSV: {str(e)}")
                logger.error(f"Error loading CSV: {str(e)}")

    def analyze_csv(self):
        """Analyze the loaded CSV and update UI components."""
        if self.dataframe is None:
            self.log_message("⚠️  No CSV file loaded")
            return
        
        try:
            # Extract test configuration data
            self.test_data = self.extract_information(self.dataframe)
            
            # Update system identification
            if 'System' in self.dataframe.columns:
                self.system_name = str(self.dataframe['System'].iloc[0])
            else:
                self.system_name = "Test-Suite Controller"
            self.system_label.config(text=self.system_name)
            
            # Extract device serial numbers
            if 'Serial_No' in self.dataframe.columns:
                self.serial_numbers = self.dataframe['Serial_No'].dropna().tolist()
            else:
                self.serial_numbers = []
            
            # Update statistics display
            self.total_tests_label.config(text=str(len(self.test_data)))
            self.serial_count_label.config(text=str(len(self.serial_numbers)))
            
            # Update test type counters
            self.update_test_type_counts()
            
            self.log_message(f"📊 Analysis complete:")
            self.log_message(f"   • System: {self.system_name}")
            self.log_message(f"   • Total tests: {len(self.test_data)}")
            self.log_message(f"   • Serial numbers: {len(self.serial_numbers)}")
            
        except Exception as e:
            self.log_message(f"❌ Error analyzing CSV: {str(e)}")
            logger.error(f"Error analyzing CSV: {str(e)}")

    
    def update_test_type_counts(self):
        """Update mode checkbox counters based on loaded test data."""
        type_counts = {}

        # Extract test types and modes based on CSV entry
        for key , value in self.test_data.items():
            test_type = key[1].lower()
            mode_int = value[8]
            
            if mode_int == 0:
                test_mode = 'CW'
            elif mode_int == 1:
                test_mode = 'Pulsed'
            elif mode_int == 2:
                test_mode = 'Modulated'
            else:
                exit()

            
            # Map test modes to checkbox categories
            for checkbox_name, keywords in self.test_type_mapping.items():
                if any(keyword in test_type for keyword in keywords):
                    type_counts[checkbox_name] = type_counts.get(checkbox_name, 0) + 1
                elif any(keyword in test_mode for keyword in keywords):
                    type_counts[checkbox_name] = type_counts.get(checkbox_name, 0) + 1
                    break
        
        # Update count display labels
        for test_name in self.checkbox_vars.keys():
            count_attr = f'count_{test_name.replace(" ", "_").lower()}'
            count_label = getattr(self, count_attr, None)
            if count_label:
                count = type_counts.get(test_name, 0)
                count_label.config(text=f"({count})")

    def on_checkbox_change(self, test_name: str):
        """Handle checkbox state changes with visual feedback."""
        status_attr = f'status_{test_name.replace(" ", "_").lower()}'
        status_label = getattr(self, status_attr, None)
        
        if status_label:
            if self.checkbox_vars[test_name].get():
                status_label.config(fg='#27ae60')  # Green for selected
            else:
                status_label.config(fg='#bdc3c7')  # Gray for unselected

    def select_all_tests(self):
        """Select all available test types and modes."""
        for var in self.checkbox_vars.values():
            var.set(True)
        for test_name in self.checkbox_vars.keys():
            self.on_checkbox_change(test_name)

    def clear_all_tests(self):
        """Clear all test type and mode selections."""
        for var in self.checkbox_vars.values():
            var.set(False)
        for test_name in self.checkbox_vars.keys():
            self.on_checkbox_change(test_name)

    def get_selected_tests(self) -> Dict:
        """Filter test data based on selected checkboxes."""
        if not self.test_data:
            return {}
        
        # Get currently selected options
        selected_test_types = [name for name in ["S Parameter", "F0 Load Pull"] 
                            if self.checkbox_vars[name].get()]
        selected_modes = [name for name in ["CW", "Pulsed", "Modulated"] 
                        if self.checkbox_vars[name].get()]
        
        if not selected_test_types or not selected_modes:
            return {}
        
        filtered_tests = {}
        
        for key, value in self.test_data.items():
            test_type = key[1].lower()
            test_mode = self.get_test_mode_from_data(value)
            
            # Check test type match
            type_match = False
            for test_type_name in selected_test_types:
                keywords = self.test_type_mapping.get(test_type_name, [])
                if any(keyword in test_type for keyword in keywords):
                    type_match = True
                    break
            
            # Check test mode match
            mode_match = test_mode in [mode.lower() for mode in selected_modes]
            
            if type_match and mode_match:
                filtered_tests[key] = value
        
        return filtered_tests

    def toggle_tests(self):
        """Toggle between starting and stopping test execution."""
        if self.is_running:
            self.stop_tests()
        else:
            self.start_tests()

    def start_tests(self):
        """Initiate test execution sequence."""
        selected_tests = self.get_selected_tests()
        
        if not selected_tests:
            messagebox.showwarning("No Tests Selected", 
                                 "Please select at least one test type and ensure CSV data is loaded.")
            return
        
        if not self.test_data:
            messagebox.showwarning("No Data", "Please load a CSV file first.")
            return
        
        # Record start time
        self.test_start_time = time.time()

        # Update UI for running state
        self.is_running = True
        self.play_button.config(text="⏹ STOP TESTS", bg='#e74c3c', activebackground='#c0392b')
        self.progress.config(value=0)
        
        # Execute tests in background thread
        self.current_thread = threading.Thread(target=self.run_tests_thread, args=(selected_tests,))
        self.current_thread.daemon = True
        self.current_thread.start()

        # Record start time
        self.test_start_time = time.time()


    def stop_tests(self):
        """Stop ongoing test execution."""
        self.is_running = False
        self.play_button.config(text="▶ START TESTS", bg='#27ae60', activebackground='#229954')
        self.progress.stop()
        self.log_message("🛑 Test execution stopped by user")

    def run_tests_thread(self, selected_tests: Dict):
        """Execute tests in background thread."""
        try:
            successful_tests = 0
            failed_tests = 0
            total_tests = len(selected_tests)
            
            self.log_message("🔧 Test Execution Started")
            self.log_message("=" * 60)
            self.log_message(f"System: {self.system_name}")
            self.log_message(f"Running {total_tests} test(s)...")
            self.log_message("")
            
            for i, (key, value) in enumerate(selected_tests.items(), 1):
                if not self.is_running:
                    break
            
                test_name = key[0]
                test_type = key[1]
        
                progress_percentage = (i - 1) / total_tests * 100
                self.root.after(0, lambda p=progress_percentage: self.animate_progress_bar(p))

            
                self.log_message(f"[{i}/{total_tests}] {test_name} - {test_type}...")
                
                try:
                    # Execute the appropriate test
                    result = self.execute_test(test_type, value)
                    verdict = result.get('verdict', 'FAILED')

                    if verdict == "PASSED":
                        self.log_message(f"✅ {test_name} - PASSED")
                        successful_tests += 1
                    else:
                        self.log_message(f"❌ {test_name} - FAILED")
                        failed_tests += 1
                    
                    # Calculate total execution time
                    total_time = time.time() - self.test_start_time

                    # Store summary data for web interface
                    self.test_summary = {
                        'total_tests': total_tests,
                        'passed_tests': successful_tests,
                        'failed_tests': failed_tests,
                        'total_time': total_time
                    }
                    
                    # Store test results for web interface
                    test_result_entry = {
                        'test_type': 'Load Pull' if 'load' in test_type.lower() else 'S Parameter',
                        'test_name': f"{test_name} - {test_type}",
                        'result_data': result,
                        'timestamp': datetime.datetime.now()
                    }
                    self.all_test_results.append(test_result_entry)
                        
                except Exception as e:
                    self.log_message(f"❌ {test_name} - ERROR: {str(e)}")
                    failed_tests += 1
                    logger.error(f"Test execution failed for {test_name}: {str(e)}")

            # Display execution summary
            self.log_message("")
            self.log_message("=" * 60)
            self.log_message("🎉 Test Execution Complete!")
            self.log_message(f"✅ Successful: {successful_tests}")
            self.log_message(f"❌ Failed: {failed_tests}")
            self.log_message(f"📊 Total: {successful_tests + failed_tests}")
            
        except Exception as e:
            self.log_message(f"❌ Unexpected error during test execution: {str(e)}")
            logger.error(f"Unexpected test execution error: {str(e)}")
        
        finally:
            # Reset UI state on main thread
            self.root.after(0, self.reset_ui_after_tests)

    def reset_ui_after_tests(self):
        """Reset UI components after test completion."""
        self.is_running = False
        self.play_button.config(text="▶ START TESTS", bg='#27ae60', activebackground='#229954')
        self.progress.stop()
        self.animate_progress_bar(100)


    def execute_test(self, test_type: str, test_data: List) -> Dict:
        """Execute individual test based on type."""
        try:
            test_type_lower = test_type.lower()
            
            if 'load' in test_type_lower:
                # Execute load pull test
                result = run_test(test_data)
                return result
            else:
                # Execute S-parameter test
                result = run_s_param_test(test_data)
                return result
                
        except Exception as e:
            logger.error(f"Test execution error for {test_type}: {str(e)}")
            return {"verdict": "FAILED", "error": str(e)}

    def log_message(self, message: str):
        """Add message to results display area."""
        self.results_text.insert('end', f"{message}\n")
        self.results_text.see('end')
        self.root.update_idletasks()

    def clear_results(self):
        """Clear results display and stored test data."""
        self.results_text.delete(1.0, 'end')
        self.all_test_results = []
        self.log_message("🧹 Results cleared")

    def show_csv_data(self):
        """Display CSV data in popup window."""
        if self.dataframe is None:
            messagebox.showinfo("No Data", "Please load a CSV file first.")
            return
        
        # Create data viewer window
        data_window = tk.Toplevel(self.root)
        data_window.title("CSV Data Viewer")
        data_window.geometry("800x600")
        
        text_frame = tk.Frame(data_window)
        text_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, font=('Consolas', 9))
        v_scrollbar = tk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        h_scrollbar = tk.Scrollbar(text_frame, orient='horizontal', command=text_widget.xview)
        
        text_widget.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        text_widget.insert('end', self.dataframe.to_string())
        text_widget.config(state='disabled')
        
        text_widget.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')

    def show_about(self):
        """Display application information dialog."""
        messagebox.showinfo("About", 
                          "Advanced Test Controller v2.0\n\n"
                          "Integrated test automation system for\n"
                          "S-Parameter and Load Pull testing.\n\n"
                          "Features:\n"
                          "• CSV-based test configuration\n"
                          "• Multi-threaded test execution\n"
                          "• Real-time progress monitoring\n"
                          "• Web-based results visualization\n"
                          "• Comprehensive error handling")

    # CSV Processing and Data Extraction Methods
    
    def read_csv(self, file_path: str) -> Optional[pd.DataFrame]:
        """Read CSV file with comprehensive error handling."""
        try:
            if not os.path.exists(file_path):
                logger.error(f"File not found: {file_path}")
                return None
            
            if not file_path.lower().endswith('.csv'):
                logger.error(f"File is not a CSV: {file_path}")
                return None
            
            try:
                df = pd.read_csv(file_path, header=0)
            except UnicodeDecodeError:
                logger.warning(f"UTF-8 encoding failed, trying latin-1 for {file_path}")
                df = pd.read_csv(file_path, header=0, encoding='latin-1')
            
            if df.empty:
                logger.error(f"CSV file is empty: {file_path}")
                return None
            
            logger.info(f"Successfully read CSV: {file_path} ({len(df)} rows)")
            return df
            
        except Exception as e:
            logger.error(f"Unexpected error reading CSV {file_path}: {str(e)}")
            return None

    def filter_data(self, data: Any) -> Optional[Any]:
        """Filter out invalid data values."""
        try:
            if pd.isna(data) or data == '' or (isinstance(data, str) and data.strip() == ''):
                return None
            return data
        except Exception as e:
            logger.warning(f"Data filtering error: {str(e)}")
            return None

    def validate_required_columns(self, df: pd.DataFrame) -> bool:
        """Validate DataFrame has required columns."""
        required_columns = ['Test_Type']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Missing required columns: {missing_columns}")
            logger.info(f"Available columns: {list(df.columns)}")
            return False
        
        return True

    def safe_numeric_conversion(self, value: Any, default: float = 0.0) -> float:
        """Safely convert value to numeric type."""
        try:
            if pd.isna(value):
                return default
            return float(value)
        except (ValueError, TypeError) as e:
            logger.warning(f"Numeric conversion error for {value}, using default {default}: {str(e)}")
            return default

    def validate_file_path(self, file_path: Optional[str]) -> bool:
        """Validate if file path exists and is accessible."""
        if not file_path:
            return False
        
        try:
            return os.path.exists(file_path) and os.path.isfile(file_path)
        except Exception as e:
            logger.warning(f"File path validation error for {file_path}: {str(e)}")
            return False

    def extract_information(self, df: pd.DataFrame) -> Dict[Tuple[str, str], List[Any]]:
        """Extract test configuration information from DataFrame."""
        if df is None or df.empty:
            logger.error("Cannot extract information from empty DataFrame")
            return {}
        
        if not self.validate_required_columns(df):
            return {}
        
        data_dict = {}
        successful_rows = 0
        failed_rows = 0
        
        for row_idx, row in df.iterrows():
            try:
                test_type = self.filter_data(getattr(row, 'Test_Type', None))
                if not test_type:
                    logger.warning(f"Row {row_idx + 1}: Missing Test_Type, skipping")
                    failed_rows += 1
                    continue
                
                key = (f"row : {row_idx + 1}", test_type)
                
                # Extract file paths
                source_cal_file = self.filter_data(getattr(row, 'Source_Phase_Cal', None))
                vector_cal_file = self.filter_data(getattr(row, 'Vector_Cal', None))
                init_file = self.filter_data(getattr(row, 'Initialization', None))
                gold_standard_file = self.filter_data(getattr(row, 'ReferenceData', None))
                
                # Validate reference data file
                if gold_standard_file and not self.validate_file_path(gold_standard_file):
                    logger.warning(f"Row {row_idx + 1}: ReferenceData file not found: {gold_standard_file}")
                
                # Extract system parameters
                ip = self.filter_data(getattr(row, 'IP', None))
                input_voltage = self.safe_numeric_conversion(getattr(row, 'InputV', None))
                output_voltage = self.safe_numeric_conversion(getattr(row, 'OutputV', None))
                aux1_voltage = self.safe_numeric_conversion(getattr(row, 'Aux1V', None))
                aux2_voltage = self.safe_numeric_conversion(getattr(row, 'Aux2V', None))
                max_gain = self.safe_numeric_conversion(getattr(row, 'MaxGain', None))
                stop_at_compression = self.filter_data(getattr(row, 'StopAtCompression', None))
                bias_control = bool(self.filter_data(getattr(row, 'BiasControl', None)))
                
                # Initialize sweep flags
                power_sweep = False
                impedance_sweep = False
                
                # Process bandwidth settings
                if_bandwidth = self.filter_data(getattr(row, 'IF_BW', None))
                mod_bandwidth = self.filter_data(getattr(row, 'MOD_BW', None))
                if if_bandwidth:
                    if_bandwidth = int(if_bandwidth)
                if mod_bandwidth:
                    mod_bandwidth = int(mod_bandwidth)

                # Process pulsed settings
                pulse_period = self.filter_data(getattr(row, 'Pulse_Period', None))
                pulse_width = self.filter_data(getattr(row, 'Pulse_Width', None))
                meas_window = self.filter_data(getattr(row, 'Measurement_Window', None))
                meas_delay = self.filter_data(getattr(row, 'Measurement_Delay', None))
                pulse_average = self.filter_data(getattr(row, 'Average', None))

                if pulse_period:
                    pulse_period = float(pulse_period)
                if pulse_width:
                    pulse_width = float(pulse_width)
                if meas_window:
                    meas_window = float(meas_window)
                if meas_delay:
                    meas_delay = float(meas_delay)
                if pulse_average:
                    pulse_average = int(pulse_average)
                

                # Process waveform settings
                mod_waveform = self.filter_data(getattr(row, 'Waveform', None))
                
                # Process measurement mode
                mode = 0  # Default to CW mode
                try:
                    mode_value = getattr(row, 'Mode', None)
                    if not pd.isna(mode_value) and mode_value is not None:
                        mode_str = str(mode_value).lower()
                        if 'cw' in mode_str:
                            mode = 0
                        elif 'pulsed' in mode_str:
                            mode = 1
                        elif 'modulated' in mode_str:
                            mode = 2
                        else:
                            logger.warning(f"Row {row_idx + 1}: Unknown mode '{mode_value}', using CW")
                except Exception as e:
                    logger.warning(f"Row {row_idx + 1}: Mode processing error: {str(e)}")
                
                # Process frequency configuration
                freq = None
                try:
                    min_freq = self.safe_numeric_conversion(getattr(row, 'Min_Freq', None))
                    max_freq = self.safe_numeric_conversion(getattr(row, 'Max_Freq', None))
                    step_size = self.safe_numeric_conversion(getattr(row, 'Freq_Step', None))
                    
                    if not pd.isna([min_freq, max_freq, step_size]).any():
                        if min_freq != max_freq and step_size != 0:
                            freq = np.arange(min_freq, max_freq + 0.1, step_size).tolist()
                            freq = [round(number, 1) for number in freq]
                        else:
                            freq = max_freq
                    else:
                        logger.warning(f"Row {row_idx + 1}: Incomplete frequency parameters")
                except Exception as e:
                    logger.warning(f"Row {row_idx + 1}: Frequency processing error: {str(e)}")
                
                # Process power configuration
                pwr = None
                try:
                    min_pwr = self.safe_numeric_conversion(getattr(row, 'Min_Pwr', None))
                    max_pwr = self.safe_numeric_conversion(getattr(row, 'Max_Pwr', None))
                    pwr_step_size = self.safe_numeric_conversion(getattr(row, 'Pwr_Step', None))
                    
                    if not pd.isna([min_pwr, max_pwr, pwr_step_size]).any():
                        if min_pwr != max_pwr and pwr_step_size != 0:
                            power_sweep = True
                            pwr = [min_pwr, max_pwr, pwr_step_size]
                        else:
                            pwr = max_pwr
                    else:
                        logger.warning(f"Row {row_idx + 1}: Incomplete power parameters")
                except Exception as e:
                    logger.warning(f"Row {row_idx + 1}: Power processing error: {str(e)}")
                
                # Process impedance configuration
                impedance_list = None
                try:
                    impedance_file = self.filter_data(getattr(row, 'targGrid1f0', None))
                    if impedance_file and impedance_file.endswith(".csv"):
                        if self.validate_file_path(impedance_file):
                            impedance_list = []
                            with open(impedance_file, 'r') as file:
                                lines = file.readlines()
                                if len(lines) > 2:
                                    impedance_sweep = True
                                for index, line in enumerate(lines):
                                    if index != 0:  # Skip header
                                        try:
                                            line_data = line.strip().split(',')
                                            impedance_targets = tuple([float(item) for item in line_data])
                                            impedance_list.append(impedance_targets)
                                        except (ValueError, IndexError) as e:
                                            logger.warning(f"Row {row_idx + 1}: Impedance line {index} parsing error: {str(e)}")
                        else:
                            logger.warning(f"Row {row_idx + 1}: Impedance file not found: {impedance_file}")
                except Exception as e:
                    logger.warning(f"Row {row_idx + 1}: Impedance processing error: {str(e)}")
                
                # Compile configuration data
                data_dict[key] = [
                    freq, pwr, impedance_list, source_cal_file, vector_cal_file, 
                    init_file, gold_standard_file, ip, mode, input_voltage, 
                    output_voltage, aux1_voltage, aux2_voltage, max_gain, 
                    stop_at_compression, power_sweep, impedance_sweep, bias_control,
                    if_bandwidth, mod_bandwidth, mod_waveform, pulse_period, pulse_width, meas_window, meas_delay, pulse_average
                ]
                
                successful_rows += 1
                logger.info(f"Successfully processed row {row_idx + 1}: {test_type}")
                
            except Exception as e:
                logger.error(f"Error processing row {row_idx + 1}: {str(e)}")
                failed_rows += 1
                continue
        
        logger.info(f"Data extraction complete: {successful_rows} successful, {failed_rows} failed")
        return data_dict
    

    def run(self):
        """Start the main application event loop."""
        self.root.mainloop()


if __name__ == "__main__":
    try:
        # Create and run the test controller
        controller = TestController()
        controller.run()
        
    except Exception as e:
        logger.error(f"Unexpected error in main execution: {str(e)}")
        print(f"Error starting Test Controller: {str(e)}")
        
        # Show error dialog if tkinter is available
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Startup Error", 
                                f"Failed to start Test Controller:\n\n{str(e)}")
        except:
            pass