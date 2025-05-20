import sys
import subprocess
import traceback # For detailed exception logging
import re # For parsing winget output
import tempfile # Added for temporary directory creation
import os # Added for path operations, if needed later
import shutil # Added for directory cleanup
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QTableView,
    QLabel,
    QStatusBar,
    QTextEdit,
    QFileDialog,
    QGroupBox,
    QAbstractItemView # Added for table view options
)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon
from PySide6.QtCore import Qt, QSettings

# Helper function to get correct path for bundled resources
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Winget2Intunewin GUI Packer")
        self.setGeometry(100, 100, 900, 700) # Adjusted size for better layout

        # Set Application Window Icon
        try:
            icon_path = resource_path("assets/logo.png")
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
            else:
                # This print might not be visible in a bundled app without console
                print(f"Warning: Window Icon file not found at {icon_path}. Using default icon.")
        except Exception as e:
            print(f"Error setting window icon: {e}")

        self.selected_app_data = None # To store data of the selected app
        self.current_temp_dir = None # To store the path of the current temporary directory
        self.downloaded_installer_path = None # To store the path to the downloaded installer
        self.install_script_path = None # To store the path to the generated install.ps1
        self.uninstall_script_path = None # To store the path to the generated uninstall.ps1
        self.detection_script_path = None # To store the path to the generated detection.ps1
        self.intunewin_util_path = None # To store path to IntuneWinAppUtil.exe

        # Main widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- Search & Selection Area ---
        search_group = QGroupBox("Search & Selection")
        search_layout = QVBoxLayout()

        # Application Search Input
        search_input_layout = QHBoxLayout()
        search_input_layout.addWidget(QLabel("Search for Application:"))
        self.search_input = QLineEdit()
        search_input_layout.addWidget(self.search_input)
        self.search_button = QPushButton("Search")
        search_input_layout.addWidget(self.search_button)
        search_layout.addLayout(search_input_layout)

        # Search Results Display
        self.search_results_table = QTableView()
        self.table_model = QStandardItemModel()
        self.table_model.setHorizontalHeaderLabels(["Name", "ID", "Version", "Source"])
        self.search_results_table.setModel(self.table_model)
        self.search_results_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.search_results_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.search_results_table.setEditTriggers(QAbstractItemView.EditTriggers.NoEditTriggers)
        self.search_results_table.horizontalHeader().setStretchLastSection(True)
        # self.search_results_table.resizeColumnsToContents() # Call after populating
        search_layout.addWidget(self.search_results_table)
        search_group.setLayout(search_layout)
        main_layout.addWidget(search_group)
        
        # --- Packaging Configuration Area ---
        packaging_group = QGroupBox("Packaging Configuration")
        packaging_layout = QVBoxLayout()

        # Selected Application Display (simplified for now)
        self.selected_app_label = QLabel("Selected App: None")
        packaging_layout.addWidget(self.selected_app_label)

        # Output Folder Selection
        output_folder_layout = QHBoxLayout()
        output_folder_layout.addWidget(QLabel("Output Folder for .intunewin:"))
        self.output_folder_input = QLineEdit()
        self.output_folder_input.setReadOnly(True) # User selects via browse
        output_folder_layout.addWidget(self.output_folder_input)
        self.browse_button = QPushButton("Browse...")
        output_folder_layout.addWidget(self.browse_button)
        packaging_layout.addLayout(output_folder_layout)

        # IntuneWinAppUtil.exe Path Selection
        intunewin_util_layout = QHBoxLayout()
        intunewin_util_layout.addWidget(QLabel("Path to IntuneWinAppUtil.exe:"))
        self.intunewin_util_input = QLineEdit()
        self.intunewin_util_input.setReadOnly(True)
        intunewin_util_layout.addWidget(self.intunewin_util_input)
        self.browse_intunewin_util_button = QPushButton("Browse...")
        intunewin_util_layout.addWidget(self.browse_intunewin_util_button)
        packaging_layout.addLayout(intunewin_util_layout)

        packaging_group.setLayout(packaging_layout)
        main_layout.addWidget(packaging_group)

        # --- Action & Status Area ---
        action_status_group = QGroupBox("Action & Status")
        action_status_layout = QVBoxLayout()

        # Package Button
        self.package_button = QPushButton("Create .intunewin Package")
        action_status_layout.addWidget(self.package_button, alignment=Qt.AlignmentFlag.AlignCenter) # Center the button

        # Log Window (Optional detailed log)
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        action_status_layout.addWidget(self.log_window)
        action_status_group.setLayout(action_status_layout)
        main_layout.addWidget(action_status_group)
        
        # Status Bar
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Ready")

        # Connect signals to slots
        self.browse_button.clicked.connect(self.open_output_folder_dialog)
        self.search_button.clicked.connect(self.handle_search_button_clicked)
        self.search_results_table.selectionModel().selectionChanged.connect(self.handle_table_selection_changed)
        self.browse_intunewin_util_button.clicked.connect(self._browse_for_intunewin_util) # Connect new button
        self.package_button.clicked.connect(self.handle_package_button_clicked) # Connect package button

        self._load_settings() # Load settings on startup

        # TODO: Connect other signals to slots (e.g., package_button)
        # TODO: Implement dark mode theme (initial version applied, can be refined)

    def open_output_folder_dialog(self):
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self.output_folder_input.text() # Start at current path if any
        )
        if folder_path:
            self.output_folder_input.setText(folder_path)
            self.statusBar().showMessage(f"Output folder set to: {folder_path}")

    def parse_winget_search_output(self, output_text):
        apps = []
        lines = output_text.strip().split('\n')
        
        header_index = -1
        for i, line in enumerate(lines):
            if re.search(r"Name\s+Id\s+Version\s+(?:Match\s+)?Source", line):
                header_index = i
                break
        
        if header_index == -1 or header_index + 1 >= len(lines) or not lines[header_index+1].startswith("---"):
            self.log_window.append("INFO: Winget search output format not recognized or no data rows found.")
            return apps

        # Find column start indices from the header
        header_line = lines[header_index]
        try:
            # More robust way to find columns based on typical winget output
            # Name | ID | Version | Source
            # Assumes ID, Version, Source are single words after the initial Name block
            # This is a heuristic and might need adjustment if winget output changes drastically
            
            # Find the start of "Id " (with space) to avoid matching "Id" within a name
            id_col_start = header_line.find(" Id ") 
            # Find " Version "
            version_col_start = header_line.find(" Version ", id_col_start)
            # Find " Source"
            source_col_start = header_line.find(" Source", version_col_start)

            if not all([id_col_start > 0, version_col_start > 0, source_col_start > 0]):
                 self.log_window.append("INFO: Could not reliably determine column starts from header.")
                 return apps # Fallback or error

            # Data rows start 2 lines after the header (header, then "----" line)
            for line in lines[header_index + 2:]:
                line = line.strip()
                if not line or line.startswith("---"): # Skip empty lines or other separators
                    continue

                # Extract based on found column positions
                name = line[:id_col_start].strip()
                # For ID, Version, Source, we split the remainder and take the first 3 parts
                # This handles variable spacing better than fixed slicing for the rightmost columns.
                remaining_part = line[id_col_start:].strip()
                parts = re.split(r'\s{2,}', remaining_part) # Split by 2 or more spaces
                
                app_id = ""
                version = ""
                source = ""

                if len(parts) >= 2: # We need at least ID and Version
                    app_id = parts[0].strip()
                    raw_version = parts[1].strip()
                    
                    # Clean up the version string: remove any " Tag: ..." suffix
                    tag_suffix_index = raw_version.find(" Tag: ")
                    if tag_suffix_index != -1:
                        version = raw_version[:tag_suffix_index].strip()
                    else:
                        version = raw_version

                    if len(parts) >= 4: # Name, Id, Version, Match, Source
                        # If "Match" column is present, Source is the 4th element from remaining_part split
                        source = parts[3].strip() 
                    elif len(parts) == 3: # Name, Id, Version, Source (No "Match" column)
                        source = parts[2].strip()
                    else: # Only ID and Version found reliably from parts (and Match might be the 3rd part if it exists without source)
                        # This case means we have ID and Version, but Source might be missing or part of a 3rd unexpected column.
                        # If parts has only 2 elements (ID, Version), source will be N/A
                        source = "N/A" 
                    
                    if name and app_id: # Ensure essential fields are present
                         apps.append({"Name": name, "ID": app_id, "Version": version, "Source": source})
                elif name: # Sometimes only name and ID might appear if the line is malformed or short
                    # Try a simpler split if the above fails for some lines
                    parts_simple = line.split()
                    if len(parts_simple) >=2:
                        # This fallback is very heuristic and might not be reliable.
                        # It assumes ID is the last significant word if primary parsing fails.
                        # Consider removing or refining if it causes incorrect parsing.
                        app_id_simple = parts_simple[-1] 
                        name_simple = " ".join(parts_simple[:-1]).strip() # Reconstruct name
                        # A very basic check to see if app_id_simple looks like an ID
                        if name_simple and app_id_simple and '.' in app_id_simple: 
                             apps.append({"Name": name_simple, "ID": app_id_simple, "Version": "N/A", "Source": "N/A"})


            if not apps and lines[header_index + 2:]: # If no apps parsed but there were data lines
                self.log_window.append("INFO: Attempted to parse data rows but no applications were extracted. Check parsing logic against winget output.")

        except Exception as e:
            self.log_window.append(f"ERROR: Exception during parsing winget output: {e}")
            self.log_window.append(traceback.format_exc())
        
        return apps

    def handle_search_button_clicked(self):
        search_term = self.search_input.text().strip()
        
        # Clear previous results from table
        self.table_model.setRowCount(0) # Clears data but keeps headers
        self.selected_app_label.setText("Selected App: None")
        self.selected_app_data = None

        if not search_term:
            self.statusBar().showMessage("Please enter a search term.")
            self.log_window.append("INFO: Search attempt with empty term.")
            return

        self.statusBar().showMessage(f"INFO: Attempting to search for: {search_term}...")
        # Log the command with the added flag
        self.log_window.append(f"CMD: winget search --accept-source-agreements \"{search_term}\"")

        try:
            self.log_window.append("INFO: Preparing to execute winget command...")
            
            command_list = ['winget', 'search', '--accept-source-agreements', search_term]

            process = subprocess.run(
                command_list,
                capture_output=True,
                text=True, 
                encoding='utf-8',  # Specify UTF-8 encoding
                errors='replace',   # Replace undecodable characters
                check=False, 
                # shell=False by default
            )
            
            self.log_window.append(f"INFO: Winget command executed. Return code: {process.returncode}")

            if process.stdout and process.stdout.strip():
                self.log_window.append("--- stdout ---")
                self.log_window.append(process.stdout.strip())
            else:
                self.log_window.append("INFO: Winget command produced no stdout or stdout was empty.")

            if process.stderr and process.stderr.strip():
                self.log_window.append("--- stderr ---")
                self.log_window.append(process.stderr.strip())
            else:
                self.log_window.append("INFO: Winget command produced no stderr or stderr was empty.")

            if process.returncode == 0:
                if process.stdout and process.stdout.strip():
                    self.statusBar().showMessage("Search successful. Parsing results...")
                    self.log_window.append("INFO: Parsing winget output...")
                    parsed_apps = self.parse_winget_search_output(process.stdout)
                    
                    if parsed_apps:
                        self.log_window.append(f"INFO: Parsed {len(parsed_apps)} applications.")
                        for app_info in parsed_apps:
                            row = [
                                QStandardItem(app_info.get("Name", "N/A")),
                                QStandardItem(app_info.get("ID", "N/A")),
                                QStandardItem(app_info.get("Version", "N/A")),
                                QStandardItem(app_info.get("Source", "N/A"))
                            ]
                            self.table_model.appendRow(row)
                        self.search_results_table.resizeColumnsToContents()
                        self.statusBar().showMessage(f"Search complete. {len(parsed_apps)} applications found.")
                    else:
                        self.log_window.append("INFO: No applications were parsed from the output.")
                        self.statusBar().showMessage("Search complete. No applications found or output not parseable.")
                else:
                    self.statusBar().showMessage("Search successful, but no applications found or no output from winget.")
                    self.log_window.append("INFO: Search returned success (0), but stdout was empty. Check winget behavior directly in terminal.")
            else:
                error_message = f"Winget search failed. Return code: {process.returncode}."
                self.log_window.append(f"ERROR: {error_message}")
                self.statusBar().showMessage(error_message + " Check log for details.")

        except FileNotFoundError:
            error_message = "ERROR: winget command not found. Please ensure it's installed and in your PATH."
            self.log_window.append(error_message)
            self.statusBar().showMessage(error_message)
        except Exception as e:
            error_message = f"ERROR: An unexpected error occurred during search: {e}"
            self.log_window.append(error_message)
            self.log_window.append(f"Exception type: {type(e)}")
            self.log_window.append("--- Traceback ---")
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage("An unexpected error occurred. Check log.")

    def handle_table_selection_changed(self, selected, deselected):
        selected_indexes = self.search_results_table.selectionModel().selectedRows()
        if selected_indexes:
            # Assuming first column is Name (index 0) and second is ID (index 1)
            # For more robustness, you might want to store full row data or query by header name
            selected_row = selected_indexes[0].row()
            name_item = self.table_model.item(selected_row, 0) # Name column
            id_item = self.table_model.item(selected_row, 1)   # ID column
            version_item = self.table_model.item(selected_row, 2) # Version
            source_item = self.table_model.item(selected_row, 3) # Source

            if name_item and id_item:
                name = name_item.text()
                app_id = id_item.text()
                self.selected_app_label.setText(f"Selected App: {name} ({app_id})")
                self.selected_app_data = {
                    "Name": name,
                    "ID": app_id,
                    "Version": version_item.text() if version_item else "N/A",
                    "Source": source_item.text() if source_item else "N/A"
                }
                self.statusBar().showMessage(f"Selected: {name}")
            else:
                self.selected_app_label.setText("Selected App: Error retrieving details")
                self.selected_app_data = None
        else:
            self.selected_app_label.setText("Selected App: None")
            self.selected_app_data = None
            self.statusBar().showMessage("Selection cleared.")

    def _create_temp_packaging_dir(self):
        """Creates a unique temporary directory for the packaging process."""
        try:
            # Create a unique temporary directory
            # The directory will be created in the default temporary location for the OS
            self.current_temp_dir = tempfile.mkdtemp(prefix="winget_pkg_")
            self.log_window.append(f"INFO: Created temporary directory: {self.current_temp_dir}")
            # We could also create subdirectories here if a specific structure is needed immediately
            # For example:
            # os.makedirs(os.path.join(self.current_temp_dir, "scripts"), exist_ok=True)
            # os.makedirs(os.path.join(self.current_temp_dir, "installer"), exist_ok=True)
            return self.current_temp_dir
        except Exception as e:
            error_message = f"ERROR: Failed to create temporary packaging directory: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage("Error creating temporary directory. Check log.")
            self.current_temp_dir = None # Ensure it's None on failure
            return None

    def _find_installer_file(self, download_dir, app_id):
        """Attempts to find the downloaded installer file in the given directory."""
        if not download_dir or not os.path.isdir(download_dir):
            self.log_window.append(f"ERROR: Download directory '{download_dir}' is invalid or does not exist.")
            return None

        self.log_window.append(f"INFO: Searching for installer in: {download_dir}")
        potential_installers = []
        common_installer_extensions = ('.exe', '.msi', '.msix', '.msixbundle', '.appx', '.appxbundle', '.zip') # Added .zip

        # Scan the primary download directory
        for item in os.listdir(download_dir):
            item_path = os.path.join(download_dir, item)
            if os.path.isfile(item_path) and item.lower().endswith(common_installer_extensions):
                potential_installers.append(item_path)
        
        # Winget sometimes creates a subdirectory named after the App ID or a similar structure.
        # Let's check if such a subdirectory exists and contains the installer.
        # Example: <download_dir>/<app_id_folder>/installer.exe
        # This part is heuristic.
        if not potential_installers:
            possible_subdir_name_parts = app_id.split('.') # e.g., "Microsoft.Edge" -> ["Microsoft", "Edge"]
            # Check for subdirectories that might match parts of the app_id or common names like 'install', 'setup'
            for item in os.listdir(download_dir):
                item_path = os.path.join(download_dir, item)
                if os.path.isdir(item_path):
                    # A simple check: if directory name is part of app_id or a generic installer name
                    if any(part.lower() in item.lower() for part in possible_subdir_name_parts) or \
                       item.lower() in ['installer', 'install', 'setup', app_id.lower()]:
                        self.log_window.append(f"INFO: Checking potential sub-directory: {item_path}")
                        for sub_item in os.listdir(item_path):
                            sub_item_path = os.path.join(item_path, sub_item)
                            if os.path.isfile(sub_item_path) and sub_item.lower().endswith(common_installer_extensions):
                                potential_installers.append(sub_item_path)
                        if potential_installers: # Found in this subdir, break from checking other subdirs
                            break 

        if not potential_installers:
            self.log_window.append(f"ERROR: No installer file found in '{download_dir}' (or relevant subdirectories) with extensions: {common_installer_extensions}")
            return None
        
        if len(potential_installers) == 1:
            installer_path = potential_installers[0]
            self.log_window.append(f"INFO: Found installer: {installer_path}")
            return installer_path
        else:
            # If multiple installers, this is ambiguous. Log and try to pick one.
            # A common scenario is an .exe and a .msi. Heuristics can be complex.
            # For now, let's log a warning and pick the first one found, or one that contains app_id in its name.
            self.log_window.append(f"WARNING: Multiple potential installers found: {potential_installers}")
            # Attempt to find one that contains the app_id in its name (more specific match)
            for p_path in potential_installers:
                if app_id.lower() in os.path.basename(p_path).lower():
                    self.log_window.append(f"INFO: Selected installer (contains app_id in name): {p_path}")
                    return p_path
            self.log_window.append(f"INFO: Selecting the first one found as fallback: {potential_installers[0]}")
            return potential_installers[0] # Fallback, might need refinement

    def _sanitize_filename(self, name):
        """Sanitizes a string to be suitable for use as a filename component."""
        if not name: 
            return "DefaultApp"
        # Remove characters that are invalid in Windows filenames or problematic
        name = re.sub(r'[\\/:*?\"<>|\s+]', '', name)
        # Remove any leading/trailing non-alphanumeric characters that might remain after sanitization
        name = re.sub(r'^[^a-zA-Z0-9]+|\[^a-zA-Z0-9]+$', '', name)
        if not name: # Fallback if sanitization results in an empty string
            return "DefaultAppInstall"
        return name

    def _generate_install_script(self, app_id, app_name, app_version, target_dir):
        """Generates the install.ps1 script (dynamically named) and saves it to the target directory."""
        if not all([app_id, app_name, app_version, target_dir]):
            self.log_window.append("ERROR: Missing data for generating install script (app_id, app_name, app_version, or target_dir).")
            self.install_script_path = None # Ensure it's reset
            return None

        sanitized_app_name = self._sanitize_filename(app_name)
        script_name = f"{sanitized_app_name}.ps1" # Dynamic script name
        script_path = os.path.join(target_dir, script_name)
        # Update self.install_script_path early so it's available even if generation fails below for some reason
        self.install_script_path = script_path 

        self.log_window.append(f"INFO: Generating {script_name} for {app_name} (ID: {app_id}) in {target_dir}")

        # Using f-string with triple quotes for multiline script content
        # Ensure PowerShell special characters like $ are escaped by doubling them ($$) if they are meant to be literal in the script
        # and not part of Python's f-string interpolation. Here, $AppId etc. are PS variables.
        script_content = f"""# install.ps1 - Generated by Winget2Intunewin GUI Packer

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$AppId = "{app_id}"
$AppName = "{app_name}" # Used for logging
$AppVersion = "{app_version}" # Used for specific version install

Write-Host "Starting installation process for $AppName (ID: $AppId, Version: $AppVersion)..."

$WingetPath = ""
# Attempt to find winget.exe in common locations
if (Test-Path "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe") {{
    $WingetPath = Get-Item "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
}} elseif (Test-Path "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe") {{
    $WingetPath = "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe"
}} else {{
    Write-Error "Winget executable not found in common paths. Please ensure Winget is installed and accessible."
    exit 1
}}

Write-Host "Using Winget executable at: $WingetPath"

try {{
    Write-Host "Executing: `"$WingetPath`" install --id `"$AppId`" --version `"$AppVersion`" --exact --scope machine --accept-package-agreements --accept-source-agreements --disable-interactivity"
    # Using Start-Process to better handle execution and capture exit codes if needed, though direct call is also fine.
    # Add --disable-interactivity for robust silent execution
    & $WingetPath install --id "$AppId" --version "$AppVersion" --exact --scope machine --accept-package-agreements --accept-source-agreements --disable-interactivity
    
    if ($LASTEXITCODE -ne 0) {{
        Write-Error "Winget install command failed for $AppId with exit code $LASTEXITCODE."
        exit $LASTEXITCODE
    }}
    Write-Host "Winget install command for $AppId (Version: $AppVersion) completed successfully."
    
    # Optional: Add further verification steps here if the application creates a specific registry key or file.

}} 
catch {{
    Write-Error "An error occurred during the installation of $AppName (ID: $AppId, Version: $AppVersion)."
    Write-Error $_.Exception.Message
    # Attempt to get more details if it's a Winget error specifically
    if ($_.Exception.InnerException) {{
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }}
    exit 1 # General error code
}}

Write-Host "Installation script for $AppName (ID: $AppId, Version: $AppVersion) finished."
exit 0
"""

        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            self.log_window.append(f"INFO: Successfully generated {script_path}")
            self.statusBar().showMessage(f"{script_name} generated for {app_name}.")
            # self.install_script_path = script_path # Already set above
            return script_path
        except Exception as e:
            error_message = f"ERROR: Failed to write install script {script_path}: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage(f"Error generating {script_name}. Check log.")
            # self.install_script_path remains as it might have been set, or None if error before path creation
            return None

    def _generate_uninstall_script(self, app_id, app_name, target_dir):
        """Generates the uninstall.ps1 script and saves it to the target directory."""
        if not all([app_id, app_name, target_dir]):
            self.log_window.append("ERROR: Missing data for generating uninstall script (app_id, app_name, or target_dir).")
            return None

        script_name = "uninstall.ps1"
        script_path = os.path.join(target_dir, script_name)
        self.log_window.append(f"INFO: Generating {script_name} for {app_name} (ID: {app_id}) in {target_dir}")

        script_content = f"""# uninstall.ps1 - Generated by Winget2Intunewin GUI Packer

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$AppId = "{app_id}"
$AppName = "{app_name}" # Used for logging

Write-Host "Starting uninstallation process for $AppName (ID: $AppId)..."

$WingetPath = ""
# Attempt to find winget.exe in common locations
if (Test-Path "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe") {{
    $WingetPath = Get-Item "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
}} elseif (Test-Path "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe") {{
    $WingetPath = "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe"
}} else {{
    Write-Error "Winget executable not found in common paths. Please ensure Winget is installed and accessible."
    # For uninstall, if winget is not found, we might not want to fail the whole script if the app isn't there anyway.
    # However, for an explicit uninstall command, it is an error if winget itself is missing.
    exit 1 
}}

Write-Host "Using Winget executable at: $WingetPath"

# Check if the application is installed before attempting to uninstall
Write-Host "Checking if $AppName (ID: $AppId) is installed..."
$InstalledApp = & $WingetPath list --id "$AppId" --accept-source-agreements

if ($InstalledApp -match $AppId) {{
    Write-Host "$AppName (ID: $AppId) is installed. Proceeding with uninstall."
    try {{
        Write-Host "Executing: `"$WingetPath`" uninstall --id `"$AppId`" --accept-package-agreements --accept-source-agreements --disable-interactivity"
        & $WingetPath uninstall --id "$AppId" --accept-package-agreements --accept-source-agreements --disable-interactivity
        
        if ($LASTEXITCODE -ne 0) {{
            Write-Error "Winget uninstall command failed for $AppId with exit code $LASTEXITCODE."
            exit $LASTEXITCODE
        }}
        Write-Host "Winget uninstall command for $AppId completed successfully."
    }}
    catch {{
        Write-Error "An error occurred during the uninstallation of $AppName (ID: $AppId)."
        Write-Error $_.Exception.Message
        if ($_.Exception.InnerException) {{
            Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
        }}
        exit 1 # General error code
    }}
}} else {{
    Write-Host "$AppName (ID: $AppId) is not found or not installed via Winget. Uninstall script will consider this a success (nothing to uninstall)."
    # In Intune, an uninstall script exiting with 0 when the app is not found is often desired.
    exit 0 
}}

Write-Host "Uninstallation script for $AppName (ID: $AppId) finished."
exit 0
"""

        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            self.log_window.append(f"INFO: Successfully generated {script_path}")
            self.statusBar().showMessage(f"{script_name} generated for {app_name}.")
            self.uninstall_script_path = script_path
            return script_path
        except Exception as e:
            error_message = f"ERROR: Failed to write uninstall script {script_path}: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage(f"Error generating {script_name}. Check log.")
            self.uninstall_script_path = None
            return None

    def _generate_detection_script(self, app_id, app_name, app_version, target_dir):
        """Generates the detection.ps1 script and saves it to the target directory."""
        if not all([app_id, app_name, app_version, target_dir]):
            self.log_window.append("ERROR: Missing data for generating detection script (app_id, app_name, app_version, or target_dir).")
            return None

        script_name = "detection.ps1"
        script_path = os.path.join(target_dir, script_name)
        self.log_window.append(f"INFO: Generating {script_name} for {app_name} (ID: {app_id}, Version: {app_version}) in {target_dir}")

        # Ensure PowerShell $ variables are escaped as $$ if they are part of the f-string literal template.
        # Variables like {app_id} are Python f-string interpolations.
        script_content = f"""# detection.ps1 - Generated by Winget2Intunewin GUI Packer

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$AppId = "{app_id}"
$AppName = "{app_name}"
$AppVersion = "{app_version}"

Write-Host "Starting detection for $AppName (ID: $AppId, Version: $AppVersion)..."

$WingetPath = ""
if (Test-Path "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe") {{
    $WingetPath = Get-Item "$($env:ProgramFiles)\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
}} elseif (Test-Path "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe") {{
    $WingetPath = "$($env:LOCALAPPDATA)\Microsoft\WindowsApps\winget.exe"
}} else {{
    Write-Host "Winget executable not found. Cannot perform detection."
    exit 1
}}

Write-Host "Using Winget executable at: $WingetPath"

$ExitCode = 1 # Default to not detected
$TempOutFile = Join-Path $env:TEMP "winget_detect_out_$(Get-Random).txt"
$TempErrFile = Join-Path $env:TEMP "winget_detect_err_$(Get-Random).txt"

try {{
    Write-Host "Executing: `"$WingetPath`" list --id `"$AppId`" --version `"$AppVersion`" --exact --accept-source-agreements"
    $Process = Start-Process -FilePath $WingetPath -ArgumentList "list --id \`"$AppId\`" --version \`"$AppVersion\`" --exact --accept-source-agreements" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $TempOutFile -RedirectStandardError $TempErrFile
    $CmdExitCode = $Process.ExitCode

    if ($CmdExitCode -eq 0) {{
        $DetectionOutput = Get-Content $TempOutFile -ErrorAction SilentlyContinue
        $Found = $false
        foreach ($line in $DetectionOutput) {{
            if (($line -match [regex]::Escape($AppId)) -and ($line -match [regex]::Escape($AppVersion))) {{
                Write-Host "Detected: $AppName (ID: $AppId Version: $AppVersion) - Matched Line: $line"
                $ExitCode = 0 # Detected
                $Found = $true
                break
            }}
        }}
        if (-not $Found) {{
             Write-Host "Application $AppId with version $AppVersion not found in winget list output (command exit code 0)."
        }}
    }} else {{
        Write-Host "Winget list command failed with exit code $CmdExitCode."
        $ErrorOutput = Get-Content $TempErrFile -ErrorAction SilentlyContinue
        if ($ErrorOutput) {{
            Write-Host "Winget list stderr: $ErrorOutput"
        }}
    }}
}} catch {{
    Write-Host "An error occurred during the detection process for $AppName (ID: $AppId, Version: $AppVersion). Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {{
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)"
    }}
    # ExitCode remains 1 (default)
}} finally {{
    Remove-Item $TempOutFile -ErrorAction SilentlyContinue
    Remove-Item $TempErrFile -ErrorAction SilentlyContinue
}}

if ($ExitCode -eq 0) {{
    Write-Host "Final Detection Status: Application Found - $AppName (ID: $AppId, Version: $AppVersion). This output is used by Intune."
}} else {{
    Write-Host "Final Detection Status: Application Not Found - $AppName (ID: $AppId, Version: $AppVersion)."
}}

exit $ExitCode
"""

        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            self.log_window.append(f"INFO: Successfully generated {script_path}")
            self.statusBar().showMessage(f"{script_name} generated for {app_name}.")
            self.detection_script_path = script_path
            return script_path
        except Exception as e:
            error_message = f"ERROR: Failed to write detection script {script_path}: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage(f"Error generating {script_name}. Check log.")
            self.detection_script_path = None
            return None

    def _download_installer(self, app_id, app_version, download_dir):
        """Downloads the installer for the given app_id to the specified directory."""
        self.downloaded_installer_path = None # Reset before attempt
        if not app_id or not download_dir:
            self.log_window.append("ERROR: App ID or download directory missing for download.")
            self.statusBar().showMessage("Error: Missing information for download.")
            return False

        self.log_window.append(f"INFO: Attempting to download installer for ID: {app_id}, Version: {app_version}")
        self.statusBar().showMessage(f"Downloading {app_id}...")

        command = [
            'winget', 'download',
            '--id', app_id,
            '--version', app_version, # Specify the exact version
            '--exact', # Ensure only this version is targeted
            '--download-directory', download_dir,
            '--accept-package-agreements',
            '--accept-source-agreements'
            # Consider adding --scope machine if applicable and supported by download
        ]

        self.log_window.append(f"CMD: {' '.join(command)}")

        try:
            process = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                check=False 
            )

            self.log_window.append(f"INFO: Winget download command executed. Return code: {process.returncode}")

            if process.stdout and process.stdout.strip():
                self.log_window.append("--- stdout (download) ---")
                self.log_window.append(process.stdout.strip())
            else:
                self.log_window.append("INFO: Winget download produced no stdout or stdout was empty.")

            if process.stderr and process.stderr.strip():
                self.log_window.append("--- stderr (download) ---")
                self.log_window.append(process.stderr.strip())
            else:
                self.log_window.append("INFO: Winget download produced no stderr or stderr was empty.")

            if process.returncode == 0:
                # Winget download might not produce significant stdout on success, 
                # but stderr might contain progress or verbose logging.
                # The main indicator is the return code.
                self.log_window.append(f"INFO: Winget download command for {app_id} completed.")
                
                # Now, try to find the actual downloaded installer file
                found_installer = self._find_installer_file(download_dir, app_id)
                if found_installer:
                    self.downloaded_installer_path = found_installer
                    self.log_window.append(f"INFO: Successfully verified installer: {self.downloaded_installer_path}")
                    self.statusBar().showMessage(f"Installer for {app_id} downloaded and verified.")
                    return True # Download and verification successful
                else:
                    self.log_window.append(f"ERROR: Winget download command succeeded for {app_id}, but the installer file could not be found in '{download_dir}'.")
                    self.statusBar().showMessage(f"Download for {app_id} complete, but installer not found. Check logs.")
                    return False # Download succeeded, but file not found
            else:
                error_message = f"Winget download failed for {app_id}. Return code: {process.returncode}."
                self.log_window.append(f"ERROR: {error_message}")
                self.statusBar().showMessage(error_message + " Check log for details.")
                return False

        except FileNotFoundError:
            error_message = "ERROR: winget command not found. Please ensure it's installed and in your PATH."
            self.log_window.append(error_message)
            self.statusBar().showMessage(error_message)
            return False
        except Exception as e:
            error_message = f"ERROR: An unexpected error occurred during winget download: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage("An unexpected error occurred during download. Check log.")
            return False

    def apply_dark_mode(self):
        # Modern Dark Theme with Blue/Purple Gradients
        self.setStyleSheet('''
            QWidget {
                background-color: #1e1e2f; /* Dark blue-ish grey */
                color: #e0e0e0; /* Light grey text */
                font-size: 10pt;
                font-family: "Segoe UI", Arial, sans-serif;
            }
            QMainWindow {
                background-color: #1e1e2f;
            }
            QGroupBox {
                background-color: transparent; /* Make GroupBox background transparent */
                border: 1px solid #4a00e0; /* Purple border */
                border-radius: 8px;
                margin-top: 1.5ex; /* Increased top margin for title space */ 
                font-weight: bold;
                padding-top: 1.5ex; /* Add padding on top of groupbox content area to push it down */
            }
            QGroupBox::title {
                subcontrol-origin: margin; /* Position relative to the margin */
                subcontrol-position: top left; 
                padding: 5px 15px; /* Increased padding for title text */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8A2BE2, stop:1 #4A00E0); /* BlueViolet to darker Purple gradient */
                color: #ffffff;
                border-radius: 6px; /* Slightly less than GroupBox for a mild inset visual, or match GroupBox's 8px */
                font-weight: bold;
                /* Additional offset if needed, negative values pull it up/left */
                left: 10px; /* Nudge title to the right a bit from the very edge */
            }
            QLineEdit, QTextEdit, QTableView {
                background-color: #252538;
                color: #e0e0e0;
                border: 1px solid #8A2BE2; /* Lighter Purple border for inputs */
                border-radius: 5px;
                padding: 6px;
            }
            QTableView {
                gridline-color: #4169E1; /* Royal Blue grid lines */
            }
            QHeaderView::section { /* For QTableView header */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4169E1, stop:1 #3A5FCD); /* RoyalBlue to slightly darker blue */
                color: #ffffff;
                padding: 5px;
                border: 1px solid #3A5FCD;
                font-weight: bold;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8A2BE2, stop:1 #4169E1); /* BlueViolet to RoyalBlue gradient */
                color: #ffffff;
                border: none; /* Remove default border */
                border-radius: 5px;
                padding: 8px 15px;
                font-weight: bold;
                min-height: 22px;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #9B30FF, stop:1 #5179FF); /* Lighter gradient on hover */
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4A00E0, stop:1 #2A4FAD); /* Darker gradient on press */
            }
            QLabel {
                color: #f0f0f0; /* Brighter white for labels */
                font-weight: bold; /* Make labels bold */
            }
            QStatusBar {
                color: #cccccc;
            }
            QStatusBar::item {
                border: none; /* Remove borders from status bar items */
            }
            QScrollBar:horizontal {
                border: none;
                background: #252538;
                height: 10px;
                margin: 0px 20px 0 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8A2BE2, stop:1 #4169E1);
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                width: 20px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            QScrollBar:vertical {
                border: none;
                background: #252538;
                width: 10px;
                margin: 20px 0 20px 0;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #8A2BE2, stop:1 #4169E1);
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 20px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            /* Style for the QLineEdit used in the QFileDialog to make it match the theme */
            QFileDialog QLineEdit {
                 background-color: #252538;
                 color: #e0e0e0;
                 border: 1px solid #8A2BE2;
                 border-radius: 5px;
                 padding: 6px;
            }
        ''')

    def _browse_for_intunewin_util(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select IntuneWinAppUtil.exe",
            os.path.dirname(self.intunewin_util_input.text()) if self.intunewin_util_input.text() else os.path.expanduser("~"), # Start in current dir or home
            "Executable files (*.exe)" # Filter for .exe files
        )
        if file_path:
            # Basic validation: check if filename is IntuneWinAppUtil.exe
            if os.path.basename(file_path).lower() == "intunewinapputil.exe":
                self.intunewin_util_path = file_path
                self.intunewin_util_input.setText(file_path)
                self._save_settings()
                self.statusBar().showMessage(f"IntuneWinAppUtil.exe path set to: {file_path}")
                self.log_window.append(f"INFO: IntuneWinAppUtil.exe path set to: {file_path}")
            else:
                self.statusBar().showMessage("Selected file is not IntuneWinAppUtil.exe. Please select the correct file.")
                self.log_window.append(f"WARNING: User selected '{os.path.basename(file_path)}\' instead of IntuneWinAppUtil.exe")

    def _load_settings(self):
        settings = QSettings("WingetGUIOrg", "Winget2IntunewinPacker")
        loaded_path = settings.value("intunewin_util_path")
        if loaded_path and isinstance(loaded_path, str) and os.path.isfile(loaded_path): # Check if path exists and is a file
            self.intunewin_util_path = loaded_path
            self.intunewin_util_input.setText(self.intunewin_util_path)
            self.log_window.append(f"INFO: Loaded IntuneWinAppUtil.exe path: {self.intunewin_util_path}")
        else:
            self.log_window.append("INFO: IntuneWinAppUtil.exe path not set or invalid. Please configure it.")

    def _save_settings(self):
        settings = QSettings("WingetGUIOrg", "Winget2IntunewinPacker")
        if self.intunewin_util_path:
            settings.setValue("intunewin_util_path", self.intunewin_util_path)
            self.log_window.append(f"INFO: Saved IntuneWinAppUtil.exe path: {self.intunewin_util_path}")
        else:
             settings.remove("intunewin_util_path") # Or settings.setValue("intunewin_util_path", None)
             self.log_window.append("INFO: Cleared IntuneWinAppUtil.exe path from settings.")

    def _run_intunewin_app_util(self, source_folder, setup_file_path, output_package_dir):
        """Executes IntuneWinAppUtil.exe to package the application."""
        if not self.intunewin_util_path or not os.path.isfile(self.intunewin_util_path):
            self.log_window.append("ERROR: Path to IntuneWinAppUtil.exe is not set or invalid. Please configure it.")
            self.statusBar().showMessage("Error: IntuneWinAppUtil.exe path not configured.")
            return False

        if not all([source_folder, setup_file_path, output_package_dir]):
            self.log_window.append("ERROR: Missing source folder, setup file, or output directory for IntuneWinAppUtil.exe.")
            return False
        
        if not os.path.isdir(source_folder):
            self.log_window.append(f"ERROR: Source folder '{source_folder}\' does not exist.")
            return False
        if not os.path.isfile(setup_file_path):
            self.log_window.append(f"ERROR: Setup file '{setup_file_path}\' does not exist.")
            return False
        if not os.path.isdir(output_package_dir):
            self.log_window.append(f"ERROR: Output package directory '{output_package_dir}\' does not exist. Creating it...")
            try:
                os.makedirs(output_package_dir, exist_ok=True)
                self.log_window.append(f"INFO: Created output directory: {output_package_dir}")
            except Exception as e:
                self.log_window.append(f"ERROR: Failed to create output directory '{output_package_dir}': {e}")
                return False

        # Derive the expected output filename for logging/checking later, though the tool creates it.
        # e.g., if setup_file_path is .../install.ps1, output will be install.intunewin
        expected_output_filename = os.path.splitext(os.path.basename(setup_file_path))[0] + ".intunewin"
        expected_output_filepath = os.path.join(output_package_dir, expected_output_filename)

        self.log_window.append(f"INFO: Preparing to package '{os.path.basename(setup_file_path)}\' using IntuneWinAppUtil.exe.")
        self.log_window.append(f"   Source Folder: {source_folder}")
        self.log_window.append(f"   Setup File: {setup_file_path}")
        self.log_window.append(f"   Output Directory: {output_package_dir}")
        self.log_window.append(f"   Expected output file: {expected_output_filepath}")
        self.statusBar().showMessage(f"Packaging with IntuneWinAppUtil.exe...")

        command = [
            self.intunewin_util_path,
            "-c", source_folder,    # Source folder containing all setup files
            "-s", setup_file_path,  # The setup file (e.g., install.ps1)
            "-o", output_package_dir, # Output directory for the .intunewin file
            "-q"                    # Quiet mode, suppress UI of the tool
        ]

        self.log_window.append(f"CMD: {' '.join(command)}")

        try:
            process = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                check=False,
                # shell=False is default and recommended
            )

            self.log_window.append(f"INFO: IntuneWinAppUtil.exe executed. Return code: {process.returncode}")

            # IntuneWinAppUtil.exe logs to its own log file, but stdout/stderr might have summary/errors
            if process.stdout and process.stdout.strip():
                self.log_window.append("--- stdout (IntuneWinAppUtil.exe) ---")
                self.log_window.append(process.stdout.strip())
            # else:
                # self.log_window.append("INFO: IntuneWinAppUtil.exe produced no stdout or stdout was empty.")

            if process.stderr and process.stderr.strip():
                self.log_window.append("--- stderr (IntuneWinAppUtil.exe) ---")
                self.log_window.append(process.stderr.strip())
            # else:
                # self.log_window.append("INFO: IntuneWinAppUtil.exe produced no stderr or stderr was empty.")

            if process.returncode == 0:
                self.log_window.append(f"INFO: IntuneWinAppUtil.exe completed successfully.")
                # Verify the output file was created
                if os.path.isfile(expected_output_filepath):
                    self.log_window.append(f"SUCCESS: Package '{expected_output_filepath}\' created successfully.")
                    self.statusBar().showMessage(f"Package '{expected_output_filename}\' created successfully!")
                    return True
                else:
                    self.log_window.append(f"ERROR: IntuneWinAppUtil.exe reported success, but output file '{expected_output_filepath}\' was not found.")
                    self.statusBar().showMessage("Packaging succeeded, but output file missing. Check logs.")
                    return False # Success from tool, but file not found
            else:
                error_message = f"IntuneWinAppUtil.exe failed. Return code: {process.returncode}. Check its log file for details (usually in %TEMP%\\MicrosoftIntuneAppUtil.log or similar)."
                self.log_window.append(f"ERROR: {error_message}")
                self.log_window.append(f"   Attempted command: {' '.join(command)}") # Log the command again on error for easy copy-paste
                self.statusBar().showMessage(error_message + " Check log for details.")
                return False

        except FileNotFoundError:
            error_message = f"ERROR: IntuneWinAppUtil.exe not found at '{self.intunewin_util_path}\'. Please check the path."
            self.log_window.append(error_message)
            self.statusBar().showMessage(error_message)
            return False
        except Exception as e:
            error_message = f"ERROR: An unexpected error occurred while running IntuneWinAppUtil.exe: {e}"
            self.log_window.append(error_message)
            self.log_window.append(traceback.format_exc())
            self.statusBar().showMessage("An unexpected error occurred during packaging. Check log.")
            return False

    def handle_package_button_clicked(self):
        """Orchestrates the entire packaging process when the package button is clicked."""
        self.log_window.append("INFO: 'Create .intunewin Package' button clicked.")
        self.statusBar().showMessage("Starting packaging process...")

        # 1. Validation Checks
        if not self.selected_app_data:
            self.log_window.append("ERROR: No application selected for packaging.")
            self.statusBar().showMessage("Error: Please select an application first.")
            return
        app_id = self.selected_app_data.get("ID")
        app_name = self.selected_app_data.get("Name")
        app_version = self.selected_app_data.get("Version")
        if not all([app_id, app_name, app_version]):
            self.log_window.append("ERROR: Selected application data is incomplete (missing ID, Name, or Version).")
            self.statusBar().showMessage("Error: Selected application data incomplete.")
            return

        output_package_dir = self.output_folder_input.text()
        if not output_package_dir or not os.path.isdir(output_package_dir): # Also check if it's a valid dir
            self.log_window.append("ERROR: Output folder for .intunewin package is not set or invalid.")
            self.statusBar().showMessage("Error: Please set a valid output folder for the package.")
            return

        if not self.intunewin_util_path or not os.path.isfile(self.intunewin_util_path):
            self.log_window.append("ERROR: Path to IntuneWinAppUtil.exe is not set or invalid. Configure it in Packaging Configuration.")
            self.statusBar().showMessage("Error: IntuneWinAppUtil.exe path not configured.")
            return
        
        self.log_window.append(f"INFO: Starting packaging for: {app_name} - {app_id} - {app_version}")

        # 2. Create Temporary Directory
        temp_dir = self._create_temp_packaging_dir()
        if not temp_dir:
            self.log_window.append("ERROR: Failed to create temporary directory. Aborting packaging.")
            self.statusBar().showMessage("Error: Failed to create temp directory. Check logs.")
            return

        try:
            # 3. Download Installer
            self.log_window.append("INFO: Step 1: Downloading installer...")
            if not self._download_installer(app_id, app_version, temp_dir):
                self.log_window.append("ERROR: Failed to download installer. Aborting packaging.")
                # StatusBar message is set by _download_installer
                return # No cleanup here yet, will be handled by a finally block or dedicated method later
            self.log_window.append("INFO: Installer download step completed.")

            # 4. Generate Scripts
            self.log_window.append("INFO: Step 2: Generating PowerShell scripts...")
            if not self._generate_install_script(app_id, app_name, app_version, temp_dir):
                self.log_window.append("ERROR: Failed to generate main install script. Aborting packaging.")
                return
            if not self._generate_uninstall_script(app_id, app_name, temp_dir):
                self.log_window.append("ERROR: Failed to generate uninstall.ps1. Aborting packaging.")
                return
            if not self._generate_detection_script(app_id, app_name, app_version, temp_dir):
                self.log_window.append("ERROR: Failed to generate detection.ps1. Aborting packaging.")
                return
            self.log_window.append("INFO: PowerShell script generation completed.")

            # 5. Run IntuneWinAppUtil.exe
            self.log_window.append("INFO: Step 3: Packaging with IntuneWinAppUtil.exe...")
            if not self.install_script_path: # Should have been set by _generate_install_script
                self.log_window.append("CRITICAL ERROR: install_script_path not set after script generation. Aborting.")
                self.statusBar().showMessage("Critical error: Install script path missing.")
                return
            
            packaging_successful = False # Flag to track overall success for cleanup decision
            if self._run_intunewin_app_util(temp_dir, self.install_script_path, output_package_dir):
                self.log_window.append(f"SUCCESS: Successfully packaged {app_name} to {output_package_dir}")
                self.statusBar().showMessage(f"Successfully packaged {app_name}!")
                packaging_successful = True
            else:
                self.log_window.append("ERROR: Failed to package with IntuneWinAppUtil.exe. Check logs.")
                # StatusBar message set by _run_intunewin_app_util
                # packaging_successful remains False
                return # Abort if packaging itself fails

        finally:
            if temp_dir and os.path.isdir(temp_dir):
                if packaging_successful: # Only cleanup on full success
                    self._cleanup_temp_directory(temp_dir)
                else:
                    self.log_window.append(f"INFO: Packaging failed or was aborted. Temporary directory '{temp_dir}\' will be kept for debugging.")
                    self.statusBar().showMessage(f"Packaging failed. Temp files kept at: {temp_dir}")

    def _cleanup_temp_directory(self, temp_dir_path):
        """Recursively deletes the specified temporary directory."""
        if temp_dir_path and os.path.isdir(temp_dir_path):
            self.log_window.append(f"INFO: Attempting to cleanup temporary directory: {temp_dir_path}")
            try:
                shutil.rmtree(temp_dir_path)
                self.log_window.append(f"INFO: Successfully cleaned up temporary directory: {temp_dir_path}")
                # Clear the instance variable if it matches the one being cleaned
                if self.current_temp_dir == temp_dir_path:
                    self.current_temp_dir = None 
            except Exception as e:
                error_message = f"ERROR: Failed to cleanup temporary directory '{temp_dir_path}': {e}"
                self.log_window.append(error_message)
                self.log_window.append(traceback.format_exc())
                self.statusBar().showMessage(f"Warning: Failed to cleanup temp directory '{os.path.basename(temp_dir_path)}\'.")
        else:
            self.log_window.append(f"INFO: Temporary directory '{temp_dir_path}\' not found or already cleaned up.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.apply_dark_mode() # Apply dark mode
    window.show()
    sys.exit(app.exec()) 