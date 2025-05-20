## **Project Overview: Winget to Intunewin Packaging Tool**

**Project Title:** Winget2Intunewin GUI Packer

**Project Goal:** To develop a user-friendly graphical interface (GUI) application that streamlines the process of finding and packaging the latest versions of applications available via Winget into .intunewin files, making them ready for manual upload to Microsoft Intune.

**Target Audience:** Infrastructure Engineers and IT Administrators who frequently need to create .intunewin packages for Win32 application deployment in Microsoft Intune.

**Problem Statement:** The current manual workflow for packaging applications for Intune, which involves finding installers, writing installation scripts, and using the command-line IntuneWinAppUtil.exe tool, is inefficient and repetitive. Automating these steps will save significant time and reduce potential errors.

Solution Overview:  
This project will deliver a desktop application with a simple interface. Users will be able to search the Winget repository, select an application, and with minimal further input, the tool will handle the download of the application installer, generate standard PowerShell scripts for installation, uninstallation, and detection, and finally bundle everything into a .intunewin file using the Microsoft Win32 Content Prep Tool. The resulting .intunewin file can then be manually uploaded to the Intune portal by the administrator.

### **Key Features:**

1. **Winget Application Search:**  
   * A dedicated area in the GUI to input an application name or keyword.  
   * A button to execute a winget search command in the background.  
   * A display table or list to present the search results, including key information like Application Name, ID, Version, and Source.  
2. **Application Selection:**  
   * Enable users to easily select a specific application from the search results list for packaging.  
3. **Automated Installer Download:**  
   * Functionality to automatically download the selected application's installer using the winget download command. The installer will be saved to a temporary, structured directory.  
4. **Automated PowerShell Script Generation:**  
   * Generate standard install.ps1, uninstall.ps1, and detection.ps1 scripts tailored for the selected Winget application ID.  
   * These scripts will utilize winget install and winget uninstall for core actions and winget list for detection.  
5. **.intunewin File Creation:**  
   * Integrate the Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe).  
   * Provide an option for the user to specify the output folder for the final .intunewin file.  
   * Execute IntuneWinAppUtil.exe programmatically, passing the temporary source folder (containing the installer and scripts) and the chosen output folder.  
6. **Status and Logging:**  
   * A status area or log window within the GUI to display the progress of each step (searching, downloading, scripting, packaging).  
   * Report success or failure messages and any relevant output or errors from Winget or IntuneWinAppUtil.exe.

### **Technical Stack:**

* **Programming Language:** Python 3.x  
* **GUI Framework:** PyQt6 or PySide6 (Provides robust widgets and a designer tool for building the UI)  
* **External Dependencies:**  
  * **Winget (Windows Package Manager):** Must be installed and accessible on the operating system where the tool runs.  
  * **Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe):** The tool will need to be downloaded and its path configured within the application.  
* **Python Libraries:**  
  * subprocess: To execute external commands (winget, IntuneWinAppUtil.exe).  
  * os, shutil: For managing file paths, creating temporary directories, and cleaning up.  
  * logging: For tracking application events and errors.

### **High-Level Milestones:**

1. **GUI Foundation:** Set up the sleek modern application window in dark mode, layout, and essential UI elements (input fields, buttons, lists, status display).  
2. **Winget Search Implementation:** Add the functionality to search Winget and display results in the GUI.  
3. **Winget Download Integration:** Implement the logic to download the selected application's installer to a temporary location.  
4. **Script Generation Logic:** Develop the Python code to generate the standard PowerShell scripts with the correct Winget ID.  
5. **IntuneWinAppUtil Execution:** Integrate the call to IntuneWinAppUtil.exe, passing the necessary parameters.  
6. **Error Handling and User Feedback:** Add comprehensive error handling for all external process calls and update the GUI with clear status information.  
7. **File Output Configuration:** Allow the user to select the destination folder for the .intunewin file.  
8. **Temporary File Cleanup:** Implement cleanup of the temporary download and script directories after successful packaging.

### **Potential Future Enhancements:**

* **Custom Script Integration:** Allow users to provide their own installation/uninstallation/detection scripts.  
* **Metadata Editing:** Provide fields to edit application metadata (Name, Publisher, Description) before packaging.  
* **Requirements Configuration:** Add options to configure minimum requirements (OS version, architecture, disk space, etc.) which could potentially be included in the detection script or noted for manual Intune configuration.  
* **Batch Packaging:** Allow selecting and packaging multiple applications in sequence.

This project structure provides a clear path to building a valuable automation tool for your Intune packaging tasks.