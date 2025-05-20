# Winget2Intunewin GUI Packer

A utility to streamline the creation of `.intunewin` packages for Microsoft Intune by leveraging Winget for application discovery and download.

## Features

*   **Application Search**: Search for applications using Winget right from the GUI.
*   **Version Selection**: View available versions and select the specific one needed.
*   **Automated Download**: Downloads the selected application installer using Winget.
*   **PowerShell Script Generation**: Automatically generates:
    *   `install.ps1`: For installing the application via Winget (dynamically named based on the app).
    *   `uninstall.ps1`: For uninstalling the application via Winget.
    *   `detection.ps1`: For Intune to detect if the application is installed correctly.
*   **.intunewin Creation**: Integrates with `IntuneWinAppUtil.exe` to package the downloaded installer and generated scripts into an `.intunewin` file.
*   **Persistent Configuration**: Remembers the path to your `IntuneWinAppUtil.exe`.
*   **Dark Mode UI**: A sleek, modern user interface.
*   **Logging**: Provides detailed logs of its operations in the "Action Status" area.
*   **Temporary File Management**: Creates a temporary directory for packaging and cleans it up on success, or preserves it on failure for debugging.

## Prerequisites

*   **Python**: Python 3.8 or newer is recommended.
*   **Winget**: The Windows Package Manager (`winget.exe`) must be installed and accessible in your system's PATH.
*   **Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`)**:
    *   This tool is required for creating `.intunewin` packages.
    *   You can download it from [Microsoft's GitHub repository](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool).
    *   The application will prompt you to specify the path to this executable on its first use or if the path is not configured.

## Setup

1.  **Clone the Repository (Optional)**:
    If you have cloned this project from a Git repository:
    ```bash
    git clone <repository_url>
    cd winget-automate 
    ```
    If you have the files directly, navigate to the project directory.

2.  **Create a Virtual Environment (Recommended)**:
    ```bash
    python -m venv venv
    ```
    Activate the virtual environment:
    *   Windows (Command Prompt/PowerShell):
        ```cmd
        venv\Scripts\activate
        ```
    *   Linux/macOS (bash/zsh):
        ```bash
        source venv/bin/activate
        ```

3.  **Install Dependencies**:
    Ensure your `requirements.txt` file is up-to-date with all necessary packages (e.g., `PySide6`).
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Run the Application**:
    ```bash
    python main_window.py
    ```

2.  **Configure `IntuneWinAppUtil.exe` Path**:
    *   If not already configured, the application might prompt or you can set it in the "Packaging Configuration" section.
    *   Click "Browse..." next to "Path to IntuneWinAppUtil.exe" and locate the `IntuneWinAppUtil.exe` file on your system. This path will be saved for future sessions.

3.  **Search for an Application**:
    *   Enter the name or keyword of the application you want to package in the "Search for Application:" input field (e.g., "vscode", "7zip").
    *   Click the "Search" button.
    *   Search results will appear in the table below, showing Name, ID, Version, and Source.

4.  **Select an Application**:
    *   Click on a row in the search results table to select an application.
    *   The "Selected App" field in the "Packaging Configuration" section will update.

5.  **Set Output Folder**:
    *   In the "Packaging Configuration" section, click "Browse..." next to "Output Folder for .intunewin:".
    *   Select a directory where you want the final `.intunewin` package to be saved.

6.  **Create Package**:
    *   Click the "Create .intunewin Package" button.
    *   The application will:
        1.  Create a temporary working directory.
        2.  Download the selected application installer using Winget.
        3.  Generate the necessary PowerShell scripts (`<AppName>.ps1` for install, `uninstall.ps1`, `detection.ps1`).
        4.  Run `IntuneWinAppUtil.exe` to package these files into an `<AppName>.intunewin` file in your specified output folder.
    *   Monitor the "Action Status" log window for detailed progress and any errors.
    *   The status bar will also provide brief updates.

7.  **Upload to Intune**:
    *   The generated `<AppName>.intunewin` file is ready to be uploaded to Microsoft Intune as a Windows app (Win32).
    *   When configuring the app in Intune:
        *   **Install command**: `powershell.exe -ExecutionPolicy Bypass -File .\<AppName>.ps1` (replace `<AppName>.ps1` with the actual script name).
        *   **Uninstall command**: `powershell.exe -ExecutionPolicy Bypass -File .\uninstall.ps1`
        *   **Detection rules**: Use the "Use a custom detection script" option and upload the generated `detection.ps1` file.

## Project Structure Notes

*   The main application logic and GUI are currently in `main_window.py` located in the project root.
*   The `src/` directory exists and contains an `__init__.py`, indicating it's a Python package, though current UI-centric logic is primarily in the root. Future non-GUI logic might be refactored here.
*   `requirements.txt` lists the Python dependencies.

## Troubleshooting

*   **"Winget command not found"**: Ensure Winget is installed correctly and that its installation directory is part of your system's PATH environment variable.
*   **"IntuneWinAppUtil.exe path not configured"**: Use the "Browse..." button in the "Packaging Configuration" section to set the correct path to `IntuneWinAppUtil.exe`.
*   **Packaging Fails**:
    *   Check the "Action Status" log in the application for detailed error messages from Winget or `IntuneWinAppUtil.exe`.
    *   If packaging fails, the temporary working directory (path shown in the status bar/logs) is usually preserved. Inspect its contents (downloaded installer, generated scripts) for clues.
    *   `IntuneWinAppUtil.exe` also creates its own log file, typically in `%TEMP%\MicrosoftIntuneAppUtil.log` or a similarly named file in that directory, which can provide more detailed packaging errors.

## License

This project is currently not under a specific license. Please contact the maintainer for licensing inquiries. 