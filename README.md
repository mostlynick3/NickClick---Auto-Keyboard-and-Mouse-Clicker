# NickClick - Auto Keyboard and Mouse Clicker

**Automate repetitive mouse and keyboard tasks with a user-friendly GUI for Windows.**

## Description

NickClick is a Python application designed to automate mouse clicks, keyboard inputs, and delays on Windows machines. It allows users to record, edit, save, and schedule macros for repetitive tasks, improving productivity and reducing manual effort. The application features a graphical user interface (GUI) built with tkinter, making it accessible to both technical and non-technical users.

> **Note:** This project description was written by Mistral AI.

## Features

- Record and Playback Macros: Record mouse clicks and keyboard inputs, then replay them as needed.
- Manual Action Creation: Add mouse clicks, keyboard inputs, and delays manually for precise control.
- Script Management: Save, load, and delete scripts in JSON format for easy reuse.
- Script Scheduling: Schedule scripts to run at specific times, with support for one-time, daily, weekly, monthly, or custom interval execution.
- Editable and Lockable Scripts: Lock scripts to prevent accidental edits, with optional password protection.
- Multi-Repetition Execution: Run scripts multiple times with a single click.
- Abort Hotkey: Set a hotkey to stop script execution at any time.
- Dark/Light Theme: Switch between themes for comfortable use in any environment.
- Cross-Platform: Works on Windows, macOS, and Linux (with minor adjustments).

## Requirements

- Python 3.6 or higher
- Required Python packages:
  - tkinter (usually included with Python)
  - pyautogui
  - pynput
  - psutil
  - sv_ttk
  - tkcalendar (optional, for enhanced date selection)

Install dependencies using pip:
pip install pyautogui pynput psutil sv_ttk tkcalendar

## Installation

1. Clone this repository or download the source code.
2. Install the required dependencies (see above).
3. Run the application:
   python nickclick.py

## Usage

### Recording a Macro
1. Click the "Record Macro" button.
2. Press any key to start recording.
3. Perform the actions you want to automate.
4. Press the same key again to stop recording.
5. Save your script for future use.

### Running a Script
1. Load a saved script or record a new one.
2. Set the number of repetitions.
3. Click "Run Script" and follow the prompts to set an abort hotkey.
4. The script will execute according to your settings.

### Scheduling a Script
1. Click the "Schedule Script" button.
2. Set the first execution date and time, repetitions, and frequency.
3. Keep the application running to execute the script as scheduled.

### Managing Scripts
- New Script: Start a new script from scratch.
- Load Script: Open a previously saved script.
- Save Script: Save your current script to a file.
- Delete Script: Remove a script file from your system.
- Lock/Unlock Script: Protect scripts from accidental changes.

## Screenshots
<img width="1052" height="724" alt="image" src="https://github.com/user-attachments/assets/1067aa8a-eb68-46a9-9643-f722c671c054" />
<img width="1052" height="682" alt="image" src="https://github.com/user-attachments/assets/1ec01ce8-10e3-4373-824e-3f990d44b992" />
<img width="1049" height="682" alt="image" src="https://github.com/user-attachments/assets/f7f6c1e6-2e97-432f-b1e8-2fadb20145ec" />
<img width="502" height="569" alt="image" src="https://github.com/user-attachments/assets/a3602ef8-3bab-495d-af84-5a93d31bf634" />


## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License
GPL 3.

## Contact

For questions or feedback, please open an issue on the GitHub repository.
