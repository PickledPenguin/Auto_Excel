# Excel Automation Script

## Description

This script automates data extraction from Excel files located in the `Excel` directory. It simulates mouse movements and keypresses based on user-defined configurations.

## Project Background

Developed for the Rochester Institute of Technology's Software Engineering department, this script eliminates manual input of student employees' hours from Excel into an external website. It was designed to save hours of work every week previously spent on repetitive tasks. It is also adaptable to other tasks that rely on manually entering Excel data. 
For example, I personally used this script as a Course Assistant at RIT to automate grade entry from Excel into MyCourses after updates by another assistant and professor confirmation.

## Architecture Details

The script relies on a waypoint system set by the user on the screen, enabling flexibility across different websites without code redesign. Initial setup is required, and the screen must remain static during operation.

## Installation Requirements

1. **Python3**: [Install Python3](https://www.python.org/downloads/)
2. **pip3**: [Install pip3](https://pip.pypa.io/en/stable/installation/)
3. **Modules**: Install dependencies listed in `requirements.txt`. [Guide here](https://note.nkmk.me/en/python-pip-install-requirements/)

## Configuration

1. Open a terminal or cmd window.
2. Navigate to the project directory.
3. Run the setup command:
   
   ```sh
   python3 ./setup.py
   ```

4. Configure the following options:
   - **Wait Time**: Set time (in seconds) between actions.
   - **Datetime Format**: Define datetime display format. [Format codes](https://strftime.org/)
   - **Waypoints**: Define actions (e.g., click, enter) using keyboard shortcuts.

### Waypoint Actions:

| Key | Action |
|----|----------------------|
| `c` | Click |
| `d` | Double Click |
| `t` | Tab |
| `e` | Enter |
| `p` | Paste (input text) |
| `i` | Insert Data (Excel input) |
| `w` | Wait (pause execution) |

Press `esc` to save configurations and exit.

## Execution Types

Choose between:
- **File-based**: Process each Excel file in the directory.
- **Row-based**: Process each row of data in each Excel file.

## Running the Script

1. Ensure Excel files are in the `Excel` directory.
2. Run the script:
   
   ```sh
   python3 ./auto_excel.py
   ```

3. Position your mouse on the target screen/window.
4. The script will execute configured actions to input data from Excel into the website.

## Common Issues

- **Non-Static Screen**: Make sure the screen remains untouched and static during operation.
- **Non-Looping Configuration**: Configure the script to return to the starting position after each cycle, so it can loop properly
- **Can't Accessing Excel Files**: Verify correct file path and permissions.
- **Handling Unusual Excel Data**: Script converts most data types to strings but may encounter issues with unconventional data.
