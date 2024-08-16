# User-Debt-Inquiry-Automation-Script

This Python script automates the process of checking whether users are in debt and records the results. It simulates mouse actions and keyboard inputs to batch process a list of user IDs from an Excel file, checking the status of each user and recording those who are in debt.

## Key Features

- **Batch Processing**: Capable of reading user IDs from an Excel file in bulk.
- **Automated Inquiry**: Automatically simulates mouse clicks and keyboard inputs to check user status.
- **Result Recording**: Records the IDs of users who are in debt into a log file.
- **Checkpoint Resumption**: Supports resuming from the last stopped point for improved efficiency.
- **Logging**: Detailed logging of each step's execution and results.

## Prerequisites

- Ensure that the required Python libraries are installed on your system.
- The script directory should contain the following files:
  - An Excel file named `AAA.xlsx`, containing a list of user IDs.
  - An image file named `out_money.png`, used for image recognition to determine if a user is in debt.

## Usage

1. Install the required Python libraries (see below).
2. Save the script as a `.py` file and ensure it is in the directory with the mentioned files.
3. Run the script. If necessary, run it as an administrator to avoid permission issues.

## Dependencies

- `openpyxl`: For reading and writing Excel files.
- `pyautogui`: For simulating mouse and keyboard operations.
- `time`: For controlling the wait time between operations.
- `logging`: For logging information.

## Installation

Install the required dependencies with the following pip command:

```bash
pip install openpyxl pyautogui
```


Please adjust the content according to the actual situation of your project. If you need further assistance, such as adding specific sections or providing more detailed explanations for certain parts, feel free to let me know.
