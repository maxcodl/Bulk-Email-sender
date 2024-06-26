

# Gmail Bulk Email Sender

This project automates the process of sending personalized emails to a list of contacts using Gmail. It uses `undetected_chromedriver`, `selenium`, and `pyautogui` to interact with the Gmail web interface and send emails.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Contributing](#contributing)
- [License](#license)

## Features

- Sends personalized emails to a list of contacts from an Excel file.
- Automates the process of composing and sending emails in Gmail.
- Attaches a file to each email.
- Uses `undetected_chromedriver` to avoid detection by Gmail's bot protection.

## Requirements

- Python 3.x
- Google Chrome
- `undetected_chromedriver`
- `selenium`
- `pandas`
- `pyautogui`
- An Excel file with the contacts information

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/gmail-bulk-email-sender.git
    cd Bulk-Email-sender
    ```

2. **Prepare your Excel file:**

    Ensure you have an Excel file named `contacts.xlsx` in the same directory with the following columns:
    - `User Emails`: The email address of the recipient.
    - `First Name`: The first name of the recipient.

## Usage

1. **Run the Program:**

    ```bash
    start.bat
    ```

2. **Process:**
    - The script will open a Chrome browser and navigate to the Gmail inbox page.
    - It will read the contacts from `contacts.xlsx`.
    - For each contact, it will compose an email with a personalized message and attach a file.
    - It will then send the email and proceed to the next contact.

## Configuration

- **Email Subject:** You can change the subject of the email by modifying the `subject` variable in the script.
- **Email Body:** You can customize the body of the email by modifying the `body` variable in the script.
- **Attachment Path:** Ensure the file path in `pyautogui.typewrite` matches the location of your attachment file.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any changes or improvements.

## View Counter 
![Counter](https://profile-counter.glitch.me/Bulk-Email-sender/count.svg)
