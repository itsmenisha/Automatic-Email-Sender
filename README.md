# 📧 Automatic Email Sender - CineMarga Edition

This Python-based tool automates the process of composing and sending HTML-rich emails using Microsoft Outlook. Designed for **CineMarga**, this sender includes a professional HTML template, dynamic tables linking to CineMarga's online presence, and a personal message for outreach and collaboration. It supports both test messages and fully formatted marketing emails, either sent automatically or displayed for review before sending.

This script uses the `win32com.client` module to interface directly with the Outlook application installed on Windows, making it efficient for internal and external communications without needing third-party email libraries or servers.

---

## 🚀 Features

- Compose professional HTML emails from the command line
- Choose between test or CineMarga marketing email templates
- Dynamic email fields (To, CC, Subject)
- Automatically open the email in Outlook or send it directly
- Built-in safety checks for formatting email addresses
- Custom CineMarga branding, links, and contact information

---

## 🛠️ Requirements

- Windows OS
- Microsoft Outlook (must be installed and configured)
- Python 3.x
- `pywin32` library  
  *(Install using `pip install pywin32`)*

---

## 🔐 Before You Start

1. **Login to Your Outlook Account**  
   Ensure you're logged into Outlook on your system and it is configured with a valid email account.

2. **Install Dependencies**  
   Open Command Prompt or terminal and run:
   ```
   pip install pywin32
   ```

3. **Verify Outlook is Running**  
   Start the Outlook desktop app before running this script to avoid any COM initialization errors.

---

## ▶️ How to Use

1. **Run the script**  
   In your terminal, navigate to the script's directory and run:
   ```
   python email_sender.py
   ```

2. **Choose Email Type**  
   You will be prompted:
   ```
   Choose what kind of email to build:
     1. CineMarga
     2. test
   ```

3. **Enter Email Addresses**  
   You'll be asked to input "TO" and "CC" addresses. You can:
   - Leave blank to use default values
   - Enter `e` to skip the field
   - Enter multiple addresses separated by `;`

4. **Choose Auto-Send Option**  
   ```
   automatic sending? [y/n] (default no):
   ```
   - `y` to send the email immediately
   - `n` to preview it in Outlook first

---

## 📂 File Structure

```
email_sender.py       # Main Python script
README.md             # This documentation
```
