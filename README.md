# WhatsApp Bulk Sender (Selenium)

A Python script to automate sending WhatsApp messages to a list of numbers from an Excel file. This script uses Selenium to automate the web browser and simulates human typing speeds to reduce the risk of detection.

## 🚀 Features
- **Excel Support:** Reads numbers directly from `num.xlsx`.
- **Intelligent Formatting:** Automatically adds country codes (+91) and cleans formatting.
- **Human-like Behavior:** Random delays between messages and typing simulation.
- **Status Logging:** Saves successful and failed numbers to text files.
- **Session Persistence:** Saves your WhatsApp login session locally so you don't have to scan the QR code every time.

## 🛠️ Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/AlenSarangSatheesh/whatsapp-bulk-sender.git
   cd whatsapp-bulk-sender
   
2. Install dependencies:

   pip install -r requirements.txt


3. Important: You must have Google Chrome installed.

   📋 Usage
   Create an Excel file named num.xlsx in the project folder.

   Put phone numbers in the first column (no header needed).

   Run the script: python main.py

   Scan the QR code when the browser opens (only needed the first time).

Disclaimer
This script is for educational purposes only. Automated messaging may violate WhatsApp's Terms of Service. Use responsibly.