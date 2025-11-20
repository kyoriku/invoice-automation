# Invoice Automation

Python CLI tool for automating bi-weekly invoice generation from schedule emails

**Tech Stack:** Python, Google Sheets API, Gmail IMAP/SMTP

**Key Features:** Email parsing • Auto hour calculation • Google Sheets integration • PDF/Excel export • Email delivery • Manual shift editing • Tax summaries

**Why I built it:** Reduced invoicing time by ~80%. Replaced manual process of copying shifts, calculating hours, updating spreadsheets, exporting files, and sending emails.

![Invoice Automation System](assets/banner.webp)

<details>
<summary><b>Built With</b></summary>

[![Python](https://img.shields.io/badge/Python-3776AB.svg?style=for-the-badge&logo=Python&logoColor=white)](https://www.python.org/)
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853.svg?style=for-the-badge&logo=Google-Sheets&logoColor=white)](https://developers.google.com/sheets/api)
[![Gmail](https://img.shields.io/badge/Gmail-EA4335.svg?style=for-the-badge&logo=Gmail&logoColor=white)](https://developers.google.com/gmail/api)

</details>

## Technical Details

**Email Processing**
- IMAP connection to Gmail
- Regex parsing for old and new Homebase email formats
- Automatic shift extraction with duplicate detection

**Invoice Generation**
- Google Sheets API for spreadsheet creation
- Template copying via Google Drive API
- Auto hour calculation with HST (13%)
- Sequential invoice numbering

**Data Management**
- JSON storage for shifts and invoice tracking
- Configurable bi-weekly pay period detection
- Manual shift editing, adding, deletion

**Export & Delivery**
- PDF/Excel export with custom formatting
- SMTP email with attachments
- Interactive CLI with shift review before sending

## License
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge&logo=mit)](https://opensource.org/licenses/MIT)

This project is licensed under the [MIT](https://opensource.org/licenses/MIT) license.

## Questions
For questions, email me at devkyoriku@gmail.com.