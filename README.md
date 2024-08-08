# Email Automation Tool

## Overview

This project is a Python-based email automation tool that reads email data from an Excel file and sends personalized emails with embedded images via the Gmail API. It is designed to help users send bulk emails efficiently, embedding images directly within the email body to enhance visual appeal.

## Features

- Read recipient details, custom paragraphs, and subject lines from an Excel file.
- Send personalized emails with HTML content.
- Embed images directly within the email body.
- Uses the Gmail API for secure and reliable email delivery.

## Prerequisites

Before you begin, ensure you have met the following requirements:

- Python 3.x installed on your machine.
- Required Python packages installed (see `requirements.txt`).
- A Google Cloud project with the Gmail API enabled.
- OAuth 2.0 credentials (`credentials.json`) for accessing the Gmail API.

## File Structure

Here’s the structure of the project directory:

```
EmailAutomationTool/
├── emailer.py
├── credentials.json
├── test.xlsx
├── your_image.jpeg
├── requirements.txt
└── README.md
```

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/email-automation-tool.git
   cd email-automation-tool
   ```

2. Install the required packages:

   ```bash
   pip install -r requirements.txt
   ```

3. Ensure you have the following files in the project directory:
   - `credentials.json` (OAuth 2.0 credentials from Google Cloud)
   - `test.xlsx` (Excel file with email data)
   - `your_image.jpeg` (Image to embed in the email)

## Configuration

1. **Excel File Structure**:

   The Excel file (`test.xlsx`) should be structured as follows:

   | Column | Data                    |
   |--------|-------------------------|
   | A      | Email Address           |
   | B      | Receiver Name           |
   | C      | Subject Line            |
   | D      | Custom Paragraph        |

2. **Google Cloud Setup**:
   
   - Create a project in the [Google Cloud Console](https://console.cloud.google.com/).
   - Enable the Gmail API for your project.
   - Set up OAuth 2.0 credentials and download the `credentials.json` file.
   - Place the `credentials.json` file in the project directory.

## Usage

1. Run the Python script:

   ```bash
   python emailer.py
   ```

   The script will:
   - Read the email data from `test.xlsx`.
   - Send personalized emails to each recipient.
   - Embed `your_image.jpeg` in the email body.

## Customization

- **Common Text**: You can modify the `common_text` variable in `emailer.py` to change the repeated content for all emails.
- **Image**: Replace `your_image.jpeg` with the desired image file. Ensure the file name is updated in the script.

## Troubleshooting

- Ensure you have an active internet connection.
- Verify that the `credentials.json`, `test.xlsx`, and image files are in the correct directory.
- Check the Google Cloud Console for any issues with the Gmail API.

## Contributing

Contributions are welcome! Please fork the repository and create a pull request with your changes. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [Google APIs Client Library for Python](https://github.com/googleapis/google-api-python-client)
- [Pandas](https://pandas.pydata.org/)

---

Replace `"https://github.com/your-username/email-automation-tool.git"` with the actual URL of your GitHub repository. You can also add or modify sections based on your specific needs.
