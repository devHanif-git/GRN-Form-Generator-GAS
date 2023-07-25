<h1 align="center">Google Apps Script - GRN Form Generator</h1>

<p align="center">
  <strong>A Google Apps Script project to automatically generate GRN (Good Receiving Note) Forms from spreadsheet data.</strong>
</p>

<p align="center">
  <a href="https://your-live-demo-link.com/">View Demo</a>
  ·
  <a href="https://github.com/devHanif-git/GRN-Form-Generator/issues">Report Bug</a>
  ·
  <a href="https://github.com/devHanif-git/GRN-Form-Generator/issues">Request Feature</a>
</p>

## Table of Contents

- [About](#about)
- [How it Works](#how-it-works)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## About

The Google Apps Script - GRN Form Generator is a script that automates the process of generating GRN (Good Receiving Note) Forms from data in a Google Sheets spreadsheet. It allows you to create multiple GRN Forms for different suppliers and organize them in a designated destination folder.

## How it Works

The script reads data from a Google Sheets spreadsheet with specific column configurations:
- Column A: Date
- Column B: Supplier
- Column C: Customer
- Column D: Delivery Order Number (DONO)
- Column E: GRN Number
- Column F: Purchase Order Number (PO)
- Column G: Model Number

Each row in the spreadsheet represents an item received, and the script generates a separate GRN Form for each supplier containing the relevant item details.

The generated GRN Forms are based on predefined Google Docs templates, stored in Google Drive, with placeholders like `{{DATE}}`, `{{GRNNO}}`, `{{SUPPLIER}}`, and others. The script replaces these placeholders with actual data from the spreadsheet.

If the number of items in a GRN Form exceeds a certain threshold, the script creates additional GRN Forms (e.g., "GRN Form (2)", "GRN Form (3)", etc.) and spreads the item details across these forms to avoid exceeding the Google Docs' document length limit.

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/devHanif-git/GRN-Form-Generator.git
    ```
2. Open the `code.gs` file in the Google Apps Script editor.
3. Set up your Google Sheets spreadsheet with the required column configurations mentioned in the "How it Works" section.
4. Create Google Docs templates for GRN Forms with placeholders for item details.
5. Obtain the ID of each template file from Google Drive and replace the template IDs in the `code.gs` file.

## Usage 

1. Open your Google Sheets spreadsheet containing the item data.
2. In the toolbar, click on "Add-ons" > "GRN" > "Generate GRN Form 2023" (The menu name can be customized).
3. The script will automatically generate GRN Forms based on the data in the spreadsheet and save them in the designated destination folder in Google Drive.
4. The status column in the spreadsheet will be updated to "FINISHED" for each processed row.
5. If the number of items in a GRN Form exceeds the limit, the script will create additional forms and spread the items across them.

## Contributing

Contributions are welcome! If you have any ideas, suggestions, or bug reports, please [create an issue](https://github.com/devHanif-git/Digital-Clock/issues).

1. Fork the project.
2. Create your feature branch (`git checkout -b feature/YourFeature`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Open a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
