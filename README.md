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
