# spreadsheet-app

Overview:
This project is a simple spreadsheet application that mimic Google Sheets UI, built using React. It provides features such as data entry, basic mathematical functions, cell formatting, and the ability to save and reload spreadsheets using Excel files.

Features:
- Data entry with support for text, numbers, and dates
- Drag function and cell dependencies
- Mathematical functions: SUM, AVERAGE, MAX, MIN, COUNT
- Cell formatting: Bold, Italic, Font Size, Color
- Column and Row resizing
- Add and Delete Rows/Columns
- Trim, Uppercase, and Lowercase operations
- Remove duplicate rows
- Find and Replace functionality
- Save spreadsheet as an Excel file
- Load spreadsheet from an Excel file

Installation:
Prerequisites:
Ensure you have the following installed:
- Node.js (>= 14.x)
- npm or yarn

Setup:
1. Clone the repository: git clone https://github.com/AKI-28/spreadsheet-app.git
2. Navigate to the project directory: cd spreadsheet-app/frontend
3. Install dependencies: npm install
4. Start the development server: npm start

Usage:
Once the application is running, you can perform various spreadsheet operations such as entering data, formatting, and applying formulas.

Components:
'Spreadsheet.js'
This file contains the main logic for the spreadsheet interface. It includes:
- State management for cells, formatting, and file operations
- Event handlers for user interactions
- Functionality for formula evaluation

'Spreadsheet.css'
Contains styles for the spreadsheet layout and formatting.

Dependencies:
'react': Frontend library
'xlsx': Library for handling Excel file operations

Build:
To create a production build: npm run build

Contributing:
Feel free to fork this project and contribute by submitting pull requests.

License:
This project is licensed under the MIT License.

