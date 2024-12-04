# Warehouse Pivot Application

A robust Windows Forms-based application designed for inventory management, filtering, and exporting data. This tool is specifically crafted to assist inventory departments in 3PL (Third-Party Logistics) environments, enabling seamless cycle counting and data analysis workflows.

---

## 🚀 Features

### ✅ Excel File Loader
- Load `.xlsx` files using the EPPlus library for efficient Excel file handling.
- Automatically detects important columns, such as:
  - `Location`
  - `SKU`
  - `Movable Unit Label`
- Displays all data in an easy-to-read `ListView`.

### ✅ Dynamic Filtering
- Filter data dynamically by entering criteria in combo boxes for:
  - `Location`
  - `SKU`
  - `Movable Unit Label`
- Real-time filtering is intelligently delayed (2 seconds) to improve performance and prevent excessive searches.

### ✅ Export Options
- Export filtered data to an HTML table for printing or further use.
- Two export styles available:
  - **Basic**: Clean and simple.
  - **Modern**: Professionally styled with CSS and optimized for printing on a single page.

### ✅ Progress Bar Feedback
- Provides a progress indicator during Excel file loading to enhance user experience.

### ✅ Debugger Panel
- Logs application events and errors in a `RichTextBox` with color-coded messages:
  - Blue: Information
  - Orange: Warnings
  - Red: Errors

### ✅ Custom Action Column
- Adds an empty `Action` column in the export for employees to manually input data, such as checkmarks or notes.

---

## 🛠 Requirements

- **Operating System**: Windows
- **.NET Framework**: Version 4.7.2 or higher
- **Dependencies**:
  - [EPPlus](https://github.com/EPPlusSoftware/EPPlus) for Excel file handling.

---

## 📖 How It Works

### Loading an Excel File
1. Click on **File > Open** to load an `.xlsx` file.
2. The application will:
   - Parse the first worksheet of the file.
   - Display all rows and columns in `ListView1`.
   - Populate unique values in the combo boxes (`Location`, `SKU`, `Movable Unit Label`) for filtering.

### Filtering Data
1. Use the combo boxes to filter data dynamically.
2. Filters are applied to the data in `ListView1` based on:
   - `Location`
   - `SKU`
   - `Movable Unit Label`
3. The filtering logic waits 2 seconds after typing to avoid triggering multiple searches unnecessarily.

### Exporting Filtered Data
1. Enter the `Location` prefix (e.g., `A1`, `B2`) in the `Location Filter` textbox.
2. Click the **Export Location** button.
3. Choose the desired export style:
   - **Basic**: A plain and clean table.
   - **Modern**: Enhanced with CSS for professional presentation.
4. The filtered data will be:
   - Exported as an HTML file to your **Desktop**.
   - Automatically opened in your default web browser.

---

## 🎯 Example Use Case: Cycle Counting in a 3PL Environment

1. **Load Data**:
   - Import an Excel inventory report exported from your Warehouse Management System (WMS).
2. **Filter**:
   - Use the combo boxes to focus on specific `Location` ranges (e.g., `A1-01` to `A1-50`) or filter by `SKU` and `Movable Unit Label`.
3. **Export**:
   - Export the filtered data into a clean HTML table for printing or sharing with your team.
   - The table includes an empty `Action` column for employees to manually mark items as counted or inspected.

---

## 📷 Screenshots

### Application Interface
![Application Interface](./screenshots/app_interface.png)

### Filtered Data Export (Modern Style)
![HTML Export - Modern Style](./screenshots/html_export_modern.png)

---

## 💻 Installation and Setup

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/YourUsername/WarehousePivot.git
