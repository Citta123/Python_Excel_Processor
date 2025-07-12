
# 🧠 Python Excel Processor: Automated .xls to .xlsx Converter with Data Cleaning and Batch Column Logic

This project is a robust Python automation tool built to handle large-scale Excel processing workflows — from converting outdated `.xls` files to `.xlsx` format, to cleaning and updating critical column values based on business rules. Designed for operations teams, analysts, and developers who routinely handle Excel-based reporting and reconciliation.

---

## 🚀 Features

- **🔁 .XLS to .XLSX Converter**  
  Seamlessly converts legacy Excel `.xls` files into modern `.xlsx` format using the Windows COM interface (requires Microsoft Excel).

- **🧹 Data Sanitization**  
  Removes invisible whitespaces and cleans cell values to avoid issues caused by inconsistent manual input.

- **📊 Column Value Automation**  
  Automatically updates:
  - `BL Akhir` values with custom input.
  - `BL Awal` values based on a configuration.
  - `LBR` values computed from comma-separated logic.

- **⚙️ YAML-Based Configuration**  
  Use the `input_config.yaml` file to drive flexible, reusable processing logic — no need to touch the script for every run.

- **🪵 Logging System**  
  All actions and errors are logged to `script_log.txt` to help you troubleshoot with ease.

---

## 📁 Project Structure

```
ProjekXLSX/
├── Last.py                  # Main processing script
├── input_config.yaml        # Dynamic processing configuration
├── Merged_Data.xlsx         # Output Excel file
├── Folder1/                 # Input XLS files
├── Folder2/                 # Input XLSX files
└── script_log.txt           # Auto-generated log file
```

---

## ⚙️ How It Works

1. Ensures all Excel processes are cleanly closed before execution.
2. Reads input file locations and config values from `input_config.yaml`.
3. Converts `.xls` → `.xlsx` files as needed.
4. Applies cleaning rules and updates the following columns:
   - `BL Awal`
   - `BL Akhir`
   - `LBR`
5. Merges and writes final results into `Merged_Data.xlsx`.

---

## 🖥️ Requirements

- Windows OS
- Python 3.x
- Microsoft Excel installed
- Required packages:
  ```bash
  pip install pyyaml pywin32
  ```

---

## 🚦 Usage

1. Place your `.xls` and `.xlsx` files in the appropriate folders (`Folder1`, `Folder2`).
2. Edit `input_config.yaml` to match your desired parameters.
3. Run the script:
   ```bash
   python Last.py
   ```

---

## 🧑‍💻 About the Developer

I'm a freelance Python developer focused on automation, data engineering, and process optimization. This tool is one example of how Python can be used to replace repetitive Excel tasks with smart logic that saves hours of manual work.

If you'd like to hire me to build similar automation systems or help you with messy Excel data — [plusenergi77@gmail.com](mailto:plusenergi77@gmail.com)

---

## 📄 License

This project is licensed under the Apache License 2.0 – you are free to use, modify, and distribute it with proper attribution.
