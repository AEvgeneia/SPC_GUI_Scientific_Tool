# GPR-SPC Analyzer: Statistical Process Control for Gamma Passing Rate Data


> Version: 1.1.0 – December 2025

**GPR-SPC Analyzer** is a Python-based GUI application (available as executables for Windows and macOS) for performing Statistical Process Control (SPC) on Gamma Passing Rate (GPR) data. Developed as a research tool for analyzing and exploring SPC methods in radiotherapy QA, it supports both standard and heuristic approaches, allows interactive outlier detection and elimination, and exports analysis results as CSV files and PDF charts.

---

## 📌 Features

- Load data from `.xlsx`, `.xls`, or `.csv` files using a predefined template.
- Once loaded:
  - The data is automatically sorted by QA date.
  - A preview of the loaded dataset is shown, along with a summary of which data entries are selected for analysis.
  - Basic descriptive statistics are calculated for selected Gamma metrics.
  - Histograms can be generated to visualize data distributions.
  - The Anderson-Darling test can be performed to assess normality.
- SPC Analysis (core functionality):
  - Choose one of four SPC methods:
    - Shewhart I - Control Charts (for normally distributed data)
    - Weighted Standard Deviation (WSD) I - Control Charts
    - Scaled Weighted Variance (SWV) I - Control Charts
    - Skewness Correction (SC) I - Control Charts
  - Detect and visualize outliers on control charts.
  - Interactive I‑charts: click any point to instantly see its ID and inspect it before elimination.
  - Use the interactive elimination loop to remove selected outliers and recalculate control and specification limits in real-time.
- Export results:
  - PDF files with SPC charts and histograms.
  - CSV files containing normality test results, summary statistics, and elimination logs.

---

## 🆕 What’s New in Version 1.1.0

This update introduces improvements in structure, usability and stability:

### ✔ Code & Architecture
- Rewrote the entire GUI using an **Object-Oriented Programming (OOP)** architecture.
- Improved structure for easier maintenance and future feature expansion.

### ✔ User Interface Enhancements
- Replaced list-selection widgets with **checkbox-based selectors** for clearer metric selection.
- Improved layout aesthetics and visual consistency across all tabs.

### ✔ Bug Fixes
- Fixed a minor issue affecting internal event handling when switching SPC methods.
- Fixed an issue affecting the file loading.

### ✔ Documentation
- Added a new **detailed analytical manual** (“User Manual v_1_1_0”), included in the repository.

---


## 🎯 Who Should Use This?

- Researchers and academics working on QA processes or SPC techniques in radiotherapy
- Medical physicists analyzing Gamma Passing Rate (GPR) data
- Students and scientists studying applied statistics or quality control in medical physics

---

## 🖥️ How to Use the App

### Option 1: Run the Executable (Windows & macOS)

- **Windows**  
  - Download the `.zip` package from the [Releases](#-releases) section.  
  - Extract the contents and run the executable file.  
  - If Windows SmartScreen or antivirus shows a warning, choose **"Run anyway"** (the app is safe).  

- **macOS (Apple Silicon – M1/M2)**  
  - Download the `.zip` package from the [Releases](#-releases) section.  
  - Extract the contents and run the application.  
  - If macOS Gatekeeper warns that the app is from an unidentified developer, right-click the app, select **"Open"**, and confirm.

---

### Option 2: Run from Python Source (cross-platform)

Ensure Python 3 is installed. Then:

```bash
git clone https://github.com/AEvgeneia/SPC_GUI_Scientific_Tool.git
cd SPC_GUI_Scientific_Tool
pip install -r requirements.txt
python SPC_GUI.py
```
---

## 📖 User Manual

A complete, step-by-step manual is available in the repository:

👉 **[Download the SPC_GUI User Manual (PDF)](Manual_v1.1.0.pdf)**

---

## 📦 Releases

You can find the latest installable executables for **Windows** and **macOS (M1/M2)** in the [Releases section](https://github.com/AEvgeneia/SPC_GUI_Scientific_Tool/releases) of this repository.

---

## 📂 Input File Format

Your input file must follow a specific format with required columns:
- ID
- Site of cancer
- QA Date
- One or more GPR columns (e.g., `Global 3%2mm`, `Global 2%2mm`, `Local 3%3mm`, etc.)

If you want to perform SPC on the **mean gamma (γ)** value, the column `MedianDoseDev` is also required.

A sample Excel template is available in the `example_data/` folder:
`input_example/template.xlsx`

---

## 📦 Dependencies

All dependencies are listed in requirements.txt. Main ones include:
- pandas
- numpy
- scipy
- matplotlib
- seaborn
- ttkthemes
- openpyxl

---

## 📤 Output

- **On-screen display**:
  - Interactive GUI showing loaded data, statistical summaries, histograms, normality test results, and SPC I-Charts.
  - Options to select metrics, run different SPC methods, and eliminate outliers.
  - All plots are displayed live within the GUI.

- **Exports**:
  - CSV files with summaries of SPC statistics, normality test results, and outlier elimination logs.
  - PDF files containing all selected plots (SPC I-charts and histograms).

---

## 🪪 License

This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute this software with attribution.

---

## 🤝 Contributing

If you find a bug, have a suggestion, or want to improve this tool, feel free to open an issue or submit a pull request.  
All contributions are welcome and appreciated!

---

## 🧠 Acknowledgements
This tool was developed as part of ongoing research at the Laboratory of Medical Physics, Department of Medicine, National and Kapodistrian University of Athens.

---

## 📧 Contact

If you use this app in your research or have questions, feel free to reach out via aspaeug[at]med.uoa.gr or cite the repository using the provided [CITATION.cff](CITATION.cff) file.

