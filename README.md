# 🎯 Question Paper Formatter  

## 📌 Overview  
This project automates the process of formatting a raw Excel file into a structured question paper format with well-defined styling, institution details, images, and proper alignment. It also supports conversion of the formatted Excel file into PDF format for easy sharing and printing.  

---

## 🚀 Features  
✅ Reads raw data from an input Excel file  
✅ Rearranges the data into a structured format  
✅ Adds institutional headers (College name, address, accreditation details, etc.)  
✅ Inserts images (College logo, USN header)  
✅ Formats text (Bold, font changes, text wrapping, alignments)  
✅ Applies borders and background colors  
✅ Adds special formatting (Merged cells, question separators, RBT classification)  
✅ Exports the final output as an Excel file and converts it to a PDF  

---

## 🛠️ Technologies Used  
- **Python** 🐍  
- `pandas` (For reading and processing Excel data)  
- `openpyxl` (For modifying Excel sheets)  
- `os` (For handling file paths)  
- `excel_to_pdf.py` (For converting Excel to PDF)  

---

## 📂 Project Structure  
```
📂 Question-Paper-Formatter
│-- 📄 input_excel_file.xlsx # Input file (raw question data)
│-- 📄 output_qp_format.xlsx # Output file (formatted question paper)
│-- 🖼️ rnsit_logo.jpeg # College logo
│-- 🖼️ usn.jpeg # USN header image
│-- 📄 excel_to_pdf.py # Script for converting Excel to PDF
│-- 📄 main.py # Main script for formatting Excel
│-- 📄 README.md # Project documentation
```

## 📖 How It Works  
1️⃣ Load raw data from **input_excel_file.xlsx**  
2️⃣ Extract necessary details like semester, subject, exam date, etc.  
3️⃣ Reformat the content into a structured format  
4️⃣ Insert institutional details and images  
5️⃣ Apply proper styling and formatting  
6️⃣ Save the output as **output_qp_format.xlsx**  
7️⃣ Convert the formatted Excel file to PDF  

## 🏃‍♂️ How to Run the Project  

### 1️⃣ Install dependencies:  
```
pip install pandas openpyxl
```

### 2️⃣ Place the raw Excel file (input_excel_file.xlsx) in the project folder.

### 3️⃣ Run the script:

```
python Run_this_file.py
```

The formatted Excel file (output_qp_format.xlsx) and the PDF version will be generated automatically.
