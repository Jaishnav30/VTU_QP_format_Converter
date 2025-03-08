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
│-- 📄 input_excel_file.xlsx                   # Input file (raw question data)
│-- 📄 output_qp_format.xlsx                   # Output file (formatted question paper)
│-- 🖼️ rnsit_logo.jpeg                         # College logo
│-- 🖼️ usn.jpeg                                # USN header image
│-- 📄 Run_this_file.py                        # Main script for formatting Excel
│-- 📄 excel_to_pdf.py                         # Script for converting Excel to PDF
│-- 📄 README.md                               # Project documentation
```

## 📖 How It Works  
1️⃣ Load raw data from **input_excel_file.xlsx**  
2️⃣ Extract necessary details like semester, subject, exam date, etc.  
3️⃣ Reformat the content into a structured format  
4️⃣ Insert institutional details and images  
5️⃣ Apply proper styling and formatting  
6️⃣ Save the output as **output_qp_format.xlsx**  
7️⃣ Convert the formatted Excel file to PDF  

## 🚀 Getting Started
### 1️⃣ Clone the Repository
To get a copy of this project on your local machine, run the following command:

```
git clone https://github.com/your-username/Question-Paper-Formatter.git
cd Question-Paper-Formatter
```

### 2️⃣ Install Dependencies
Make sure you have Python installed, then install the required dependencies:

```
pip install pandas openpyxl
```

### 3️⃣ Place the raw Excel file (input_excel_file.xlsx) in the project folder.

### 4️⃣ Run the script:

```
python Run_this_file.py
```

The formatted Excel file (output_qp_format.xlsx) and the PDF version will be generated automatically.

---

## 🚧 Limitations
1️⃣ The USN (University Seat Number) in the top-right corner is a static .jpeg image, fixed as 1RN__**CS**___, and does not dynamically update based on the input branch.
2️⃣ Certain sections of the code can be optimized to reduce redundancy and improve efficiency.

---

## 🚀 Potential Improvements
✅ Convert the USN image into selectable text, allowing it to dynamically change based on the input branch.
✅ Refactor and optimize the code to enhance readability, modularity, and maintainability.
✅ Automate column width adjustments and row insertions to eliminate hardcoded values.
✅ Improve image handling by dynamically loading images from a specified directory instead of using static paths.

---

💡 If you have any suggestions for further improvements, feel free to contribute! 🤝

## 📌 Example Output

After execution, the script will generate a professionally formatted question paper with proper headers, images, and alignment.

## 📝 Contributions

Feel free to fork this repository and enhance the formatting options or improve the PDF export functionality!

## 🔗 Connect with Me

For any suggestions, reach out via GitHub Issues.

## ⭐ If you find this project useful, don't forget to star the repository! ⭐
