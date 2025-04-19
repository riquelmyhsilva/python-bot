
# 🛠️ Python Bot – Product Registration Automation

This project automates the process of registering products on a web form using data from an Excel file. The automation is done with **PyAutoGUI**, which simulates keyboard and mouse actions to fill out the form fields.

🔗 **Live Website for Automation**: [cadastro-produtos-devaprender.netlify.app](https://cadastro-produtos-devaprender.netlify.app/)

📦 **Excel File Used**: `products.xlsx`

---

## 💻 How It Works

The script performs the following steps:

- Loads product data from an Excel spreadsheet (`products.xlsx`) using `openpyxl`.
- Iterates through each row and uses `PyAutoGUI` and `pyperclip` to:
  - Click on form fields.
  - Paste data (to avoid keyboard layout issues).
- Selects product size using conditional logic (e.g., dropdowns for "Small", "Medium", "Large").
- Includes logging to monitor script activity and help with debugging.
- Allows stopping the script anytime using `ESC` or `CTRL+C`.

---

## 🛑 How to Stop the Script

- Press `ESC` during execution to exit safely.
- Use `CTRL+C` in the terminal to interrupt the script.

---

## 📌 Notes

- You may need to adjust screen coordinates based on your screen resolution and browser zoom level.
- This script is tailored to work with the layout of the website mentioned above.
- The same approach can be adapted to other systems or web forms.

---

## 📄 Files Included

- `app.py`: Main automation script
- `products.xlsx`: Sample product data (mocked)

