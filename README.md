
# 🇧🇩 VBA Number to Bangladeshi Taka and Poisha Converter

This repository contains a custom Excel VBA function (`NumberToWords`) that converts any numeric value into words using the **Bangladeshi currency system**, i.e., Taka and Poisha.  
Supports large numbers with proper formatting like **Lac** and **Crore**.

---

## 🔁 Example Output

| Input     | Output                                                       |
|-----------|--------------------------------------------------------------|
| `1234.56` | One Thousand Two Hundred Thirty-Four Taka and Fifty-Six Poisha Only |
| `7500000` | Seventy-Five Lac Taka Only                                   |
| `10000000.75` | One Crore Taka and Seventy-Five Poisha Only                 |

---

## 📥 How to Use

### In Excel:

1. Press `ALT + F11` to open the VBA editor.
2. Insert a new Module: `Insert > Module`.
3. Paste the code from `NumberToWords.bas` file.
4. Save and return to Excel.
5. Use the function in any cell:
   ```excel
   =NumberToWords(123456.78)
