# Water Data Transformer

A **C# project** that automates the transformation of water data from Excel files into a structured format for analysis and reporting.

---

## 🔹 Overview

This tool reads **raw water data** from Excel files, performs daily, monthly, yearly, and hydrological calculations, and outputs a clean, structured Excel file.  

It is ideal for **hydrology engineers, data analysts, and environmental researchers**.

---

## 🔹 Features

- **Automatic Excel processing** using C# and Microsoft Excel Interop.
- Calculates:
  - Daily values
  - Monthly aggregation
  - Yearly totals
  - Hydrological year values
- Saves the transformed data into a **ready-to-use Excel file**.
- Handles missing data safely.
- Uses a **template Excel file** to maintain structure.

---

## 🔹 Technology Stack

- **C# (.NET Framework)**
- **Microsoft Excel Interop**
- **Windows Forms (for GUI)**
- **Visual Studio**

---

## 🔹 How It Works

1. **Open Source File:** The program asks you to select your raw Excel file.  
2. **Read & Transform Data:**  
   - Daily, monthly, yearly, and hydrological year calculations are performed automatically.  
   - Missing data is handled gracefully.  
3. **Write to Template:** Data is written into a template Excel file with pre-defined sheets for days, months, years, and hydrological years.  
4. **Save Output:** The transformed Excel file is saved in the same folder as the source with a `_Transformed` suffix.

---

## 🔹 Example Output

### 1️⃣ Daily Data (`Days` sheet)
| Date       | Value (m³) |
|------------|------------|
| 2026-01-01 | 120        |
| 2026-01-02 | 115        |

### 2️⃣ Monthly Data (`Months` sheet)
| Month-Year | Total (m³) |
|------------|------------|
| Jan-2026   | 3500       |
| Feb-2026   | 3620       |

### 3️⃣ Yearly Data (`Years` sheet)
| Year | Total (m³) |
|------|------------|
| 2026 | 43000      |
| 2027 | 45000      |

### 4️⃣ Hydrological Year (`Hydro Years` sheet)
| Year-Range | Total (m³) |
|------------|------------|
| 2025-2026  | 44000      |

---

## 🔹 Screenshots

![Open File Dialog](./images/openfile.png)
*Select the Excel file to transform*

![Transformation Progress](./images/progress.png)
*Shows progress of daily/monthly/yearly calculations*

![Result Excel Sheet](./images/result.png)
*Output template with structured data*

> **Tip:** Screenshots are optional but make your repository look professional.

---

## 🔹 Example Usage

```csharp
using System;

namespace WaterDataTransformerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourcePath = @"C:\WaterData\datapass.xls";
            string templatePath = @"C:\WaterData\template.xlsx";

            ExcelTransformer transformer = new ExcelTransformer(sourcePath, templatePath);
            transformer.Transform();

            Console.WriteLine("Transformation completed!");
        }
    }
}
