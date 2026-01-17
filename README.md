# Water Data Transformer

A C# project for transforming and processing water data from Excel files using Microsoft Excel Interop.

---

## 🔹 Overview

This project automates the extraction, calculation, and transformation of water-related data (daily, monthly, yearly, and hydrological year) from source Excel files into a structured format suitable for analysis or reporting.

It is designed for **hydrology engineers, data analysts, and environmental researchers**.

---

## 🔹 Features

- Reads source Excel files containing raw water data.
- Performs **daily, monthly, yearly, and hydrological year calculations**.
- Outputs a **transformed Excel file** ready for further analysis.
- Modular C# design with separate methods for each calculation type.
- Handles missing data safely with default values.

---

## 🔹 Technologies Used

- **C# (.NET Framework)**
- **Microsoft Excel Interop**
- **Visual Studio (recommended)**

---

## 🔹 How It Works

1. The program opens the **source Excel file**.
2. It reads each row of data, performing calculations:
   - **Daily values (`dayVal`)**
   - **Monthly values (`monthVal`)**
   - **Yearly values (`yearVal`)**
   - **Hydrological year values (`hydroVal`)**
3. The transformed data is written to a **new Excel file** with the same structure as a template.
4. The file is saved with `_Transformed` suffix in the same folder as the source.

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
