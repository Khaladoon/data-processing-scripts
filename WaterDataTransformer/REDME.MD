# Water Data Transformer

![.NET](https://img.shields.io/badge/.NET-Framework-blue)
![License](https://img.shields.io/badge/License-MIT-green)

---

## 🔹 Overview

**Water Data Transformer** is a C# project for transforming and processing water data from Excel files using Microsoft Excel Interop.

It automates the extraction, calculation, and transformation of water-related data (daily, monthly, yearly, hydrological year) from source Excel files into a structured format suitable for analysis or reporting.

Ideal for **hydrology engineers, data analysts, and environmental researchers**.

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

1. Open the **source Excel file**.
2. Read each row of data and perform calculations:
   - **Daily values**
   - **Monthly values**
   - **Yearly values**
   - **Hydrological year values**
3. Write the transformed data to a **new Excel file** with the same template structure.
4. Save the file with a `_Transformed` suffix in the same folder as the source.

---

## 🔹 Screenshots

**Open File Dialog**  
![Open File Dialog](./images/openfile.png)

**Transformation Progress**  
![Transformation Progress](./images/progress.png)

**Resulting Excel Sheet**  
![Result Excel Sheet](./images/result.png)

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
            string sourcePath = @"./data/datapass.xls";
            string templatePath = @"./data/template.xlsx";

            ExcelTransformer transformer = new ExcelTransformer(sourcePath, templatePath);
            transformer.Transform();

            Console.WriteLine("Transformation completed!");
        }
    }
}
