# Data_Extraction_Tool
Extract data you desired from selected files and combine into a new sheet

## Business Problem
As a former Risk Consultant, I frequently faced the challenge of consolidating data from multiple Excel files. This manual process took 2-3 days each month and was prone to errors.

## Solution
This VBA tool automates the extraction of specific KPIs from multiple Excel files into a standardized format, reducing processing time from days to minutes.

## Key Features
- **Configurable Mapping**: Define which KPIs to extract via simple Excel table
- **Batch Processing**: Processes entire folders of Excel files automatically  
- **Error Handling**: Identifies missing data or invalid cell references
- **Audit Trail**: Tracks source files and extraction results

## Business Impact
- **Time Savings**: Reduced monthly process from 3 days to 30 minutes
- **Accuracy**: Eliminated manual copy-paste errors
- **Scalability**: Easily adapts to new KPIs or file structures

## 🛠 Installation & Usage

### For Excel Users:
1. Open Excel and press `Alt + F11` to open VBA Editor
2. Insert a new Module
3. Copy and paste the code from `src/KPI_Extraction_Tool.vba`
4. Create a "KPI_Mapping" sheet with your KPI configurations
5. Run the `ExtractByCellMapping` macro

### KPI Mapping Format:
Create an Excel sheet named "KPI_Mapping":
```csv
KPI_Name,Cell_Address
Revenue,B5
Profit,C8
Growth_Rate,F12
```
