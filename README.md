\# Excel Master



A local AI-powered Excel agent that understands, analyzes, modifies, and validates Excel workbooks through natural language.



\## Vision



Upload an Excel workbook and interact with it using prompts.



Examples:



\- Explain this workbook.

\- Clean the dataset.

\- Find missing or duplicate records.

\- Add useful features.

\- Create formulas.

\- Build charts and dashboards.

\- Reformat a worksheet.

\- Validate formulas and references.



\## Architecture



User Prompt

↓

Hermes Agent

↓

Local LLM / Ollama

↓

Excel Tools

↓

pandas + openpyxl

↓

Validation

↓

Output Workbook



\## Planned Tools



\- Workbook Inspector

\- Data Cleaner

\- Feature Engineer

\- Formula Engine

\- Chart Builder

\- Formatting Engine

\- Workbook Validator



\## v0.1 Goal



Excel file → Inspect → Understand → Explain

