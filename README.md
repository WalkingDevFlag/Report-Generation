# Report Generation

## Overview
This project is designed to generate personalized Word documents based on a predefined template and data from a CSV file. It uses the `python-docx` library to modify the Word document content dynamically and `pandas` to handle the CSV data. The generated documents are customized invitations for the "Annual Innovation Summit 2024."

## Features
- Generate personalized Word documents for multiple recipients.
- Replace placeholders in a Word template with values from a CSV file.
- Easily customizable template and input data.
- Automates batch document creation for efficient workflow.

## Requirements
The project requires the following libraries:
- `python-docx`
- `pandas`

Additionally, it requires Python 3.10 or higher.

## Conda Environment Setup
To set up the environment and install the required libraries, follow these steps:
1. Create a new Conda environment:
   ```bash
   conda create -n report python=3.10
   ```
2. Activate the Conda environment:
   ```bash
   conda activate report
   ```
3. Install the necessary Python packages:
   ```bash
   pip install python-docx pandas
   ```

## Usage
1. **Prepare the Input Data:**
   - Create a CSV file (`contacts.csv`) with the following columns:
     - `Salutations`
     - `First Name`
     - `Last Name`
     - `Last Contacted`
     - `Company Name`

2. **Create a Word Template:**
   - Design a Word document (`template2.docx`) with placeholders enclosed in square brackets (e.g., `[Salutation]`, `[First Name]`, etc.).

3. **Run the Script:**
   - Execute the Python script to generate personalized documents:
     ```bash
     python script.py
     ```
   - The script will generate one Word document per row in the CSV file, naming the output files sequentially (e.g., `report_1.docx`, `report_2.docx`, etc.).

4. **Verify Output:**
   - Check the output directory for the generated Word documents.

## Example
**Input CSV (`contacts.csv`):**
```csv
Salutations,First Name,Last Name,Last Contacted,Company Name
Mr.,John,Doe,2024-12-01,Acme Corp
Ms.,Jane,Smith,2024-11-20,Innovate LLC
```

**Template Word Document (`template2.docx`):**
```text
Annual Innovation Summit 2024
Dear [Salutation] [First Name] [Last Name],
We hope this message finds you well. It’s been wonderful collaborating with [Company Name] since [Last Contacted], and we are excited to continue our journey together.
```

**Generated Document (`report_1.docx`):**
```text
Annual Innovation Summit 2024
Dear Mr. John Doe,
We hope this message finds you well. It’s been wonderful collaborating with Acme Corp since 2024-12-01, and we are excited to continue our journey together.
```

## License
This project is licensed under the MIT License.
