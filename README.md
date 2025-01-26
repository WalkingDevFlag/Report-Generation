# Report Generation

## Overview
This project is designed to automate the generation of reports using Python. It leverages various libraries to create dynamic documents that can be customized based on user input or data sources. The project includes three main folders: `src`, `car-damage-report`, and `Better Template`, each serving a specific purpose in report generation.

## Features
- **Dynamic Report Generation**: Generate reports by replacing placeholders with actual data.
- **Car Damage Reporting**: Create reports specifically tailored for car damage assessments, adapting to various dent conditions.
- **Invoice Generation**: Utilize a more sophisticated template for generating invoices that include tabular data.
- **User-Friendly Setup**: Easy environment setup using Conda.

## Requirements
Before running the project, ensure you have the following Python packages installed:

- `python-docx`: For creating and editing Word documents.
- `pandas`: For data manipulation and analysis.
- `faker`: For generating fake data for testing purposes.
- `docx`: For handling .docx files.
- `docx2pdf`: For converting .docx files to PDF format.

## Conda Environment Setup
To set up the Conda environment for this project, follow these steps:

1. Create a new Conda environment:
   ```bash
   conda create -n report python=3.10
   ```

2. Activate the newly created environment:
   ```bash
   conda activate report
   ```

3. Install the required packages:
   ```bash
   conda install python-docx pandas faker docx docx2pdf
   ```

## Usage
After setting up your environment and installing the necessary packages, you can start using the project. 

1. Navigate to the desired folder based on your needs:
   - **`src`**: Use this folder for basic report generation by replacing placeholders in templates.
   - **`car-damage-report`**: This folder contains scripts specifically designed to generate reports based on car damage assessments that dynamically change according to the number and type of dents.
   - **`Better Template`**: This folder includes templates and scripts for generating invoices with tabular data.

2. Run the appropriate script in your terminal or command prompt:
   ```bash
   python main.py
   ```

## License
This project is licensed under the MIT License.
