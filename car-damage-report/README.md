# Dynamic Damage Report Generator

## Overview

The Dynamic Damage Report Generator is a Python-based tool designed to create customized damage reports from a Word template and a CSV dataset. It automates the process of analyzing vehicle damage, including classification, area measurement, and reporting in a structured format.

This tool processes damage data, groups it by `Class ID`, and dynamically populates a Word document template with detailed breakdowns of damage types, instance counts, and mask areas.

## Features

- **Dynamic Content Generation**: Automatically populates a Word template with data from a CSV file.
- **Grouped Data Presentation**: Groups damage data by `Class ID` for better organization.
- **Custom Formatting**: Applies bold fonts, bullet points, and specific font sizes for a professional look.
- **Error Handling**: Ensures the required columns are present in the CSV file and validates the template structure.
- **Reusable Template**: Works with any Word template that includes predefined placeholders for dynamic content.

## Requirements

- **Python**: Version 3.10 or higher
- **Dependencies**: 
  - `python-docx`
  - `pandas`



## Conda Environment Setup

To set up the environment and install the required dependencies:

1. **Create the environment**:
    ```bash
    conda create -n report python=3.10
    ```
2. **Activate the environment**:
    ```bash
    conda activate report
    ```
3. **Install the required packages**:
    ```bash
    pip install python-docx pandas
    ```

## Usage

1. **Prepare the Input Files**:
   - CSV File: Ensure your data file has the following columns:
     - `Class ID`
     - `Damage Type`
     - `Area`
   - Word Template: Use a `.docx` template with placeholders for `Damage Breakdown:` and content structure.

2. **Update Paths in the Code**:
   - Replace the paths for `csv_path`, `template_path`, and `output_path` in the `if __name__ == '__main__'` section with your file locations.

3. **Run the Script**:
    ```bash
    python script_name.py
    ```
   Replace `script_name.py` with the name of your Python file.

4. **Output**:
   - The generated Word report will be saved at the specified `output_path`.


## Example

### Input CSV
| Class ID | Damage Type | Area      |
|----------|-------------|-----------|
| 1        | Dent        | 150.50    |
| 1        | Dent        | 200.75    |
| 2        | Scratch     | 80.00     |

### Word Template
```
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.

Damage Breakdown:
• Class ID [Class ID]: [Damage Type]
• Number of Instances: [Number of Instances]
• Mask Areas:
• [Damage Type] [Instance Number]: [Area in sq. mm]
```

### Generated Report
```
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.

Damage Breakdown:
• Class ID 1: Dent
  Number of Instances: 2
  Mask Areas:
    Dent 1: 150.5 sq. mm
    Dent 2: 200.75 sq. mm

• Class ID 2: Scratch
  Number of Instances: 1
  Mask Areas:
    Scratch 1: 80.0 sq. mm
```
## License

This project is licensed under the MIT License.
