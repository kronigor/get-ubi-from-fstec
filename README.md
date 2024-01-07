# FSTEC Threat List (BDU) Convertor

## Description

This Python script is designed to automate the process of fetching, processing, and documenting cybersecurity threats from the FSTEC database. It aims to simplify the task of monitoring and analyzing security threats.

## Key Features

- **Data Fetching**: Downloads the latest threat list in Excel format from the FSTEC website if it's not available locally.
- **Data Processing**: Parses the downloaded Excel file to extract relevant information using pandas.
- **Document Generation**: Converts the processed data into a well-formatted Word document. It includes changing document orientation and formatting text.

## Installation
Ensure you have Python version 3.x installed. To install necessary dependencies, use the provided `requirements.txt`:
```
pip install -r requirements.txt
```

## Usage

To run the script with Python, use the following command:
```
python get_ubi_from_fstec.py
```

## License
This project is licensed under the GNU General Public License (GPL).
