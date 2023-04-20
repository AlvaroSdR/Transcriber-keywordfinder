# Keyword Analysis Tool

This Python script is designed to analyze keywords from a text file and generate an Excel file with summary and detailed information.

## Description

The keyword analysis tool reads a text file containing keywords, groups, and sets, and performs the following tasks:

1. Identifies the 10 most frequently used keywords.
2. Groups the keywords based on the provided groups and sets.
3. Generates a summary sheet in the Excel file with the keyword count, group, and set.
4. Generates additional sheets for "Pains" and "Benefits" that show sentences containing the keywords from the respective sets, along with the corresponding group and keyword.

## Dependencies

The script requires the following dependencies:

- Python 3.x
- openpyxl (Python library for working with Excel files)

You can install the openpyxl library using the following command:

```bash
pip install openpyxl
```

## Usage
Place the text file containing keywords in the same folder as the script.
Update the input_file variable in the script with the filename of your input text file.
Run the script using Python 3.x.
The script will generate an Excel file named "outputs.xlsx" in the same folder as the input file.
Open "outputs.xlsx" to view the summary and detailed information in the different sheets.
Note: Please be aware that you need LinkedIn Develper API keys for the correct usage of the application.

## Notes
The script assumes that the input text file contains keywords, groups, and sets in the specified format (one entry per line, separated by tabs).
The script uses the first sheet for the summary and creates additional sheets for "Pains" and "Benefits" information in the Excel file.
The script auto-sizes the columns in the generated Excel file for better readability.

## Contributing
If you would like to contribute to this project, please contact in advance. Any contributions are welcome!

## License
This project is licensed under the My License. In case of use, please contact in advance.

## Contact
If you have any questions or suggestions, feel free to contact me at asanchezdrio@gmail.com.

Thanks for using the keyword finder!
