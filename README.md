# tk\_data\_cleaner

**Python Tkinter GUI tool for cleaning and preprocessing Excel data, exporting to TXT.**

---

## Overview

`tk_data_cleaner` is a user-friendly Python application designed to clean, preprocess, and export Excel data to TXT format. The tool is suitable for data analysts, developers, and professionals who need to quickly prepare structured data for analysis or reporting.

Key features:

- **Import Excel files** (.xls, .xlsx)
- **Interactive GUI** built with Tkinter (`app.py`)
- **Clean data** by removing extra spaces and duplicates (`logic.py`)
- **Filter rows** based on minimum character length
- **Export clean data** to TXT format
- **Dynamic tooltips** for real-time feedback (`tooltip.py`)

---

## Project Structure

```
tk_data_cleaner/
├── app.py             # Main application script (GUI)
├── logic.py           # Data cleaning and preprocessing logic
├── tooltip.py         # Tooltip utility for GUI components
├── requirements.txt   # Python dependencies
└── README.md          # Project documentation
```

---

## Installation

1. Make sure Python 3.x is installed: [Download Python](https://www.python.org/downloads/)

2. Clone the repository:

```bash
git clone https://github.com/lcspires/tk_data_cleaner.git
cd tk_data_cleaner
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

> Tkinter is included with standard Python installations.

---

## Usage

1. Run the GUI:

```bash
python app.py
```

2. Steps in the application:

- Click **"Select File"** to load an Excel spreadsheet.
- Organize columns using move up/down or remove unwanted columns.
- Set the **minimum number of characters** for the first column.
- Click **"Generate TXT"** to export the cleaned dataset.

3. Tooltips provide real-time feedback:

- Spaces removed
- Duplicates removed
- Rows removed due to length

---

## Example

Before cleaning:

| Name | Email                                    | Phone |
| ---- | ---------------------------------------- | ----- |
| John | [john@email.com](mailto\:john@email.com) | 12345 |
| Mary | [mary@email.com](mailto\:mary@email.com) | 67890 |
| John | [john@email.com](mailto\:john@email.com) | 12345 |

After cleaning (minimum 4 chars, duplicates removed, spaces trimmed):

| Name | Email                                    | Phone |
| ---- | ---------------------------------------- | ----- |
| John | [john@email.com](mailto\:john@email.com) | 12345 |
| Mary | [mary@email.com](mailto\:mary@email.com) | 67890 |

---

## Contributing

Contributions are welcome! You can:

- Report bugs
- Suggest new features
- Submit pull requests

Please follow standard GitHub workflow with feature branches.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Contact

For questions or collaboration:

- GitHub: [lcspires](https://github.com/lcspires)
- Email: [ferreira.l@ufba.br](mailto\:ferreira.l@ufba.br)

