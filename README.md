# DataDispatcher

## Overview

DataDispatcher is a Python program designed to facilitate efficient data transfer and management across Windows applications to improve workflow. It provides methods to transfer files between directories, list files in a directory, and perform simple automation tasks with Microsoft Excel.

## Features

- **File Transfer**: Copy files from a source directory to a target directory.
- **Directory Listing**: List all files present in a specified directory.
- **Excel Automation**: Automate simple tasks in Microsoft Excel such as reading cell values.

## Requirements

- Python 3.x
- `pywin32` library for Windows COM object interaction.
- Logging support for tracing and debugging purposes.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/DataDispatcher.git
   ```
2. Navigate to the project directory:
   ```bash
   cd DataDispatcher
   ```
3. Install the required Python libraries:
   ```bash
   pip install pywin32
   ```

## Usage

1. Set up the source and target directories in the Python script:
   ```python
   dispatcher = DataDispatcher(source_folder='C:/source', target_folder='C:/target')
   ```

2. Run the script:
   ```bash
   python data_dispatcher.py
   ```

3. Check the `data_dispatcher.log` file for logs related to file transfers and operations.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For questions or support, please contact [your-email@example.com](mailto:your-email@example.com).