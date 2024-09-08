# Excel File Creator

This Python program allows you to quickly create multiple Excel files with customizable names and locations. It uses the `openpyxl` library to generate Excel files and provides options for directory and file naming and placement.

## Features

- **Custom File Names**: Choose a custom name or let the program generate a random name for your Excel files.
- **Custom Location**: Specify a directory to save the files, or use the current working directory.
- **Create Directory**: Option to create specified directory if it does not exist.
- **Custom Separators**: Add a separator between the file name and the iteration number.
- **Predefined Iteration**: Start the file naming from a specific iteration number.
- **Validation**: Ensures that forbidden characters are not used in file names and separators.

## Requirements

- Python 3.x
- `openpyxl` library

You can install `openpyxl` using pip:

```
pip install openpyxl
```

## How to Use

1. **Run the Program**: Execute the Python script.

2. **Provide Inputs**:
   - **Specify Directory**: Choose specific directory or accept the default placement.
   - **Number of Files**: Enter the number of Excel files you want to create.
   - **Custom File Name**: Decide whether to use a custom file name or a random one.
   - **Starting Iteration**: Specify a starting iteration number, or accept the default value `1`.
   - **Separator**: Decide whether to add a separator between the file name and the iteration number.
   - **Confirmation**: Review the file names and directory, then confirm to create the files.

3. **File Creation**: The program will create the specified number of Excel files with the chosen naming and placement options.

## Example

To create 5 Excel files with a custom name, starting from iteration 3, and with a separator "_":

1. Run the script.
2. Choose file placement or accept the default.
3. Enter the number of files: `5`
4. Choose custom file name: `yes`
5. Enter the file name: `Example`
6. Choose predefined starting iteration: `yes`
7. Enter starting iteration: `3`
8. Add separator: `yes`
9. Enter separator: `_`
10. Confirm the details and the files will be created.

## Important Notes

- Avoid using forbidden characters in file names and separators: `/`, `\`, `.`, `>`, `<`, `*`, `?`, `!`, `|`, ` "`, `:`, `\0`.
- Ensure that the specified directory exists or is writable.
   - *Check now added for nonexistent directories. If directory path is not found it can be created.*
