# DORA File Renamer

This PowerShell script provides a graphical user interface (GUI) for generating and renaming files based on a custom naming convention. The UI lets you:

- Enter a file code identifier.
- Select a module from a predefined list.
- Specify a period in the `yyyyMM` format.
- Choose a file and automatically extract a timestamp from its last modification time.
- Select a file extension.
- Generate a new file name and rename the original file.
- Copy the generated file name to the clipboard.

## Features

- **GUI**: Built with WPF and XAML.
- **File Selection**: Uses Windows Forms to open a file browser.
- **Timestamp Extraction**: Generates a timestamp based on file modification.
- **Validation**: Ensures required fields are filled and period format is valid.
- **Clipboard**: Copies the new file name for easy pasting.

## Requirements

- PowerShell 5.0 or later.
- .NET Framework supporting WPF and Windows Forms.

## Usage

1. Open PowerShell and navigate to the script directory:
    ```powershell
    cd 'c:\Users\jomar\Downloads\Dora Naming'
    ```

2. Execute the script:
    ```powershell
    .\naming.ps1
    ```

3. Follow the prompts in the generated GUI window to rename your file.

## License

This project is licensed under the MIT License.

## Acknowledgments

This script leverages .NET's PresentationFramework for the GUI and Windows Forms for the file dialog.