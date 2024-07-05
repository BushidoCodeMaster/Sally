# Sally - Simple Auto Illustrator

https://github.com/BushidoCodeMaster/Sally/tree/main/dist/SimpleAutoillustrator

## Overview
Sally automates specific tasks with Excel files and integrates them with Adobe Illustrator. This tool enables users to insert a new column, generate QR codes with customizable data, and create executable JSX scripts for Adobe Illustrator.
Sally is written in Python primary. Including some acknowledgments of the Java Script and jsx

## Features
- **Excel Manipulation:** Automatically inserts an "id" column at the beginning of Excel files without overwriting the original.
- **QR Code Generation:** Enables the creation of QR codes encoded with either URLs or plain text, selectable via a user-friendly GUI.
- **Adobe Illustrator Integration:** Automates task execution within Adobe Illustrator by generating and executing JSX scripts.

## Usage
1. **Prepare Your Files:** Ensure your Excel files are set up for processing. Sample files can be found in the `examples` folder.
2. **Running Sally:** Utilize the GUI to navigate and set preferences.
3. **Generate QR Codes:** Choose data types for encryption and generate QR codes directly.
4. **Run Adobe illustrator** launching the app before pressing final "Execute!" button, will incrise the overall productivity and efficienty of the executed scripts.
5. **Execute in Adobe Illustrator:** Generate and execute JSX scripts in Adobe Illustrator to automate graphic tasks by pressing "Execute!".

### Detailed Steps:
- **Excel Setup:**
  - Navigate to "QR Code Generator" tab.
  - Load your Excel file.
  - Automatically check and insert IDs if necessary.
  - Reload your new Excel file (*_ids.xlsx).
  - Select columns whose data will be encrypted into QR codes (id is necesarry).
  - Specify the output folder for generated QR codes.
  - Click "Generate QArrr Codes!" to execute QR generation.
  
- **Adobe Illustrator Scripting:**
  - Load Adobe Illustrator.
  - Choose your template and Excel files.
  - Specify columns for data import into Illustrator.
  - Execute the script to process and output files in the desired format (PDF or AI).

## Requirements
- A licensed version of **Adobe Illustrator** is recommended for optimal functionality.

## Disclaimer
Sally creates new files rather than overwriting existing ones. We are not responsible for data loss or file damage.

## Future Plans
- **MacOS Support:** To make Sally available to MacOS users.
- **Custom Folder Structure:** Users can create folder structures based on Excel data.
- **Flexible JSX Scripts:** More dynamic script generation based on the Excel cell content, supporting up to five custom rules and different ai templates usages based on such content. (Student Score, Winners Place and etc.)

## License
This project is licensed under the GNU GPLv3 License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments
Created by Timur Mustafin, aka BushidoCoder - © 2024.



---
