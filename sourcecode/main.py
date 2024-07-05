import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Canvas, PhotoImage
import pandas as pd
import qrcode
from qrcode.image.svg import SvgPathImage
import os, sys
import subprocess

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('SimpleAutoillustrator')
        self.geometry('1000x600')

        # Determine if the application is running as a bundled executable
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        icon_path = os.path.join(application_path, 'sai.ico')
        welcome_image_path = os.path.join(application_path, 'welcome_transparent.png')

        # Set the window icon
        self.iconbitmap(icon_path)

        self.project_directory = os.getcwd()  # Default directory for file dialogs
        self.illustrator_path = "C:/Program Files/Adobe/Adobe Illustrator 2024/Support Files/Contents/Windows/Illustrator.exe"
        self.template_path = None
        self.output_base_path = None
        self.excel_path = None

        self.file_type = tk.StringVar(value="PDF")
        self.os_type = tk.StringVar(value="Win")
        self.column_vars = {}
        self.filename_vars = {}
        self.qr_code_enabled = tk.StringVar(value="no")
        self.qr_column_vars = {}

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=1, fill="both")

        self.init_os_tab(welcome_image_path)
        self.init_qr_code_tab()
        self.init_badges_tab()

    def init_os_tab(self, welcome_image_path):
        os_tab = ttk.Frame(self.notebook)
        self.notebook.add(os_tab, text='Choose OS')

         # Welcome text
        ttk.Label(os_tab, text="Welcome to SimpleAutoillustrator! To begin Chose your OS :", font=("Arial", 12)).pack(side=tk.TOP, padx=20, pady=20)
        # Create a Canvas widget that matches the size of the os_tab
        canvas = Canvas(os_tab)
        canvas.pack(fill="both", expand=True)

        # Load the image with transparency already applied
        welcome_image = PhotoImage(file=welcome_image_path)
        
        # Resize and place the image as the window size changes
        def resize_image(event):
            # Resize the image to match the event width and height
            resized_image = welcome_image.zoom(max(1, event.width//welcome_image.width()), max(1, event.height//welcome_image.height()))
            resized_image = resized_image.subsample(max(1, resized_image.width()//event.width), max(1, resized_image.height()//event.height))
            
            # Create a new image in the canvas
            canvas.delete("all")  # Remove previous elements
            canvas.create_image(event.width // 2, event.height // 2, image=resized_image, anchor="center")
            canvas.image = resized_image  # Update the reference to prevent garbage collection

            # Re-create other elements like radio buttons and text
            canvas.create_window(200, 150, window=ttk.Radiobutton(os_tab, text="Windows", variable=self.os_type, value="Win"))
            canvas.create_window(300, 150, window=ttk.Radiobutton(os_tab, text="MacOS", variable=self.os_type, value="Mac"))
            canvas.create_text(500, 450, text="Created by Timur Mustafin aka BushidoCoder - © 2024", fill="black", font=("Arial", 10))

        # Bind the resize event of the os_tab to the resize_image function
        os_tab.bind("<Configure>", resize_image)

    def init_qr_code_tab(self):
        qr_tab = ttk.Frame(self.notebook)
        self.notebook.add(qr_tab, text='QR Code Generator')
        
        qr_main_frame = ttk.Frame(qr_tab)  # Central frame for all elements
        qr_main_frame.pack(fill='x', expand=True, pady=20)

        self.create_file_selector("Step1. Select Excel :", self.select_excel_file, qr_main_frame)
        ttk.Button(qr_main_frame, text="Step 2. Check/Create Excel IDs", command=self.handle_excel_ids).pack(fill='x', pady=10)

        ttk.Label(qr_main_frame, text="Step 3. Select data types for QR code encryption :").pack(fill='x')
        self.qr_data_frame = ttk.Frame(qr_main_frame)  # Frame for column checkboxes
        self.qr_data_frame.pack(fill='x', expand=True)

        ttk.Label(qr_main_frame, text="Step 4. Does your Excel have right ids?").pack(fill='x', pady=5)
        qr_code_control_frame = ttk.Frame(qr_main_frame)
        qr_code_control_frame.pack(fill='x', expand=True, pady=5, anchor='center')
        ttk.Radiobutton(qr_code_control_frame, text="No", variable=self.qr_code_enabled, value="no").pack(side=tk.LEFT, expand=True)
        ttk.Radiobutton(qr_code_control_frame, text="Yes", variable=self.qr_code_enabled, value="yes").pack(side=tk.LEFT, expand=True)

        self.create_file_selector("Step 5. Select Output Base Path", self.select_output_base_path, qr_main_frame)
        ttk.Button(qr_main_frame, text="Generate QR Codes! Arrr!", command=self.generate_qr_codes).pack(fill='x', pady=10)

    def init_badges_tab(self):
        badges_tab = ttk.Frame(self.notebook)
        self.notebook.add(badges_tab, text='Badge/Cert Generator')

        self.file_path_section = ttk.Frame(badges_tab)
        self.file_path_section.pack(fill='x', expand=True)
        self.create_file_selector("Step 1. Select Illustrator Path :", self.select_illustrator_path, self.file_path_section)
        self.create_file_selector("Step 2. Select *.ai Template :", self.select_template_file, self.file_path_section)
        self.create_file_selector("Step 3. Select Second *.ai  Template :", self.select_template_file, self.file_path_section)
        self.create_file_selector("Step 4. Select your *.xlsx Excel :", self.select_excel_file, self.file_path_section)

        ttk.Label(self.file_path_section, text="Step 5. Pick Column Names to be Processed :").pack(fill='x', pady=10)
        ttk.Label(self.file_path_section, text="Chose All Placeholder names matching your *.ai and 'ID' if you want to use QR Codes").pack(fill='x', pady=10)

        self.columns_frame = ttk.Frame(badges_tab)
        self.columns_frame.pack(fill='x', expand=True)

        self.output_data_logic_section = ttk.Frame(badges_tab)
        self.output_data_logic_section.pack(fill='x', expand=True)
        ttk.Label(self.output_data_logic_section, text="Step 6. Select the Excel columns to modify Output saving logic:").pack()
        self.filename_logic_frame = ttk.Frame(self.output_data_logic_section)
        self.filename_logic_frame.pack(fill='x', expand=True)

        self.qr_code_section = ttk.Frame(badges_tab)
        self.qr_code_section.pack(fill='x', expand=True)
        ttk.Label(self.qr_code_section, text="Step 7. Use QR code?").pack()
        qr_code_control_frame = ttk.Frame(self.qr_code_section)
        qr_code_control_frame.pack(fill='x', expand=True, pady=5, anchor='center')  # Center the frame
        ttk.Radiobutton(qr_code_control_frame, text="No", variable=self.qr_code_enabled, value="no").pack(side=tk.LEFT, expand=True)
        ttk.Radiobutton(qr_code_control_frame, text="Yes", variable=self.qr_code_enabled, value="yes").pack(side=tk.LEFT, expand=True)

        self.create_file_selector("Step 8. Select Output Base Path", self.select_output_base_path, badges_tab)

        self.output_type_section = ttk.Frame(badges_tab)
        self.output_type_section.pack(fill='x', expand=True)
        ttk.Label(self.output_type_section, text="Step 9. Select Output File Type:").pack(pady=5)
        self.file_type_frame = ttk.Frame(self.output_type_section)
        self.file_type_frame.pack(fill='x', expand=True)
        ttk.Radiobutton(self.file_type_frame, text='PDF', value='PDF', variable=self.file_type).pack(side=tk.LEFT, expand=True)
        ttk.Radiobutton(self.file_type_frame, text='AI', value='AI', variable=self.file_type).pack(side=tk.LEFT, expand=True)

        ttk.Button(badges_tab, text="Execute!", command=self.generate_badges).pack(side=tk.BOTTOM, pady=10)
            
    def create_file_selector(self, label, command, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(frame, text=label).pack(side=tk.LEFT)
        path_entry = ttk.Entry(frame, width=50)
        path_entry.pack(side=tk.LEFT, expand=True, padx=10)

        if label == "Select Illustrator Path":
            path_entry.insert(0, self.illustrator_path)
            ttk.Button(frame, text="Browse", command=lambda: command(path_entry, self.illustrator_path)).pack(side=tk.LEFT)
        else:
            ttk.Button(frame, text="Browse", command=lambda: command(path_entry, self.project_directory)).pack(side=tk.LEFT)

    def select_path(self, entry, initial_dir):
            filepath = filedialog.askopenfilename(initialdir=initial_dir)
            if filepath:
                entry.delete(0, tk.END)
                entry.insert(0, filepath)

    def select_illustrator_path(self, entry, initial_dir):
        filepath = filedialog.askopenfilename(title="Select Illustrator Executable", filetypes=[("Executable files", "*.exe")], initialdir=initial_dir)
        if filepath:
            entry.delete(0, tk.END)
            entry.insert(0, filepath)
            self.illustrator_path = filepath
        else:
            # If cancelled, restore the previous path
            entry.delete(0, tk.END)
            entry.insert(0, self.illustrator_path)

    def select_excel_file(self, entry, initial_dir=None):
        if not initial_dir:  # Fallback to a default directory if not provided
            initial_dir = os.getcwd()
        filepath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], initialdir=initial_dir)
        if filepath:
            entry.delete(0, tk.END)
            entry.insert(0, filepath)
            self.excel_path = filepath
            self.load_excel_columns(filepath)

    def load_excel_columns(self, filepath):
        try:
            df = pd.read_excel(filepath, nrows=0)
            columns = df.columns.tolist()
            self.display_column_checkboxes(self.columns_frame, columns, self.column_vars)
            self.display_column_checkboxes(self.filename_logic_frame, columns, self.filename_vars)
            self.display_column_checkboxes(self.qr_data_frame, columns, self.qr_column_vars)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")

    def display_column_checkboxes(self, frame, columns, vars_dict):
        for widget in frame.winfo_children():
            widget.destroy()
        for column in columns:
            vars_dict[column] = tk.BooleanVar()
            ttk.Checkbutton(frame, text=column, variable=vars_dict[column]).pack(anchor='w', side=tk.LEFT)

    def select_template_file(self, entry, initial_dir=None):
        if not initial_dir:  # Fallback to a default directory if not provided
            initial_dir = os.getcwd()
        filepath = filedialog.askopenfilename(title="Select Template File", filetypes=[("AI files", "*.ai"), ("All files", "*.*")], initialdir=initial_dir)
        if filepath:
            entry.delete(0, tk.END)
            entry.insert(0, filepath)
            self.template_path = filepath

    def select_output_base_path(self, entry, initial_dir=None):
        if not initial_dir:  # Fallback to a default directory if not provided
            initial_dir = os.getcwd()
        directory = filedialog.askdirectory(title="Select Output Base Path", initialdir=initial_dir)
        if directory:
            entry.delete(0, tk.END)
            entry.insert(0, directory)
            self.output_base_path = directory

    def handle_excel_ids(self):
        if self.excel_path:
            df = pd.read_excel(self.excel_path)
            if 'id' not in df.columns:
                df.insert(0, 'id', range(1, len(df) + 1))
                new_path = os.path.splitext(self.excel_path)[0] + '_ids.xlsx'
                df.to_excel(new_path, index=False)
                messagebox.showinfo("Info", f"ID column added and saved as {new_path}")
                self.excel_path = new_path  # Update the path in the system
            else:
                messagebox.showinfo("Info", "ID column already exists.")

    def generate_qr_codes(self):
        if self.qr_code_enabled.get() == "yes" and self.excel_path:
            df = pd.read_excel(self.excel_path)
            qr_output_dir = os.path.join(os.path.dirname(self.excel_path), "QR_Codes")
            os.makedirs(qr_output_dir, exist_ok=True)
            for index, row in df.iterrows():
                qr_data = '\n'.join(str(row[col]) for col in df.columns if self.qr_column_vars[col].get())
                qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4, image_factory=SvgPathImage)
                qr.add_data(qr_data)
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")
                file_name = f"{row['id']}.svg"
                img.save(os.path.join(qr_output_dir, file_name))
            messagebox.showinfo("Info", "QR codes generated successfully.")

    def generate_badges(self):
        if not all([self.illustrator_path, self.template_path, self.output_base_path, self.excel_path]):
            messagebox.showerror("Error", "Please select all required paths and at least one column")
            return
        selected_columns = {col: var.get() for col, var in self.column_vars.items() if var.get()}
        filename_parts = [col for col, var in self.filename_vars.items() if var.get()]
        file_extension = '.ai' if self.file_type.get() == 'AI' else '.pdf'
        df = pd.read_excel(self.excel_path, usecols=list(selected_columns.keys()))
        df['filename'] = df[filename_parts].apply(lambda row: '_'.join(row.values.astype(str)), axis=1)
        df['output_path'] = df['filename'].apply(lambda x: os.path.join(self.output_base_path, x + file_extension))

        # Save the temporary JSON in the script directory
        script_directory = os.path.dirname(__file__)
        temp_json_path = os.path.join(script_directory, 'temp_data.json')
        df.to_json(temp_json_path, orient='records')

        # Save the temporary JSX script in the script directory
        script_path = os.path.join(script_directory, 'temp_io_badge.jsx')
        self.write_jsx_script(script_path, temp_json_path, df['filename'])

        self.run_illustrator(script_path)

        # Clean up temporary files
        os.remove(temp_json_path)
        os.remove(script_path)

    def write_jsx_script(self, script_path, json_path, filenames):
        ai_option = 'true' if self.file_type.get() == 'AI' else 'false'
        qr_folder = os.path.join(os.path.dirname(self.excel_path), "QR_Codes") if self.qr_code_enabled.get() == "yes" else ""
        use_qr = 'true' if self.qr_code_enabled.get() == "yes" else 'false'
        script_content = f"""
#target illustrator
function readJSON(filePath) {{
    var file = new File(filePath);
    file.open('r');
    var jsonString = file.read();
    file.close();
    return eval("(" + jsonString + ")");
}}

var jsonFilePath = "{json_path.replace("\\", "\\\\")}";
var dataList = readJSON(jsonFilePath);

var templateFilePath = "{self.template_path.replace("\\", "\\\\")}";
var templateFile = new File(templateFilePath);

for (var i = 0; i < dataList.length; i++) {{
    var data = dataList[i];
    var doc = app.open(templateFile);

    for (var key in data) {{
        if (key !== 'svg_path' && data.hasOwnProperty(key)) {{
            var textFrames = doc.textFrames;
            for (var j = 0; j < textFrames.length; j++) {{
                var textFrame = textFrames[j];
                if (textFrame.name === key) {{
                    textFrame.contents = data[key];
                }}
            }}
        }}
    }}

    var outputFilePath = data["output_path"];
    var outputFile = new File(outputFilePath);
    if (outputFile.exists) {{
        var counter = 1;
        do {{
            var newFileName = outputFile.name.replace(/\\.[^\\.]+$/, '') + '_duplicate' + counter + outputFile.name.match(/\\.[^\\.]+$/)[0];
            outputFile = new File(outputFile.parent + '/' + newFileName);
            counter++;
        }} while (outputFile.exists);
    }}

    if ({use_qr}) {{
        var qrFilePath = "{qr_folder.replace("\\", "\\\\")}" + "/" + data.id + ".svg";
        var qrFile = new File(qrFilePath);
        if (qrFile.exists) {{
            var qrDoc = app.open(qrFile);
            app.executeMenuCommand('selectall');
            app.executeMenuCommand('copy');
            qrDoc.close(SaveOptions.DONOTSAVECHANGES);

            doc.activate();
            app.executeMenuCommand('paste');
            var pastedItem = doc.selection[0];

            var placeholder = doc.pageItems.getByName("QRPlaceholder");
            pastedItem.width = placeholder.width;
            pastedItem.height = placeholder.height;
            pastedItem.position = [placeholder.position[0], placeholder.position[1] - pastedItem.height + 52];
        }} else {{
            alert("QR code for the ID " + data.id + " is not found.");
        }}
    }}

    if ({ai_option}) {{
        var saveOptions = new IllustratorSaveOptions();
        saveOptions.pdfCompatible = false;
    }} else {{
        var saveOptions = new PDFSaveOptions();
    }}
    doc.saveAs(outputFile, saveOptions);
    doc.close(SaveOptions.DONOTSAVECHANGES);


}}
alert("Processing files completed.");
"""
        with open(script_path, 'w') as f:
            f.write(script_content)

    def run_illustrator(self, script_path):
        try:
            subprocess.run([self.illustrator_path, '-run', script_path], check=True)
            messagebox.showinfo("Success", "Illustrator script executed successfully.")
            self.restart_app()
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Execution Failed", f"Illustrator execution failed: {e}")

    def restart_app(self):
        self.destroy()
        self.__init__()    

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
