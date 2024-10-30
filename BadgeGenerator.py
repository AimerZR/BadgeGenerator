import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF
import pandas as pd
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue

# Constants
CONFIG_FILE = 'config.json'
OUTPUT_FOLDER = "Badge_output"
PRINT_FOLDER = "print"
DEFAULT_DPI = 300
HINT_TEXT = "Please enter the code after 'No.'."

# Queue for thread-safe GUI operations
gui_queue = queue.Queue()

def load_config(config_file):
    """
    Load and parse the configuration JSON file.

    :param config_file: Path to the configuration file.
    :return: Configuration dictionary.
    """
    try:
        with open(config_file, "r", encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        raise Exception(f"The configuration file '{config_file}' does not exist.")
    except json.JSONDecodeError as e:
        raise Exception(f"Error parsing '{config_file}': {e}")

def get_image_dimensions(image_path):
    """
    Calculate image dimensions in millimeters based on DPI.

    :param image_path: Path to the image file.
    :return: Tuple of (width_mm, height_mm).
    """
    with Image.open(image_path) as img:
        dpi = img.info.get('dpi', (DEFAULT_DPI, DEFAULT_DPI))[0]
        width_mm = (img.width / dpi) * 25.4
        height_mm = (img.height / dpi) * 25.4
    return width_mm, height_mm

def sanitize_filename(filename):
    """
    Sanitize the filename by removing or replacing illegal characters.

    :param filename: Original filename.
    :return: Sanitized filename.
    """
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    return filename

def clean_employee_id(employee_id):
    """
    Clean the employee ID by removing trailing '.0' if it's an integer.

    :param employee_id: Original employee ID (string).
    :return: Cleaned employee ID.
    """
    if employee_id.endswith('.0') and employee_id[:-2].isdigit():
        return employee_id[:-2]
    return employee_id

def generate_badge(side, data, suppress_message=False):
    """
    Generate a single badge image for the specified side ('front' or 'back').

    :param side: 'front' or 'back'.
    :param data: Dictionary containing badge data.
    :param suppress_message: If True, suppress success message.
    """
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Handle name splitting if necessary
    name = data.get('name', '')
    char_limit = data.get('name_char_limit', 10)

    if len(name) > char_limit:
        if ' ' in name:
            first, last = name.split(' ', 1)
            data.update({
                'first_name': first,
                'last_name': last,
                'text_elements': ['first_name', 'last_name', 'id', 'department', 'position']
            })
        else:
            error_msg = (f"The name '{name}' is too long. "
                         "Please use spaces to separate first and last names or edit it manually.")
            if not suppress_message:
                gui_queue.put(lambda: messagebox.showerror("Error", error_msg))
            else:
                raise Exception(error_msg)
            return
    else:
        data.update({
            'text_elements': ['name', 'id', 'department', 'position']
        })

    # Load background image
    try:
        with Image.open(data['background_img']) as img:
            img = img.convert("RGBA")  # Ensure image is in RGBA mode
    except Exception as e:
        error_msg = f"Failed to load {side} image: {e}"
        if not suppress_message:
            gui_queue.put(lambda: messagebox.showerror("Error", error_msg))
        else:
            raise Exception(error_msg)
        return

    # Insert photo if front side and photo provided
    if side == 'front' and data.get('photo_img'):
        try:
            with Image.open(data['photo_img']) as photo:
                photo = photo.resize(tuple(data['photo_size']), Image.LANCZOS)
                img.paste(photo, tuple(data['photo_position']), photo if photo.mode == 'RGBA' else None)
        except Exception as e:
            error_msg = f"Failed to insert photo: {e}"
            if not suppress_message:
                gui_queue.put(lambda: messagebox.showerror("Error", error_msg))
            else:
                raise Exception(error_msg)
            return

    draw = ImageDraw.Draw(img)

    # Load fonts and colors
    try:
        fonts = {
            key: ImageFont.truetype(data[f'font_{key}'], int(data[f'{key}_size']))
            for key in data['text_elements']
        }
        colors = {
            key: data.get(f'{key}_color', 'black')
            for key in data['text_elements']
        }
    except IOError as e:
        error_msg = f"Font not found: {e}"
        if not suppress_message:
            gui_queue.put(lambda: messagebox.showerror("Error", error_msg))
        else:
            raise Exception(error_msg)
        return

    # Draw text elements
    for element in data['text_elements']:
        position = tuple(data.get(f'{element}_position', (0, 0)))
        text = data.get(element, '')
        font = fonts[element]
        color = colors[element]
        draw.text(position, text, font=font, fill=color)

    # Generate output filename
    employee_id = data.get('id', '')
    name_for_filename = employee_id[4:] if employee_id.startswith('No.') else employee_id
    name_for_filename = sanitize_filename(name_for_filename)  # Ensure filename is safe
    output_filename = os.path.join(OUTPUT_FOLDER, f"badge_{side}_{name_for_filename}.png")

    # Save the badge image
    img.save(output_filename, quality=100)

    # Notify user
    if not suppress_message:
        gui_queue.put(lambda: messagebox.showinfo(
            "Success",
            f"{side.capitalize()} Badge saved as {output_filename}"
        ))

def generate_badge_pdf(config):
    """
    Generate a PDF containing all badge front and back images based on the configuration.

    :param config: Configuration dictionary.
    :return: Path to the generated PDF.
    """
    try:
        badge_folder = config["badge_folder"]
        front_prefix = config["badge_front_prefix"]
        back_prefix = config["badge_back_prefix"]
        output_pdf = config["output_pdf"]
        group_spacing_x = config["group_spacing_x"]
        group_spacing_y = config["group_spacing_y"]
        paper_width = config["paper_width"]
        paper_height = config["paper_height"]
        start_y = config["start_y_position"]
        badge_width = config["badge_width"]
    except KeyError as e:
        raise Exception(f"Missing key in config.json: {e}")

    try:
        # Initialize PDF
        pdf = FPDF(orientation='P', unit='mm', format=(paper_width, paper_height))
        pdf.set_auto_page_break(auto=True, margin=15)

        # Retrieve and sort badge files
        try:
            files = os.listdir(badge_folder)
        except FileNotFoundError:
            raise Exception(f"Badge folder '{badge_folder}' does not exist.")

        front_files = sorted([f for f in files if f.startswith(front_prefix)])
        back_files = sorted([f for f in files if f.startswith(back_prefix)])

        # Create dictionaries for front and back files keyed by their identifier
        front_files_dict = {f[len(front_prefix):-4]: f for f in front_files}
        back_files_dict = {f[len(back_prefix):-4]: f for f in back_files}

        # Get all unique keys
        all_keys = sorted(set(front_files_dict.keys()).union(back_files_dict.keys()))

        # Arrange badges in groups of three per page
        for i in range(0, len(all_keys), 3):
            pdf.add_page()
            for j in range(3):
                if i + j >= len(all_keys):
                    break

                key = all_keys[i + j]
                front_img_path = front_files_dict.get(key)
                back_img_path = back_files_dict.get(key)

                x_pos = (badge_width + group_spacing_x) * j + group_spacing_x
                y_pos = start_y

                # Insert front image if available
                if front_img_path:
                    front_full_path = os.path.join(badge_folder, front_img_path)
                    try:
                        front_width_mm, front_height_mm = get_image_dimensions(front_full_path)
                        pdf.image(front_full_path, x=x_pos, y=y_pos, w=front_width_mm, h=front_height_mm)
                        y_pos += front_height_mm + group_spacing_y
                    except Exception as e:
                        raise Exception(f"Failed to insert front image '{front_full_path}': {e}")

                # Insert back image if available
                if back_img_path:
                    back_full_path = os.path.join(badge_folder, back_img_path)
                    try:
                        back_width_mm, back_height_mm = get_image_dimensions(back_full_path)
                        pdf.image(back_full_path, x=x_pos, y=y_pos, w=back_width_mm, h=back_height_mm)
                    except Exception as e:
                        raise Exception(f"Failed to insert back image '{back_full_path}': {e}")

        # Ensure output directories exist
        os.makedirs(PRINT_FOLDER, exist_ok=True)

        output_pdf_path = os.path.join(os.getcwd(), output_pdf)
        try:
            pdf.output(output_pdf_path)
        except Exception as e:
            raise Exception(f"Failed to save PDF to '{output_pdf_path}': {e}")

        return output_pdf_path
    except Exception as e:
        raise Exception(f"An error occurred during PDF generation: {e}")

def select_image(entry_widget):
    """
    Open a file dialog to select an image file and set the selected path in the entry widget.

    :param entry_widget: Tkinter Entry widget to display the selected image path.
    """
    file_path = filedialog.askopenfilename(
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")],
        title="Select Image File"
    )
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def create_side_frame(parent, side, selected_preset, presets):
    """
    Create a frame for either the front or back badge side with all necessary input fields.

    :param parent: Parent tkinter widget.
    :param side: 'front' or 'back'.
    :param selected_preset: Variable storing the selected preset name.
    :param presets: Dictionary of presets from config.json.
    :return: Configured tkinter Frame.
    """
    frame = tk.Frame(parent)
    current_row = 0

    # Header label
    tk.Label(
        frame,
        text=f"{side.capitalize()} Badge Content",
        font=("Arial", 14)
    ).grid(row=current_row, column=0, columnspan=3, padx=10, pady=10)
    current_row += 1

    # Input fields
    entries = {}
    fields = [
        ("Name:", "name"),
        ("Employee ID:", "id"),
        ("Department:", "department"),
        ("Position:", "position")
    ]

    for label_text, field_key in fields:
        tk.Label(frame, text=label_text).grid(row=current_row, column=0, padx=10, pady=5, sticky='e')
        entry = tk.Entry(frame, width=30)
        entry.grid(row=current_row, column=1, padx=10, pady=5, sticky='w')
        entries[field_key] = entry

        # Special handling for Employee ID hint
        if field_key == 'id':
            entry.insert(0, HINT_TEXT)

            def on_focus_in(event, e=entry):
                if e.get() == HINT_TEXT:
                    e.delete(0, tk.END)

            entry.bind("<FocusIn>", on_focus_in)

        current_row += 1

    # Photo selection for front side
    if side == 'front':
        tk.Label(frame, text="Photo Image:").grid(row=current_row, column=0, padx=10, pady=5, sticky='e')
        photo_entry = tk.Entry(frame, width=30)
        photo_entry.grid(row=current_row, column=1, padx=10, pady=5, sticky='w')
        tk.Button(
            frame,
            text="Browse",
            command=lambda: select_image(photo_entry)
        ).grid(row=current_row, column=2, padx=5, pady=5)
        entries['photo_img'] = photo_entry
        current_row += 1
    else:
        entries['photo_img'] = None

    def generate():
        """
        Callback function to generate the badge based on user inputs.
        """
        preset_name = selected_preset.get()
        preset = presets[preset_name][side]
        data = preset.copy()

        # Gather user input
        data.update({
            'name': entries['name'].get().strip(),
            'id': f"No. {entries['id'].get().strip().upper()}",
            'department': entries['department'].get().strip(),
            'position': entries['position'].get().strip(),
        })

        if side == 'front' and entries['photo_img']:
            data['photo_img'] = entries['photo_img'].get().strip()

        # Generate the badge in a separate thread to keep GUI responsive
        threading.Thread(target=generate_badge, args=(side, data)).start()

    # Generate button
    tk.Button(
        frame,
        text=f"Generate {side.capitalize()} Badge",
        command=generate
    ).grid(row=current_row, column=0, columnspan=3, padx=10, pady=10)

    return frame

def batch_generate_badges(presets, selected_preset_name, error_log_path):
    """
    Batch generate badges from an Excel file using the selected preset.

    :param presets: Dictionary of presets from config.json.
    :param selected_preset_name: The name of the currently selected preset.
    :param error_log_path: Path to the error log file.
    """
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_file:
        return  # User cancelled the dialog

    def task():
        try:
            # Specify 'Employee Number' column as string type
            df = pd.read_excel(excel_file, dtype={'Employee Number': str})

            # Validate required columns
            required_columns = [
                'Local Name',        # Local Name
                'English Name',      # English Name
                'Employee Number',   # Employee Number
                'Department',        # Department (Local)
                'Position',          # Position (Local)
                'Department_en',     # Department (English)
                'Position_en'        # Position (English)
            ]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                error_message = f"The Excel file is missing the following required columns: {', '.join(missing_columns)}"
                gui_queue.put(lambda: messagebox.showerror(
                    "Error",
                    error_message
                ))
                # Write to error log
                write_errors_to_log([error_message], error_log_path)
                return

            preset_name = selected_preset_name.get()
            if preset_name not in presets:
                error_message = f"The selected preset '{preset_name}' does not exist."
                gui_queue.put(lambda: messagebox.showerror("Error", error_message))
                # Write to error log
                write_errors_to_log([error_message], error_log_path)
                return

            errors = []  # List to collect error messages
            errors_lock = threading.Lock()  # Lock to ensure thread-safe access to errors

            # Check for duplicate Employee Numbers
            employee_numbers = df['Employee Number'].astype(str).str.strip()
            duplicates = employee_numbers[employee_numbers.duplicated(keep=False)].unique()
            if len(duplicates) > 0:
                duplicate_message = f"Duplicate Employee Numbers found: {', '.join(duplicates)}"
                errors.append(duplicate_message)

            # Helper function to check if a value is missing
            def is_missing(value):
                return pd.isna(value) or str(value).strip() == ''

            # Function to process each row
            def process_row(index, row):
                try:
                    # Extract and validate required fields
                    local_name = row['Local Name']
                    english_name = row['English Name']
                    employee_id = row['Employee Number']
                    department_local = row['Department']
                    position_local = row['Position']
                    department_en = row['Department_en']
                    position_en = row['Position_en']

                    # Check for missing values
                    missing_fields = []
                    if is_missing(local_name):
                        missing_fields.append('Local Name')
                    if is_missing(english_name):
                        missing_fields.append('English Name')
                    if is_missing(employee_id):
                        missing_fields.append('Employee Number')
                    if is_missing(department_local):
                        missing_fields.append('Department')
                    if is_missing(position_local):
                        missing_fields.append('Position')
                    if is_missing(department_en):
                        missing_fields.append('Department_en')
                    if is_missing(position_en):
                        missing_fields.append('Position_en')

                    if missing_fields:
                        error_message = f"Row {index + 2}: Missing fields - {', '.join(missing_fields)}."
                        with errors_lock:
                            errors.append(error_message)
                        return

                    # Clean employee ID
                    employee_id = clean_employee_id(str(employee_id).strip())

                    # Convert to string and strip
                    local_name = str(local_name).strip()
                    english_name = str(english_name).strip()
                    department_local = str(department_local).strip()
                    position_local = str(position_local).strip()
                    department_en = str(department_en).strip()
                    position_en = str(position_en).strip()

                    # Check if Employee Number is duplicated
                    if employee_id in duplicates:
                        error_message = f"Row {index + 2}: Employee Number '{employee_id}' is duplicated."
                        with errors_lock:
                            errors.append(error_message)
                        return

                    # Generate front badge
                    front_preset = presets[preset_name]['front'].copy()
                    front_data = {
                        'name': local_name,
                        'id': f"No. {employee_id}",
                        'department': department_local,
                        'position': position_local,
                    }
                    front_data.update(front_preset)
                    generate_badge('front', front_data, suppress_message=True)

                    # Generate back badge
                    back_preset = presets[preset_name]['back'].copy()
                    back_data = {
                        'name': english_name,
                        'id': f"No. {employee_id}",
                        'department': department_en,
                        'position': position_en,
                    }
                    back_data.update(back_preset)
                    generate_badge('back', back_data, suppress_message=True)

                except Exception as e:
                    error_message = f"Row {index + 2}: Failed to generate badge for Employee ID {employee_id}: {e}"
                    with errors_lock:
                        errors.append(error_message)

            # Use ThreadPoolExecutor for parallel processing
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(process_row, idx, row) for idx, row in df.iterrows()]
                for future in as_completed(futures):
                    pass  # Errors are handled within process_row

            # Write all errors to the error log
            if errors:
                write_errors_to_log(errors, error_log_path)
                # Notify user via message box
                gui_queue.put(lambda: messagebox.showerror(
                    "Batch Generation Errors",
                    f"Errors were encountered during batch generation. Please check the error log at:\n{error_log_path}"
                ))
            else:
                # No errors, notify success
                gui_queue.put(lambda: messagebox.showinfo(
                    "Success",
                    "All badges have been successfully generated and saved in the 'Badge_output' folder."
                ))

        except Exception as e:
            error_message = f"Failed to generate badges: {e}"
            gui_queue.put(lambda: messagebox.showerror("Error", error_message))
            # Write to error log
            write_errors_to_log([f"General Error: {e}"], error_log_path)

    # Run the task in a separate thread to keep GUI responsive
    threading.Thread(target=task).start()

def write_errors_to_log(errors, error_log_path):
    """
    Write all error messages to the specified error log file.

    :param errors: List of error messages.
    :param error_log_path: Path to the error log file.
    """
    try:
        # Ensure the directory exists
        os.makedirs(os.path.dirname(error_log_path), exist_ok=True)
        with open(error_log_path, 'w', encoding='utf-8') as f:
            for error in errors:
                f.write(error + '\n')
    except Exception as e:
        # If writing to the log fails, notify the user
        error_message = f"Failed to write to error log: {e}"
        gui_queue.put(lambda: messagebox.showerror("Error", error_message))

def create_gui():
    """
    Create and launch the main GUI for the Badge Generator application.
    """
    window = tk.Tk()
    window.title("Badge Generator")
    window.geometry('1200x400')  # Increased height for better layout

    # Load presets from config.json
    try:
        config = load_config(CONFIG_FILE)
        presets = config['presets']
        error_log_path = config.get('error_log', 'error_logs/generation_errors.txt')
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load {CONFIG_FILE}: {e}")
        window.destroy()
        return

    # Variable to store selected preset
    selected_preset = tk.StringVar(value=list(presets.keys())[0])

    # Preset selection frame
    preset_frame = tk.Frame(window)
    preset_frame.pack(fill='x', padx=10, pady=10)

    # Create sub-frames for better layout management
    left_subframe = tk.Frame(preset_frame)
    left_subframe.pack(side=tk.LEFT, fill='x', expand=True)

    right_subframe = tk.Frame(preset_frame)
    right_subframe.pack(side=tk.RIGHT)

    # Left Sub-Frame: Preset Selection
    tk.Label(left_subframe, text="Select Preset:").pack(side=tk.LEFT)

    for preset_name in presets:
        tk.Radiobutton(
            left_subframe,
            text=preset_name,
            variable=selected_preset,
            value=preset_name
        ).pack(side=tk.LEFT, padx=5)

    # Right Sub-Frame: Action Buttons
    tk.Button(
        right_subframe,
        text="Generate PDF",
        command=lambda: generate_pdf_action(config)
    ).pack(side=tk.RIGHT, padx=5)

    tk.Button(
        right_subframe,
        text="Batch Generate Badges",
        command=lambda: batch_generate_badges(presets, selected_preset, error_log_path)
    ).pack(side=tk.RIGHT, padx=5)

    def generate_pdf_action(config):
        """
        Action to generate PDF and handle exceptions.
        """
        def task():
            try:
                output_pdf = generate_badge_pdf(config)
                gui_queue.put(lambda: messagebox.showinfo("Success", f"PDF successfully generated at:\n{output_pdf}"))
            except Exception as e:
                error_message = f"Failed to generate PDF: {e}"
                gui_queue.put(lambda: messagebox.showerror("Error", error_message))

        # Run PDF generation in a separate thread
        threading.Thread(target=task).start()

    # Main content frame for front and back badge configurations
    content_frame = tk.Frame(window)
    content_frame.pack(fill='both', expand=True, padx=10, pady=10)

    # Configure grid for equal expansion
    content_frame.grid_rowconfigure(0, weight=1)
    content_frame.grid_columnconfigure(0, weight=1)
    content_frame.grid_columnconfigure(1, weight=1)

    # Create frames for front and back badge configurations
    front_frame = create_side_frame(content_frame, "front", selected_preset, presets)
    back_frame = create_side_frame(content_frame, "back", selected_preset, presets)

    front_frame.grid(row=0, column=0, sticky='nsew', padx=5)
    back_frame.grid(row=0, column=1, sticky='nsew', padx=5)

    def process_gui_queue():
        """
        Process any pending GUI operations from the queue.
        """
        try:
            while True:
                task = gui_queue.get_nowait()
                task()
        except queue.Empty:
            pass
        window.after(100, process_gui_queue)  # Check the queue every 100ms

    # Start processing the GUI queue
    window.after(100, process_gui_queue)

    window.mainloop()

if __name__ == "__main__":
    create_gui()
