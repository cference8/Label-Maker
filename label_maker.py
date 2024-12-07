import os
import customtkinter as ctk
from tkinter import TclError, filedialog, messagebox, colorchooser, Menu
import logging
from PIL import Image, ImageDraw, ImageFont, features
import qrcode
from io import BytesIO

# Set up logging to log errors to a file
logging.basicConfig(filename='file_processing_errors.log',
                    level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Function to locate resource files, works for both PyInstaller executable and dev environment
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    import sys
    try:
        # PyInstaller creates a temporary folder and stores the path in sys._MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    
    # Ensure backslashes for Windows paths
    return os.path.join(base_path, relative_path).replace('\\', '/')


# Ensure the history file is stored in a user-writable location (not inside the executable)
def get_history_file_path():
    """ Get the writable path for the history file in a fixed directory on your computer. """
    # Define the fixed path where you want the history file to be saved
    base_path = r"G:\Shared drives\Scribe Workspace\Scribe Master Folder\Scribe Label Maker"
        
    # Ensure the directory exists (creates it if it doesn't exist)
    os.makedirs(base_path, exist_ok=True)
    
    # Construct the full path for the history file
    history_file_path = os.path.join(base_path, "order_history.json")
    
    return history_file_path

# Set CustomTkinter appearance mode and color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

# Global variables
labels_data = []
displayed_envelope_files = set()
displayed_letter_files = set()
order_colors = {}
qr_codes = {}

# Function to load order history from JSON file
def load_order_history():
    import json
    history_file_path = get_history_file_path()
    if os.path.exists(history_file_path):
        with open(history_file_path, 'r') as file:
            return json.load(file)
    return []

# Function to save order history to JSON file
def save_order_history(history):
    import json
    history_file_path = get_history_file_path()
    with open(history_file_path, 'w') as file:
        json.dump(history, file)

# Function to update the order history with the last 10 entries
def update_order_history(order_name, color):
    history = load_order_history()

    # Remove any existing entries for the same order_name to ensure uniqueness
    history = [entry for entry in history if entry['order_name'] != order_name]

    # Add the new entry
    history.append({"order_name": order_name, "color": color})

    # Ensure only the last 20 unique entries are kept
    if len(history) > 20:
        history = history[-20:]

    save_order_history(history)

# Function to display the last 20 order_name and color combinations in the GUI
def display_order_history():
    # Clear previous displayed history
    for widget in history_label_frame.winfo_children():
        widget.destroy()

    history = load_order_history()

    # Reverse the history to show newest entries first
    history.reverse()

    for entry in history:
        order_name = entry['order_name']
        color = entry['color']

        # Populate the order_colors dictionary
        order_colors[order_name] = color

        # Determine the appropriate text color based on background brightness
        def is_light_color(hex_color):
            hex_color = hex_color.lstrip("#")
            r, g, b = (
                int(hex_color[0:2], 16),
                int(hex_color[2:4], 16),
                int(hex_color[4:6], 16),
            )
            brightness = (r * 299 + g * 587 + b * 114) / 1000  # Luminance formula
            return brightness > 186

        text_color = "black" if is_light_color(color) else "white"

        # Create a label for each history entry with the background color set to the assigned color
        order_label = ctk.CTkLabel(
            history_label_frame,
            text=order_name,
            font=("Helvetica", 12),
            text_color=text_color,
        )
        order_label.configure(fg_color=f"#{color}")  # Set the background color to the assigned color
        order_label.pack(pady=2, padx=5, anchor="w", fill="x")

        # Adjust bindtags to include history_label_frame
        order_label.bindtags((str(order_label), str(history_label_frame), "all"))

        # Bind click event to change color
        order_label.bind(
            "<Button-1>",
            lambda event, name=order_name, label=order_label: change_order_history_color(event, name, label),
        )

# Function to select Envelope Chip files
def select_envelope_files():
    files = filedialog.askopenfilenames(title="Select Envelope Chip Files")
    if files:
        generate_labels_data(files, "Envelopes")
    else:
        messagebox.showerror("Error", "No Envelope Chip files selected!")

# Function to select Letter Chip files
def select_letter_files():
    files = filedialog.askopenfilenames(title="Select Letter Chip Files")
    if files:
        generate_labels_data(files, "Letters")
    else:
        messagebox.showerror("Error", "No Letter Chip files selected!")

# Function to generate the labels data based on selected files
def generate_labels_data(files, chip_type):
    import logging
    from collections import defaultdict
    import re
    global labels_data, displayed_envelope_files, displayed_letter_files, order_colors
    valid_files = []  # To store valid files for display in the labels
    order_type_count = defaultdict(list)  # To group by order_name and card_envelope type
    invalid_files = []  # List to track invalid files

    # Convert labels_data into a set of existing (order_name, card_envelope) pairs to check for duplicates
    existing_labels = {(entry['order_name'], entry['card_envelope']) for entry in labels_data}

    for file_path in files:
        try:
            # Validate if file exists and is accessible
            if not os.path.isfile(file_path):
                logging.error(f"File not found or inaccessible: {file_path}")
                continue

            file_name = os.path.basename(file_path)
            base_name = os.path.splitext(file_name)[0]  # Remove the extension (e.g., ".bin")

            # Check if the file name contains "Envelopes" or "Letters"
            if "Envelopes" in base_name:
                card_envelope = "Envelope"
                parts = base_name.split("Envelopes")
                order_name = parts[0].strip()
                rest = parts[1].strip()
            elif "Letters" in base_name:
                card_envelope = "Card"
                parts = base_name.split("Letters")
                order_name = parts[0].strip()
                rest = parts[1].strip()
            else:
                # Prompt the user to select the type
                selected_type = ask_envelope_or_letter(file_name)
                if not selected_type:
                    # User cancelled or closed the dialog
                    continue
                if selected_type == "Envelopes":
                    card_envelope = "Envelope"
                elif selected_type == "Letters":
                    card_envelope = "Card"
                else:
                    # Invalid selection
                    continue
                # Since we don't have "Envelopes" or "Letters" in the base_name,
                # we'll assume the entire base_name is the order_name
                order_name = base_name.strip()
                rest = ''

            # Now try to extract the range from rest or base_name
            if rest:
                range_match_string = rest
            else:
                range_match_string = base_name

            # Extract the range using regex
            match = re.search(r'(\d+)-(\d+)', range_match_string)
            if match:
                start_num = int(match.group(1))
                end_num = int(match.group(2))
                num_records = end_num - start_num + 1
                # Remove the range from order_name if it was in order_name
                if not rest:
                    order_name = order_name[:match.start()].strip()
            else:
                num_records = None

            # Check if this label has already been generated
            if (order_name, card_envelope) in existing_labels:
                continue  # Skip if this order_name and card_envelope already exist

            # Append valid files to labels_data
            label_entry = {
                "order_name": order_name,
                "batch_chip": "1 of 1",  # Default batch_chip for single files
                "card_envelope": card_envelope,
                "num_records": num_records  # Store the number of records
            }
            labels_data.append(label_entry)

            # Add valid files to the appropriate set for preventing duplicates
            if card_envelope == "Envelope" and file_name not in displayed_envelope_files:
                valid_files.append(file_name)
                displayed_envelope_files.add(file_name)
            elif card_envelope == "Card" and file_name not in displayed_letter_files:
                valid_files.append(file_name)
                displayed_letter_files.add(file_name)

            # Prompt for color if order_name doesn't already have one
            if order_name not in order_colors:
                assign_color_for_order(order_name)

        except Exception as e:
            # Log any unexpected exceptions during file processing
            logging.error(f"Failed to process file '{file_path}': {str(e)}")

    # Notify the user if there were any invalid files
    if invalid_files:
        invalid_file_list = "\n".join(invalid_files)
        messagebox.showerror("Invalid Files", f"The following files were invalid for {chip_type}:\n{invalid_file_list}")

    # Update the labels_data entries to assign batch numbers if there are multiple files per order
    order_type_files = defaultdict(list)
    for label in labels_data:
        key = (label['order_name'], label['card_envelope'])
        order_type_files[key].append(label)

    for (order_name, card_envelope), labels in order_type_files.items():
        total_batches = len(labels)
        for i, label in enumerate(labels):
            label['batch_chip'] = f"{i + 1} of {total_batches}"

    # Display only valid and non-duplicate files in the appropriate label
    if valid_files:
        file_names = "\n".join(valid_files)
        if chip_type == "Envelopes":
            current_text = envelope_label.cget("text")
            envelope_label.configure(text=current_text + file_names + "\n")
        elif chip_type == "Letters":
            current_text = letter_label.cget("text")
            letter_label.configure(text=current_text + file_names + "\n")

def ask_envelope_or_letter(file_name):
    import customtkinter as ctk
    result = [None]  # Use a list to hold the result due to scope

    dialog = ctk.CTkToplevel()
    dialog.title("File Type")
    dialog.geometry("400x200")
    dialog.attributes('-topmost', True)
    dialog.grab_set()

    label = ctk.CTkLabel(
        dialog,
        text=f"File '{file_name}' does not contain 'Envelopes' or 'Letters'.\nPlease select the type:",
        font=("Helvetica", 14),
        wraplength=380,
        justify="left"
    )
    label.pack(pady=20, padx=10)

    def on_select(choice):
        result[0] = choice
        dialog.grab_release()
        dialog.destroy()

    button_frame = ctk.CTkFrame(dialog)
    button_frame.pack(pady=10)

    envelope_button = ctk.CTkButton(
        button_frame,
        text="Envelopes",
        command=lambda: on_select("Envelopes"),
        width=120
    )
    envelope_button.pack(side="left", padx=10)

    letters_button = ctk.CTkButton(
        button_frame,
        text="Letters",
        command=lambda: on_select("Letters"),
        width=120
    )
    letters_button.pack(side="left", padx=10)

    dialog.wait_window()

    return result[0]

# Function to prompt user for a color for each unique order_name
def assign_color_for_order(order_name):
    color = colorchooser.askcolor(title=f"Choose color for {order_name}")
    if color[1]:
        order_colors[order_name] = color[1][1:]
        display_order_color(order_name, color[1])

# Function to change the color when the label is clicked
def change_color(order_name, color_label):
    color = colorchooser.askcolor(title=f"Choose a new color for {order_name}")
    if color[1]:
        order_colors[order_name] = color[1][1:]
        color_label.configure(fg_color=color[1])

# Function to display the order color
def display_order_color(order_name, color_hex):
    def is_light_color(hex_color):
        hex_color = hex_color.lstrip("#")
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        brightness = (r * 299 + g * 587 + b * 114) / 1000  # Luminance formula
        return brightness > 186

    text_color = "black" if is_light_color(color_hex) else "white"

    envelope_label.pack_forget()
    letter_label.pack_forget()

    # Create the color label
    color_label = ctk.CTkLabel(
        scrollable_frame,
        text=f"{order_name} assigned color",
        fg_color=color_hex,
        text_color=text_color
    )
    
    # Bind the left-click event to change the label color
    color_label.bind("<Button-1>", lambda event: change_label_color_on_click(event, color_label))

    # Pack the label
    color_label.pack(pady=5, padx=20)

    letter_label.pack(pady=10, padx=20, fill="x", side="bottom")
    envelope_label.pack(pady=10, padx=20, fill="x", side="bottom")

def change_label_color_on_click(event, label):
    # Prompt the user to choose a new color
    color = colorchooser.askcolor(title="Choose a new color for the label")
    if color[1]:
        # Update the label's background color (fg_color) with the selected color
        label.configure(fg_color=color[1])

        # Adjust text color based on the brightness of the selected color
        def is_light_color(hex_color):
            hex_color = hex_color.lstrip("#")
            if len(hex_color) == 6:  # Ensure we have a valid 6-digit hex color
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                brightness = (r * 299 + g * 587 + b * 114) / 1000
                return brightness > 186
            else:
                return False  # Default to dark if the color is invalid or not as expected

        # Set text color to black if light color, white if dark color
        text_color = "black" if is_light_color(color[1]) else "white"
        label.configure(text_color=text_color)

        # Update the order_colors dictionary with the new color (remove '#')
        order_name = label.cget("text").split(" assigned color")[0]
        order_colors[order_name] = color[1][1:]

        # Update the order history with the new color
        update_order_history(order_name, order_colors[order_name])

        # Refresh the display order history to reflect the color change
        display_order_history()

def change_order_history_color(event, order_name, order_label):
    color = colorchooser.askcolor(title=f"Choose new color for {order_name}")
    if color[1]:
        new_color = color[1][1:]  # Remove the '#' from the color code
        order_colors[order_name] = new_color
        order_label.configure(fg_color=color[1])

        # Update text color based on brightness
        def is_light_color(hex_color):
            hex_color = hex_color.lstrip("#")
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            return brightness > 186

        text_color = "black" if is_light_color(new_color) else "white"
        order_label.configure(text_color=text_color)

        # Update the order history
        update_order_history(order_name, new_color)

def generate_labels_pdf(labels_data, qr_codes, output_pdf="labels_with_qr.pdf"):
    """
    Generates a multi-page PDF of labels with a 2x6 layout and QR codes for standard letter-sized paper (8.5x11 inches).

    :param labels_data: List of dictionaries with label data (order_name, batch_chip, card_envelope, color, num_records).
    :param qr_codes: Dictionary mapping order_name to its QR code URL.
    :param output_pdf: Output file name for the generated PDF.
    """

    # Define page dimensions for 8.5 x 11 inches at 300 DPI
    PAGE_WIDTH, PAGE_HEIGHT = 2550, 3300  # 8.5 x 11 inches at 300 DPI
    MARGIN = 80  # Margin in pixels
    LABEL_WIDTH, LABEL_HEIGHT = 1100, 450  # Label dimensions for 2x6 grid
    GAP_X, GAP_Y = 150, 80  # Gaps between labels horizontally and vertically

    # Load fonts using the resource_path function from your existing code
    font_path = resource_path('resources/Arial_Bold.ttf')  # Adjust the path as needed
    if not os.path.exists(font_path):
        messagebox.showerror("Error", f"Font file not found at {font_path}")
        return

    font_large = ImageFont.truetype(font_path, 60)
    font_medium = ImageFont.truetype(font_path, 50)
    font_small = ImageFont.truetype(font_path, 40)  # Added a smaller font size for num_records

    # Prepare to store pages
    pages = []

    # Number of rows and columns for the 2x6 layout
    num_rows = 6
    num_cols = 2
    labels_per_page = num_rows * num_cols

    for page_num in range((len(labels_data) + labels_per_page - 1) // labels_per_page):
        # Create a blank canvas for the page
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)

        # Calculate start and end indices for labels on this page
        start_index = page_num * labels_per_page
        end_index = min(start_index + labels_per_page, len(labels_data))

        # Calculate label positions column-first
        positions = []
        for col in range(num_cols):
            for row in range(num_rows):
                x_start = MARGIN + col * (LABEL_WIDTH + GAP_X)
                y_start = MARGIN + row * (LABEL_HEIGHT + GAP_Y)
                positions.append((x_start, y_start))

        # Draw labels for this page
        for i, label_index in enumerate(range(start_index, end_index)):
            x_start, y_start = positions[i]
            x_end = x_start + LABEL_WIDTH
            y_end = y_start + LABEL_HEIGHT

            # Draw label box
            draw.rectangle([x_start, y_start, x_end, y_end], outline="black", width=3)
    
            # Calculate the position of the vertical line
            x_line = x_start + (LABEL_WIDTH - 450) // 2  # Line in the middle of the label (adjusted with - 500)

            # Draw the vertical line
            draw.line([(x_line, y_start + 200), (x_line, y_end - 100)], fill="black", width=5)

            # Add label text
            label = labels_data[label_index]
            order_name = label["order_name"]
            batch_chip = label["batch_chip"]
            card_envelope = label["card_envelope"]
            num_records = label.get('num_records', None)
            text_color = label.get('color', "black")

            # Ensure text_color is in a format PIL can use
            if isinstance(text_color, str):
                if not text_color.startswith('#') and len(text_color) == 6:
                    text_color = '#' + text_color  # Add '#' if missing
                # PIL can handle color names and hex strings with '#'

            # Draw the text onto the label
            draw.text((x_start + 20, y_start + 20), f"Order Name & Number:", fill='black', font=font_medium)
            draw.text((x_start + 30, y_start + 90), order_name, fill=text_color, font=font_large)
            draw.text((x_start + 20, y_start + 200), "Chip #:", fill='black', font=font_medium)
            draw.text((x_start + 30, y_start + 265), batch_chip, fill=text_color, font=font_large)
            if num_records is not None:
                draw.text((x_start + 400, y_start + 200), f"# of Records:", fill='black', font=font_medium)
                draw.text((x_start + 480, y_start + 265), str(num_records), fill=text_color, font=font_medium)
                # Adjust positions as needed to fit all text
            else:
                # If num_records is not available, adjust the positions of "Type:"
                y_type = y_start + 320
            draw.text((x_start + 20, y_start + 360), "Type:", fill='black', font=font_medium)
            draw.text((x_start + 170, y_start + 360), card_envelope, fill=text_color, font=font_large)

            # Add QR code if it exists for the order
            if order_name in qr_codes:
                qr = qrcode.QRCode()
                qr.add_data(qr_codes[order_name])
                qr.make(fit=True)

                # Generate QR code image
                qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
                qr_size = 250  # QR code size
                qr_img = qr_img.resize((qr_size, qr_size))
                qr_position = (x_end - qr_size - 20, y_start + 190)  # Position near top-right of label
                page.paste(qr_img, qr_position)

        # Append this page to the pages list
        pages.append(page)

    # Save pages as a single PDF
    if pages:
        pages[0].save(output_pdf, save_all=True, append_images=pages[1:], resolution=300)
        print(f"Labels saved to {output_pdf}")
    else:
        messagebox.showerror("Error", "No pages were created. Check if labels_data is populated correctly.")

# Function to create and save the PDF file and display the "Open File" button if successful
def create_pdf():
    if not labels_data:
        messagebox.showerror("Error", "No valid files selected!")
        return

    try:
        import logging
        from tkinter import simpledialog

        # Prompt the user for a file name
        file_name = simpledialog.askstring("Input", "Enter the file name for the PDF:", parent=root)
        
        if not file_name:
            messagebox.showerror("Error", "File name not specified.")
            return

        # Default save path (ensure it doesn't end with a backslash)
        save_directory = r"G:\Shared drives\Scribe Workspace\Scribe Master Folder\Batch Labels"

        # Ensure the directory exists
        os.makedirs(save_directory, exist_ok=True)

        # Construct the full save path using os.path.join and os.path.normpath
        save_path = os.path.normpath(os.path.join(save_directory, f"{file_name}.pdf"))

        # Print the save path for debugging
        print(f"PDF will be saved to: {save_path}")

        # Update labels_data to include color information
        for label in labels_data:
            label['color'] = order_colors.get(label['order_name'], 'black')

        # Call generate_labels_pdf
        generate_labels_pdf(labels_data, qr_codes, output_pdf=save_path)

        # Check if the file was created
        if not os.path.exists(save_path):
            messagebox.showerror("Error", f"The PDF file was not created at {save_path}")
            return

        messagebox.showinfo("Success", f"Labels saved to {save_path}")
        open_button.configure(command=lambda: open_pdf_file(save_path))
        open_button.pack(pady=10, padx=20, before=reset_button)

        # Update order history
        for label in labels_data:
            update_order_history(label['order_name'], order_colors[label['order_name']])

        display_order_history()

    except Exception as e:
        logging.error(f"Failed to create the PDF file: {str(e)}")
        messagebox.showerror("Error", f"Failed to create the PDF file: {str(e)}")

def open_pdf_file(file_path):
    import platform
    import subprocess
    import logging
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":
            subprocess.Popen(['open', file_path])
        else:
            subprocess.Popen(['xdg-open', file_path])
    except Exception as e:
        logging.error(f"Failed to open the PDF file: {str(e)}")
        messagebox.showerror("Error", f"Failed to open the PDF file: {str(e)}")

# Scroll event for mouse (Windows)
def on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

# Scroll event for macOS (Mac uses different event types for scrolling)
def on_mousewheel_mac(event):
    canvas.yview_scroll(-1 if event.num == 5 else 1, "units")

# Function to reset label data, displayed files, and order colors
def reset_data():
    global labels_data, displayed_envelope_files, displayed_letter_files, order_colors

    # Clear all relevant data
    labels_data.clear()
    displayed_envelope_files.clear()
    displayed_letter_files.clear()
    order_colors.clear()

    # Reset label texts
    envelope_label.configure(text="Selected Envelope Files:\n")
    letter_label.configure(text="Selected Letter Files:\n")

    # Destroy any color assignment labels
    for widget in scrollable_frame.winfo_children():
        if isinstance(widget, ctk.CTkLabel) and "assigned color" in widget.cget("text"):
            widget.destroy()

    # Hide the open button
    open_button.pack_forget()

qr_window = None  # Initialize the global variable

def add_qr_code_window():
    global qr_window, qr_codes

    if not labels_data:
        messagebox.showerror("Error", "No order label to add QR code.")
        return

    # Check if the QR Code window is already open
    try:
        if qr_window is not None and qr_window.winfo_exists():
            # Bring the existing window to the front
            qr_window.lift()
            qr_window.focus_force()
            return
    except (AttributeError, TclError):
        pass

    # Create a new pop-up window
    qr_window = ctk.CTkToplevel(root)
    qr_window.title("Add QR Code")
    qr_window.geometry("400x350")
    qr_window.configure(bg="#3A3A3A")

    # Ensure the window is on top and grabs focus
    qr_window.attributes('-topmost', True)
    qr_window.focus_force()
    qr_window.grab_set()

    # Dropdown for order selection
    order_label = ctk.CTkLabel(
        qr_window,
        text="Select Order Name:",
        font=("Helvetica", 14),
        text_color="black"
    )
    order_label.pack(pady=10, padx=10)

    # Generate a list of unique order names, preserving order
    order_names = []
    seen_order_names = set()
    display_name_to_order_name = {}  # Mapping from display name to actual order name
    for label in labels_data:
        order_name = label['order_name']
        if order_name not in seen_order_names:
            seen_order_names.add(order_name)
            # Check if the order has a QR code
            if order_name in qr_codes:
                display_name = f"{order_name} - has QR Code"
            else:
                display_name = order_name
            order_names.append(display_name)
            display_name_to_order_name[display_name] = order_name

    selected_order = ctk.StringVar(qr_window)
    dropdown = ctk.CTkOptionMenu(qr_window, variable=selected_order, values=order_names)
    dropdown.pack(pady=10, padx=10, fill="x")

    # Text field for QR code URL
    url_label = ctk.CTkLabel(
        qr_window,
        text="Enter URL for QR Code:",
        font=("Helvetica", 14),
        text_color="black"
    )
    url_label.pack(pady=10, padx=10)

    # Updated input box with placeholder text
    url_entry = ctk.CTkEntry(qr_window, placeholder_text="Right click here to paste URL")
    url_entry.pack(pady=10, padx=10, fill="x")

    # Bind the right-click event to paste from clipboard
    def paste_clipboard(event):
        try:
            # Get the clipboard content
            url = qr_window.clipboard_get()
            # Clear the current content of the input box
            url_entry.delete(0, 'end')
            # Insert the clipboard content into the input box
            url_entry.insert(0, url)
        except:
            # Handle exceptions (e.g., clipboard is empty or contains non-text data)
            messagebox.showerror("Error", "Clipboard does not contain valid text.")

    url_entry.bind("<Button-3>", paste_clipboard)  # For Windows and Linux
    # url_entry.bind("<Button-2>", paste_clipboard)  # For macOS (optional)

    # Clear button to clear the input box
    clear_button = ctk.CTkButton(
        qr_window,
        text="Clear",
        command=lambda: url_entry.delete(0, 'end'),
        fg_color="#ff4d4d",  # Optional: set a color for the button
        hover_color="#ff1a1a"  # Optional: set a hover color
    )
    clear_button.pack(pady=10, padx=10, fill="x")

    # Add QR code button
    def add_qr_code(event=None):
        global qr_window

        order_name_display = selected_order.get()
        order_name = display_name_to_order_name.get(order_name_display)
        url = url_entry.get()

        if not order_name:
            messagebox.showerror("Error", "Please select an order.")
            return
        if not url:
            messagebox.showerror("Error", "Please enter a valid URL.")
            return

        if order_name in qr_codes:
            # Optionally prompt the user whether to overwrite the existing QR code
            overwrite = messagebox.askyesno(
                "Overwrite QR Code",
                f"A QR Code already exists for {order_name}.\nDo you want to overwrite it?"
            )
            if not overwrite:
                return

        qr_codes[order_name] = url

        # After adding the QR code, close the window and reset focus to the main window
        qr_window.destroy()
        root.focus_set()
        qr_window = None

    add_button = ctk.CTkButton(
        qr_window,
        text="Add QR Code",
        command=add_qr_code,
        fg_color="#133d8e",
        hover_color="#266cc3"
    )
    add_button.pack(pady=20, padx=10, fill="x")

    # Bind the Enter key to the add_qr_code function
    qr_window.bind('<Return>', add_qr_code)

    # Ensure that when the window is closed, the reference is removed
    def on_close():
        global qr_window
        qr_window.grab_release()
        qr_window.destroy()
        qr_window = None
        root.focus_set()

    qr_window.protocol("WM_DELETE_WINDOW", on_close)

# GUI Setup
root = ctk.CTk()
root.title("Label Maker")
root.geometry("700x600")  # Increased the window width to accommodate both sides
root.configure(bg="#3A3A3A")

# Function to create and display a context menu with "Refresh" option
def show_context_menu(event):
    context_menu = Menu(root, tearoff=0)  # Create a context menu using tkinter's Menu
    context_menu.add_command(label="Refresh", command=display_order_history)  # Add "Refresh" option
    context_menu.tk_popup(event.x_root, event.y_root)  # Display the menu at the cursor's position

# Bind right-click event to the entire root window to show the context menu
root.bind("<Button-3>", show_context_menu)

root.iconbitmap(resource_path('resources/scribe-icon.ico'))

from PIL import Image
logo_image = Image.open(resource_path('resources/scribe-logo-final.webp'))
logo_image = logo_image.resize((258, 100), Image.Resampling.LANCZOS)

logo_ctk_image = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(258, 100))
logo_label = ctk.CTkLabel(root, image=logo_ctk_image, text="")
logo_label.pack(pady=10)

# Left side scrollable frame for main content
left_frame = ctk.CTkFrame(root, fg_color="#3A3A3A", corner_radius=15)
left_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)

# Create an inner frame to hold the canvas and leave space for the corner radius
inner_frame = ctk.CTkFrame(left_frame, fg_color="#3A3A3A", corner_radius=15)
inner_frame.pack(expand=True, fill="both", padx=10, pady=10)

canvas = ctk.CTkCanvas(inner_frame, bg="#3A3A3A", highlightthickness=0)
scrollbar = ctk.CTkScrollbar(inner_frame, orientation="vertical", command=canvas.yview)
scrollable_frame = ctk.CTkFrame(canvas, fg_color="#3A3A3A", corner_radius=15)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=450)
canvas.configure(yscrollcommand=scrollbar.set)

canvas.bind_all("<MouseWheel>", on_mousewheel)
canvas.bind_all("<Button-4>", on_mousewheel_mac)
canvas.bind_all("<Button-5>", on_mousewheel_mac)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Right side frame for order history
right_frame = ctk.CTkFrame(root, fg_color="#2E2E2E", corner_radius=15, width=250)
right_frame.pack(side="right", fill="y", padx=(5, 10), pady=10)

history_label_frame = ctk.CTkScrollableFrame(right_frame, width=230, fg_color="#2E2E2E", height=500)
history_label_frame.pack(pady=10, padx=10, fill="both", expand=True)

# Add widgets to left scrollable frame
instruction = ctk.CTkLabel(scrollable_frame, text="Select files to generate labels", font=("Helvetica", 16), text_color="white")
instruction.pack(pady=1, padx=20, expand=False)

# Add a new label for the file format below the instruction label
file_format_example = ctk.CTkLabel(scrollable_frame, text="Examples: Chris LaVigne T1 Copy 1 Envelopes-1-167.bin\n" + 
                                                                "Chris LaVigne T1 Copy 2 Letters-1-167.bin\n" +
                                                                    "Chris LaVigne T1 Envelopes-1-167.bin", 
                                 font=("Helvetica", 12), text_color="gray")
file_format_example.pack(pady=1, padx=20, expand=False)

envelope_button = ctk.CTkButton(scrollable_frame, text="Select Envelope Chip Files", command=select_envelope_files)
envelope_button.pack(pady=10, padx=20, fill="x", expand=True)

letter_button = ctk.CTkButton(scrollable_frame, text="Select Letter Chip Files", command=select_letter_files)
letter_button.pack(pady=10, padx=20, fill="x", expand=True)

add_qr_button = ctk.CTkButton(scrollable_frame, text="Add QR Code", command=add_qr_code_window, fg_color="#6c757d", hover_color="#adb5bd")
add_qr_button.pack(pady=10, padx=20, fill="x")

create_button = ctk.CTkButton(scrollable_frame, text="Create PDF", command=create_pdf, fg_color="#133d8e", hover_color="#266cc3")
create_button.pack(pady=10, padx=20, fill="x", expand=True)

reset_button = ctk.CTkButton(scrollable_frame, text="Reset", command=reset_data, width=100, fg_color="#8e1313", hover_color="#c32626")
reset_button.pack(pady=15, padx=20)

open_button = ctk.CTkButton(scrollable_frame, text="Open Created PDF File", width=300, fg_color="#133d8e", hover_color="#266cc3")
open_button.pack_forget()

envelope_label = ctk.CTkLabel(scrollable_frame, text="Selected Envelope Files:\n", font=("Helvetica", 12), text_color="white", anchor="w", justify="left")
envelope_label.pack(pady=10, padx=20, fill="x", side="top")

letter_label = ctk.CTkLabel(scrollable_frame, text="Selected Letter Files:\n", font=("Helvetica", 12), text_color="white", anchor="w", justify="left")
letter_label.pack(pady=10, padx=20, fill="x", side="top")

# Display the order history on the right side
display_order_history()

root.mainloop()