import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, colorchooser, Menu
import logging
from PIL import Image, ImageDraw, ImageFont
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
            
            # Ignore hyphens in base_name
            base_name = base_name.replace("-", "")

            # Ensure the files match the correct chip_type (Envelopes or Letters)
            if chip_type == "Envelopes" and "Letters" in base_name:
                invalid_files.append(file_name)  # Track invalid files
                continue  # Skip this file as it's not valid for Envelopes
            elif chip_type == "Letters" and "Envelopes" in base_name:
                invalid_files.append(file_name)  # Track invalid files
                continue  # Skip this file as it's not valid for Letters

            # Extract the order_name and card_envelope type
            if "Envelopes" in base_name:
                order_name = base_name.split("Envelopes")[0].strip()  # Get everything before "Envelopes"
                card_envelope = "Envelope"
            elif "Letters" in base_name:
                order_name = base_name.split("Letters")[0].strip()  # Get everything before "Letters"
                card_envelope = "Card"

            # Check if this label has already been generated
            if (order_name, card_envelope) in existing_labels:
                continue  # Skip if this order_name and card_envelope already exist

            # The rest of the logic remains unchanged
            order_type_count[(order_name, card_envelope)].append(file_name)

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

    # Assign batch numbers for each unique order_name and type
    for (order_name, card_envelope), file_list in order_type_count.items():
        for i, file_name in enumerate(file_list):
            batch_chip = f"{i + 1} of {len(file_list)}"

            # Append valid files to labels_data if not already present
            if (order_name, card_envelope) not in existing_labels:
                labels_data.append({
                    "order_name": order_name,
                    "batch_chip": batch_chip,
                    "card_envelope": card_envelope
                })

                # Add valid files to the appropriate set for preventing duplicates
                if chip_type == "Envelopes" and file_name not in displayed_envelope_files:
                    valid_files.append(file_name)
                    displayed_envelope_files.add(file_name)
                elif chip_type == "Letters" and file_name not in displayed_letter_files:
                    valid_files.append(file_name)
                    displayed_letter_files.add(file_name)

    # Display only valid and non-duplicate files in the appropriate label
    if valid_files:
        file_names = "\n".join(valid_files)
        if chip_type == "Envelopes":
            current_text = envelope_label.cget("text")
            envelope_label.configure(text=current_text + file_names + "\n")
        elif chip_type == "Letters":
            current_text = letter_label.cget("text")
            letter_label.configure(text=current_text + file_names + "\n")

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

    :param labels_data: List of dictionaries with label data (order_name, batch_chip, card_envelope, color).
    :param qr_codes: Dictionary mapping order_name to its QR code URL.
    :param output_pdf: Output file name for the generated PDF.
    """
    from PIL import Image, ImageDraw, ImageFont
    import qrcode
    import os

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

    font_large = ImageFont.truetype(font_path, 70)
    font_medium = ImageFont.truetype(font_path, 50)

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

            # Add label text
            label = labels_data[label_index]
            order_name = label["order_name"]
            batch_chip = label["batch_chip"]
            card_envelope = label["card_envelope"]
            text_color = label.get('color', "black")

            # Ensure text_color is in a format PIL can use
            if isinstance(text_color, str):
                if not text_color.startswith('#') and len(text_color) == 6:
                    text_color = '#' + text_color  # Add '#' if missing
                # PIL can handle color names and hex strings with '#'

            # Draw the text onto the label
            draw.text((x_start + 20, y_start + 20), f"Order Name & Number:", fill='black', font=font_medium)
            draw.text((x_start + 20, y_start + 100), order_name, fill=text_color, font=font_large)
            draw.text((x_start + 20, y_start + 200), "Batch / Chip Number:", fill='black', font=font_medium)
            draw.text((x_start + 20, y_start + 265), batch_chip, fill=text_color, font=font_large)
            draw.text((x_start + 20, y_start + 360), "Type:", fill='black', font=font_medium)
            draw.text((x_start + 180, y_start + 350), card_envelope, fill=text_color, font=font_large)

            # Add QR code if it exists for the order
            if order_name in qr_codes:
                qr = qrcode.QRCode()
                qr.add_data(qr_codes[order_name])
                qr.make(fit=True)

                # Generate QR code image
                qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
                qr_size = 200  # QR code size
                qr_img = qr_img.resize((qr_size, qr_size))
                qr_position = (x_end - qr_size - 20, y_start + 230)  # Position near top-right of label
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

def add_qr_code_window():
    global qr_codes

    if not labels_data:
        messagebox.showerror("Error", "No order label to add QR code.")
        return

    # Create a new pop-up window
    qr_window = ctk.CTkToplevel(root)
    qr_window.title("Add QR Code")
    qr_window.geometry("400x300")
    qr_window.configure(bg="#3A3A3A")

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
    for label in labels_data:
        order_name = label['order_name']
        if order_name not in seen_order_names:
            seen_order_names.add(order_name)
            order_names.append(order_name)

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

    url_entry = ctk.CTkEntry(qr_window, placeholder_text="Enter URL here...")
    url_entry.pack(pady=10, padx=10, fill="x")

    # Add QR code button
    def add_qr_code():
        order_name = selected_order.get()
        url = url_entry.get()

        if not order_name:
            messagebox.showerror("Error", "Please select an order.")
            return
        if not url:
            messagebox.showerror("Error", "Please enter a valid URL.")
            return

        qr_codes[order_name] = url
        messagebox.showinfo("Success", f"QR Code added for order: {order_name}")
        qr_window.destroy()

    add_button = ctk.CTkButton(
        qr_window,
        text="Add QR Code",
        command=add_qr_code,
        fg_color="#133d8e",
        hover_color="#266cc3"
    )
    add_button.pack(pady=20, padx=10, fill="x")

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

icon_path = resource_path('resources/scribe-icon.ico')
root.iconbitmap(icon_path)

logo_path = resource_path('resources/scribe-logo-final.png')

from PIL import Image
logo_image = Image.open(logo_path)
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