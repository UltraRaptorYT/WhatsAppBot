import time
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from playwright.sync_api import sync_playwright
import pyperclip
from urllib.parse import quote
from io import BytesIO
from PIL import Image
import win32clipboard
from platform import system
import pyautogui as pg

root = None

namelist_label_text = "Namelist Excel File"
message_label_text = "Message Text File"
image_label_text = "Image Files"
document_label_text = "Document File"
failed_numbers = []


def copy_image_to_clipboard(image_path):
    image = Image.open(image_path)

    # Save image to a bytes buffer
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    # Open clipboard and set the image
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()


def reset_paths():
    global namelist_path, message_path, image_paths, document_path
    namelist_path = None
    message_path = None
    image_paths = []
    document_path = None
    namelist_label.config(text=namelist_label_text)
    message_label.config(text=message_label_text)
    image_label.config(text=image_label_text)
    document_label.config(text=document_label_text)


def upload_image_file():
    global image_paths
    image_paths = filedialog.askopenfilenames(
        filetypes=(("Image Files", "*.png *.jpg *.jpeg *.gif *.bmp *.ico *.webp"),)
    )
    if image_paths:
        image_label.config(text=f"Selected {len(image_paths)} image(s)")
    else:
        image_label.config(text=image_label_text)


def upload_document_file():
    file_path = upload_file()
    if file_path:
        global document_path
        document_path = file_path
        document_label.config(text=f"Selected {document_label_text}: {file_path}")


def process_data():
    global failed_numbers
    failed_numbers = []
    global namelist_path
    global message_path
    global image_path
    global document_path

    global error_label
    error_label.config(text="", fg="red")
    if not namelist_path:
        error_label.config(text="Namelist File not selected")
        return
    try:
        f = open(namelist_path, "r")
        f.close()
    except FileNotFoundError:
        error_label.config(text="Namelist File not found")
        return

    if not message_path:
        error_label.config(text="Message File not selected")
        return
    try:
        f = open(message_path, "r")
        f.close()
    except FileNotFoundError:
        error_label.config(text="Message File not found")
        return

    df = pd.read_excel(namelist_path)
    root.attributes("-topmost", False)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, timeout=0)
        context = browser.new_context()
        login_page = context.new_page()
        login_page.goto("https://web.whatsapp.com")
        login_page.wait_for_selector(".xcgk4ki", timeout=0)  # Check for logged in

        for _, row in df.iterrows():
            name = str(row["Name"])
            phone_number = str(row["Mobile Number"])

            if "+" not in str(phone_number):
                phone_number = f"+65{phone_number}"

            message = ""
            with open(message_path, "r", encoding="utf-8") as f:
                message = f.read()
                f.close()

            message = message.replace("{name}", name)

            page = context.new_page()
            page.goto(
                f"https://web.whatsapp.com/send?phone={phone_number}&text={quote(message)}",
                wait_until="domcontentloaded",
            )
            page.wait_for_selector('[data-icon="menu"]', timeout=0)
            time.sleep(5)
            # Check if phone number is valid
            element = page.query_selector(
                '[aria-label="Phone number shared via url is invalid."]'
            )
            if element:
                failed_numbers.append(phone_number)
                print(phone_number, "Failed!")
                page.close()
                continue

            time.sleep(5)
            page.wait_for_selector('[data-icon="send"]', timeout=0)
            if document_path:
                time.sleep(2)
                page.wait_for_selector('[data-icon="plus"]', timeout=0)
                page.click('[data-icon="plus"]')
                page.wait_for_selector('.xuxw1ft:has-text("Document")', timeout=0)
                page.click('.xuxw1ft:has-text("Document")')
                time.sleep(2)
                pyperclip.copy(document_path)
                if system().lower() in ("windows", "linux"):
                    pg.hotkey("ctrl", "v")
                elif system().lower() == "darwin":
                    pg.hotkey("command", "v")
                else:
                    raise Warning(f"{system().lower()} not supported!")
                time.sleep(2)
                pg.press("enter")
                print("SENDING DOCUMENT")

            if image_paths:
                for path in image_paths:
                    copy_image_to_clipboard(path)
                    time.sleep(1)
                    page.press("[aria-activedescendant]", "ControlOrMeta+v")
                    print(f"SENDING IMAGE: {path}")
                    time.sleep(2)

            time.sleep(2)
            page.click('[data-icon="send"]')
            time.sleep(2)
            page.close()
        time.sleep(5)
        # Log Out
        # Click on Use here
        login_page.wait_for_selector(".x1v8p93f", timeout=0)
        login_page.click(".x1v8p93f")
        time.sleep(1)
        # Open the Menu Dropdown
        login_page.wait_for_selector('[data-icon="menu"]', timeout=0)
        login_page.click('[data-icon="menu"]')
        time.sleep(1)
        # Click on Log out btn
        login_page.wait_for_selector('[aria-label="Log out"]', timeout=0)
        login_page.click('[aria-label="Log out"]')
        time.sleep(1)
        # Confirm logout btn
        login_page.wait_for_selector(".x1v8p93f", timeout=0)
        login_page.click(".x1v8p93f")
        time.sleep(1)
        # Check Logout is confirmed
        login_page.wait_for_selector(
            '[aria-label="Scan this QR code to link a device!"]', timeout=0
        )
        time.sleep(5)
        browser.close()
    time.sleep(1)
    print(failed_numbers)
    print("DONE!")
    error_label.config(fg="green", text="DONE!")
    root.attributes("-topmost", True)
    return


def upload_file(file_type=()):
    file_types = file_type + (("All files", "*.*"),)
    file_path = filedialog.askopenfilename(filetypes=file_types)
    global error_label
    error_label.config(fg="red")
    error_label.config(text="")
    if not file_path:
        print("No file selected.")
        error_label.config(text="No file selected.")
        return

    try:
        with open(file_path, "r"):
            print(f"File {file_path} is accessible.")
    except IOError as e:
        print(f"Cannot access file {file_path}: {e}")
        error_label.config(text="No file selected.")
        return
    return file_path.replace("/", "\\")


def upload_namelist_file():
    file_path = upload_file(file_type=(("Excel Spreadsheet", "*.xlsx"),))
    if file_path:
        global namelist_path
        namelist_path = file_path
        namelist_label.config(text=f"Selected {namelist_label_text}: {file_path}")


def upload_message_file():
    file_path = upload_file(file_type=(("Text File", "*.txt"),))
    if file_path:
        global message_path
        message_path = file_path
        message_label.config(text=f"Selected {message_label_text}: {file_path}")


def main():
    global root

    root = tk.Tk()
    root.title("BWM Whatsapp Mass Sender")

    main_frame = tk.Frame(root)
    main_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    global namelist_label
    namelist_label = tk.Label(
        main_frame, text=namelist_label_text, wraplength=350, justify="left"
    )
    namelist_label.pack(pady=(10, 2.5))

    namelist_button = tk.Button(
        main_frame, text="Upload Namelist File", command=upload_namelist_file
    )
    namelist_button.pack(pady=(2.5, 10))

    global message_label
    message_label = tk.Label(
        main_frame, text=message_label_text, wraplength=350, justify="left"
    )
    message_label.pack(pady=(10, 2.5))

    message_button = tk.Button(
        main_frame, text="Upload Message File", command=upload_message_file
    )
    message_button.pack(pady=(2.5, 10))

    global image_label
    image_label = tk.Label(
        main_frame, text=image_label_text, wraplength=350, justify="left"
    )
    image_label.pack(pady=(10, 2.5))

    image_button = tk.Button(
        main_frame, text="Upload Image Files", command=upload_image_file
    )
    image_button.pack(pady=(2.5, 10))

    global document_label
    document_label = tk.Label(
        main_frame, text=document_label_text, wraplength=350, justify="left"
    )
    document_label.pack(pady=(10, 2.5))

    document_button = tk.Button(
        main_frame, text="Upload Document File", command=upload_document_file
    )
    document_button.pack(pady=(2.5, 10))

    button_frame = tk.Frame(main_frame)
    button_frame.pack(pady=10)

    process_button = tk.Button(button_frame, text="Send Message", command=process_data)
    process_button.pack(side="left", padx=5)

    reset_button = tk.Button(button_frame, text="Reset", command=reset_paths)
    reset_button.pack(side="left", padx=5)

    global error_label
    error_label = tk.Label(
        main_frame, text="", wraplength=350, justify="left", fg="red"
    )
    error_label.pack(pady=(2.5))
    # Instructional Frame
    instruction_frame = tk.Frame(root, bd=2, padx=10, relief="solid")
    instruction_frame.pack(side="right", fill="both", padx=10, pady=10)

    instruction = """
Instructions:

1. Upload the Namelist Excel file. [Must contain the following COLUMN NAME ("Name", "Mobile Number")]
2. Upload the Message Template file. [If want to be personalised, text should contain "{name}" so that it will be replaced by the "Name"]
3. Upload the Image file(s) [Less than 16MB]. [OPTIONAL]
4. Upload the Document file. [Less than 16MB] [OPTIONAL, Note using this mode will not allow for background running of bot]
5. Click 'Send Message' to send message to namelist.
6. Click 'Reset' to clear the selected files.
"""

    instruction_label = tk.Label(
        instruction_frame,
        text=instruction.strip(),
        justify="left",
        wraplength=300,
    )
    instruction_label.pack(padx=10, pady=10)

    # Set the window size
    window_width = 910
    window_height = 715

    # Get the screen dimension
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Find the center position
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)

    reset_paths()
    # Set the position of the window to the center of the screen
    root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
    root.attributes("-topmost", True)
    root.mainloop()


if __name__ == "__main__":
    main()
