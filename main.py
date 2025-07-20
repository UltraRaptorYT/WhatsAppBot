import os
import time
import tkinter as tk
from datetime import datetime
from io import BytesIO
from tkinter import filedialog
from urllib.parse import quote

import pandas as pd
import requests
import tempfile
import win32clipboard
from PIL import Image
from playwright.sync_api import sync_playwright

root = None

namelist_label_text = "Namelist Excel File"
message_label_text = "Message Text File"
image_label_text = "Image Files"
document_label_text = "Document File"
failed_numbers = []


def ensure_log_file():
    if not os.path.exists("log.txt"):
        with open("log.txt", "w", encoding="utf-8") as f:
            print("[INFO] Log file created.")
            f.write("[INFO] Log file created.\n")


def logger(message, logType="INFO", log_file="log.txt"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_message = f"[{timestamp}] [{logType}] : {message}"
    print(formatted_message)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(formatted_message + "\n")


def copy_image_to_clipboard(image_path):
    try:
        image = Image.open(image_path)
    except Exception as e:
        logger(f"Failed to open image at {image_path}: {e}", "ERROR")
        raise ValueError(f"Failed to open image at {image_path}: {e}")

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
    logger("Processing Data...")
    global failed_numbers
    failed_numbers = []
    global namelist_path
    global message_path
    global image_path
    global document_path

    global error_label
    error_label.config(text="", fg="red")
    if not namelist_path:
        logger("Namelist File not selected", "ERROR")
        error_label.config(text="Namelist File not selected")
        return
    try:
        f = open(namelist_path, "r")
        f.close()
    except FileNotFoundError:
        logger("Namelist File not found", "ERROR")
        error_label.config(text="Namelist File not found")
        return

    if not message_path:
        logger("Message File not selected", "ERROR")
        error_label.config(text="Message File not selected")
        return
    try:
        f = open(message_path, "r")
        f.close()
    except FileNotFoundError:
        logger("Message File not found", "ERROR")
        error_label.config(text="Message File not found")
        return

    df = pd.read_excel(namelist_path, dtype={"Mobile Number": str})

    if "Mobile Number" not in df.columns:
        logger("Missing 'Mobile Number' column in excel", "ERROR")
        error_label.config(text="Missing 'Mobile Number' column in excel")
        return

    root.attributes("-topmost", False)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, timeout=0)
        context = browser.new_context()
        login_page = context.new_page()
        login_page.goto("https://web.whatsapp.com", timeout=0)
        login_page.wait_for_selector("header", timeout=0)  # Check for logged in

        buttons = login_page.query_selector_all("button")

        for button in buttons:
            if button.inner_text().strip() == "Continue":  # Unable To Change To Chinese
                button.click()
                break
        else:
            print("No 'Continue' button found.")

        with open(message_path, "r", encoding="utf-8") as f:
            base_message = f.read()
            f.close()

        for _, row in df.iterrows():
            phone_number = str(row["Mobile Number"]).strip()

            if not phone_number:
                continue

            if "+" not in str(phone_number):
                phone_number = f"+65{phone_number}"

            row_image_paths = []
            if "ImageURL" in row and pd.notna(row["ImageURL"]):
                img_path = str(row["ImageURL"]).strip()
                try:
                    if img_path.startswith("http://") or img_path.startswith(
                        "https://"
                    ):
                        response = requests.get(img_path)
                        response.raise_for_status()
                        with tempfile.NamedTemporaryFile(
                            delete=False, suffix=".png"
                        ) as tmp_img:
                            tmp_img.write(response.content)
                            tmp_img_path = tmp_img.name
                        row_image_paths.append(tmp_img_path)
                        logger(
                            f"Downloaded image from URL for {phone_number}: {img_path}"
                        )
                    else:
                        Image.open(img_path)  # check validity
                        row_image_paths.append(img_path)
                        logger(
                            f"Appended local image path for {phone_number}: {img_path}"
                        )
                except Exception as e:
                    logger(
                        f"Failed to load image for {phone_number}: {img_path} | Error: {e}",
                        "ERROR",
                    )

            message = base_message[:]

            for col in df.columns:
                placeholder = f"{{{col}}}"
                value = str(row[col])
                message = message.replace(placeholder, value)

            page = context.new_page()
            page.goto(
                f"https://web.whatsapp.com/send?phone={phone_number}&text={quote(message)}",
                wait_until="domcontentloaded",
            )
            page.wait_for_selector("header", timeout=0)
            time.sleep(1)
            # Check if phone number is valid
            element = page.query_selector(
                '[aria-label="Phone number shared via url is invalid."]'
            )
            if element:
                failed_numbers.append(phone_number)
                logger(f"{phone_number} Failed!", "ERROR")
                page.close()
                continue

            time.sleep(1)
            page.wait_for_selector(
                'button[data-tab="11"] > span', timeout=0
            )  # Unable To Change To Chinese

            if document_path:
                logger(f"Uploading document: {document_path}")
                with page.expect_file_chooser() as fc_info:
                    page.click('[button[data-tab="10"]]')
                    page.click(
                        '.xuxw1ft:has-text("Document")'
                    )  # Unable To Change To Chinese
                file_chooser = fc_info.value
                file_chooser.set_files(document_path)
                logger(f"SENDING DOCUMENT to {phone_number}")
                time.sleep(1)

            row_image_paths += image_paths

            if row_image_paths:
                for path in row_image_paths:
                    copy_image_to_clipboard(path)
                    time.sleep(1)
                    page.press("[aria-activedescendant]", "ControlOrMeta+v")
                    logger(f"SENDING IMAGE to {phone_number}: {path}")
                    time.sleep(1)

            time.sleep(1)
            page.click(
                'button[data-tab="11"] > span', force=True
            )  # Unable To Change To Chinese
            time.sleep(1)
            max_wait = 20
            for _ in range(max_wait):
                last_message = page.locator('[data-tab="8"] div.message-out').last
                if last_message:
                    check = last_message.locator('[data-icon="msg-check"]')
                    dbl_check = last_message.locator('[data-icon="msg-dblcheck"]')
                    if dbl_check.count() > 0 or check.count() > 0:
                        logger(f"{phone_number} Message sent and delivery confirmed")
                        break
                time.sleep(1)
            else:
                logger(
                    f"{phone_number} Message may not be delivered (no msg-check icon found).",
                    logType="ERROR",
                )

            page.close()
        time.sleep(5)
        # Log Out
        # Click on Use here
        login_page.wait_for_selector(
            '.x1v8p93f:has-text("Use here")', timeout=0
        )  # Unable To Change To Chinese
        login_page.click(
            '.x1v8p93f:has-text("Use here")'
        )  # Unable To Change To Chinese
        time.sleep(1)
        # Open the Menu Dropdown
        login_page.wait_for_selector(
            '[title="Menu"]', timeout=0
        )  # Unable To Change To Chinese
        login_page.click('[title="Menu"]')  # Unable To Change To Chinese
        time.sleep(1)
        # Click on Log out btn
        login_page.wait_for_selector(
            'text="Log out"', timeout=0
        )  # Unable To Change To Chinese
        login_page.click('text="Log out"')  # Unable To Change To Chinese
        time.sleep(1)
        # Confirm logout btn
        login_page.wait_for_selector(
            '.x1v8p93f:has-text("Log out")', timeout=0
        )  # Unable To Change To Chinese
        login_page.click('.x1v8p93f:has-text("Log out")')  # Unable To Change To Chinese
        time.sleep(1)
        # Check Logout is confirmed
        login_page.wait_for_selector(
            '[aria-label="Scan this QR code to link a device!"]', timeout=0
        )  # Unable To Change To Chinese
        time.sleep(1)
        browser.close()
    time.sleep(1)
    logger(
        f"The following numbers has failed to send {failed_numbers}",
        logType="ERROR",
    )
    logger("Process COMPLETED.")
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
        logger("No file selected.", "ERROR")
        error_label.config(text="No file selected.")
        return

    try:
        with open(file_path, "r"):
            logger(f"File {file_path} is accessible.")
    except IOError as e:
        logger(f"Cannot access file {file_path}: {e}", "ERROR")
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
    ensure_log_file()
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

1. Upload the Namelist Excel file. [Must contain the following COLUMN NAME ("Mobile Number")]
2. Upload the Message Template file. [If want to be personalised, text should contain "{Name}" so that it will be replaced by the "Name"]
3. Upload the Image file(s) [Less than 16MB]. [OPTIONAL]
4. Upload the Document file. [Less than 16MB] [OPTIONAL]
5. Click 'Send Message' to send message to namelist.
6. Click 'Reset' to clear the selected files.

NOTE:
CUSTOM IMAGES requires the PATH OF IMAGE and "ImageURL" column on excel
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
