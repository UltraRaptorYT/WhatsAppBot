# WhatsAppBot

Automatically sends WhatsApp Messages to people using Playwright in the background.

# Setup

Step 1: Install pipenv

```shell
pipenv install
```

Done!

# Run Program

Step 1: Type in command

```shell
pipenv run python sch.py
```

Step 2: Upload Namelist Excel

Step 3: Upload Message Text file

[Optional]: Upload Image/Document file

Step 4: Click on "Send Message" button

Step 5: Scan Whatsapp Popup

# Packaging

```shell
pyinstaller --onefile --add-data="C:\Users\sohho\AppData\Local\ms-playwright;ms-playwright" --icon="BWM.ico" main.py
```
