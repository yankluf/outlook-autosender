
# Outlook AutoSender

This Python script automates sending emails from Outlook using a list of email addresses in an Excel file. It uses Selenium for automation.

This version (1.0) is designed for personal use, so it may not be entirely ready-to-use for all scenarios. I didn't take the time to consider every possible situation.

Feel free to modify the code to suit your needs!

## Configuration Required

Create the environment with `python -m venv env` and install dependencies using `pip install -r requirements.txt`.

You'll also need to create a `.env` file to store your email credentials in the variables `EMAIL_ACCOUNT` and `EMAIL_PASSWORD`. You may also need to change the URL for your Outlook access or comment out the lines related to 2FA if you don't use it.

In any case, I strongly recommend reviewing the code to see what else you might need to customize before using the tool.