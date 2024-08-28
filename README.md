
# Outlook AutoSender

This Python script automates sending emails from Outlook using a list of email addresses in an Excel file. It uses Selenium for web browser automation.

The current version (1.x) is designed for personal use, so it may not be entirely ready-to-use for all scenarios. I didn't take the time to consider every possible situation, so feel free to modify the code to suit your needs!

## Installation (with Windows script)

Open a Terminal (PowerShell) in the root folder and execute the installer with `.\install.ps1`. This will create the environment with `python -m venv env` and install all dependencies using `pip install -r requirements.txt`. Also, it will install Pyhon 3.12 if it's not installed.

## Configuration Required

You'll need to edit the `.env` file and type your email credentials. You can change the URL where you access your web version of Outlook too (may be different depending on your organization).

In any case, I strongly recommend reviewing the code to see if you need to customize something else before using the tool.

## First Use (IMPORTANT!)

Since this tool isn't designed for every scenario, you'll need to check what your _login process_ looks like. This usually depends on your organization if you're using Outlook in a business environment.

For this reason, you'll be asked several questions when launching the tool. To find the best fit for you, select "yes" (y) the first time you use it, to see how your _login process_ is. You can then change to "no" (n) the answers you need, until you're able to access your inbox with the tool.

If the options provided aren't enough to automate your login process, you might need to modify the code yourself or wait for a newer version of the tool.