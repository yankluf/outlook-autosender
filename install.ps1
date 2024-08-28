# OUTLOOK AUTOSENDER V1.0 - SIMPLE INSTALLER
# ******************************************

# Checks if Python is installed, and installs Python 3.12 if not
try {
    python --version
} catch {
    winget install -e -i --id=9ncvdn91xzqp --source=msstore
}

python -m venv env
.\env\Scripts\Activate.ps1
pip install -r .\requirements.txt

$EnvFileContent = @'
OUTLOOK_URL=https://outlook.office.com/mail/
EMAIL_ADDRESS=address@example.com
EMAIL_PASSWORD=yourpasswordhere
'@

New-Item .env
Set-Content .env $EnvFileContent

