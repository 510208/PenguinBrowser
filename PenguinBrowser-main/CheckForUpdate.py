import json
import requests
import tkinter as tk
from tkinter import messagebox
import os.path
import subprocess
import webbrowser 

def check_for_updates():
    # Get the latest release information from GitHub
    url = "https://api.github.com/repos/510208/PenguinBrowser/releases/latest"
    response = requests.get(url)
    release_info = json.loads(response.text)

    # Compare the current version with the latest release
    current_version = "1.0.3"
    latest_version = release_info["tag_name"]
    if latest_version > current_version:
        msg_YNUpdate = messagebox.askyesno('Update Checker','有一個新版本可以更新，該版本版本號為'+latest_version+'，按下"確認"下載更新版本。')
        if msg_YNUpdate == 1:
            subprocess.run(["UpdateSite.bat"])
    else:
        messagebox.showinfo('Update Checker','感謝您安裝最新版本，您目前的版本為'+current_version+'，最新版本為'+latest_version+'按下確定以運行PenguinBrowser。此軟體為免費開源，如果您付費購買，請給賣家負評。')
        subprocess.run(["PGBrowser.exe"])

subprocess.run(["Dev.bat"])
check_for_updates()
