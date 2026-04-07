# XLSForm Quality Reviewer — Local Setup Guide

This tool runs entirely on your computer. No data is uploaded to any server.
Your XLSForm files are processed locally and never leave your machine.

---

## What You Need

Before starting, make sure you have received the following two files:

| File | Purpose |
|---|---|
| `app.py` | The application |
| `requirements.txt` | List of software dependencies |

Place both files in the same folder (e.g. a folder named `XLSForm Reviewer` on your Desktop).

---

## macOS Instructions

### Step 1 — Check Python is installed

Open **Terminal** (press `⌘ + Space`, type `Terminal`, press Enter) and run:

```bash
python3 --version
```

You should see something like `Python 3.10.x` or `Python 3.11.x`.

- **If Python is not found:** Download and install it from [https://www.python.org/downloads/](https://www.python.org/downloads/). Select the latest **3.11.x** release. Run the installer and follow all prompts.
- **If the version shown is 3.9 or earlier:** We recommend installing 3.11 from the link above. Both can coexist on your machine.

---

### Step 2 — Navigate to the app folder

In Terminal, type the following (replace the path if your folder is in a different location):

```bash
cd ~/Desktop/"XLSForm Reviewer"
```

To confirm you are in the right place, run `ls` — you should see `app.py` and `requirements.txt` listed.

---

### Step 3 — Create a virtual environment *(first time only)*

```bash
python3 -m venv venv
```

This creates an isolated environment so the app's dependencies do not affect any other software on your computer. This only needs to be done once.

---

### Step 4 — Activate the virtual environment

```bash
source venv/bin/activate
```

Your terminal prompt will change to show `(venv)` at the beginning. This means the environment is active.

---

### Step 5 — Install dependencies *(first time only)*

```bash
pip install -r requirements.txt
```

This may take 1–2 minutes. You will see a list of packages being downloaded and installed.

---

### Step 6 — Launch the app

```bash
streamlit run app.py
```

A browser window will open automatically at `http://localhost:8501`.
If it does not open, copy that address and paste it into any browser manually.

---

### Every subsequent time (returning users)

Only Steps 2, 4, and 6 are needed:

```bash
cd ~/Desktop/"XLSForm Reviewer"
source venv/bin/activate
streamlit run app.py
```

Or simply double-click **`launch_mac.command`** if it was included with your files.

---

### To stop the app

Press `Ctrl + C` in Terminal.

---

---

## Windows Instructions

### Step 1 — Check Python is installed

Open **Command Prompt** (press `Win + R`, type `cmd`, press Enter) and run:

```cmd
python --version
```

You should see something like `Python 3.10.x` or `Python 3.11.x`.

- **If Python is not found or the version is below 3.10:**
  1. Go to [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)
  2. Download the latest **3.11.x** Windows installer (choose the 64-bit version)
  3. Run the installer. **Important:** on the first screen, tick the box that says **"Add Python to PATH"** before clicking Install Now
  4. After installation, close and reopen Command Prompt and run `python --version` again to confirm

---

### Step 2 — Navigate to the app folder

In Command Prompt, type (replace the path if your folder is elsewhere):

```cmd
cd %USERPROFILE%\Desktop\XLSForm Reviewer
```

To confirm you are in the right place, run `dir` — you should see `app.py` and `requirements.txt` listed.

---

### Step 3 — Create a virtual environment *(first time only)*

```cmd
python -m venv venv
```

---

### Step 4 — Activate the virtual environment

```cmd
venv\Scripts\activate
```

Your prompt will change to show `(venv)` at the beginning.

> **If you see an error about script execution being disabled**, run the following command once in PowerShell (search "PowerShell" in the Start menu, right-click, select "Run as Administrator"):
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```
> Then close PowerShell and return to Command Prompt.

---

### Step 5 — Install dependencies *(first time only)*

```cmd
pip install -r requirements.txt
```

---

### Step 6 — Launch the app

```cmd
streamlit run app.py
```

A browser window will open automatically at `http://localhost:8501`.

---

### Every subsequent time (returning users)

```cmd
cd %USERPROFILE%\Desktop\XLSForm Reviewer
venv\Scripts\activate
streamlit run app.py
```

Or simply double-click **`launch_windows.bat`** if it was included with your files.

---

### To stop the app

Press `Ctrl + C` in Command Prompt.

---

---

## Linux Instructions

> These instructions apply to Ubuntu, Debian, and most Debian-based distributions.
> Adjust package manager commands (`apt`) for other distributions (e.g. `dnf` for Fedora).

### Step 1 — Check Python is installed

Open a terminal and run:

```bash
python3 --version
```

- **If Python 3.10 or higher is not installed**, run:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv -y
```

---

### Step 2 — Navigate to the app folder

```bash
cd ~/Desktop/"XLSForm Reviewer"
```

Confirm with `ls` — you should see `app.py` and `requirements.txt`.

---

### Step 3 — Create a virtual environment *(first time only)*

```bash
python3 -m venv venv
```

---

### Step 4 — Activate the virtual environment

```bash
source venv/bin/activate
```

Your prompt will change to show `(venv)`.

---

### Step 5 — Install dependencies *(first time only)*

```bash
pip install -r requirements.txt
```

---

### Step 6 — Launch the app

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.
If it does not open automatically, navigate to that address manually.

---

### Every subsequent time (returning users)

```bash
cd ~/Desktop/"XLSForm Reviewer"
source venv/bin/activate
streamlit run app.py
```

---

### To stop the app

Press `Ctrl + C` in the terminal.

---

---

## Troubleshooting

| Problem | Likely cause | Solution |
|---|---|---|
| `python3: command not found` | Python is not installed | Follow Step 1 for your operating system |
| `pip: command not found` | pip not on PATH | Use `python3 -m pip install -r requirements.txt` instead |
| `ModuleNotFoundError: No module named 'streamlit'` | Virtual environment not active | Run the activate command (Step 4) before launching |
| Browser does not open | Streamlit cannot detect browser | Navigate manually to `http://localhost:8501` |
| Port 8501 already in use | Another Streamlit instance is running | Stop it first with `Ctrl + C`, or run `streamlit run app.py --server.port 8502` |
| Windows: `activate` gives an execution policy error | PowerShell security setting | Follow the PowerShell note in Windows Step 4 |
| File upload fails with `openpyxl` error | Missing dependency | Ensure Step 5 completed without errors; re-run if needed |
| App shows 0 questions after upload | Wrong sheet structure | The file must have a `survey` sheet with `type` and `name` columns |

---

## Frequently Asked Questions

**Is my data safe?**
Yes. The application runs entirely on your local machine. No XLSForm data, results, or any other information is transmitted to any external server.

**Can I run this without an internet connection?**
Yes, once the dependencies are installed (Step 5), the app runs fully offline. An internet connection is only needed for the initial dependency installation.

**Can multiple people use the same installation?**
The app runs on a single local port. Multiple users on the same machine would need to share a single session. For independent use, each user should follow this guide to create their own installation.

**How do I update the app when a new version is provided?**
Replace `app.py` with the new file. Dependencies rarely change; if they do, a new `requirements.txt` will be provided. Re-run `pip install -r requirements.txt` with the virtual environment active to update.

**Where is the downloaded HTML report saved?**
The report is saved to whichever location your browser's default download folder points to.

---

*XLSForm Quality Reviewer · Standards: World Bank DIME ietestform & IPA ipacheckscto*
