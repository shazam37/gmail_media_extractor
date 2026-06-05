# Gmail Media Extractor

[![Build & Package](https://github.com/shazam37/gmail_media_extractor/actions/workflows/build.yml/badge.svg)](https://github.com/shazam37/gmail_media_extractor/actions/workflows/build.yml)

A desktop app that downloads all your Gmail attachments in one click — organised by sender, date, and file type. No technical knowledge required.

---

## Features

- **One-click sign-in** — authorises with Google in your browser; no passwords stored
- **Smart filtering** — choose file types (PDF, Images, Videos, Word, Excel, etc.) and an optional date range
- **Auto-organised output** — files are saved in folders by sender → date → category
- **Crash-safe** — if interrupted, re-running resumes from where it left off; nothing is re-downloaded
- **Google Sheets log** — optionally logs every file to a spreadsheet you own
- **OneDrive links** — also downloads files shared via OneDrive links inside emails
- **Headless / server mode** — run on a cloud VM overnight without keeping your laptop on

---

## Download

Go to the [**Actions tab**](https://github.com/shazam37/gmail_media_extractor/actions), open the latest successful **Build & Package** run, and download the zip for your OS from the **Artifacts** section at the bottom:

| Platform | Artifact name |
|----------|---------------|
| Windows  | `GmailMediaExtractor_Windows` |
| macOS    | `GmailMediaExtractor_macOS` |
| Linux    | `GmailMediaExtractor_Linux` |

---

## Installation

### Windows

1. Unzip `GmailMediaExtractor_Windows.zip`
2. Double-click `GmailMediaExtractor.exe`
3. If Windows shows **"Windows protected your PC"** → click **More info** → **Run anyway**
   *(This appears because the app is not commercially code-signed. It is safe.)*

### macOS

1. Unzip `GmailMediaExtractor_macOS.zip`
2. Double-click `GmailMediaExtractor.app`
3. If macOS says **"can't be opened because it is from an unidentified developer"**:
   → Right-click (or Ctrl-click) the app → **Open** → **Open**
   *(You only need to do this once.)*

### Linux

1. Unzip `GmailMediaExtractor_Linux.zip`
2. Make the file executable and run it:
   ```bash
   chmod +x GmailMediaExtractor
   ./GmailMediaExtractor
   ```

---

## Using the App

1. **Open the app** — a window appears with all options.
2. **Sign in to Google** — your browser opens once. Log in and approve the permissions. The app continues automatically.
3. **Set a date range** *(optional)* — leave blank to scan all emails.
4. **Choose file types** — tick the categories you want, or leave **Select All** checked.
5. **Choose an output folder** — or leave the default (`~/GmailMedia`).
6. **Click Start Extraction** — a progress bar shows how many emails have been scanned.
7. **When done**, click **Open Folder** to see your files.

### Output folder structure

```
GmailMedia/
├── sender@example.com/
│   ├── 2024-03-15/
│   │   ├── PDF Documents/
│   │   │   └── report.pdf
│   │   └── Images/
│   │       └── photo.jpg
│   └── 2024-04-01/
│       └── Spreadsheets/
│           └── budget.xlsx
└── another@example.com/
    └── ...
```

### Re-running the app

Running the app a second time is completely safe — it skips everything already downloaded and only fetches new attachments that arrived since the last run.

### Switching Google accounts

Click **Sign out / switch account** inside the app, then restart. Your browser will prompt you to sign in with any Google account.

---

## Headless Mode (Run on a Server / Cloud VM)

If you have a large inbox (thousands of emails) and don't want to keep your laptop running for hours, you can offload the job to a free cloud VM.

### Recommended: Oracle Cloud Free Tier

Oracle offers a **permanently free** Ubuntu VM — no credit card expiry, no time limit.

**One-time setup:**

1. Sign up at [cloud.oracle.com](https://cloud.oracle.com) → choose **Free Tier**
2. Create an instance: **Compute → Instances → Create Instance**
   - Image: Ubuntu 22.04
   - Shape: VM.Standard.E2.1.Micro (Always Free)
   - Download the SSH key when prompted
3. Note the VM's **Public IP address**

**Copy files to the VM** *(from your local machine):*

```bash
# Copy the extractor files
scp -i your_key.key \
  files/core.py files/run_headless.py \
  files/requirements_headless.txt files/credentials.json \
  ubuntu@YOUR_VM_IP:~/extractor/

# Copy your Google auth token (run the desktop app once first to create it)
#   Windows token: %APPDATA%\GmailMediaExtractor\token.json
#   Mac/Linux token: <app folder>/.app_data/token.json
scp -i your_key.key token.json \
  ubuntu@YOUR_VM_IP:~/.config/gmail_media_extractor/token.json
```

**Set up and run on the VM:**

```bash
ssh -i your_key.key ubuntu@YOUR_VM_IP

sudo apt update && sudo apt install -y python3 python3-pip
mkdir -p ~/.config/gmail_media_extractor
pip3 install -r ~/extractor/requirements_headless.txt

# Run in background — logs to run.log
nohup python3 ~/extractor/run_headless.py \
  --output-dir ~/GmailMedia \
  > ~/extractor/run.log 2>&1 &

# Watch progress from anywhere
tail -f ~/extractor/run.log
```

**Available options:**

```
--output-dir DIR        Where to save files (default: ~/GmailMedia)
--categories "A,B"      Comma-separated types to download (default: all)
--after  YYYY/MM/DD     Only emails after this date
--before YYYY/MM/DD     Only emails before this date
--no-sheets             Disable Google Sheets logging
--list-categories       Print all available category names and exit
```

**Example with filters:**

```bash
python3 ~/extractor/run_headless.py \
  --after 2023/01/01 \
  --categories "PDF Documents,Images,Spreadsheets" \
  --output-dir ~/GmailMedia
```

**Copy files back when done:**

```bash
scp -ri your_key.key ubuntu@YOUR_VM_IP:~/GmailMedia ./GmailMedia
```

If the run is interrupted for any reason, just re-run the same command — it resumes from where it stopped.

---

## Privacy & Security

- The app requests **read-only** Gmail access — it cannot send, delete, or modify emails
- Your Google sign-in token is stored **only on your own machine** (or the VM you control)
- No data is sent to any third-party server — only Google's own APIs are contacted
- The Google Sheets log (if enabled) is written to a spreadsheet **in your own Google account**

---

## Building from Source

**Prerequisites:** Python 3.11+, pip

```bash
git clone https://github.com/shazam37/gmail_media_extractor.git
cd gmail_media_extractor/files

pip install -r requirements.txt
python app.py          # run directly
```

To build a standalone binary:

```bash
# Place your credentials.json in the files/ folder first
python -m PyInstaller build.spec --noconfirm
# Output: dist/GmailMediaExtractor (or .exe / .app)
```

### GitHub Actions (automated builds)

The repository uses GitHub Actions to automatically build for all three platforms on every push to `main`.

**Required secret:**

| Secret name | Where to get it |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | Paste the full contents of your `credentials.json` file (from [Google Cloud Console](https://console.cloud.google.com) → APIs & Services → Credentials → OAuth 2.0 Client) |

To add the secret: **repo → Settings → Secrets and variables → Actions → New repository secret**

---

## License

MIT
