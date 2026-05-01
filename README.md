# Real-Debris STRM Maker

A powerful desktop GUI tool for Windows that turns Real-Debrid torrents, downloads, direct links, and magnet links into organized `.strm` files for Kodi, Plex-compatible libraries, or any media manager that supports streaming pointers.

Built with Python + Tkinter, this app goes beyond basic STRM creation by adding automated torrent monitoring, keyword-based auto-sorting, Real-Debrid integration, retry logic, system tray support, and hands-free scheduled downloads.

---

# Features

## Core STRM Creation

* Convert Real-Debrid torrents into `.strm` files
* Add magnet links directly to Real-Debrid
* Auto-select video files only from torrents
* Resolve Real-Debrid unrestricted links automatically
* Support for direct links and redirect resolution
* Smart filename cleaning + duplicate-safe naming
* Batch process multiple torrents at once

---

# Advanced Automation

* Scheduled automatic torrent scans
* Check every X hours OR daily at a set time
* Keyword-based download rules
* Auto-route matching torrents to specific folders
* Default fallback folder for unmatched torrents
* Tracks processed torrent IDs to avoid duplicate scans
* One-click “Run Keyword Scan Now”

---

# Smart Matching

Keyword rules normalize:

* Dots (`.`)
* Underscores (`_`)
* Hyphens (`-`)

Example:
`WWE.Raw.S35E14` → matches `WWE Raw`

Perfect for:

* TV shows
* Wrestling events
* Anime releases
* Sports archives
* Ongoing episodic content

---

# User Experience

* Clean Tkinter desktop GUI
* System tray minimize-to-tray support (via pystray)
* Local settings persistence
* Real-Debrid token saving
* Detailed scan reports
* Retry logic for unstable links/API hiccups
* Folder browser UI
* Scrollable magnet input
* Torrent list management

---

# Requirements

## Python Version

* Python 3.10+

## Required Packages

```bash
pip install pillow pystray
```

---

# Setup

## 1. Clone the repository

```bash
git clone https://github.com/yourusername/real-debris-strm-maker.git
cd real-debris-strm-maker
```

## 2. Install dependencies

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pillow pystray
```

## 3. Run the app

```bash
python strm_maker.py
```

---

# Real-Debrid API Token

You’ll need your personal Real-Debrid API token.

Inside the app:

1. Paste token into `Real-Debrid API Token`
2. Click `Save Token`

---

# How To Use

## Manual STRM Creation

### From Existing Torrents

* Load Recent Torrents
* Select one or more torrents
* Choose output folder
* Create `.strm` files

### From Magnet Links

* Paste one magnet per line
* Add Magnets to Real-Debrid
* Wait for processing
* Create STRM files

---

# Auto-Download Workflow

1. Open Auto-Download Settings
2. Set default auto-download folder
3. Add keyword rules
4. Enable scheduler
5. Choose interval or daily time
6. Leave running in tray

Example Rules:

```txt
WWE Raw → D:\Media\Wrestling\Raw
AEW Dynamite → D:\Media\Wrestling\AEW
One Piece → D:\Anime\One Piece
```

---

# Example Use Cases

## Kodi Library Automation

Automatically build `.strm` libraries for:

* WWE weekly shows
* Anime simulcasts
* TV episodes
* Sports replays

## Plex / Jellyfin Hybrid

Use `.strm` files as lightweight media placeholders.

## Archive Management

Organize Real-Debrid downloads without manually handling files.

---

# File Naming Logic

* Uses original filenames when available
* Cleans invalid Windows characters
* Prevents duplicate overwrites
* Falls back to indexed names

---

# Scheduler Details

* Polls every minute
* Supports:

  * Every N hours
  * Daily HH:MM
* Prevents duplicate same-day scans
* Background thread safe

---

# Settings Storage

Saved locally to:

```txt
~/.strm_maker_settings.json
```

Stores:

* Output folder
* RD token
* Auto settings
* Keyword rules
* Processed torrents

---

# Safety / Reliability

* Multiple write retries
* Multiple link retries
* API error parsing
* Download fallback logic
* Handles:

  * HTTP errors
  * Redirect failures
  * Missing links
  * Duplicate filenames

---

# Optional Tray Support

If `pystray` + `Pillow` are installed:

* Minimize to tray
* Close to tray
* Restore instantly
* Run scans from tray
* Quit from tray menu

---

# Known Limitations

* Windows-focused GUI
* Requires Real-Debrid account
* Tkinter UI styling is functional over flashy
* Torrent success depends on Real-Debrid availability

---

# Future Ideas

* TMDB metadata scraping
* Sonarr/Radarr integration
* Better episode parsing
* STRM folder templates
* Download history export
* Dark mode
* Portable EXE build

---

# Why This Exists

Real-Debrid is powerful, but managing torrents → links → `.strm` manually is tedious.

Real-Debris STRM Maker streamlines that entire pipeline into:
Magnet → Real-Debrid → Organized `.strm` Library

---

# License

MIT License (recommended — add LICENSE file)

---

# Disclaimer

This tool is not affiliated with Real-Debrid.
Users are responsible for complying with all platform terms and local laws.

---

# Contributing

Pull requests, improvements, and automation ideas are welcome.

---

# Quick Summary

**Best for users who want:**

* Automated Real-Debrid library building
* Kodi/Plex/Jellyfin `.strm` workflows
* Wrestling/anime/show auto-sorting
* Hands-off scheduled STRM generation

If you live in Real-Debrid and hate repetitive setup, this is built for exactly that.
