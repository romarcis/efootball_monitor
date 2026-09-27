# eFootball Wi-Fi Monitor

**English** | [Italiano](README.it.md)

A Python script for Windows that monitors your connection while you play eFootball over Wi-Fi.
At the end of each session it saves everything to a **color-coded Excel file**. With `--sniff` it also beeps before kick-off to tell you whether the match connection will be good or bad.

> The program's messages and Excel column names are currently in Italian.

## What it measures

| Item | What it tells you |
|---|---|
| **router** (ping, jitter, packet loss) | Wi-Fi quality between PC and router: should stay under 5 ms, with low jitter and 0% loss |
| **internet** (default 1.1.1.1) | quality of your ISP line |
| **Wi-Fi** (signal %, dBm, band, channel) | signal strength, with warnings if you are on 2.4 GHz or the PC switches access point |
| **match** (only with `--sniff`) | server or opponent, country, ping, packets per second and "stutters" (gaps over 150 ms in incoming data = lag) |

Green = good, yellow = borderline, red = problem.

## Installation (once)

1. Install Python 3 from https://www.python.org/downloads/ and tick **"Add python.exe to PATH"**.
2. Copy `efootball_monitor.py` into a folder, for example `C:\efootball-monitor`.
3. In the Command Prompt run `pip install openpyxl scapy`.
4. Install **Npcap** from https://npcap.com with the default options (only needed for `--sniff`).
5. Keep **Windows location turned on** (Settings > Privacy & security > Location). Without it, Windows blocks reading the Wi-Fi signal.

## .exe version (for people without Python)

You can turn the script into a single program, `eFootballMonitor.exe`, to share with people who don't have Python.

1. On **your** PC, where Python is installed, put `build_exe.bat` in the same folder as the script and **double-click** it.
2. After a couple of minutes the program is in `dist\eFootballMonitor.exe`.
3. Whoever uses it only needs to install **Npcap** and keep Windows location on.

When you double-click the exe, Windows asks for administrator rights and match analysis (`--sniff`) starts automatically. Excel files are saved next to the exe. When you're done, press Ctrl+C and then Enter to close.

The exe is not signed, so on first launch Windows may show "Windows protected your PC": click **More info > Run anyway**. Some antivirus programs may flag it by mistake, which is common for programs built this way.

## How to use it

Open the **Command Prompt as administrator**, go to the folder and start the script:

```
cd C:\efootball-monitor
python efootball_monitor.py --sniff
```

Leave it open while you play: it prints a line every 5 seconds. When you're done, go back to the window and press **Ctrl+C**.

Without `--sniff` (and without administrator rights) it still works, but only measures Wi-Fi, router and internet, not the match.

## Pre-match alert (with `--sniff`)

As soon as the game connects to the match server, within about 1 second the PC beeps and prints a line like:

`>>> PARTITA IN ARRIVO (0.1 s)  server dedicato  34.154.0.13:5735  Milano, Italia  ping 8 ms -> connessione BUONA`

- **1 high beep** = good connection
- **2 beeps** = so-so, or not measurable
- **3 low beeps** = bad
- **+ 1 long beep** = the opponent (or the server) is in **another country**

**Dedicated server** means the match goes through a server (for example Google Cloud in Milan). In that case the opponent's IP is not visible, so the country shown is the server's.
**P2P** means you are connected directly to the opponent, so the country is theirs.

The country comes from the free ip-api.com service. Discord or other voice calls can trigger false alerts. Add `--muto` to turn the sound off.

## Excel files

When you press **Ctrl+C** the script creates two things in its folder.

**1. The session file**, for example `efootball_log_20260927_214045.xlsx`, with three sheets:
- **Riepilogo** (summary): what percentage of the time the connection was OK, ATTENZIONE (warning) or PROBLEMA (problem), session averages and the most frequent problems.
- **Partite** (matches): matches found in this session, with server, country and ping.
- **Andamento** (timeline): one row every 5 seconds. The **Giudizio** (verdict) and **Cosa non va** (what's wrong) columns explain in words what was happening. Every number is color-coded and filters are already on: filter Giudizio to `PROBLEMA` to jump to the worst moments.

**2. The history file `efootball_match_history.xlsx`**, a single file to which every session appends its matches. Use it to see over time which servers and countries you get and how the connection behaves. Keep it **closed** when you press Ctrl+C, otherwise it can't be updated.


**If you close the window instead of pressing Ctrl+C**, no Excel file is created. A backup `.csv` with the same name stays in the folder, and you can convert it:

```
python efootball_monitor.py --excel efootball_log_20260927_214045.csv
```

## Reading the results

- **Router with jitter or packet loss**: the problem is your home Wi-Fi. Move closer to the router, use 5 GHz or a cable.
- **Router fine but internet bad**: the problem is your line or ISP.
- **Everything fine but many stutters in the match**: the server or the opponent is far away or unstable. Not your fault.
- **All green but your inputs feel delayed**: usually a distant opponent, because the game slows both players down to keep them in sync. It can also be your PC: try turning off V-Sync, using a wired controller and putting the screen in "game mode".
- **Server "no ping"**: many game servers don't answer ping, that's normal. Look at the stutters instead.

## Options

| Option | What it does |
|---|---|
| `--sniff` | analyses the match and gives the pre-kick-off alert (needs administrator) |
| `--muto` | no sound when a match is found |
| `--window 2` | one line every 2 seconds instead of 5 |
| `--gateway 192.168.1.1` | set the router IP yourself if it isn't detected |
| `--internet 8.8.8.8` | ping a different internet host |
| `--excel FILE.csv` | create the Excel file from a leftover backup `.csv` |

## How it works

Inspired by [eFootball_Network_Monitor_Tool](https://github.com/SuNingXJBT/eFootball_Network_Monitor_Tool). Pings use the Windows ICMP API, Wi-Fi data comes from `netsh wlan show interfaces`, and match traffic is captured with Npcap + scapy: the remote host sending a steady UDP stream is treated as the match server or opponent.
