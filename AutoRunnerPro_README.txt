Automated Task Runner (Pro)
Author: xTheRedShirtx

Purpose
Automates a repeatable 8 step desktop workflow with safe delays and a long timed wait. Sends OBS start and stop hotkeys. Provides coordinate capture, live logs, a watchdog, and an emergency stop.

Requirements
- Windows 10 or 11
- Python 3.9 or newer
- Packages: PyQt5, pyautogui, pygetwindow

Install
1) Open Command Prompt.
2) Run:
   python -m pip install --upgrade pip
   python -m pip install PyQt5 pyautogui pygetwindow

OBS Hotkeys (optional but supported)
1) Open OBS.
2) Settings > Hotkeys.
3) Set Start Streaming to Up Arrow.
4) Set Stop Streaming to Down Arrow.
5) Click Apply. Keep OBS open when you want hotkeys to work.

Launch
1) Save the script as main.py in a folder of your choice.
2) Open Command Prompt in that folder.
3) Run:
   python main.py

UI Overview
- Runner tab: set times and safety controls. Start or Stop.
- Coordinates tab: capture or edit click points for steps.
- Debug tab: live logs. Open or copy the log file.
- Help tab: quick reference.

First Run Checklist
1) Open the Coordinates tab.
2) For each step row click Pick. Move mouse to the target point. Press Left Ctrl to capture. Press ESC to cancel if needed.
3) Click Test Click to verify each point.
4) Click Save Coordinates.
5) Switch to the Runner tab.
6) Set Hours and Minutes for the long wait. Start with 0 hours 1 minute for testing.
7) Set Step delay to 2 to 5 seconds.
8) Leave Dry run checked for the first test.
9) Click Start. Confirm logs and timer update.
10) Uncheck Dry run and click Start again when ready for real clicks.

What the loop does
1) Click Step 1.
2) Click Step 2 then type today’s date in YYYY-MM-DD.
3) Click Step 3.
4) Wait for Step 4 wait seconds.
5) Click Step 5.
6) Wait for the main timer Hours and Minutes. Live countdown updates every second. The app auto clears OBS broadcast error dialogs if detected and restarts streaming via hotkeys.
7) Click Step 7.
8) Click Step 8.
9) Short pause then repeat until you stop.

Controls and Hotkeys
- Start begins the loop.
- Stop requests a clean stop.
- Delete triggers an emergency stop.
- Move mouse to the top left corner to trigger PyAutoGUI failsafe.
- Capture overlay: Left Ctrl captures. ESC cancels.

Settings
- Hours and Minutes control the long wait in Step 6.
- Step delay is the pause after each click.
- Retries is the number of times to retry a failed click.
- Watchdog stops the run if it becomes unresponsive.
- Step 4 wait adds a short pause after Step 3.
- Dry run logs actions without clicking.
- Always on top keeps the window visible.

Persistence
- Coordinates and settings are saved between runs using QSettings for the organization name and app name set in the code.

Logs
- File: automation_log.txt with rotation.
- Debug tab shows the live log.
- Menu: File > Open Log.

Troubleshooting
- Clicks do nothing: run Command Prompt as Administrator. Increase Step delay. Verify the target window is visible.
- Wrong click location: recapture coordinates after moving or resizing windows. Multi monitor layouts change absolute coordinates.
- Watchdog timeout: raise the Watchdog value. The app pauses the watchdog during intended waits.
- OBS not starting: confirm hotkeys in OBS. Keep OBS window open. Remove conflicting global hotkeys.
- Script exits early: check automation_log.txt for errors.

Safe Use Tips
- Disable screen savers or sleep mode during long waits.
- Close popups that could steal focus.
- Verify each point with Test Click before long unattended runs.
- Use Dry run to validate a new workflow.

Uninstall
- To remove packages run:
  python -m pip uninstall pyautogui pygetwindow PyQt5

Support
Owned and authored by xTheRedShirtx.
