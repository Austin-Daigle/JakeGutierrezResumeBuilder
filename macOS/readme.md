This is the mac version of the program.

To exempt a specific .dmg file from macOS Gatekeeper's "damaged" quarantine warning (especially for one you created), use this Terminal command, replacing YourApp.dmg with your file's name:

text
xattr -d com.apple.quarantine YourApp.dmg
Run it from the directory containing the .dmg (use cd /path/to/directory first if needed). This removes the quarantine attribute Gatekeeper flags on unsigned or developer-built files downloaded or created outside official channels.

Afterward, double-click the .dmg to mount it normally. For apps inside, you may still need to right-click → Open or use System Settings → Privacy & Security → Open Anyway on first launch. Note that this doesn't disable Gatekeeper system-wide—it's per-file and safer than broader
