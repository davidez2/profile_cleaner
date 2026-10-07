Remove-StaleProfiles
A PowerShell GUI tool that connects to a remote Windows PC, finds user profiles that have not been used for a configurable number of days (default 90), backs each one up with 7-Zip on the remote PC, and then deletes the stale profiles.
Dry run mode is enabled by default. Nothing is archived or deleted until you untick it.
---
Features
WinForms GUI: scan, review, tick/untick profiles, run.
Staleness is based on the newer of two signals: `Win32_UserProfile.LastUseTime` and the `NTUSER.DAT` last-write time. `LastUseTime` alone is unreliable, so this errs on the side of keeping profiles.
7-Zip runs on the remote PC (PowerShell remoting). Profile data does not travel over the network, and archives are stored on the remote PC.
A profile is deleted only if its archive was created cleanly (7z exit code 0) and passed `7z t` (integrity test).
Deletion uses `Remove-CimInstance` on `Win32_UserProfile`, which removes both the folder and the registry entry.
Dry run checks the connection, 7-Zip, the backup drive and free space, and estimates profile sizes.
Cancel button: stops the job, kills the remote 7z process and removes the partial archive.
Optional alternate credentials.
Log file saved locally, with a copy placed in the backup folder on the remote PC.
Requirements
Item	Details
Local PC	Windows PowerShell 5.1 or later, run as Administrator
Account	Local administrator on the remote PC (current login or entered via Credentials...)
Remote PC	WinRM enabled (`Enable-PSRemoting -Force`) for the backup step. Scanning and deletion fall back to DCOM, but the backup does not.
7-Zip	Installed on the remote PC, or a local `7z.exe` / standalone `7za.exe` that the script copies to the remote temp folder and removes afterwards
Disk space	Enough free space on the remote backup drive for the archives
Admin share	`C$` access is only used to read `NTUSER.DAT` dates during the scan. If it is unreachable, those dates show as `-` and `LastUseTime` is used alone.
Running the script
If Windows blocks it with a "not digitally signed" error, use one of these:
```powershell
# Option 1: unblock the downloaded file
Unblock-File -Path .\Remove-StaleProfiles.ps1
.\Remove-StaleProfiles.ps1

# Option 2: bypass the policy for this window only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Remove-StaleProfiles.ps1

# Option 3: one-liner
powershell.exe -ExecutionPolicy Bypass -File .\Remove-StaleProfiles.ps1
```
If the policy is enforced by Group Policy, Options 2 and 3 are overridden. Ask your administrator or sign the script with `Set-AuthenticodeSignature`.
Usage
Enter the Remote computer name. Click Credentials... if you need a different admin account.
Set Idle more than (days).
Set 7-Zip path on remote PC (default `C:\Program Files\7-Zip\7z.exe`). Optionally set a Local 7za.exe to copy over if 7-Zip is missing there.
Set Backup folder ON REMOTE PC (a local path on the target, e.g. `C:\ProfileBackups` or `D:\ProfileBackups`).
Review the Never touch list of profile names.
Click 1. Scan. Stale profiles are highlighted and pre-ticked.
Adjust the ticks, then click 2. Run. Do a dry run first.
When the dry run looks right, untick DRY RUN and run again. A confirmation prompt appears before any deletion.
Grid statuses
Status	Meaning
`STALE`	Idle longer than the threshold; pre-selected
`Active`	Used within the threshold; not selected
`Loaded / in use - skipped`	A user is currently logged on; cannot be selected
`Excluded by name`	On the "Never touch" list; cannot be selected
`No date found - review manually`	No usable timestamp; not pre-selected
Safety behavior
Special/system profiles are ignored.
Loaded profiles are skipped, and this is re-checked immediately before each profile is processed.
The backup folder cannot be inside a profile selected for deletion.
If 7-Zip returns anything other than exit code 0 (including code 1, "warning", typically locked files), the profile is kept.
If the archive fails `7z t`, the profile is kept.
Archives are named `<Computer>_<User>_<yyyyMMdd-HHmmss>.7z`.
Backup details
Compression: `-mx=3` (fast), `-mmt=2` (limits CPU use on the user's PC), `-snl` (junctions/symlinks stored as links, not followed).
Skip temp / browser cache folders (on by default) excludes `AppData\Local\Temp`, `INetCache` and the Chrome/Edge `Cache` folders. This relies on 7-Zip path matching; inspect the first archive's contents, or untick the option for a full backup.
Backups are stored on the same PC being cleaned. Copy them to another location afterwards if you need them to survive a disk failure.
Logs
Live runs write `ProfileCleanup_<Computer>_<timestamp>.log` next to the script (or in `%TEMP%` if the script path is unknown).
A copy is placed in the backup folder on the remote PC at the end of the run.
The GUI log pane shows the same entries.
Troubleshooting
Problem	Fix
"Cannot open a PowerShell remoting session"	Run `Enable-PSRemoting -Force` on the target; check the firewall allows WinRM (TCP 5985); for non-domain PCs, add the target to `TrustedHosts`
"7-Zip not found ... and no valid local 7z/7za.exe"	Install 7-Zip on the target, correct the path, or browse to a local `7za.exe`
NTUSER.DAT dates show `-`	`C$` share not reachable; the scan still works using `LastUseTime`
"Backup NOT clean" with exit code 1	Some files were locked or unreadable. The profile is kept; retry later or investigate the 7z output in the log
"Deleted (leftovers)"	The profile registration was removed but some files remain in the folder; remove them manually
Script blocked as unsigned	See Running the script
Disclaimer
This tool permanently deletes user profiles. Test on a non-critical machine first, always run a dry run before a live run, and verify that backups are restorable before relying on them.
