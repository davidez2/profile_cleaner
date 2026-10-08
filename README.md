# Remove-StaleProfiles

A PowerShell GUI tool that connects to a remote Windows PC, finds user profiles that have not been used for a configurable number of days (default 90), backs each one up with 7-Zip **on the remote PC**, and then deletes the stale profiles.

**Dry run mode is enabled by default.** Nothing is archived or deleted until you untick it.

---

## Features

- WinForms GUI: scan, review, tick/untick profiles, run.
- Staleness basis is selectable (**Idle time based on**). Default is **LastUseTime only** (the `Win32_UserProfile.LastUseTime` value). If that value looks wrong on your PCs (Windows updates, AV and backup agents can reset it to "today"), switch to **User activity files**: the newest of (a) files in Desktop/Documents/Downloads, (b) the Recent items folder, (c) Chrome/Edge/Firefox history and Outlook OST files. These only change when a real person uses the profile **on that PC**. Hive timestamps (`LastUseTime`, `NTUSER.DAT`) are ignored by this mode because Windows updates, AV and backup agents that load profile hives reset them to "today". If no activity evidence exists at all, the profile folder's creation date is used. Other options: newest of all signals (safest; includes the home folder date), `NTUSER.DAT` only, and the home folder's last-modified date only.
- Domain `lastLogon` is intentionally not used: it is domain-wide, so a logon on another PC would hide a stale profile here.
- The grid shows every signal side by side (including the last-modified date of the profile/home folder itself) plus which one decided the result. The SID is not displayed. Evidence is collected on the remote PC over WinRM, so it doesn't depend on the `C$` share. **Export grid to CSV** saves the table for review.
- Only users that actually have a profile folder under `\Users\` are checked. Registry-only profile entries (no folder) are ignored and named in the log.
- 7-Zip runs **on the remote PC** (PowerShell remoting). Profile data does not travel over the network, and archives are stored on the remote PC.
- A profile is deleted only if its archive was created cleanly (7z exit code 0) **and** passed `7z t` (integrity test).
- Deletion uses `Remove-CimInstance` on `Win32_UserProfile`, which removes both the folder and the registry entry.
- The main table shows **Home size (MB)** (the whole profile folder) and **Backup size (MB, no AppData)** (what will actually be archived), measured on the remote PC. Both columns sort numerically. Junctions and symlinks are skipped so nothing is counted twice. Untick **Calculate home folder sizes** to make the scan faster.
- Dry run checks the connection, 7-Zip, the backup drive and free space, and compares the backup size (AppData excluded) with the free space.
- Cancel button: stops the job, kills the remote 7z process and removes the partial archive.
- Optional alternate credentials.
- Log file saved locally, with a copy placed in the backup folder on the remote PC.

## Requirements

| Item | Details |
|---|---|
| Local PC | Windows PowerShell 5.1 or later, run **as Administrator** |
| Account | Local administrator on the remote PC (current login or entered via **Credentials...**) |
| Remote PC | WinRM enabled (`Enable-PSRemoting -Force`) for the backup step. Scanning and deletion fall back to DCOM, but the backup does not. |
| 7-Zip | Installed on the remote PC, **or** a local `7z.exe` / standalone `7za.exe` that the script copies to the remote temp folder and removes afterwards |
| Disk space | Enough free space on the remote backup drive for the archives |
| Admin share | `C$` access is only a fallback for reading `NTUSER.DAT` dates if remote evidence collection over WinRM fails. |

## Running the script

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

## Usage

1. Enter the **Remote computer** name. Click **Credentials...** if you need a different admin account.
2. Set **Idle more than (days)**.
3. Set **7-Zip path on remote PC** (default `C:\Program Files\7-Zip\7z.exe`). Optionally set a **Local 7za.exe** to copy over if 7-Zip is missing there.
4. Set **Backup folder ON REMOTE PC** (a local path on the target, e.g. `C:\ProfileBackups` or `D:\ProfileBackups`).
5. Review the **Never touch** list of profile names.
6. Click **1. Scan**. Stale profiles are highlighted and pre-ticked.
7. Adjust the ticks, then click **2. Run**. Do a dry run first.
8. When the dry run looks right, untick **DRY RUN** and run again. A confirmation prompt appears before any deletion.

### Grid statuses

| Status | Meaning |
|---|---|
| `STALE` | Idle longer than the threshold; pre-selected |
| `Active` | Used within the threshold; not selected |
| `Loaded / in use - skipped` | A user is currently logged on; cannot be selected |
| `Excluded by name` | On the "Never touch" list; cannot be selected || `No date found - review manually` | No usable timestamp; not pre-selected |

## Safety behavior

- Special/system profiles are ignored, and only profiles whose path contains `\Users\` are listed.
- Loaded profiles are skipped, and this is **re-checked immediately before each profile is processed**.
- The backup folder cannot be inside a profile selected for deletion.
- If 7-Zip returns anything other than exit code 0 (including code 1, "warning", typically locked files), the profile is **kept**.
- If the archive fails `7z t`, the profile is **kept**.
- Archives are named `<Computer>_<User>_<yyyyMMdd-HHmmss>.7z`.

## Backup details

- Compression: `-mx=3` (fast), `-mmt=2` (limits CPU use on the user's PC), `-snl` (junctions/symlinks stored as links, not followed).
- **The `AppData` folder is never backed up.** It is excluded from every archive, but it is still deleted along with the rest of the profile, so browser profiles, Outlook settings, app configuration and similar data are lost. The confirmation prompt for a live run says so. This relies on 7-Zip path matching; inspect the first archive's contents before trusting it.
- Backups are stored on the same PC being cleaned. Copy them to another location afterwards if you need them to survive a disk failure.

## Logs

- Live runs write `ProfileCleanup_<Computer>_<timestamp>.log` next to the script (or in `%TEMP%` if the script path is unknown).
- A copy is placed in the backup folder on the remote PC at the end of the run.
- The GUI log pane shows the same entries.

## Troubleshooting

| Problem | Fix |
|---|---|
| "Cannot open a PowerShell remoting session" | Run `Enable-PSRemoting -Force` on the target; check the firewall allows WinRM (TCP 5985); for non-domain PCs, add the target to `TrustedHosts` |
| "7-Zip not found ... and no valid local 7z/7za.exe" | Install 7-Zip on the target, correct the path, or browse to a local `7za.exe` |
| Many profiles show today's date in `LastUseTime` / `NTUSER.DAT` | Expected: Windows, AV or backup agents reset them. Switch to the **User activity files** basis, which ignores them |
| Many profiles show "No date found" | Evidence couldn't be read. Make sure WinRM works (it is collected remotely); the scan falls back to the `C$` share (NTUSER.DAT only) otherwise. Use **Export grid to CSV** to inspect the signals |
| A user from another report is missing from the grid | Only profiles with a folder under `\Users\` are checked. Registry-only entries (no folder) are skipped and named in the log |
| NTUSER.DAT dates show `-` | Remoting and `C$` both failed for that profile, or the file is missing |
| "Backup NOT clean" with exit code 1 | Some files were locked or unreadable. The profile is kept; retry later or investigate the 7z output in the log |
| "Deleted (leftovers)" | The profile registration was removed but some files remain in the folder; remove them manually |
| Script blocked as unsigned | See [Running the script](#running-the-script) |

## Disclaimer

This tool permanently deletes user profiles. Test on a non-critical machine first, always run a dry run before a live run, and verify that backups are restorable before relying on them.
