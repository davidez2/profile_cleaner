    GUI tool: find stale user profiles on a remote PC, back them up with 7-Zip ON THAT PC, then delete them.

.DESCRIPTION
    - Scan: lists profiles via CIM (WinRM with DCOM fallback). Last activity = the newer of
      Win32_UserProfile.LastUseTime and the NTUSER.DAT last-write time (read over the C$ share).
    - Backup: 7-Zip runs ON THE REMOTE PC (PowerShell remoting / WinRM) against the local profile folder
      and writes the .7z archive to a folder on the remote PC. No profile data crosses the network.
    - A profile is deleted only if its archive was created AND passed "7z t".
    - Deletion uses Remove-CimInstance on Win32_UserProfile (removes folder + registry entry).
    - DRY RUN is ON by default: nothing is archived or deleted. It does check remoting, 7-Zip,
      the backup drive and estimates profile sizes against free space.

    Requirements: run elevated, as (or with credentials of) a local admin on the remote PC.
    WinRM (Enable-PSRemoting) must be enabled on the remote PC for the backup step.
    7-Zip must be installed on the remote PC, OR give a local 7za.exe (standalone "7-Zip Extra")
    / 7z.exe and the script copies it to the remote temp folder for the run and removes it afterwards.
