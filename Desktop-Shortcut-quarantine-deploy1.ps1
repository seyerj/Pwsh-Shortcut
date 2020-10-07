#################################
# Quarantine Icon Deployment script 1 v0.1
# 10/05/2020 Josh Seyer
# Douglas County, OR
################################

# Register log event source for Windows event log
New-EventLog -LogName 'Quarantine-Shortcut-Deploy' -Source 'DC_PS_Script'

# Set target for Shortcut
$TargetFile = "https://website.com/"

# Set location for shortcut.  Public Desktop makes it available for all users
$ShortcutFile = "$env:Public\Desktop\Quarantine_Email.lnk"

# Create the shortcut
$WScriptShell = New-Object -ComObject WScript.Shell
# icon location
$iconlocation = "\\share\path\folder\iconfile.ico"
$iconfile = "iconfile=" + $iconlocation
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.IconLocation = $iconlocation # icon index 0
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

# Set Read Only
Set-ItemProperty -Path $ShortcutFile -Name IsReadOnly -Value $true

# Log Outcome
if (Test-Path $ShortcutFile) 
{
    Write-EventLog -LogName 'Quarantine-Shortcut-Deploy' -Source 'DC_PS_Script' -EntryType 'Information' -EventID 1 -Message 'Quarantine shortcut creation attempt completed'    
}
else 
{
    Write-EventLog -LogName 'Quarantine-Shortcut-Deploy' -Source 'DC_PS_Script' -EntryType 'Information' -EventID 1 -Message 'Quarantine shortcut creation attempt failed, no shortcut created'    
}

# Remove log event source for Windows event log
Remove-EventLog -LogName 'Quarantine-Shortcut-Deploy'