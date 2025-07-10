# Base UNC path (configurable)
$baseTarget = "\\BOSCH.COM\DfsRB\DfsCZ\LOC\Bj\DNOX_LAB\60_Transfer\INTERNAL\Carda_Zdenek"
if ($env:USERNAME -eq "mrkure"){
    $baseTarget = "C:\"
}
# Get the script's own file name (without extension)
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)

# Assume script name format: rcopy_foldername_lastdiroftargetfolder.ps1
$parts = $scriptName -split "-"

if ($parts.Length -ne 3 -or $parts[0] -ne "rcopy") {
    Write-Error "Script name must follow pattern: rcopy_foldername_lastdiroftargetfolder.ps1"
    exit 1
}

# Extract folder and target subdirectory
$folderName = $parts[1] -replace '-', ' '  # optional: turn underscores into spaces
$lastDir = $parts[2]

# Construct source and destination paths
$source = Join-Path (Get-Location) $folderName
$target = Join-Path $baseTarget $lastDir

# Check source folder exists
if (-not (Test-Path -Path $source)) {
    Write-Error "Source folder '$folderName' does not exist in current directory."
    exit 1
}
# Ensure the target folder exists
if (-not (Test-Path -Path $target)) {
    New-Item -Path $target -ItemType Directory -Force | Out-Null
}
# Delete target folder
$folderPath = Join-Path $target $folderName
if (Test-Path $folderPath) {
    $response = Read-Host "Are you sure you want to delete '$folderPath' and all its contents? (y/n)"
    if ($response -eq 'y') {
        Remove-Item -Path $folderPath -Recurse -Force
        Write-Host "Deleted $folderPath"
    } else {
        Write-Host "Deletion cancelled."
    }
}

# Perform the copy
Copy-Item -Path $source -Destination $target -Recurse -Force

Write-Host "Copied '$folderName'  '$target'"

# ---------------- update shortcuts ------------------------------------------
function New-RunShortcut {
    param (
        [string]$environment,
        [string]$shortcutName,
        [string]$workingDir,
        [string]$scriptLike,
        [bool]$gui
    )

    $shortcutTarget = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $condaHook = "C:\Program Files\Anaconda3\shell\condabin\conda-hook.ps1"

    # Determine shortcut window style and icon
    if ($gui -eq $false) {
        $windowStyle = 1
        $iconIndex = 230
        $extraArg = "-noe"
    }
    else {
        $windowStyle = 7
        $iconIndex = 233
        $extraArg = "-w Hidden"
    }

    $language = "&"
    if($scriptLike -like "*.py"){
        $language = "python"
    }
    # Build the command to execute
    $shortcutArgs = "-ep Bypass $extraArg -Command `"& '$condaHook'; " +
    "conda activate $environment; " +
    "while (-not (Test-Path '$workingDir')) { Write-Host 'Waiting, target directory not available ...'; Start-Sleep -Seconds 10 }; " +
    "cd '$workingDir'; " +
    "`$script = Get-ChildItem $scriptLike -File | Select-Object -First 1; " +
    "if (-not `$script) { Write-Host 'No script found' -ForegroundColor Red; Start-Sleep -Seconds 60; exit }; " +
    "$language `$script`""

    if ($env:USERNAME -eq "mrkure"){
        $shortcutArgs = "-ep Bypass $extraArg -Command `"& '$condaHook'; " +
        "conda activate $environment; " +
        "$language (gci $scriptLike -File)[0]`""
    }

    # Create the shortcut
    $wsh = New-Object -ComObject WScript.Shell
    $shortcutPath = Join-Path $workingDir $shortcutName
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $shortcutTarget
    $shortcut.Arguments = $shortcutArgs
    $shortcut.WorkingDirectory = $workingDir
    $shortcut.IconLocation = "%systemroot%\system32\imageres.dll,$iconIndex"
    $shortcut.WindowStyle = $windowStyle
    $shortcut.Save()

    Write-Host "Shortcut '$shortcutName' created at: $shortcutPath"
}

New-RunShortcut -environment "base" `
    -shortcutName "run.lnk" `
    -workingDir $folderPath `
    -scriptLike "install*ps1" `
    -gui $false 



New-RunShortcut -environment "base" `
    -shortcutName "run_cmd.lnk" `
    -workingDir $folderPath `
    -scriptLike "install*ps1" `
    -gui $false 
