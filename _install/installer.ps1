$ErrorActionPreference = 'Stop'
trap {
    Write-Host "`n*** TERMINATING ERROR ***`n$($_.Exception.Message)" -ForegroundColor Red
    Pause
    exit 1
}
# ____________________________________________________________________________________________________________
function GetEnvNameFromFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        throw "File not found: $FilePath"
    }

    $firstLine = Get-Content -Path $FilePath -TotalCount 1

    if ($firstLine -match '^name:\s*(.+)$') {
        return $matches[1].Trim()
    }
    else {
        throw "No 'name:' line found in file."
    }
}
function closePythonProcesses {
    Write-host "`Closing all Python processes"
    while ($true) {
        $pythonProcs = Get-Process python -ErrorAction SilentlyContinue
        if (-not $pythonProcs) {
            Write-Host "All python processes closed ..."
            break
        }

        Write-Host "Killing $($pythonProcs.Count) python process(es)..."
        $pythonProcs | Stop-Process -Force

        Start-Sleep -Seconds 1
    }
    Write-Host "!!! DO NOT START ANY NEW PYTHON PROCESS DURING UNINSTALL !!!" -ForegroundColor Red
}
function removeEnvironment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$targetEnv
    )
    $envs = conda env list 
    $envPath = "$env:USERPROFILE\.conda\envs$targetEnv"
    $envExists = ($envs -match "^\s*$targetEnv(\s|$)").Count -gt 0 
    if ($envExists -eq $true) {   
        closePythonProcesses
        conda activate base
        conda env remove -n $targetEnv
        if (Test-Path $envPath){
            Write-Host "Environment folder still exists`n$envPath`ndelete folder ? (y/n):" -ForegroundColor Blue
            $choice = Read-Host
            if ($choice -eq "y"){
                Remove-Item -Path $envPath -Recurse -Force
                Write-host "Environment ""$targetEnv"", succesfully removed."
            }
            else{
                "Environment $targetEnv cannot be fulle removed"
                return
            }
        }
    }
    else {
        Write-host "Cannot remove env ""$targetEnv"", environment does not exist."
    }
}
function createEnvironment {
    param(
        [string]$targetEnv,
        [string]$envFile
    )
    Write-Host " !!! USING --solver classic !!!" -ForegroundColor Red
    $envs = conda env list 
    $envExists = ($envs -match "^\s*$targetEnv(\s|$)").Count -gt 0 
    if ($envExists -eq $false) {  

        conda activate base
        $env:MENUINST_SKIP = "1" # skip spyder menu installation to avoid errors (stejne nefunguje)
        $env:PIP_VERBOSE = "0" # taky nefunguje
        conda env create -f $envFile --solver classic # use old solver conda 4.11 for backwards compatibility

        # ✅ Conda 4.11
        # Uses the classic dependency solver, which is less strict and often lets you install conflicting or loosely compatible packages.
        # More likely to succeed with loosely defined .yml files.
        # ✅ Conda 23.x+ and newer (e.g., 25.3)
        # Uses the new libmamba solver (enabled by default in recent Conda versions).
        # Much stricter and faster, but exposes conflicts that the old solver ignored.
        # Enforces better reproducibility, but may fail with older .yml f

        conda activate $targetEnv
        $packageList = conda list | Out-String
        if ($packageList -match "(?im)^spyder\s+") {
            Write-Host "Spyder is installed -> reset settings."
            spyder --reset
        }
        Write-host "Environment ""$targetEnv"" succesfully created."
        addStartupShortcuts
        createSpyderShortcut $targetEnv
    }
    else {
        Write-host "Cannot create env ""$targetEnv"", environment already exists."
    }
}
function updateEnvironment {
    param(
        [string]$targetEnv
    )
    $envs = conda env list 
    $envExists = ($envs -match "^\s*$targetEnv(\s|$)").Count -gt 0 
    if ($envExists -eq $true) {  
        conda activate $targetEnv
        conda install certifi -y
        Write-host "Environment ""$targetEnv"", succesfully updated".
    }
    else {
        Write-host "Cannot update env ""$targetEnv"", environment does not exist."
    }
}
function createSpyderShortcut {
    param(
        [string]$targetEnv
    )
    Write-Host "Create spyder shortcut ? (y/n):" -ForegroundColor Blue
    $choice = Read-Host
    if ($choice -ne "y"){
        return
    }
    Write-Host "Creating Spyder shortcut"
    $shortcutTarget = "C:\Program Files\Anaconda3\pythonw.exe"
    $windowStyle = 1

    # Build the command to execute
    $shortcutArgs = 
    "`"C:\Program Files\Anaconda3\cwp.py`" " +
    "$env:USERPROFILE\.conda\envs\work " +
    "$env:USERPROFILE\.conda\envs\work\pythonw.exe " +
    "$env:USERPROFILE\.conda\envs\work\Scripts\spyder-script.pyw"

    # Create the shortcut
    $wsh = New-Object -ComObject WScript.Shell
    $shortcutPath = "$env:USERPROFILE\Desktop\Spyder $targetEnv.lnk"
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $shortcutTarget
    $shortcut.Arguments = $shortcutArgs
    $shortcut.WorkingDirectory = "C:\Users\Public\Documents\Python Scripts"
    $shortcut.IconLocation = "$env:USERPROFILE\.conda\envs\work\Scripts\spyder.ico"
    $shortcut.WindowStyle = $windowStyle
    $shortcut.Save()
    Write-Host "Shortcut created at: $shortcutPath"
}
function chooseResourceYmlFile {
    # Force array of .yml file names
    $ymlFiles = @(Get-ChildItem -Filter *.yml -File | Select-Object -ExpandProperty Name)

    # Raise error if no files found
    if (-not $ymlFiles -or $ymlFiles.Count -eq 0) {
        throw "No .yml files found in the current directory."
    }
    if ($ymlFiles.Count -eq 1) {
        return $ymlFiles[0]

    }
    # Display list
    Write-Host "Choose a number to select resource yml file.`n" -ForegroundColor Blue
    for ($i = 0; $i -lt $ymlFiles.Count; $i++) {
        Write-Host "$($i + 1)) $($ymlFiles[$i])" -ForegroundColor Blue
    }

    # Ask until valid input is given
    while ($true) {
        $choice = Read-Host
        # Validate input
        if ($choice -as [int] -and $choice -ge 1 -and $choice -le $ymlFiles.Count) {
            $selectedFile = $ymlFiles[$choice - 1]
            return $selectedFile
        }
        else {
            Write-Host "Invalid choice. Please enter a number ..." -ForegroundColor Blue
        }
    }
}
function addStartupShortcuts{
    Write-Host "Create startup shortcut for Loader app ? (y/n):" -ForegroundColor Blue
    $choice = Read-Host
    if ($choice -eq "y"){
        $source = Join-Path $PSScriptRoot "..\_loader\run.lnk"
        $target = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\loader.lnk"
        if (Test-Path $source){
            Copy-Item -Path $source -Destination $target -Force
            Write-Host "Succesfully copied`n$source`n$target."
        }
        else {Write-Host "File not found`n$source"}
    }
    Write-Host "Create startup shortcut for Sleep preventer app ? (y/n):" -ForegroundColor Blue
    $choice = Read-Host
    if ($choice -eq "y"){
        $source = Join-Path $PSScriptRoot "..\sleep_preventer\run.lnk"
        $target = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\sleep_preventer.lnk"
        if (Test-Path $source){
            Copy-Item -Path $source -Destination $target -Force
            Write-Host "Succesfully copied`n$source`n$target."
        }
        else {Write-Host "File not found`n$source"}
    }
}
# ======================================= MAIN ===============================================================
    
$env:CONDA_ENVS_PATH = "$env:USERPROFILE\.conda\envs"
     $envFile        = chooseResourceYmlFile                                #$selectedFile #(Get-ChildItem *.yml -File)[0].name
     $targetEnv      = GetEnvNameFromFile -FilePath $envFile                # Get list of conda environments
     $envExists      = ($envs -match "^\s*$targetEnv(\s|$)").Count -gt 0    # Check if "work" environment exists
# ____________________________________________________________________________________________________________
Clear-Host
Write-Host "Environment ""$targetEnv"" exists: $envExists"
Write-Host "Instal using resource file ""$envFile"""

while ($true) {

    Write-Host "`nCreate env 1, Remove env 2, Update env 3, Reinstall env 4: " -ForegroundColor Blue
    $choice = Read-Host

    switch ($choice) {
        '1' {
            Clear-Host
            Write-host "`n========================= Creating env ""$targetEnv"" ========================="
            createEnvironment $targetEnv $envFile
        }
        '2' {
            Clear-Host
            Write-host "`n========================= Removing env ""$targetEnv""  ========================="
            removeEnvironment $targetEnv
        }
        '3' {
            Clear-Host
            Write-host "`n========================= Updating env ""$targetEnv""  ========================="
            updateEnvironment $targetEnv
        }
        '4' {
            Clear-Host
            Write-host "`n================ Reinstall: Removing env ""$targetEnv""  ========================"
            removeEnvironment $targetEnv
            Clear-Host
            Write-host "`n================ Reinstall: Creating env ""$targetEnv"" ========================="
            createEnvironment $targetEnv $envFile
            Clear-Host
            Write-host "`n================ Reinstall: Updating env ""$targetEnv""  ========================"
            updateEnvironment $targetEnv
        }
        default {
            Clear-Host
            Write-Host "Invalid selection. Create env 1, Remove env 2, Update env 3: "
        }
    }
}

