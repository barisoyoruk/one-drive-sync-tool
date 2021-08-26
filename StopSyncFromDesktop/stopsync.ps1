Add-Type -AssemblyName PresentationCore,PresentationFramework

# Constants
$DesktopPath = "C:\Users\baris\OneDrive\Desktop"   # Desktop Path
$RedirectsFolderPath = "C:\Users\baris\Redirects"  # Redirects Path
$ShortcutExtension = ".lnk"                         # Shortcut Extension

# Get path of the file or folder from the right click context
$Path = $args[0]
$CharArray = $Path -split "\\"  # Split Path

# Seperate absolute path to components
$Name = $CharArray[$Array.Count - 1]    # Get File (Folder) Name
$BaseName = $Name.Split(".")[0]         # Get File (Folder) Name without extension. For a folder BaseName == FileName

$PathWithoutName = ""                   # Get Path Name Without File (Folder) Name
For ($i = 0; $i -lt $CharArray.Count - 1; $i++) {
    $PathWithoutName += $CharArray[$i]
    $PathWithoutName += "\" 
}

# Checking Whether File (Folder) Properties are appropriate before executing steps
if (!$PathWithoutName.StartsWith($DesktopPath)) {   # Checking Whether File (Folder) is in desktop or desktop's subdirectories. If not, exit the script
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "File is not in the subdirectories of desktop"
    $MessageTitle = "Unsuccessfull"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 

    exit 1
}

# Move item from its desktop location to the Redirects folder where sync does not occur
Move-Item -Path $path -Destination $RedirectsFolderPath

# Create shourt cut to the original location of the item
$SourceFileLocation ="$($RedirectsFolderPath)\$($Name)"
$ShortcutLocation = "$($PathWithoutName)$($BaseName)$($ShortcutExtension)"

$WScriptShell = New-Object -ComObject WScript.Shell

$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $SourceFileLocation
$Shortcut.Save()

# GUI For Operation
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageIcon = [System.Windows.MessageBoxImage]::Information	
$MessageBody = "Sync is stopped for this file"
$MessageTitle = "Successfull"

$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
Write-Host "Your choice is $Result"  