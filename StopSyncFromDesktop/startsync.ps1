Add-Type -AssemblyName PresentationCore,PresentationFramework
# Constants
$DesktopPath = "C:\Users\baris\OneDrive\Desktop"   # Desktop Path
$RedirectsFolderPath = "C:\Users\baris\Redirects"  # Redirects Path

# Get path of the file or folder from the right click context
$Path = $args[0]
$CharArray = $Path -split "\\"  # Split Path

$PathWithoutName = ""                   # Get Path Name Without File (Folder) Name
For ($i = 0; $i -lt $CharArray.Count - 1; $i++) {
    $PathWithoutName += $CharArray[$i]
    $PathWithoutName += "\" 
}

# Get target file folder from the shortcut path
$WScriptShell = New-Object -ComObject WScript.Shell
$TargetPath = $WScriptShell.CreateShortcut($Path).TargetPath

if (!$TargetPath.StartsWith($RedirectsFolderPath)) {   # Checking Whether File (Folder) is in redirect folder. If not, exit the script
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "Target file is not in the redirects folder"
    $MessageTitle = "Unsuccessful"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 

    Read-Host -Prompt "asad"
    exit 1
}

# Checking Whether File (Folder) Properties are appropriate before executing steps
if (!$PathWithoutName.StartsWith($DesktopPath)) {   # Checking Whether File (Folder) is in desktop or desktop's subdirectories. If not, exit the script
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "Shortcut is not in the subdirectories of desktop"
    $MessageTitle = "Unsuccessfull"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 

    Read-Host -Prompt "asad"
    exit 1
}

# Move item from its desktop location to the Redirects folder where sync does not occur
Move-Item -Path $TargetPath -Destination $PathWithoutName
Remove-Item $Path

Read-Host -Prompt "asad"