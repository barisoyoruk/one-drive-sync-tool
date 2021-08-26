Add-Type -AssemblyName PresentationCore,PresentationFramework

# User Name
$UserName = "baris"                     
$OneDrivePath = "OneDrive\Laptop"   # Under current users directory i.e C:\Users\baris\

# Constants
$SyncFolderPath = "C:\Users\$($UserName)\$($OneDrivePath)"  # Folder that is synced
$ShortcutExtension = ".lnk"                                 # Shortcut Extension

# Get path of the file or folder from the right click context
$AbsolutePathWithFileName = $args[0]
$PathComponents = $AbsolutePathWithFileName -split "\\"  # Split AbsolutePathWithFileName

# Seperate absolute path to components
$FileName = $PathComponents[$Array.Count - 1]    # Get File (Folder) FileName

$AbsolutePath = ""                   # Get Absolute Path FileName Without File (Folder) FileName
$RelativePath = ""                   # Get Relative Path FileName Without File (Folder) FileName
$IsInRelativePath = 0
For ($i = 0; $i -lt $PathComponents.Count - 1; $i++) {

    if ($IsInRelativePath) {
        $RelativePath += $PathComponents[$i]
        $RelativePath += "\" 
    }

    $AbsolutePath += $PathComponents[$i]
    $AbsolutePath += "\" 

    if ($PathComponents[$i].Equals("baris")) {
        $IsInRelativePath = 1
    }
}

if ($RelativePath.Equals("")) {         # Checking Whether The Folder under syncable folder
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "Target file is not in under folder C:\Users\$($UserName)"
    $MessageTitle = "Unsuccessful"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 
    
    exit 1
}

# Move item from its desktop location to the One Drive folder
$From = $AbsolutePathWithFileName
$To = "$($SyncFolderPath)\$($RelativePath)\"

# If relative path folders are not in the one driver folder, create them
if(!(Test-Path $to))
{
    New-Item -Path $to -ItemType Directory -Force | Out-Null
}

# Move item from its original location to the one drive folder
Move-Item -Path $From -Destination $To 

# Create shourt ct to the original location of the item
$SourceFileLocation = "$($SyncFolderPath)\$($RelativePath)$($FileName)"
$ShortcutLocation = "$($AbsolutePathWithFileName) - S$($ShortcutExtension)"

$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $SourceFileLocation
$Shortcut.Save()

exit 0