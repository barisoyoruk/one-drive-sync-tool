Add-Type -AssemblyName PresentationCore,PresentationFramework

# User Name
$UserName = "baris"                     
$OneDrivePath = "OneDrive\Laptop"   # Under current users directory i.e C:\Users\baris\

# Constants
$SyncFolderPath = "C:\Users\$($UserName)\$($OneDrivePath)"  # Folder that is synced

# Get path of the file or folder from the right click context
$AbsolutePathWithFileName = $args[0]
$PathComponents = $AbsolutePathWithFileName -split "\\"  # Split AbsolutePathWithFileName

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

    if ($PathComponents[$i].Equals($UserName)) {
        $IsInRelativePath = 1
    }
}

if ($RelativePath.Equals("")) {       # Checking Whether The shortcut under syncable folder
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "Shortcut is not in under folder C:\Users\$($UserName)"
    $MessageTitle = "Unsuccessful"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 

    exit 1
}

# Get target file folder from the shortcut path
$WScriptShell = New-Object -ComObject WScript.Shell
$TargetPath = $WScriptShell.CreateShortcut($AbsolutePathWithFileName).TargetPath

if (!$TargetPath.StartsWith("$SyncFolderPath")) {       # Checking Whether File (Folder) is in one drive folder. If not, exit the script
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    $MessageBody = "Target file is not in the one drive folder"
    $MessageTitle = "Unsuccessful"

    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host "Your choice is $Result" 

    exit 1
}

# Move item from one drive to its original location
Move-Item -Path $TargetPath -Destination $AbsolutePath
Remove-Item $AbsolutePathWithFileName

exit 0