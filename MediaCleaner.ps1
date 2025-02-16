﻿#Created By: Nick Kliatsko
#Last Updated: 12/22/2024

#This script is designed to clean up media files in a directory. It will move all video files to the root directory, unzip any archives, and delete any unnecessary files. It will also clean up the folder names by removing any tags that are present.
#The script will also check for the presence of 7-Zip and will install it if it is not present.
#Requirements: Windows 10, Powershell 5.1, 7-Zip
#Run the script in the directory you want to clean up.

#General SEtup
Add-Type -AssemblyName System.Windows.Forms

function Select-FolderDialog {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to clean up"
    $folderBrowser.ShowNewFolderButton = $true

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        Write-Host "No folder selected" -ForegroundColor Red
        exit
    }
}

#movies or shows
$type = read-host "1=Movies 2=Shows"

#movie tasks
if ($type -eq 1){
    $path = Select-FolderDialog
    Write-host "Movie Routine"
    
    #test for 7-zip install
    if (Test-path -Path "C:\Program Files\7-Zip") {write-host "7-Zip installed" -ForegroundColor Green}
    else {write-host "7-Zip not installed" -ForegroundColor Red
        $7zip = "https://www.7-zip.org/a/7z1900-x64.exe"
        Start-Process -FilePath $7zip
        write-host "7-Zip installed" -ForegroundColor Green
    }

    #clean unnecessary files
    Get-ChildItem -path $path -Filter *Subs* -Recurse | Remove-Item -Recurse
    Get-ChildItem -path $path -Filter *Sample* -Recurse | Remove-Item -Recurse
    Get-ChildItem -path $path -Filter *Trailer* -Recurse | Remove-Item -Recurse
    Get-ChildItem -path $path -Filter *Proof* -Recurse | Remove-Item -Recurse
    Get-ChildItem -path $path -Filter *Screens* -Recurse | Remove-Item -Recurse

    #unzip archives to root
    set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"
    $unzipQueue = Get-ChildItem -Path $path -Filter *.rar -Recurse
    $count = $unzipQueue.count
    ForEach ($zippedFile in $unzipQueue) {
        write-output $zippedFile.PSParentPath
        $destpath = $path
        $sourcepath = $zippedFile.FullName
        sz x -o"$destpath" $sourcepath -r -y ;
        $count -= 1
        write-host $count
    }
    #deletes leftover archives
    Get-ChildItem $path -Include *.r?? -Recurse | Remove-Item

    #creates folders for any loose files
    $loose = Get-ChildItem $path -Attributes !Directory | select-object name, pspath
    foreach ($file in $loose) {
        write-output $file
        $name=($file.Name)
        $dir=$name.Substring(0,$name.Length-4)
        new-item -Name $dir -path $path -type directory
        $newpath = $path+$dir
        Move-Item -path $file.PSPath -destination $newpath
    }

    #cleans tags out of folder titles
    Get-ChildItem -path $path -Filter *1080p* |Rename-Item -NewName { $($_.Name -split '1080p')[0] }
    Get-ChildItem -path $path -Filter *2160p* |Rename-Item -NewName { $($_.Name -split '2160p')[0] }
    Get-ChildItem -path $path -Filter *720p* |Rename-Item -NewName { $($_.Name -split '720p')[0] }
    Get-ChildItem -path $path -Filter *HDRip* |Rename-Item -NewName { $($_.Name -split 'HDRip')[0] }
    Get-ChildItem -path $path -Filter *DVDRip* |Rename-Item -NewName { $($_.Name -split 'DVDRip')[0] }
    Get-ChildItem -path $path -Filter *BRRip* |Rename-Item -NewName { $($_.Name -split 'BRRip')[0] }
    Get-ChildItem -path $path -Filter *BR-Rip* |Rename-Item -NewName { $($_.Name -split 'BR-Rip')[0] }
    Get-ChildItem -path $path -Filter *BDRip* |Rename-Item -NewName { $($_.Name -split 'BDRip')[0] }
    Get-ChildItem -path $path -Filter *Extended* |Rename-Item -NewName { $($_.Name -split 'Extended')[0] }
    Get-ChildItem -path $path -Filter *Unrated* |Rename-Item -NewName { $($_.Name -split 'Unrated')[0] }
    Get-ChildItem -path $path -Filter *Remastered* |Rename-Item -NewName { $($_.Name -split 'Remastered')[0] }
    Get-ChildItem -path $path -Filter *iNTERNAL* |Rename-Item -NewName { $($_.Name -split 'iNTERNAL')[0] }
    Get-ChildItem -path $path -Filter *MULTi* |Rename-Item -NewName { $($_.Name -split 'MULTi')[0] }
    Get-ChildItem -path $path -Filter *WEB-DL* |Rename-Item -NewName { $($_.Name -split 'WEB-DL')[0] }
    Get-ChildItem -path $path -Filter *DVDR* |Rename-Item -NewName { $($_.Name -split 'DVDR')[0] }
    Get-ChildItem -path $path -Filter *x264* |Rename-Item -NewName { $($_.Name -split 'x264')[0] }
    Get-ChildItem -path $path -Filter *DVDRip* |Rename-Item -NewName { $($_.Name -split 'DVDRip')[0] }
    Get-ChildItem -path $path -Filter *WEBRip* |Rename-Item -NewName { $($_.Name -split 'WEBRip')[0] }
    Get-ChildItem -path $path -Filter *HC* |Rename-Item -NewName { $($_.Name -split 'HC')[0] }
    Get-ChildItem -path $path -Filter *DVDScr* |Rename-Item -NewName { $($_.Name -split 'DVDScr')[0] }
    Get-ChildItem -path $path -Filter *REAL* |Rename-Item -NewName { $($_.Name -split 'REAL')[0] }
    Get-ChildItem -path $path -Filter *BluRay* |Rename-Item -NewName { $($_.Name -split 'BluRay')[0] }
    Get-ChildItem -path $path -Filter *LiMiTED* |Rename-Item -NewName { $($_.Name -split 'LiMiTED')[0] }
    Get-ChildItem -path $path -Filter *10bit* |Rename-Item -NewName { $($_.Name -split '10bit')[0] }
    Get-ChildItem -path $path -Filter *X265* |Rename-Item -NewName { $($_.Name -split 'X265')[0] }
    Get-ChildItem -path $path -Filter *BD-Rip* |Rename-Item -NewName { $($_.Name -split 'BD-Rip')[0] }
    Get-ChildItem -path $path -Filter *hevc-d3g* |Rename-Item -NewName { $($_.Name -split 'hevc-d3g')[0] }
    Get-ChildItem -path $path -Filter *REPACK* |Rename-Item -NewName { $($_.Name -split 'REPACK')[0] }
    Get-ChildItem -path $path -Filter *1080-hd4u* |Rename-Item -NewName { $($_.Name -split '1080-hd4u')[0] }
    Get-ChildItem -path $path -Filter *ExtCut* |Rename-Item -NewName { $($_.Name -split 'ExtCut')[0] }
    Get-ChildItem -path $path -Filter *"Anniversary Edition"* |Rename-Item -NewName { $($_.Name -split 'Anniversary Edition')[0] }
    Get-ChildItem -path $path -Filter *"Restored"* |Rename-Item -NewName { $($_.Name -split 'Restored')[0] }
    Get-ChildItem -path $path -Filter *.* | Rename-Item -NewName {$_.name -replace '[.]',' ' }
    Get-ChildItem -Path $path -filter "* 20??" | Rename-Item -newname { $_ -replace '(.*)(\d{4})', '$1($2)'}
    Get-ChildItem -Path $path -filter "* 19??" | Rename-Item -newname { $_ -replace '(.*)(\d{4})', '$1($2)'}
}

#Movie Tools
    <# Check the number of subfolders
    if ($subfolders.Count -gt 3) {
        Write-Host "There are more than three subfolders in $folderPath."
    } else {
        Write-Host "There are three or fewer subfolders in $folderPath."
    }
    #>

    #Copy all video files to root (troubleshooting)
    <#
    $videoFiles = Get-ChildItem -Path $path -Recurse -File | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }
    foreach ($file in $videoFiles) {
        $destinationPath = Join-Path -Path $path -ChildPath $file.Name
        Move-Item -Path $file.FullName -Destination $destinationPath -Force
        Write-Host "Copied $($file.Name) to $($destinationPath)"
        }
    #>

    
    #Cleans up Empty Folders
    <#
    $movieFolders = Get-ChildItem -Path $path -Directory
    foreach ($folder in $movieFolders) {
        $subfolder = Get-ChildItem -Path $folder -Directory
        # Get all files and folders inside the subfolder
        $items = Get-ChildItem -Path $subFolder.FullName -Force -Recurse
        # If there are no items, delete the subfolder
        if ($items.Count -eq 0) {
            Write-host "This folder is empty"
            Remove-Item -Path $subFolder.FullName -Force
        }
        # If there are items, delete them and then delete the subfolder
        else {
            Foreach ($file in $items){
                Move-Item -Path $file.FullName -Destination $folder -Force
                Remove-Item -Path $subFolder.FullName -Recurse -Force
            }
        }
    } 
    #>

#show tasks
elseif ($type -eq 2){
    $path = Select-FolderDialog
    Write-host "Show Routine"
    $FileType= @("*.*mp4", "*.*mkv")
    #dump loose video files to root
    Get-ChildItem -recurse ($path) -include ($FileType) | move-Item -Destination ($path)
    #unzip archives to root
    set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"
    $unzipQueue = Get-ChildItem -Path $path -Filter *.rar -Recurse
    $count = $unzipQueue.count
    ForEach ($zippedFile in $unzipQueue) {
        Write-Output $zippedFile.PSParentPath
        $destpath = $path
        $sourcepath = $zippedFile.FullName
        sz x -o"$destpath" $sourcepath -r -y ;
        $count -= 1
        write-host $count
    }
    #delete (hopefully) now-empty folders
    Remove-Item $path\* -Exclude *.* -Force
break
}

else{
    write-host "you broke something"
    break
}

#Show Tools
#TBC...