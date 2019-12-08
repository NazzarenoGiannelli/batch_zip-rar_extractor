Write-Host '                                         '
Write-Host '-----------------------------------------'
Write-Host '---------- BATCH ZIP EXTRACTOR ----------'
Write-Host '-----------------------------------------'
Write-Host '                                         '
Write-Host 'Extract multiple .zip files in new folders with the same name'
Write-Host '                                         '
Write-Host '-----------------------------------------'
Write-Host '                                         '

do {
    $FilesFolder = Read-Host -Prompt 'In which folder the zip files are located?';

    $FilesList = Get-ChildItem -Path $FilesFolder | Where-Object {$_.extension -eq '.zip'};

    foreach ($ListedFile in $FilesList) {
        mkdir -Path ('{0}\{1}' -f $ListedFile.Directory, $ListedFile.BaseName);
        $ArgumentList = 'x "{0}" -o"{1}\{2}\"' -f $ListedFile.FullName, $ListedFile.Directory, $ListedFile.BaseName;
        Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $ArgumentList -Wait;
    }
    
    Write-Host '                                         '
    Write-Host '-----------------------------------------'
    Write-Host '                                         '
    Write-Host "1 - Repeat the operation in another folder"
    Write-Host "2 - Exit the program"
    $choice = Read-Host "Choose a number to proceed. Press Enter to confirm."

    switch ($choice) {
        1 {
            Write-Host "Let's process other files!"
        }
        2 {
            Write-Host "Turning off...see you soon! :)"
            Start-Sleep 3
        }
    }
} until ($choice -eq 2)
