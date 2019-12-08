Write-Host '                                         '
Write-Host '---------------------------------------------------------------------'
Write-Host '---------------------- BATCH ZIP/RAR EXTRACTOR ----------------------'
Write-Host '---------------------------------------------------------------------'
Write-Host '                                         '
Write-Host 'Extract multiple .zip or .rar files in new folders with the same name'
Write-Host '                                         '
Write-Host '---------------------------------------------------------------------'
Write-Host '                                         '

do {
    $FilesFolder = Read-Host -Prompt 'In which folder the compressed files are located?';

    Write-Host 'Select a file extention for the files to process'
    Write-Host '   '
    Write-Host '1 - Extract .zip files'
    Write-Host '2 - Extract .rar files'
    Write-Host '3 - Extract both .zip and .rar files'
    Write-Host '   '
    $FileExtention = Read-Host -Prompt "Choose a number to proceed. Press Enter to confirm."

    switch ($FileExtention) {
        1 { 
            $FilesList = Get-ChildItem -Path $FilesFolder | Where-Object {$_.extension -eq '.zip'};

            foreach ($ListedFile in $FilesList) {
                mkdir -Path ('{0}\{1}' -f $ListedFile.Directory, $ListedFile.BaseName);
                $ArgumentList = 'x "{0}" -o"{1}\{2}\"' -f $ListedFile.FullName, $ListedFile.Directory, $ListedFile.BaseName;
                Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $ArgumentList -Wait;
            }
         }

        2 { 
            $FilesList = Get-ChildItem -Path $FilesFolder | Where-Object {$_.extension -eq '.rar'};

            foreach ($ListedFile in $FilesList) {
                mkdir -Path ('{0}\{1}' -f $ListedFile.Directory, $ListedFile.BaseName);
                $ArgumentList = 'x "{0}" -o"{1}\{2}\"' -f $ListedFile.FullName, $ListedFile.Directory, $ListedFile.BaseName;
                Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $ArgumentList -Wait;
            }
        }

        3 { 
            $FilesList = Get-ChildItem -Path $FilesFolder | Where-Object {($_.extension -eq '.zip' -or $_.extension -eq '.rar')};

            Write-Host $FilesList

            foreach ($ListedFile in $FilesList) {
                mkdir -Path ('{0}\{1}' -f $ListedFile.Directory, $ListedFile.BaseName);
                $ArgumentList = 'x "{0}" -o"{1}\{2}\"' -f $ListedFile.FullName, $ListedFile.Directory, $ListedFile.BaseName;
                Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $ArgumentList -Wait;
            }
        }
        Default {1}
    }

    Write-Host '   '
    Write-Host '   '
    Write-Host 'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv'
    Write-Host 'OPERATION COMPLETED SUCCESFULLY'
    Write-Host '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
    Write-Host '   '
    Write-Host '   '
    Write-Host '___'
    Write-Host '   '
    Write-Host 'What you wanna do next?'
    Write-Host '   '
    Write-Host "1 - Repeat the operation in another folder"
    Write-Host '   '
    Write-Host "2 - Exit the program"
    Write-Host '   '
    $choice = Read-Host -Prompt "Choose a number to proceed. Press Enter to confirm."

    switch ($choice) {
        1 {
            Write-Host '___'
            Write-Host '   '
            Write-Host "Let's process other files!"
        }
        2 {
            Write-Host '___'
            Write-Host '   '
            Write-Host "Turning off...see you soon! :)"
            Start-Sleep 3
        }
    }
} until ($choice -eq 2)