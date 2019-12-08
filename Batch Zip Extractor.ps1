Write-Host '                                         '
Write-Host '-----------------------------------------'
Write-Host '---------- BATCH ZIP EXTRACTOR ----------'
Write-Host '-----------------------------------------'
Write-Host '                                         '
Write-Host 'Extract multiple .zip files in new folders with the same name'
Write-Host '                                         '
Write-Host '-----------------------------------------'
Write-Host '                                         '

$FilesFolder = Read-Host -Prompt 'In which folder the zip files are located?';

$FilesList = Get-ChildItem -Path $FilesFolder | where {$_.extension -eq '.zip'};

foreach ($ListedFile in $FilesList) {
    mkdir -Path ('{0}\{1}' -f $ListedFile.Directory, $ListedFile.BaseName);
    $ArgumentList = 'x "{0}" -o"{1}\{2}\"' -f $ListedFile.FullName, $ListedFile.Directory, $ListedFile.BaseName;
    Start-Process -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $ArgumentList -Wait;
}