#Dialog box for folder selection
$folder = Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    RootFolder      = "Desktop"
    Description     = "PDF Page Counter Powershell - Pick a folder"   
}
    if($FolderBrowser.ShowDialog() -eq "OK")
    {
        $folder += $FolderBrowser.SelectedPath
    }
#Question box YesNo for subfolders    
$answer = [System.Windows.Forms.MessageBox]::Show("Subfolders?", "PDFPCP", "YesNo", "Question")
    if($answer -eq "Yes")
    {
        #Script for counting pages and making a report that goes in the chosen folder if subfolders is selected
        $outputFile = Join-Path -Path $folder -ChildPath "Rapport.txt"
$Total = $Files = 0

foreach($File in (Get-ChildItem -Path $folder -Recurse -Filter *.pdf)){
    $Pages = (X:\PDF_TOOLS\pdfinfo.exe $File.FullName | Select-String -Pattern '(?<=Pages:\s*)\d+').Matches.Value
    $Total += $Pages
    $Files++ 
    [PSCustomObject]@{
        PdfFile = $File.Name
        Pages   = $Pages
    } | Out-File -FilePath $outputFile -Append
}

"`nTotal Number of pages: {0} in {1} files" -f $Total,$Files | Out-File -FilePath $outputFile -Append
    }

    if($answer -eq "No")
    {
        #Script for counting pages and making a report that goes in the chosen folder if subfolders is not selected
        $outputFile = Join-Path -Path $folder -ChildPath "Rapport.txt"
$Total = $Files = 0

foreach($File in (Get-ChildItem -Path $folder -Filter *.pdf)){
    $Pages = (X:\PDF_TOOLS\pdfinfo.exe $File.FullName | Select-String -Pattern '(?<=Pages:\s*)\d+').Matches.Value
    $Total += $Pages
    $Files++ 
    [PSCustomObject]@{
        PdfFile = $File.Name
        Pages   = $Pages
    } | Out-File -FilePath $outputFile -Append
}

"`nTotal Number of pages: {0} in {1} files" -f $Total,$Files | Out-File -FilePath $outputFile -Append
    }