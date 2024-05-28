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

#Get name from selected folder for name of report
$name = Get-Item $FolderBrowser.SelectedPath

#Get date for name of report
$date = Get-Date -Format 'dd-MM-yyyy'

#Question box YesNo for subfolders    
$answer = [System.Windows.Forms.MessageBox]::Show("Subfolders?", "PDFPCP", "YesNo", "Question")
    if($answer -eq "Yes")
    {
        #Script for counting pages and making a report that goes in the chosen folder if subfolders is selected
        $outputFile = Join-Path -Path $folder -ChildPath "$($name.BaseName)_report_$($date).txt"
$Total = $Files = 0

foreach($File in (Get-ChildItem -Path $folder -Recurse -Filter *.pdf)){
    $Pages = (.\pdfinfo $File.FullName | Select-String -Pattern '(?<=Pages:\s*)\d+').Matches.Value
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
        $outputFile = Join-Path -Path $folder -ChildPath "$($name.BaseName)_report_$($date).txt"
$Total = $Files = 0

foreach($File in (Get-ChildItem -Path $folder -Filter *.pdf)){
    $Pages = (.\pdfinfo $File.FullName | Select-String -Pattern '(?<=Pages:\s*)\d+').Matches.Value
    $Total += $Pages
    $Files++ 
    [PSCustomObject]@{
        PdfFile = $File.Name
        Pages   = $Pages
    } | Out-File -FilePath $outputFile -Append
}

"`nTotal Number of pages: {0} in {1} files" -f $Total,$Files | Out-File -FilePath $outputFile -Append
    }
    
    #Finished pop up
    $wshell = New-Object -ComObject Wscript.Shell

    $wshell.Popup("Finished!",0,"PDFPCP v1",0x1)
