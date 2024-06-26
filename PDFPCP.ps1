#Dialog box for folder selection
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    RootFolder      = "Desktop"
    Description     = "PDF Page Counter Powershell - Pick a folder"   
}
    if($FolderBrowser.ShowDialog() -eq "OK")
    {
        $folder = $FolderBrowser.SelectedPath
    }
    else { break }

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
    $object = [PSCustomObject]@{
        PdfFile = $File.Name
        Pages   = $Pages
    }
    #Write process to terminal
    Write-Host $object
    #Make CSV file
    $csvData = $object | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
    $csvData | Select-Object -Skip 1 | Out-File -FilePath $outputFile -Append
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
    $object = [PSCustomObject]@{
        PdfFile = $File.Name
        Pages   = $Pages
    }
    #Write process to terminal
    Write-Host $object
    #Make CSV file
    $csvData = $object | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
    $csvData | Select-Object -Skip 1 | Out-File -FilePath $outputFile -Append
}

"`nTotal Number of pages: {0} in {1} files" -f $Total,$Files | Out-File -FilePath $outputFile -Append
    }
    
    #Finished pop up
    $wshell = New-Object -ComObject Wscript.Shell

    $wshell.Popup("Finished!",0,"PDFPCP",0x1)

#Opens selected folder in Explorer window
Invoke-Item $name
    