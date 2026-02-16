#  process_crossboarding.ps1
#   Updated for Crossboarding Documents
#   Testing notes:  
#      1)  errors in logs, merges or doc conversion steps
#      3)  handles duplicate hires in the same month & word documents
#      4)  sometimes files have font or formatting issues that Word-PDF cannot handle - simply open those in PDF or Word and resave to remove issues
#      5)  will throw an error if EE folder within month has no files at all - user can simple ignore those errors


Import-Module -Name Microsoft.PowerShell.Archive
. $PSScriptRoot\utilities.ps1

# Set logging
$log = $pwd.Path + "\work.log"
Start-Transcript -Append $log 

# Year Folder has subfolders now
$pdfCount = 0
$fileInfo = @()
$docCount = 0
$srcArchivePath = $pwd.Path + "\2019\"
$dstUnzipPath = $pwd.Path + "\Unzip\"

# track processed names to handle duplicates in same year
$processedNames = @{}

# for each month folder
Get-ChildItem $srcArchivePath -Filter "*" -Directory | ForEach-Object {
   Write-Host "Folder: $($_.FullName)"

   # if no folders in month, go to next month
   $monthFolder = $_.FullName
   $fileCount = (Get-ChildItem $srcArchivePath -Filter "*" -Directory | Measure-Object).Count
   Write-Host "File count: $($fileCount)"
   if ($fileCount -eq 0) {
      Write-Host "Skipping empty folder"
      return
   }

   # for each employee folder in the month
   Get-ChildItem $monthFolder -Directory | ForEach-Object {
      Write-Output "Employee-Folder: $($_.FullName)"

      $employeeFolderPath = $_.FullName
      $nameFolder = $_.Name
      $nameSplit = $nameFolder.Split("_")
      $eeString = "2" + $nameSplit[0]
      $eeFirstName = $nameSplit[1]
      $eeLastName = $nameSplit[2]
     

      # basename key for duplicate checking
      $baseNameKey = "$eeString`_$eeFirstName`_$eeLastName"
      $duplicateName = ""
      if ($processedNames.ContainsKey($baseNameKey)) {
         $processedNames[$baseNameKey]++
         $duplicateName = "_$($processedNames[$baseNameKey])"
         Write-Host "Duplicate name detected, adding number to differentiate"
      } else {$processedNames[$baseNameKey] = 1}

      # moved this logic down to avoid duplicate name logic
      $tgtPDFName = $eeString + "_" + $eeFirstName + "_" + $eeLastName + "_" + ".pdf"
      $tgtZipName = $eeString + "_" + $eeFirstName + "_" + $eeLastName + ".zip"
   
    
      Write-Host "Processing for $($eeFirstName) $($eeLastName) with ee# of $($eeString)"


      # for each file in folder
      # check / convert / copy
      # test for Resume, rename with FILENAME +".pdf"
      Get-ChildItem $employeeFolderPath -File | ForEach-Object {
         Write-Host "Found file: $($_.Name)"
         if ($_.Name -match "(?i)RESUME") {
            Write-Output "Found and validating $($_.FullName)"
            $newResumeFullPath = $_.FullName
            # changed to basename here for avoiding docx extension
            $newResumeDstFullPath = Join-Path $dstUnzipPath ($_.BaseName + ".pdf")
            $docType = Get-DocumentType -filePath $_.FullName
            if ($docType -eq "DOCX" -or $docType -eq "DOC" -or $docType -eq "TXT") {
               Convert-WordToPDF -wordPath $newResumeFullPath -pdfPath $newResumeDstFullPath
            }
            else {
               Copy-Item -Path $newResumeFullPath -Destination $newResumeDstFullPath
            }
         }
           # check for Word Docs
         if (($_.Name -match "\.doc$" -or $_.Name -match "\.docx") -and ($_.Name -notmatch "(?i)RESUME" )) {
            $docCount ++
            $wordDocFullPath = $_.FullName
            $wordDocDstFullPath = Join-path $dstUnzipPath ($_.BaseName + ".pdf")
            $docType = Get-DocumentType -filePath $_.FullName
            if ($docType -eq "DOCX" -or $docType -eq "DOC" -or $docType -eq "TXT") {
               Convert-WordToPDF -wordPath $wordDocFullPath -pdfPath $wordDocDstFullPath
            }
            else {
               Copy-Item -Path $wordDocFullPath -Destination $wordDocDstFullPath
            }
         }
         # copy existing pdfs
         if ($_.Name -match "\.pdf$") {
            Copy-Item -Path $_.FullName -Destination $dstUnzipPath
            $pdfCount++

            Write-Host "Saving PDFS to: $dstUnzipPath"

         }
       
         # remove resumes not matching last name
         if ($_.Name -match "RESUME") {
            if ($_.Name -notmatch $eeLastName) {
               Write-Host "Wrong name for Resume found in $($_.FullName)"
               # Remove-Item -Path $_.FullName
            }
         }
         # end file loop
      }
      # populate fileInfo csv
    
      Get-ChildItem -Path $dstUnzipPath -Filter "*" | ForEach-Object {
         # track file information in a CSV
         $fileSig = Get-FileSignature -filePath $_.FullName
         $fileExt = Get-DocumentType -filePath $_.FullName
         $fileInfo += New-Object PSObject -Property @{
            filePath = $_.FullName
            fileSig = $fileSig
            fileExt = $fileExt
            eeNumber = $eeString
            eeFirst = $eeFirstName
            eeLast = $eeLastName
         }
      }
      
      # debugging by claude
      Write-Host "=== DEBUG: Checking for PDFs before cover page creation ==="
      Write-Host "Looking in: $dstUnzipPath"
      $debugPDFs = Get-ChildItem $dstUnzipPath -Filter "*.pdf"
      Write-Host "PDF count found: $($debugPDFs.Count)"
      foreach ($pdf in $debugPDFs) {
         Write-Host "  - $($pdf.Name)"
      }
      Write-Host "=== END DEBUG ==="

      # build cover page
      $coverPDF = $pwd.path + "\Working\" + $eeString + "_cover.pdf"
      $pdfFileList = New-Object -TypeName 'System.Collections.ArrayList'
      $pdfFileList.Add($coverPDF)
      New-PDF {
         $msg = "Crossboarding Documents for " + $eeFirstName + " " + $eeLastName
         if ($duplicateName) {
            $msg += " (Instance $($processedNames[$baseNameKey]))"
         }
         New-PDFText -Text $msg
         Get-ChildItem $dstUnzipPath -Filter "*.pdf" | Sort-STNumerical | ForEach-Object {
            $pdfCount ++
            $fqPathSrc = $dstUnzipPath + $_.Name
            $fqPathDest = $fqPathSrc.replace(" ","_")
            Move-Item -Path $fqPathSrc -Destination $fqPathDest
 
            $pdfFileList.Add($fqPathDest)
            $shortPath = $fqPathDest.replace($dstUnzipPath, "")
            # $baseName = $_.BaseName
            New-PDFText -Text $shortPath
         }
      } -FilePath $coverPDF
      $pdfOutFile = $pwd.path + "\Working\" + $eeString + "_" + $eeFirstName + "_" + $eeLastName + $duplicateName + "_Crossboarding.pdf"
      
      # merge pdf block
      $pdfFileCount = $pdfFileList.Count
      if ($pdfFileCount -gt 1) {
      for ($i = 0; $i -lt $pdfFileCount - 1; $i++) {
         if ($i -eq 0) {
            $mergeOut = $pwd.Path + "\Working\" + $eeString + $duplicateName + "_workingdoc" + $i + ".pdf"
            $mergeIn = $pdfFileList[0]
         }
         else {
            $mergeOut = $pwd.Path + "\Working\" + $eeString + $duplicateName + "_workingdoc" + $i + ".pdf"
         }
         Merge-PDF -InputFile $mergeIn, $pdfFileList[$i + 1] -OutputFile $mergeOut -IgnoreProtection
         $mergeIn = $mergeOut
      }
      Move-Item -Path $mergeOut -Destination $pdfOutFile 
      } else {
         Write-Host "Only cover page exists, no pdfs exist for $($eeFirstName) $($eeLastName)"

      }
      $rmWOrkingDocPath = $pwd.Path + "\Working\" + $eeString + "_workingdoc*.*"
      Remove-Item -Path $rmWOrkingDocPath 

      # Zip PDF file
      $zipSourcePath = $pdfOutFile
      $zipDestPath = $pwd.Path + "\Zip\" + $eeString + "_" + $eeFirstName + "_" + $eeLastName + $duplicateName + "_Crossboarding.zip"
      Compress-Archive -Path $zipSourcePath -Destination $zipDestPath

      # clean up files b4 next zip loop
      $removeZipPath = $dstUnzipPath + "*"
      Remove-Item -Path $removeZipPath
      $removeStagePath = $pwd.Path + "\Working\*.*"
      Remove-Item -Path $removeStagePath
   }
   # end folder loop
}
# end month loop
$finalArchivePath = "D:\Zip"
$forUploadPath = "D:\ForUpload\"
Get-ChildItem $finalArchivePath -Filter "*.zip" | ForEach-Object {
   Move-Item -Path $_.FullName -Destination $forUploadPath
}
$fileInfoFullPath = $forUploadPath + "fileInfo.csv"
$fileInfo | Select-Object eeNumber, eeLast, eeFirst, filePath, fileSig, fileExt | Export-Csv -Path  $fileInfoFullPath -NoTypeInformation
Write-Host "Writing fileInfo to $($fileInfoFullPath)"
Write-Host "##################"
Write-Host "PDF Count: $($pdfCount)"
Write-Host "DOC Count: $($docCount)"

Stop-Transcript
