# process_onboarding documents
# CORRECTED VERSION with proper nested loop structure

# the infamous iText7 dlls :(
Add-Type -Path "\PowerShell\Modules\PSWritePDF\0.0.20\Lib\Core\itext.kernel.dll"
Add-Type -Path "\PowerShell\Modules\PSWritePDF\0.0.20\Lib\Core\itext.io.dll"
Add-Type -Path "\PowerShell\Modules\PSWritePDF\0.0.20\Lib\Core\itext.layout.dll"

Import-Module -Name Microsoft.PowerShell.Archive
. $PSScriptRoot\utilities.ps1

# logging file
$log = "\work.log"
Start-Transcript -Append $log 

# all file info csv initialized
$allFileInfo = @()

# path variables for organization purposes
$year = "2021c"
$yearFolderPath = "\$($year)\"
$dstUnzipPath = "\Unzip\"
$workingPath = "\Working\"
$zipOutputPath = "\Zip\"
$forUploadPath = "\ForUpload\"

Write-Host "========================================="
Write-Host "Starting Onboarding Document Processing .. bleep bloop blop"
Write-Host "========================================="
Write-Host "========================================="

# set variables to keep track of files
    $zipCount = 0
    $pdfCount = 0
    $imgCount = 0
    $docCount = 0


# LEVEL 0: YEAR FOLDER (starts here) (see LINE 20)
Write-Host "Working through year: $($year)"
# LEVEL 1: MONTH FOLDERS
Get-ChildItem -Path $yearFolderPath -Directory | ForEach-Object {
    $monthFolder = $_
    Write-Host "`n Looping through Month: $($monthFolder.Name)"

     
    
    # LEVEL 2: EMPLOYEE FOLDERS
    Get-ChildItem -Path $monthFolder.FullName -Directory | ForEach-Object {
        $employeeFolder = $_
        Write-Host "`n Currently working on: $($employeeFolder.Name)'s Folder"
        
        # reset name variables for each employee
        $eeString = ""
        $eeFirstName = ""
        $eeLastName = ""
        
        # LEVEL 3: EMPLOYEE FILES
        Get-ChildItem -Path $employeeFolder.FullName -File | ForEach-Object {
            $file = $_
            
            # zip files
            if ($file.Extension -eq ".zip") {
                Write-Output "found ZIP: $($file.Name)"
                $zipCount++
                
                # split 
                $zipFile = $file.Name
                $zipSplit = $zipFile.Split("_")
                # added version for older files with no 2 appended to EE
                    # $eeString = $zipSplit[0] 
                    $eeString = "2$($zipSplit[0])"
                    $eeFirstName = $zipSplit[1]
                    $eeLastName = $zipSplit[2]
                    Write-Host "Found documents for: $eeFirstName $eeLastName (ee: $eeString)"
                
                # expand the ZIPPPP archive!!
                # try / catch block to throw error instead of silently erroring out
                try {
                    Expand-Archive -LiteralPath $file.FullName -DestinationPath $dstUnzipPath
                    Write-Host "Extracting ZIP to: $dstUnzipPath"
                }
                catch {
                    Write-Error "There was an error extracting ZIP: $($_.FullName)...check manually."
                }
            }
            
            # Process RESUME files (without .pdf extension)
            if ($file.Name -match "RESUME" -and $file.Extension -ne ".pdf") {
                Write-Output "Found non pdf-RESUME: $($file.Name)"
                $newResumeDstFullPath = Join-Path $dstUnzipPath ($file.BaseName + ".pdf")
                $docType = Get-DocumentType -filePath $file.FullName
                
                if ($docType -eq "DOCX" -or $docType -eq "DOC" -or $docType -eq "TXT") {
                    Convert-WordToPDF -wordPath $file.FullName -pdfPath $newResumeDstFullPath
                    # sends extra (converting word to pdf) dialogue
                }
                else {
                    Copy-Item -Path $file.FullName -Destination $newResumeDstFullPath
                }
            }
            
            # Copy existing PDF files
            if ($file.Extension -eq ".pdf" -and $file.Name -notmatch "RESUME") {
                Write-Output "Found PDF: $($file.Name)"
                Copy-Item -Path $file.FullName -Destination $dstUnzipPath
                $pdfCount++
            }
           
            # convert if doc, and then move
            # elseif ($file.Extension -eq ".docx" -or ".doc") {
            #     Convert-WordToPDF -wordPath $file.FullName -pdfPath $dstUnzipPath
            #     $docCount++
            # }
            # all other files (counts for resumes saved as .img)
            #  -and $file.Name -notmatch "RESUME" --> extra logic in case to stop extra resume copying (but doesn't make it to final so it's fine)
            elseif ($file.Extension -ne ".zip" -and $file.Name -notmatch "RESUME") {
                Write-Output "Found / Copying other file: $($file.Name)"
                Copy-Item -Path $file.FullName -Destination $dstUnzipPath
            }         
        }
        
        # skip if no employee zip file found
        # actually this block doesn't work, and it just errors out (but not really in a bad way)
        if ([string]::IsNullOrEmpty($eeString)) {
            Write-Host "WARNING: No zip file found in $($employeeFolder.Name), skipping..."
            continue
        }
        
        Write-Host "`n"
        Write-Host "`n  Now processing unzipped files for $eeFirstName $eeLastName..."
        
        # LEVEL 4: UNZIPPED FILES VALIDATION AND IMG CONVERSION LOOP
        Get-ChildItem -Path $dstUnzipPath -File | ForEach-Object {
            $unzippedFile = $_
            
            # file information variables
            $fileSig = Get-FileSignature -filePath $unzippedFile.FullName
            $fileExt = Get-DocumentType -filePath $unzippedFile.FullName
            $allFileInfo += New-Object PSObject -Property @{
                filePath = $unzippedFile.FullName
                fileSig  = $fileSig
                fileExt  = $fileExt
                eeNumber = $eeString
                eeFirst  = $eeFirstName
                eeLast   = $eeLastName
                month    = $monthFolder.Name
            }
            
            # convert images JPG/PNG --> PDF
            if ($unzippedFile.Extension -match "^\.(jpg|jpeg|png)$") {
                $imgCount++
                $imgFullPath = $unzippedFile.FullName
                $imgDestFullPath = Join-Path $dstUnzipPath ($unzippedFile.BaseName + ".pdf")
                
                Write-Host "Converting image to PDF: $($unzippedFile.Name)"
                
                # declaring variables to be available within catch block
                $imgWriter = $null
                $imgpdf = $null
                $imgDoc = $null
                
                try {
                    # Initialize PDF writer
                    # Initialize PDF document
                    # Initialize document
                    $imgWriter = New-Object iText.Kernel.Pdf.PdfWriter($imgDestFullPath)
                    $imgpdf = New-Object iText.Kernel.Pdf.PdfDocument($imgWriter)
                    $imgDoc = New-Object iText.Layout.Document($imgpdf)
                    
                    # Compose Paragraph with images
                    $imgObject = [iText.Layout.Element.Image]::new([iText.IO.Image.ImageDataFactory]::Create([System.String]$imgFullPath))
                    
                    # Get the page size
                    $pageSize = [iText.Kernel.Geom.PageSize]::A4
                    $width = $pageSize.GetWidth()
                    $height = $pageSize.GetHeight()

                    # Rotate the image 90 degrees anti-clockwise
                    $imgObject.SetRotationAngle([Math]::PI * 1.5)
                    # Scale the image to fit the page
                    $imgObject.SetWidth($width * 0.5)
                    $imgObject.SetHeight($height * 0.5)

                    # Add to paragraph
                    $paragraph = [iText.Layout.Element.Paragraph]::new().Add([iText.Layout.Element.IElement]$imgObject)
                    
                    # Add Paragraph to document
                    $imgDoc.Add([iText.Layout.Element.IBlockElement]$paragraph)
                    # Close document
                    $imgDoc.Close()
                    
                    
                }
                catch {
                    Write-Error "Error converting image: $_"
                    # individually closes these resources to cut down on PDFs being held open for the batch run
                    if ($imgDoc) { $imgDoc.Close() }
                    if ($imgpdf) { $imgpdf.Close() }
                    if ($imgWriter) { $imgWriter.Close() }
                }

                # cleans path after try / catch to guarantee clean-up
                Remove-Item -Path $imgFullPath
            }
            
            # check for Word Docs
            # logic for if resume files get copied  -and $unzippedFile.Name -notmatch "RESUME"
            if ($unzippedFile.Extension -match "^\.(doc|docx)$") {
                $docCount++
                $pdfDestPath = Join-Path $dstUnzipPath ($unzippedFile.BaseName + ".pdf")
                Convert-WordToPDF -wordPath $unzippedFile.FullName -pdfPath $pdfDestPath
                Remove-Item -Path $unzippedFile.FullName
                
            }
            # add delay to avoid timeout?? maybe?? (timeout causing iText7 errors)
            Start-Sleep -Milliseconds 10

            
            # notify about mismatched resume files
            if ($unzippedFile.Name -match "RESUME" -and $unzippedFile.Name -notmatch $eeLastName) {
                Write-Host "This resume does not match the name of the employee. Check it to verify mismatched resume: $($unzippedFile.Name)"
            #     Remove-Item -Path $unzippedFile.FullName
            }
        }
        

        # BUILD COVER PAGE
        Write-Host "`n Creating COVER PAGE for $eeFirstName $eeLastName..."
        # COVER PAGE LOGIC
        $coverPDF = Join-Path $workingPath ($eeString + "_cover.pdf")
        $pdfFileList = New-Object -TypeName 'System.Collections.ArrayList'
        [void]$pdfFileList.Add($coverPDF)
        
        New-PDF {
            $msg = "Onboarding Documents for " + $eeFirstName + " " + $eeLastName
            New-PDFText -Text $msg
            Get-ChildItem $dstUnzipPath -Filter "*.pdf" | Sort-STNumerical | ForEach-Object {
                $pdfCount++
                $fqPathSrc = $_.FullName
                $fqPathDest = $fqPathSrc.replace(" ", "_")
                if ($fqPathSrc -ne $fqPathDest) {
                    Move-Item -Path $fqPathSrc -Destination $fqPathDest -Force
                }
                
                [void]$pdfFileList.Add($fqPathDest)
                $shortPath = $fqPathDest.replace($dstUnzipPath, "")
                New-PDFText -Text $shortPath
            }
        } -FilePath $coverPDF
        # Working ok
        
        # MERGE PDFs
        Write-Host "`n Creating MERGED PDF for $eeFirstName $eeLastName..."
        # MERGE PDF LOGIC
        if ($pdfFileList.Count -gt 1) {
            $pdfOutFile = Join-Path $workingPath ($eeString + "_" + $eeFirstName + "_" + $eeLastName + "_Onboarding.pdf")
            
            $pdfFileCount = $pdfFileList.Count
            for ($i = 0; $i -lt $pdfFileCount - 1; $i++) {
                if ($i -eq 0) {
                    # $mergeOut = Join-Path $workingPath ($eeString + "_workingdoc" + $i + ".pdf")
                    $mergeIn = $pdfFileList[0]
                } 
                $mergeOut = Join-Path $workingPath ($eeString + "_workingdoc" + $i + ".pdf")
            
                Merge-PDF -InputFile $mergeIn, $pdfFileList[$i + 1] -OutputFile $mergeOut -IgnoreProtection
                $mergeIn = $mergeOut
            }
            
            Move-Item -Path $mergeOut -Destination $pdfOutFile -Force
            Remove-Item -Path (Join-Path $workingPath ($eeString + "_workingdoc*.*")) -Force
            
            # ZIP THE FINAL PDF BOSS LOGIC 
            $zipDestPath = Join-Path $zipOutputPath ($eeString + "_" + $eeFirstName + "_" + $eeLastName + "_Onboarding.zip")
            Compress-Archive -Path $pdfOutFile -DestinationPath $zipDestPath -Force
            
            Write-Host "Signed, sealed, and delivered: $zipDestPath"
        }
        else {
            Write-Host "WARNING: No PDFs to merge for $eeFirstName $eeLastName ... that's concerning..."
        }
        
        # CLEANUP FOR NEXT EMPLOYEE LOGIC
        $removeZipPath = Join-Path $dstUnzipPath "*"
        Remove-Item -Path $removeZipPath -Force -ErrorAction SilentlyContinue
        $removeStagePath = Join-Path $workingPath "*"
        Remove-Item -Path $removeStagePath -Force -ErrorAction SilentlyContinue
        
        Write-Host "Completed: $eeFirstName $eeLastName"
    }
    # EXIT EMPLOYEE FOLDER LOOP
}
# EXIT MONTH FOLDER LOOP


# MOVE ALL ZIPS TO FINAL DESTINATION
Write-Host "`n"
Write-Host "`n Moving files to upload folder... $forUploadPath"

# MOVE FINAL ZIPS TO UPLOAD FOLDER LOGIC
Get-ChildItem -Path $zipOutputPath -Filter "*.zip" | ForEach-Object {
    Move-Item -Path $_.FullName -Destination $forUploadPath -Force
}

# EXPORT FILE INFO
$fileInfoFullPath = Join-Path $pwd "fileInfo.csv"
$allFileInfo | Select-Object eeNumber, eeLast, eeFirst, month, filePath, fileSig, fileExt | 
    Export-Csv -Path $fileInfoFullPath -NoTypeInformation

# final notes for user
Write-Host "`n"
Write-Host "========================================="
Write-Host "PROCESSING COMPLETE for year: $($year) ... bleep blop bloop"
Write-Host "========================================="
Write-Host "Total Employees Processed: $($zipCount)"
Write-Host "PDF Count: $($pdfCount)"
Write-Host "IMG Count: $($imgCount)"
Write-Host "DOC Count: $($docCount)"
Write-Host "File Info created at ... $($fileInfoFullPath)"
Write-Host "========================================="
Write-Host "`n"
Write-Host "Thanks for using the PDF merger robot .. bleep bloop"

Stop-Transcript
