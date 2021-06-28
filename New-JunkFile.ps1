function New-Junkfile {
    <#
        .SYNOPSIS
            Create temp files
        
        .DESCRIPTION
            Create email, word, text, pdf or csv files for a lab for mailflow and migration purposes
        
        .PARAMETER OutputPath
            Default save location

        .PARAMETER EmailOutputPath
            Save path for email files

        .PARAMETER ExcelOutputPath
            Save path excel documents

        .PARAMETER TextOutputPath
            Save path text documents

        .PARAMETER PdfOutputPath
            Save path for Pdf files

        .PARAMETER WordOutputPath
            Save path for word documents
        
        .PARAMETER DefaultType
            Default file type
        
        .PARAMETER FileSize
            File size you want to generate
        
        .PARAMETER NumberOfWords
            Number of words per sentance
        
        .PARAMETER NumberOfFilesToCreate
            How many files to create
        
        .EXAMPLE
            PS C:\> New-Junkfile -DefaultType txt -FileSize StupidLarge -NumberOfFilesToCreate 1 -NumberOfWords 100

            Will create 1 extremly large txt file 

        .EXAMPLE
            PS C:\> New-Junkfile -DefaultType Word, Excel -FileSize Large -NumberOfFilesToCreate 25 -NumberOfWords 5

            Will create 25 large Word and Excel documents

        .EXAMPLE
            PS C:\> New-Junkfile -DefaultType Word, Excel, Pdf, Email -FileSize Massive -NumberOfFilesToCreate 50 -NumberOfWords 5

            Will create 50 massive Word, Excel, exported Pdf's and eml files
        
        .NOTES
            Speed on file creation
            Fastest -   Emails
                        Text Files
                        Word & PDF
            Slowest     Excel
    #>

    [CmdletBinding()]
    param (
        [parameter(Position = '0')]
        [string]
        $OutputPath = "c:\temp\JunkFiles",

        [string]
        $EmailOutputPath = "c:\temp\JunkFiles\Emails\",

        [string]
        $ExcelOutputPath = "c:\temp\JunkFiles\ExcelFiles\",
    
        [string]
        $TextOutputPath = "c:\temp\JunkFiles\TextFiles\",

        [string]
        $PdfOutputPath = "c:\temp\JunkFiles\PdfFiles\",

        [string]
        $WordOutputPath = "c:\temp\JunkFiles\WordFiles\",

        [parameter(Position = '1')]
        [ValidateSet('Email', 'Excel', 'Pdf', 'Text', 'Word')]
        [object]
        $DefaultType = 'Word',

        [parameter(Position = '2', ParameterSetName = 'EmailSet')]
        [string]
        $MailFrom = "Administrator@Contoso.com",

        [parameter(Position = '3', ParameterSetName = 'EmailSet')]
        [string]
        $MailTo = "Administrator@Contoso.com",

        [parameter(Position = '4', ParameterSetName = 'FileSet')]
        [ValidateSet('Tiny', 'Small', 'Medium', 'Large', 'Massive', 'StupidLarge')]
        [string]
        $FileSize = "Tiny",

        [parameter(Position = '5', ParameterSetName = 'FileSet')]
        [ValidateRange(1, 100)]
        [int]
        $NumberOfWords = 5,

        [parameter(Position = '6', ParameterSetName = 'FileSet')]
        [Int]
        $NumberOfFilesToCreate = 1
    )
    
    begin {
        Write-Host -ForegroundColor Green "Starting file generation process"
        $directories = @($OutputPath, $EmailOutputPath, $ExcelOutputPath, $TextOutputPath, $WordOutputPath, $PdfOutputPath)
        $script:paragraph = ''
        $script:counter = 0
        $script:emailCounter = 0

        try {
            Write-Verbose "Checking directory structure of: $($OutputPath)"
            foreach ($directory in $directories) {
                if (-NOT (Test-Path -Path $directory)) {
                    $null = New-Item -Path $directory -ItemType Directory -ErrorAction Stop
                }
            }
        }
        catch {
            Write-Host -ForegroundColor Red 'ERROR: $_'
        }
    }
    
    process {
        try {
            if ($DefaultType -eq 'Word' -or $DefaultType -eq 'Excel') {
                Write-Verbose 'Searching GAC for assembly Microsoft.Office.Interop.Word'
                if (-NOT ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -Like '*Microsoft.Office.Interop.Word*' })) {
                    $currentLocation = Get-Location
                    Set-Location (Join-Path -Path $env:HOMEDRIVE -ChildPath '\')
                    Write-Verbose 'Searching for assembly Microsoft.Office.Interop.Word'
                    
                    # Find the assembly so we can load it into the GAC
                    $found = Get-ChildItem -Recurse -Filter 'Microsoft.Office.Interop.Word.dll' -ErrorAction SilentlyContinue
                    if ($found) { 
                        Write-Verbose -Message "Assembly 'Microsoft.Office.Interop.Word.dll' found! Loading assembly"
                        Add-Type -Path (Join-Path -Path $found.Directory.FullName -ChildPath 'Microsoft.Office.Interop.Word.dll') 
                    }
                    else {
                        Write-Host -ForegroundColor Red "Assembly 'Microsoft.Office.Interop.Word.dll' NOT found! This will stop you from creating word Excel and PDF documents!"
                        return
                    }
                    Set-Location $currentLocation
                }
                else {
                    Write-Verbose "Assembly 'Microsoft.Office.Interop.Word.dll' found and is already loaded"
                }
            }
        }
        catch {
            Write-Host -ForegroundColor Red 'ERROR: $_'
        }
        # Generate random document text
        Write-Host -ForegroundColor Green "File size selection: $($FileSize)`r`nNumber of words to use: $($NumberOfWords)"
        1..$NumberOfWords | ForEach-Object { 
            $text = -join ((65..90) + (97..122) | Get-Random -Count 5 | ForEach-Object { [char]$_ })
            $script:paragraph += $text + " " }

        if ($FileSize -eq "Small") { $script:paragraph += $script:paragraph * 500 }
        elseif ($FileSize -eq "Medium") { $script:paragraph += $script:paragraph * 750 }
        elseif ($FileSize -eq "Large") { $script:paragraph += $script:paragraph * 5000 }
        elseif ($FileSize -eq "Massive") { $script:paragraph += $script:paragraph * 50000 }
        elseif ($FileSize -eq "StupidLarge") { $script:paragraph += $script:paragraph * 1000000 }
   
        Write-Host -ForegroundColor Green "Number of files to create: $($NumberOfFilesToCreate)"
        foreach ($fileType in $DefaultType) {
            1..$NumberOfFilesToCreate | ForEach-Object {
                $filename = ""
                (65..90) + (97..122) | Get-Random -Count 10 | ForEach-Object { $filename += [char]$_ }
                switch ($fileType) {

                    "Email" {
                        try {
                            $mailMessage = New-Object System.Net.Mail.MailMessage
    
                            $mailMessage.From = New-Object System.Net.Mail.MailAddress($MailFrom)
                            $mailMessage.To.Add($MailTo)
                            $mailMessage.Subject = "Test .eml file-$script:counter"
                            $mailMessage.Body = $script:paragraph
                            
                            $smtpClient = New-Object System.Net.Mail.SmtpClient
                            $smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory
                            $smtpClient.PickupDirectoryLocation = $EmailOutputPath
                            $smtpClient.Send($mailMessage)
                            $smtpClient.Dispose()
                            
                            $mailMessage.Dispose()
                            Write-Verbose "Email $($emailCounter) created"
                            $emailCounter ++
                        }
                        catch {
                            Write-Host -ForegroundColor Red 'Word Error: $_'
                            return
                        }
                    }
                    
                    "Excel" {
                        try {
                            $null = Start-Job -ScriptBlock {
                                param ( $ExcelOutputPath, $filename, $script:counter)
    
                                $excel = New-Object -ComObject Excel.Application 
                                $excel.Visible = $false
                                $workbook = $excel.Workbooks.Add()
                                $worksheet = $workbook.Worksheets.Item(1)
                                $worksheet.Name = 'Worksheet 1'
                                $row = 1 
                                $column = 1 
                                $worksheet.Cells.Item($row, $column) = 'Row 1'
    
                                # Replace spaces with comma's and split the string to covert string to an array - each item separated by the comma
                                $temp = $Script:paragraph -replace ' ', ','
                                $records = $temp.split(',')
                                $i = 3
                                
                                # SLOWEST and MOST expensive operation right here
                                foreach ($record in $records) {   
                                    $excel.cells.item($i, 1) = $record 
                                    $i++ 
                                }
    
                                $xlsFile = "ExcelDoc-$fileName-$script:counter.xls"
                                [string]$savePath = Join-Path -Path $ExcelOutputPath -ChildPath $xlsFile
                                $workbook.SaveAs($savePath)
                                $excel.Quit()
                                $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
                            } -ArgumentList $ExcelOutputPath, $filename, $script:counter
                        }
                        catch {
                            Write-Host -ForegroundColor Red 'Excel Error: $_'
                            return
                        }
                    }
                                                
                    "Text" {
                        try {
                            $txtFile = "TextDoc-$fileName-$script:counter.txt"
                            $null = New-Item -Path $TextOutputPath -Name $txtFile -ItemType File -Value $paragraph -ErrorAction Stop
                            Write-PSFMessage -Level Verbose -Message "Text doc created: {0}" -StringValues $txtFile
                        }
                        catch {
                            Write-Host -ForegroundColor Red 'Text Error: $_'
                            return
                        }
                    }

                    "Pdf" {
                        try {
                            $customObject = New-Document
                            $pdFile = "testpdf-$fileName-$script:counter.pdf"
                            [string]$savePath = Join-Path -Path $PdfOutputPath -ChildPath $pdFile
                            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
        
                            $customObject.document.SaveAs($savePath, $saveFormat)
                            $customObject.document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
                            $customObject.word.Quit()
                            Write-Verbose "Pdf created: $($pdFile)"
                        }
                        catch {
                            Write-Host -ForegroundColor Red 'PDF Error: $_'
                            return
                        }
                    }

                    "Word" {
                        try {
                            $customObject = New-Document
                            $wordFile = "WordDoc-$fileName-$script:counter.docx"
                            [string]$savePath = Join-Path -Path $WordOutputPath -ChildPath $wordFile
			
                            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
                            $customObject.document.SaveAs($savePath)
                            $customObject.document.Close()
                            $customObject.word.Quit()
                            Write-Verbose "Word doc created: $($wordFile)"
                            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($customObject.word)
                        }
                        catch {
                            Write-Host -ForegroundColor Red 'Word Error: $_'
                            return
                        }
                    }
                }
                $script:counter ++
            }
        }

        # All jobs are kicked off. Clean up all completed and failed jobs
        Write-Host -ForegroundColor Green "Cleaning up background jobs"
        while (Get-Job | Where-Object PSJobTypeName -eq 'BackgroundJob') {
            foreach ($job in (Get-Job)) {
                if (($job.PsJobTypeName -eq 'BackgroundJob') -and ($job.state -eq 'Completed') -or ($job.state -eq 'Failed')) {
                    Write-Verbose "Removing Job Id: $($Job.id)"
                    Remove-Job -Id $Job.id
                }
            }
        }
    }
    
    end {
        Write-Host -ForegroundColor Green "File generation process complete!"
    }
}

function New-Document {
    <#
    .SYNOPSIS
        Create word document

    .DESCRIPTION
        Creates a word document that can be exported as pdf

    .PARAMETER EnableException
        Disables user-friendly warnings and enables the throwing of exceptions. This is less user friendly, but allows catching exceptions in calling scripts.

    .EXAMPLE
        None

    .NOTES
        Internal function
    #>
    
    [OutputType('PSCustomObject')]
    [cmdletbinding()]
    param()

    process {
        try {
            # Create a Word document instance
            $word = New-Object -ComObject word.application
            $word.Visible = $false
            $script:DisposeOfWordObject = $true
            $document = $word.documents.add()

            $selection = $word.Selection
            $selection.Style = "Normal"
            $Selection.TypeParagraph()
            $selection.TypeText($script:paragraph)
        }
        catch {
            Write-Host -ForegroundColor Red 'Word Application Error: $_'
        }
       
        # Return the document
        [PSCustomObject] @{
            Document = $document
            Word     = $word
        }
    }
}
