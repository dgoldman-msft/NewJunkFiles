# NewJunkFiles

This is a test script that you can run in a lab environment that will facilitate the creation of multiple file types (Word, Excel, PDF, Text, Email (Eml)).
## To get started with this script

1. Copy down to a local directory
2. Run . .\NewJunkFiles.ps1 - This will import the script in to the local PowerShell session

> EXAMPLE 1: New-Junkfile -DefaultType Text -FileSize StupidLarge -NumberOfFilesToCreate 1 -NumberOfWords 100

Will create 1 extremely large text file 

> EXAMPLE 2: New-Junkfile -DefaultType Word, Excel -FileSize Large -NumberOfFilesToCreate 25 -NumberOfWords 5

Will create 25 large Word and Excel documents

> EXAMPLE 3: New-Junkfile -DefaultType Word, Excel, Pdf, Email -FileSize Massive -NumberOfFilesToCreate 50 -NumberOfWords 5

Will create 50 massive Word, Excel, exported Pdf's and eml files

For more help information please reference: [New-Junkfile Help File](https://github.com/dgoldman-msft/NewJunkFiles/blob/main/docs/New-Junkfile.md)
