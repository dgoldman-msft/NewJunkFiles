---
external help file:
Module Name:
online version:
schema: 2.0.0
---

# New-Junkfile

## SYNOPSIS
Create temp files

## SYNTAX

```
New-Junkfile [[-OutputPath] <String>] [-EmailOutputPath <String>] [-ExcelOutputPath <String>]
 [-TextOutputPath <String>] [-PdfOutputPath <String>] [-WordOutputPath <String>] [[-DefaultType] <Object>]
 [[-MailFrom] <String>] [[-MailTo] <String>] [[-FileSize] <String>] [[-NumberOfWords] <Int32>]
 [[-NumberOfFilesToCreate] <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Create email, word, text, pdf or csv files for a lab for mailflow and migration purposes

## EXAMPLES

### EXAMPLE 1
```
New-Junkfile -DefaultType txt -FileSize StupidLarge -NumberOfFilesToCreate 1 -NumberOfWords 100
```

Will create 1 extremly large txt file

### EXAMPLE 2
```
New-Junkfile -DefaultType Word, Excel -FileSize Large -NumberOfFilesToCreate 25 -NumberOfWords 5
```

Will create 25 large Word and Excel documents

### EXAMPLE 3
```
New-Junkfile -DefaultType Word, Excel, Pdf, Email -FileSize Massive -NumberOfFilesToCreate 50 -NumberOfWords 5
```

Will create 50 massive Word, Excel, exported Pdf's and eml files

## PARAMETERS

### -OutputPath
Default save location

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: C:\temp\JunkFiles
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailOutputPath
Save path for email files

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: C:\temp\JunkFiles\Emails\
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelOutputPath
Save path excel documents

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: C:\temp\JunkFiles\ExcelFiles\
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextOutputPath
Save path text documents

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: C:\temp\JunkFiles\TextFiles\
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfOutputPath
Save path for Pdf files

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: C:\temp\JunkFiles\PdfFiles\
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordOutputPath
Save path for word documents

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: C:\temp\JunkFiles\WordFiles\
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultType
Default file type

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Word
Accept pipeline input: False
Accept wildcard characters: False
```

### -MailFrom
{{ Fill MailFrom Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Administrator@Contoso.com
Accept pipeline input: False
Accept wildcard characters: False
```

### -MailTo
{{ Fill MailTo Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: Administrator@Contoso.com
Accept pipeline input: False
Accept wildcard characters: False
```

### -FileSize
File size you want to generate

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: Tiny
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberOfWords
Number of words per sentance

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: 5
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberOfFilesToCreate
How many files to create

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
Speed on file creation
Fastest -   Emails
            Text Files
            Word & PDF
Slowest     Excel

## RELATED LINKS
