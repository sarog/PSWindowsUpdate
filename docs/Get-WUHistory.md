---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Get-WUHistory

## SYNOPSIS
Get list of updates history.

## SYNTAX

```
Get-WUHistory [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-MaxDate <DateTime>]
 [-Last <Int32>] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use function Get-WUHistory to get list of installed updates on specific machine.

## EXAMPLES

### Example 1
```
Get Windows Update history.

Get-WUHistory
 
ComputerName Operationname  Date                KB        Title
------------ -------------  ----                --        -----
MG-PC        Installation   30.08.2017 12:03:53 KB2267602 Definition Update for Windows Defender - KB2267602(Defini...
MG-PC        Installation   29.08.2017 11:49:50 KB3186568 Microsoft .NET Framework 4.7 for Windows 10 Version 1607 a...
MG-PC        Installation   29.08.2017 11:30:37 KB4035631 2017-08 Update for Windows Server 2016 for x64-based Syste...
MG-PC        Installation   29.08.2017 11:21:12 KB890830  Windows Malicious Software Removal Tool for Windows 8, 8.1...
MG-PC        Installation   29.08.2017 07:53:36 KB2267602 Definition Update for Windows Defender - KB2267602 (Defini...
MG-PC        Installation   27.08.2017 07:53:39 KB2267602 Definition Update for Windows Defender - KB2267602 (Defini...
MG-PC        Installation   25.08.2017 07:54:38 KB2267602 Definition Update for Windows Defender - KB2267602 (Defini...
MG-PC        Installation   23.08.2017 13:01:26 KB2267602 Definition Update for Windows Defender - KB2267602 (Defini...
MG-PC        Installation   23.08.2017 12:45:45 KB4023307 Security Update for Microsoft Silverlight (KB4023307)
MG-PC        Installation   23.08.2017 07:53:56 KB2267602 Definition Update for Windows Defender - KB2267602 (Defini...
```

### Example 2
```
Get Windows Update Agent history for last 24h.

Get-WUHistory -MaxDate (Get-Date).AddDays(-1)
 
ComputerName Operationname  Date                KB        Title
------------ -------------  ----                --        -----
MG-PC        Installation   30.08.2017 12:03:53 KB2267602 Definition Update for Windows Defender - KB2267602(Defini...
```

## PARAMETERS

### -ComputerName
Specify one or more computer names for remote connection.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Debuger
Debuger return original exceptions.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Last
Last X history entry.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxDate
Filter results by date.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 01.01.0001 00:00:00
Accept pipeline input: False
Accept wildcard characters: False
```

### -PSWUSettings
Required parameter for -SendReport.

Passes the parameters (as hashtable) necessary to send the report: @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";\[Port=25\];\[Subject="Alternative Subject"\];\[Properties="Alternative object properties"\];\[Style="Table|List"\]}

Send parameters can also be saved to a PSWUSettings.xml file in ModuleBase path: Export-Clixml @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";\[Port=25\]}"

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SendReport
Send report email to specific recipients.

Requires the parameter -PSWUSettings or declare the PSWUSettings.xml file in ModuleBase path.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the cmdlet runs. The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String[]

## OUTPUTS

### PSWindowsUpdate.History

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

