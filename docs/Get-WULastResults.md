---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Get-WULastResults

## SYNOPSIS
Get Windows Update results.

## SYNTAX

```
Get-WULastResults [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-Debuger] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Get-WULastResults cmdlet to get Windows Update LastSearchSuccessDate and LastInstallationSuccessDate.

## EXAMPLES

### Example 1
```
Get last Windows Update results.

Get-WULastResults
 
ComputerName LastSearchSuccessDate LastInstallationSuccessDate
------------ --------------------- ---------------------------
MG-PC        30.08.2017 14:02:11   03.07.2017 07:35:27
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

### -PSWUSettings
Required parameter for -SendReport.

Passes the parameters (as hashtable) necessary to send the report: \r\n@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";\[Port=25\];\[Subject="Alternative Subject"\];\[Properties="Alternative object properties"\];\[Style="Table|List"\]}

Send parameters can also be saved to a PSWUSettings.xml file in ModuleBase path: \r\nExport-Clixml @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";\[Port=25\]}"

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

### PSWindowsUpdate.LastResults

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

