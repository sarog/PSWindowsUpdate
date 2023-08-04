---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Get-WURebootStatus

## SYNOPSIS
Get Windows Update reboot status.

## SYNTAX

### ManualReboot (Default)
```
Get-WURebootStatus [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-Silent] [-Debuger]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

### AutoReboot
```
Get-WURebootStatus [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-Silent] [-AutoReboot]
 [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ScheduleReboot
```
Get-WURebootStatus [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-Silent]
 -ScheduleReboot <DateTime> [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### CancelReboot
```
Get-WURebootStatus [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-Silent]
 [-CancelReboot] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Get-WURebootStatus cmdlet to check if reboot is needed.

## EXAMPLES

### Example 1
```
Check if restart is necessary. If yes, ask to do this or don't.

Get-WURebootStatus 
 
Reboot is required. Do it now ? [Y/N] (default is 'N')
ComputerName RebootRequired RebootScheduled
------------ -------------- ---------------
MG-PC        True
```

### Example 2
```
Check if restart is necessary. If yes, then shedule it.

Get-WURebootStatus -ScheduleReboot (Get-Date -Hour 18 -Minute 0 -Second 0)
 
ComputerName RebootRequired RebootScheduled
------------ -------------- ---------------
MG-PC        True           31.08.2017 18:00:00
```

## PARAMETERS

### -AutoReboot
Do not ask for reboot if it needed, but do it now.

```yaml
Type: SwitchParameter
Parameter Sets: AutoReboot
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -CancelReboot
Cancel scheduled reboot.

```yaml
Type: SwitchParameter
Parameter Sets: CancelReboot
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ComputerName
Specify one or more computer names for remote connection.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
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

### -ScheduleReboot
Specify schedule time for reboot.

```yaml
Type: DateTime
Parameter Sets: ScheduleReboot
Aliases:

Required: True
Position: Named
Default value: 01.01.0001 00:00:00
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

### -Silent
Return true/false only.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: True (ByPropertyName, ByValue)
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

### System.Management.Automation.SwitchParameter

## OUTPUTS

### PSWindowsUpdate.RebootStatus

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

