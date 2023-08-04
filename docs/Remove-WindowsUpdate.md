---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Remove-WindowsUpdate

## SYNOPSIS
Uninstall update.

## SYNTAX

### WUSAMode
```
Remove-WindowsUpdate [-ComputerName <String[]>] [-ScheduleJob <DateTime>] -KBArticleID <String> [-AutoReboot]
 [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-WUSAMode] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### WUApiMode
```
Remove-WindowsUpdate [-ComputerName <String[]>] [-ScheduleJob <DateTime>] -UpdateID <String> [-AutoReboot]
 [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-WUSAMode] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Remove-WindowsUpdate to uninstall update.

## EXAMPLES

### Example 1
```
Try to uninstall update with specific KBArticleID = KB958830.

Get-WUUninstall -KBArticleID KB958830
```

## PARAMETERS

### -AutoReboot
Do not ask for reboot if it needed.

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

### -IgnoreReboot
Do not ask for reboot if it needed, but do not reboot automaticaly.

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

### -KBArticleID
KBArticleID that will be uninstalled.

```yaml
Type: String
Parameter Sets: WUSAMode
Aliases: HotFixID

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScheduleJob
Specify schedule time job.

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

### -ScheduleReboot
Specify time when system will be rebooted.

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

### -UpdateID
Update ID that will be uninstalled.

```yaml
Type: String
Parameter Sets: WUApiMode
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WUSAMode
Wse wusa.exe instead of WU Api.

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

### System.Object
## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

