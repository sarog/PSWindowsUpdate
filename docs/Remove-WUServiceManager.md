---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Remove-WUServiceManager

## SYNOPSIS
Remove windows update service manager.

## SYNTAX

```
Remove-WUServiceManager [-ComputerName <String[]>] [-ScheduleJob <DateTime>] -ServiceID <String> [-Debuger]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Remove-WUServiceManager cmdlet to unregister Windows Update Service Manager.

## EXAMPLES

### Example 1
```
Try unregister Microsoft Update Service.

Remove-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d"
 
Confirm
Are you sure you want to perform this action?
Performing the operation "Unregister Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".

[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y
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

### -ServiceID
Specify ServiceID of Service Manager that you want to remove.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
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

