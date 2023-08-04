---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Enable-WURemoting

## SYNOPSIS
Enable firewall rules for PSWindowsUpdate remoting.

## SYNTAX

```
Enable-WURemoting [-LocalAccountTokenFilterPolicy] [-WinRMPublic] [-Debuger] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Use Enable-WURemoting cmdlet to enable nessesery firewall rules for PSWindowsUpdate remoting.

## EXAMPLES

### Example 1
```
Enable firewall rules for PSWindowsUpdate remoting.

Enable-WURemoting -Verbose
```

## PARAMETERS

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

### -LocalAccountTokenFilterPolicy
Set LocalAccountTokenFilterPolicy registry entry to builds an elevated token on the target remote computer.

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

### -WinRMPublic
Enable WinRM Public access from all subnets.
Default access is only enabled from local subnet.
Required for workgroup computers.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

