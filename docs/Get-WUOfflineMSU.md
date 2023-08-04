---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Get-WUOfflineMSU

## SYNOPSIS
Get offline MSU package.

## SYNTAX

```
Get-WUOfflineMSU -KBArticleID <String> -Destination <String> [-AcceptAll] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Use Get-WUOfflineMSU cmdlet to download MSU package from Microsoft Update Catalog website.

## EXAMPLES

### Example 1
```
Download package for KB4551762.
PS> Get-WUOfflineMSU -KBArticleID 4551762 -Destination C:\Temp                                              

    Confirm
    Are you sure you want to perform this action?
    Performing the operation "(20.04.2020 14:27:17) 2020-03 Cumulative Update for Windows Server, version 1909 for
    x64-based Systems (KB4551762)[354.4 MB]" on target "DESKTOP-GCQBCBS".
    [Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): N

    Confirm
    Are you sure you want to perform this action?
    Performing the operation "(20.04.2020 14:27:19) 2020-03 Cumulative Update for Windows 10 Version 1909 for x64-based
    Systems(KB4551762) [354.4 MB]" on target "DESKTOP-GCQBCBS".
    [Y] Yes[A] Yes to All [N] No [L] No to All [S] Suspend [?] Help (default is "Y"): Y

    Confirm
    Are you sure you want to perform this action?
    Performing the operation "(20.04.2020 14:27:23) 2020-03 Cumulative Update for Windows 10 Version 1903 for x64-based
    Systems(KB4551762) [354.4 MB]" on target "DESKTOP-GCQBCBS".
    [Y] Yes[A]Yes to All [N] No [L] No to All [S] Suspend [?] Help (default is "Y"): L

    X Result     Title                                                                                 LastUpdated     Size
    - ------     -----                                                                                 -----------     ----
    1 Rejected   2020-03 Cumulative Update for Windows Server, version 1909 for x64-based Systems (... 3/12/2020   354.4 MB
    1 Accepted   2020-03 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB455... 3/12/2020   354.4 MB
    1 Rejected   2020-03 Cumulative Update for Windows 10 Version 1903 for x64-based Systems (KB455... 3/12/2020   354.4 MB
    1 Rejected   2020-03 Cumulative Update for Windows 10 Version 1909 for x86-based Systems (KB455... 3/12/2020   191.7 MB
    1 Rejected   2020-03 Cumulative Update for Windows 10 Version 1903 for x86-based Systems (KB455... 3/12/2020   191.7 MB
    1 Rejected   2020-03 Cumulative Update for Windows 10 Version 1909 for ARM64-based Systems (KB4... 3/12/2020   396.1 MB
    1 Rejected   2020-03 Cumulative Update for Windows Server, version 1903 for x64-based Systems (... 3/12/2020   354.4 MB
    1 Rejected   2020-03 Cumulative Update for Windows 10 Version 1903 for ARM64-based Systems (KB4... 3/12/2020   396.1 MB
    2 Downloaded 2020-03 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB455... 3/12/2020   354.4 MB

PS> Get-ChildItem C:\Temp
    Directory: C:\Temp

    Mode                LastWriteTime         Length    Name
    ----                -------------         ------    ----
    -a----              12.03.2020 07:03      371656774 windows10.0-kb4551762-x64_dacef156c781f2018d94d5a5286076610ba97279.msu
```

## PARAMETERS

### -AcceptAll
Do not ask confirmation for download updates.

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

### -Destination
Destination for downloaded files.

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

### -KBArticleID
Finds updates that contain a KBArticleID.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
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

### System.String

## OUTPUTS

### PSWindowsUpdate.OfflineMSU

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

