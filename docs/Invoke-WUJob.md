---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Invoke-WUJob

## SYNOPSIS
Invoke WUJobs by Task Schduler.

## SYNTAX

### RunNow (Default)
```
Invoke-WUJob [-ComputerName <String[]>] [-Credential <PSCredential>] [-TaskName <String>] [-Hidden]
 [-EndBoundary <DateTime>] [-RunNow] [-Script <String>] [-Force] [-Debuger] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### TriggerDate
```
Invoke-WUJob [-ComputerName <String[]>] [-Credential <PSCredential>] [-TaskName <String>] [-Hidden]
 [-EndBoundary <DateTime>] -TriggerDate <DateTime> [-Script <String>] [-Force] [-Debuger] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### TriggerAtStart
```
Invoke-WUJob [-ComputerName <String[]>] [-Credential <PSCredential>] [-TaskName <String>] [-Hidden]
 [-EndBoundary <DateTime>] [-TriggerAtStart] [-Script <String>] [-Force] [-Debuger] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Use Invoke-WUJobs cmdlet to invoke PSWindowsUpdate actions remotly.
It Based on TaskScheduler because CreateUpdateDownloader() and CreateUpdateInstaller() methods can't be called from a remote computer - E_ACCESSDENIED.

Note: Because we do not have the ability to interact, is recommended use -AcceptAll for Install-WindowsUpdate.

## EXAMPLES

### Example 1
```
Invoke Install-WindowsUpdate on remote machine today at 6:00 PM.

Invoke-WUJob -ComputerName MG-PC -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\PSWindowsUpdate.log" -TriggerDate (Get-Date -Hour 18 -Minute 0 -Second 0)
 
Confirm
Are you sure you want to perform this action?
Performing the operation "Invoke WU job: powershell.exe -Command "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\PSWindowsUpdate.log"" on target "MG-PC".
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

### -Credential
Specify alternative credential.

```yaml
Type: PSCredential
Parameter Sets: (All)
Aliases:

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

### -EndBoundary
Specify EndBoundary for Task Scheduler trigger.

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

### -Force
Force update old Task Scheduler Job if it's available.

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

### -Hidden
Specify if task must be hidden.
Default is true.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: True
Accept pipeline input: False
Accept wildcard characters: False
```

### -RunNow
Run Task Scheduler job immediately.

```yaml
Type: SwitchParameter
Parameter Sets: RunNow
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Script
Specify PowerShell script that you what to run.
Default is {ipmo PSWindowsUpdate; Get-WindowsUpdate -AcceptAll -Install | Out-File $Env:TEMP\PSWindowsUpdate.log}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Ipmo PSWindowsUpdate; Get-WindowsUpdate -AcceptAll -Install | Out-File $Env:TEMP\PSWindowsUpdate.log
Accept pipeline input: False
Accept wildcard characters: False
```

### -TaskName
Specify custom name for Task Scheduler job.
Default is 'PSWindowsUpdate'.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: PSWindowsUpdate
Accept pipeline input: False
Accept wildcard characters: False
```

### -TriggerAtStart
Specify system startup trigger for Task Scheduler job.

```yaml
Type: SwitchParameter
Parameter Sets: TriggerAtStart
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -TriggerDate
Specify Time trigger for Task Scheduler job.

```yaml
Type: DateTime
Parameter Sets: TriggerDate
Aliases:

Required: True
Position: Named
Default value: 01.01.0001 00:00:00
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

