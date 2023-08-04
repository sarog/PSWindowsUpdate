---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Set-WUSettings

## SYNOPSIS
Set Windows Update Client settings.

## SYNTAX

```
Set-WUSettings [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>]
 [-AcceptTrustedPublisherCerts] [-DisableWindowsUpdateAccess] [-NonAdministratorsElevated]
 [-TargetGroup <String>] [-TargetGroupEnabled] [-WUServer <String>] [-NotificationLevel <String>]
 [-AutoInstallMinorUpdates] [-DetectionFrequency <Int32>] [-DetectionFrequencyEnabled]
 [-IncludeRecommendedUpdates] [-NoAutoRebootWithLoggedOnUsers] [-NoAutoUpdate] [-RebootRelaunchTimeout <Int32>]
 [-RebootRelaunchTimeoutEnabled] [-RebootWarningTimeout <Int32>] [-RebootWarningTimeoutEnabled]
 [-RescheduleWaitTime <Int32>] [-ScheduledInstallDay <String>] [-ScheduledInstallTime <Int32>] [-UseWUServer]
 [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Set-WUSettings cmdlet to Set configuration of Windows Update Client.

## EXAMPLES

### Example 1
```
Enable IncludeRecommendedUpdates attributes.

Set-WUSettings -IncludeRecommendedUpdates
 
Confirm
Are you sure you want to perform this action?
Performing the operation "Set Windows Update settings" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ComputerName IncludeRecommendedUpdates
------------ -------------------------
MG-PC        True
```

## PARAMETERS

### -AcceptTrustedPublisherCerts
Enabled - The WSUS server distributes available signed non-Microsoft updates.

Disabled - The WSUS server does not distribute available signed non-Microsoft updates.

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

### -AutoInstallMinorUpdates
Silently install minor updates.

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

### -DetectionFrequency
Time between detection cycles.
Time in hours (1-22).

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

### -DetectionFrequencyEnabled
Enable/Disable detection frequency.

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

### -DisableWindowsUpdateAccess
Disables/Enables access to Windows Update.

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

### -IncludeRecommendedUpdates
Enable/Disable recommended updates.

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

### -NoAutoRebootWithLoggedOnUsers
Logged-on user can decide whether to restart the client computer.

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

### -NoAutoUpdate
Enable/Disable Automatic Updates.

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

### -NonAdministratorsElevated
Enabled - All members of the Users security group can approve or disapprove updates.

Disabled - Only members of the Administrators security group can approve or disapprove updates.

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

### -NotificationLevel
Notify mode: "Not configured", "Disabled", "Notify before download", "Notify before installation", "Scheduled installation", "Users configure"

```yaml
Type: String
Parameter Sets: (All)
Aliases: AUOptions
Accepted values: Not configured, Disabled, Notify before download, Notify before installation, Scheduled installation, Users configure

Required: False
Position: Named
Default value: None
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

### -RebootRelaunchTimeout
Time between prompts for a scheduled restart.
Time in minutes (1-1440).

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

### -RebootRelaunchTimeoutEnabled
Enable/Disable RebootRelaunchTimeout.

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

### -RebootWarningTimeout
Length, in minutes, of the restart warning countdown after updates have been installed that have a deadline or scheduled updates.
Time in minutes (1-30).

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

### -RebootWarningTimeoutEnabled
Enable/Disable RebootWarningTimeout.

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

### -RescheduleWaitTime
Time in minutes that Automatic Updates waits at startup before it applies updates from a missed scheduled installation time.
Time in minutes (1-60).

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

### -ScheduledInstallDay
Scheduled day of install: "Every day", "Every Sunday", "Every Monday", "Every Tuesday", "Every Wednesday", "Every Thursday", ""Every Friday", "EverySaturday".
Only valid if NotificationLevel (AUOptions) = "Scheduled installation"

Starting with Windows 8 and Windows Server 2012, ScheduledInstallationDay are not supported and will return unreliable values.If you try to modify these properties, the operation will appear to succeed but will have no effect.

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Every Day, Every Sunday, Every Monday, Every Tuesday, Every Wednesday, Every Thursday, Every Friday, EverySaturday

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScheduledInstallTime
Scheduled time of install in 24-hour format(0-23).

Starting with Windows 8 and Windows Server 2012, ScheduledInstallTime are not supported and will return unreliable values.If you try to modify these properties, the operation will appear to succeed but will have no effect.

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

### -TargetGroup
Name of the computer group to which the computer belongs.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetGroupEnabled
Use/Do not use client-side targeting.

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

### -UseWUServer
The computer gets its updates from a WSUS server or from Microsoft Update.

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

### -WUServer
HTTP(S) URL of the WSUS server that is used by Automatic Updates and API callers (by default).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
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

### PSWindowsUpdate.WUSettings

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

