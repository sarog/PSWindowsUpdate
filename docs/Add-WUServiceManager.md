---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Add-WUServiceManager

## SYNOPSIS
Register a new Windows Update API Service Manager.

## SYNTAX

### LocalServiceID (Default)
```
Add-WUServiceManager [-SendReport] [-PSWUSettings <Hashtable>] -ServiceID <String> [-AddServiceFlag <Int32>]
 [-AuthorizationCabPath <String>] [-Silent] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### RemoteServiceID
```
Add-WUServiceManager [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] -ServiceID <String>
 [-Silent] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### MicrosoftUpdate
```
Add-WUServiceManager [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-MicrosoftUpdate]
 [-Silent] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### OfflineSync
```
Add-WUServiceManager [-SendReport] [-PSWUSettings <Hashtable>] [-ServiceName <String>]
 -ScanFileLocation <String> [-Silent] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Add-WUServiceManager cmdlet to register new Windows Update Service Manager.

It's combination old 'Add-WUServiceManager' and 'Add-WUOfflineSync' functions to register online and offline ServiceManager

## EXAMPLES

### Example 1
```
Try register Microsoft Update Service by custom ServiceID.

Add-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d"
 
Confirm
Are you sure you want to perform this action?
Performing the operation "Register Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ServiceID                            IsManaged IsDefault Name
---------                            --------- --------- ----
7971f918-a847-4430-9279-4a52d1efe18d False     False     Microsoft Update
```

### Example 2
```
Try register Microsoft Update service as Service Manager.

Add-WUServiceManager -MicrosoftUpdate

Confirm
Are you sure you want to perform this action?
Performing the operation "Register Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ServiceID                            IsManaged IsDefault Name
---------                            --------- --------- ----
7971f918-a847-4430-9279-4a52d1efe18d False     False     Microsoft Update
```

### Example 3
```
Try register Offline Sync Service from file C:\wsusscn2.cab.

Add-WUServiceManager -ScanFileLocation C:\wsusscn2.cab
 
Confirm
Are you sure you want to perform this action?
Performing the operation "Register Offline Windows Update Service Manager: C:\wsusscn2.cab" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ServiceID                            IsManaged IsDefault Name
---------                            --------- --------- ----
7e1364ef-e30e-4f4e-9c66-84194eebcbbe False     False     Offline Sync Service
```

## PARAMETERS

### -AddServiceFlag
A combination of AddServiceFlag values: \r\n0x1 - asfAllowPendingRegistration \r\n0x2 - asfAllowOnlineRegistration \r\n0x4 - asfRegisterServiceWithAU

```yaml
Type: Int32
Parameter Sets: LocalServiceID
Aliases:

Required: False
Position: Named
Default value: 2
Accept pipeline input: False
Accept wildcard characters: False
```

### -AuthorizationCabPath
The path of the Microsoft signed local cabinet file (.cab) that has the information that is required for a service registration.
If empty, the update agent searches for the authorization cabinet file (.cab) during service registration when a network connection is available.

```yaml
Type: String
Parameter Sets: LocalServiceID
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ComputerName
Specify one or more computer names for remote connection.

```yaml
Type: String[]
Parameter Sets: RemoteServiceID, MicrosoftUpdate
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

### -MicrosoftUpdate
Register Microsoft Update Service Manager - '7971f918-a847-4430-9279-4a52d1efe18d'

```yaml
Type: SwitchParameter
Parameter Sets: MicrosoftUpdate
Aliases:

Required: True
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

### -ScanFileLocation
Path to Windows Update offline scan file (wsusscan.cab or wsusscn2.cab).

```yaml
Type: String
Parameter Sets: OfflineSync
Aliases:

Required: True
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

### -ServiceID
An identifier for the service to be registered.

Examples Of ServiceID: \r\n \r\n -- Windows Update 9482f4b4-e343-43b6-b170-9a65bc822c77 \r\n -- Microsoft Update 7971f918-a847-4430-9279-4a52d1efe18d \r\n -- Windows Store 117cab2d-82b1-4b5a-a08c-4d62dbee7782 \r\n -- Windows Server Update Service 3da21691-e39d-4da6-8a4b-b43877bcb1b7

```yaml
Type: String
Parameter Sets: LocalServiceID, RemoteServiceID
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ServiceName
Name under which it will be registered Windows Update offline service.
Default name is 'Offline Sync Service'.

```yaml
Type: String
Parameter Sets: OfflineSync
Aliases:

Required: False
Position: Named
Default value: Offline Sync Service
Accept pipeline input: False
Accept wildcard characters: False
```

### -Silent
Don't return output.

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

### PSWindowsUpdate.ServiceManager

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

