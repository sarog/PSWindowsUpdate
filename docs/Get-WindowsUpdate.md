---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Get-WindowsUpdate

## SYNOPSIS
Get list of available updates meeting the criteria.

## SYNTAX

### Default (Default)
```
Get-WindowsUpdate [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-SendHistory]
 [-ScheduleJob <DateTime>] [-AcceptAll] [-RecurseCycle <Int32>] [-Hide] [-Download] [-ForceDownload] [-Install]
 [-ForceInstall] [-AutoReboot] [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-Criteria <String>]
 [-UpdateType <String>] [-DeploymentAction <String>] [-IsAssigned] [-IsPresent] [-BrowseOnly]
 [-AutoSelectOnWebSites] [-UpdateID <String[]>] [-NotUpdateID <String[]>] [-RevisionNumber <Int32>]
 [-CategoryIDs <String[]>] [-IsInstalled] [-IsHidden] [-WithHidden] [-ShowPreSearchCriteria]
 [-RootCategories <String[]>] [-Category <String[]>] [-KBArticleID <String[]>] [-Title <String>]
 [-Severity <String[]>] [-NotCategory <String[]>] [-NotKBArticleID <String[]>] [-NotTitle <String>]
 [-NotSeverity <String[]>] [-IgnoreUserInput] [-IgnoreRebootRequired] [-AutoSelectOnly] [-MaxSize <Int64>]
 [-MinSize <Int64>] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ServiceID
```
Get-WindowsUpdate [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-SendHistory]
 [-ScheduleJob <DateTime>] [-AcceptAll] [-RecurseCycle <Int32>] [-Hide] [-Download] [-ForceDownload] [-Install]
 [-ForceInstall] [-AutoReboot] [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-ServiceID <String>]
 [-Criteria <String>] [-UpdateType <String>] [-DeploymentAction <String>] [-IsAssigned] [-IsPresent]
 [-BrowseOnly] [-AutoSelectOnWebSites] [-UpdateID <String[]>] [-NotUpdateID <String[]>]
 [-RevisionNumber <Int32>] [-CategoryIDs <String[]>] [-IsInstalled] [-IsHidden] [-WithHidden]
 [-ShowPreSearchCriteria] [-RootCategories <String[]>] [-Category <String[]>] [-KBArticleID <String[]>]
 [-Title <String>] [-Severity <String[]>] [-NotCategory <String[]>] [-NotKBArticleID <String[]>]
 [-NotTitle <String>] [-NotSeverity <String[]>] [-IgnoreUserInput] [-IgnoreRebootRequired] [-AutoSelectOnly]
 [-MaxSize <Int64>] [-MinSize <Int64>] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### WindowsUpdate
```
Get-WindowsUpdate [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-SendHistory]
 [-ScheduleJob <DateTime>] [-AcceptAll] [-RecurseCycle <Int32>] [-Hide] [-Download] [-ForceDownload] [-Install]
 [-ForceInstall] [-AutoReboot] [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-WindowsUpdate]
 [-Criteria <String>] [-UpdateType <String>] [-DeploymentAction <String>] [-IsAssigned] [-IsPresent]
 [-BrowseOnly] [-AutoSelectOnWebSites] [-UpdateID <String[]>] [-NotUpdateID <String[]>]
 [-RevisionNumber <Int32>] [-CategoryIDs <String[]>] [-IsInstalled] [-IsHidden] [-WithHidden]
 [-ShowPreSearchCriteria] [-RootCategories <String[]>] [-Category <String[]>] [-KBArticleID <String[]>]
 [-Title <String>] [-Severity <String[]>] [-NotCategory <String[]>] [-NotKBArticleID <String[]>]
 [-NotTitle <String>] [-NotSeverity <String[]>] [-IgnoreUserInput] [-IgnoreRebootRequired] [-AutoSelectOnly]
 [-MaxSize <Int64>] [-MinSize <Int64>] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### MicrosoftUpdate
```
Get-WindowsUpdate [-ComputerName <String[]>] [-SendReport] [-PSWUSettings <Hashtable>] [-SendHistory]
 [-ScheduleJob <DateTime>] [-AcceptAll] [-RecurseCycle <Int32>] [-Hide] [-Download] [-ForceDownload] [-Install]
 [-ForceInstall] [-AutoReboot] [-IgnoreReboot] [-ScheduleReboot <DateTime>] [-MicrosoftUpdate]
 [-Criteria <String>] [-UpdateType <String>] [-DeploymentAction <String>] [-IsAssigned] [-IsPresent]
 [-BrowseOnly] [-AutoSelectOnWebSites] [-UpdateID <String[]>] [-NotUpdateID <String[]>]
 [-RevisionNumber <Int32>] [-CategoryIDs <String[]>] [-IsInstalled] [-IsHidden] [-WithHidden]
 [-ShowPreSearchCriteria] [-RootCategories <String[]>] [-Category <String[]>] [-KBArticleID <String[]>]
 [-Title <String>] [-Severity <String[]>] [-NotCategory <String[]>] [-NotKBArticleID <String[]>]
 [-NotTitle <String>] [-NotSeverity <String[]>] [-IgnoreUserInput] [-IgnoreRebootRequired] [-AutoSelectOnly]
 [-MaxSize <Int64>] [-MinSize <Int64>] [-Debuger] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Get-WindowsUpdate (aka Get-WUList) cmdlet to get list of available or installed updates meeting specific criteria.

Use Download-WindowsUpdate alias to get list of updates and download it.
Equivalent Get-WindowsUpdate -Download.

Use Install-WindowsUpdate (aka Get-WUInstall) alias to get list of updates and install it.
Equivalent Get-WindowsUpdate -Install.

Use Hide-WindowsUpdate alias to get list of updates and hide it.
Equivalent Get-WindowsUpdate -Hide.

Use Show-WindowsUpdate (aka UnHide-WindowsUpdate) alias to get list of updates and unhide it.
Equivalent Get-WindowsUpdate -Hide:$false.

There are two types of filtering update: Pre search criteria, Post search criteria.

- Pre search works on server side, like example: (IsInstalled = 0 and IsHidden = 0 and CategoryIds contains '0fa1201d-4330-4fa8-8ae9-b877473b6441' )
- Post search work on client side after get the pre-filtered list of updates, like example $KBArticleID -match $Update.KBArticleIDs

Status info list:\r\n\[A|R\]DIMHUB\r\nA-IsAccetped\r\nR-IsRejected\r\n D-IsDownloaded\r\n F-DownloadFailed\r\n ?-IsInvoked\r\n I-IsInstalled\r\n F-InstallFailed\r\n ?-IsInvoked\r\n R-RebootRequired\r\n M-IsMandatory\r\n H-IsHidden\r\n U-IsUninstallable\r\n B-IsBeta

## EXAMPLES

### Example 1
```
Get windows updates available from default service manager.

Get-WindowsUpdate -Verbose
 
VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...
VERBOSE: Found[4] Updates in pre search criteria
VERBOSE: Found[4] Updates in post search criteria

ComputerName Status     KB          Size Title
------------ ------     --          ---- -----
MG-PC        -------    KB890830    44MB Narzędzie Windows do usuwania złośliwego oprogramowania dla systemów Window...
MG-PC        -------    KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...
MG-PC        -------    KB4034662   21MB 2017-08 Aktualizacja zabezpieczeń Adobe Flash Player w Windows 10 Version 1...
MG-PC        -------    KB4035631   11MB 2017-08 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...
```

### Example 2
```
Get all installed drivers that are available at Windows Update. Additionaly show pre search criteria.

Get-WindowsUpdate -WindowsUpdate -UpdateType Driver -IsInstalled -ShowPreSearchCriteria -Verbose
 
PreSearchCriteria: IsInstalled = 0 and Type = 'Driver' and IsHidden = 0
VERBOSE: MG-PC: Connecting to Windows Update server.Please wait...
VERBOSE: Found[1] Updates in pre search criteria
VERBOSE: Found[1] Updates in post search criteria

ComputerName Status     KB          Size Title
------------ ------     --          ---- -----
MGAJDALAP3   -DI----                 3MB Intel - Other hardware - Intel(R) Watchdog Timer Driver (Intel(R) WDT)
```

### Example 3
```
Get all available update on remote machine MG-PC, that contains in Title this two words 'Aktualizacja' and 'Windows 10' (as regular expression).

Get-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -Title "Aktualizacja.*Windows 10" -Verbose
 
VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...
VERBOSE: Found[14] Updates in pre search criteria
VERBOSE: Found[5] Updates in post search criteria

ComputerName Status     KB          Size Title
------------ ------     --          ---- -----
MG-PC        -------    KB3150513    2MB 2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...
MG-PC        -------    KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...
MG-PC        -------    KB4034662   21MB 2017-08 Aktualizacja zabezpieczeń Adobe Flash Player w Windows 10 Version 1...
MG-PC        -------    KB4035631   11MB 2017-08 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...
MG-PC        -------    KB4033637    4MB Aktualizacja systemu Windows 10 Version 1607 dla komputerów z procesorami x...
```

### Example 4
```
Hide update with KBArticleID: KB4034658.

Get-WindowsUpdate -KBArticleID KB4034658 -Hide -Verbose
or use alias
Hide-WindowsUpdate -KBArticleID KB4034658 -Verbose
 
VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...
VERBOSE: Found[4] Updates in pre search criteria
VERBOSE: Found[1] Updates in post search criteria

Confirm
Are you sure you want to perform this action?
Performing the operation "Hide 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB4034658)[1GB]" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ComputerName Status     KB          Size Title
------------ ------     --          ---- -----
MG-PC        ---H--     KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...
```

### Example 5
```
Unhide update with KBArticleID: KB4034658.

Get-WindowsUpdate -KBArticleID KB4034658 -WithHidden -Hide:$false -Verbose 
or use alias
Show-WindowsUpdate -KBArticleID KB4034658 -Verbose
 
VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...
VERBOSE: Found[4] Updates in pre search criteria
VERBOSE: Found[1] Updates in post search criteria

Confirm
Are you sure you want to perform this action?
Performing the operation "Show 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB4034658)[1GB]" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

ComputerName Status     KB          Size Title
------------ ------     --          ---- -----
MG-PC        ------     KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...
```

### Example 6
```
Schedule job at 6:00 PM to install update with UpdateId='ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber=200. Update will be automaticaly accepted and after all serwer will be automaticaly restarted if needed. 

Get-WindowsUpdate -MicrosoftUpdate -UpdateID ddb74579-7a1f-4d1f-80c8-e8647055314e -RevisionNumber 200 -ScheduleJob (Get-Date -Hour 18 -Minute 0 -Second 0) -Install -AcceptAll -AutoReboot -Verbose
or use alias
Install-WindowsUpdate -MicrosoftUpdate -UpdateID ddb74579-7a1f-4d1f-80c8-e8647055314e -RevisionNumber 200 -ScheduleJob (Get-Date -Hour 18 -Minute 0 -Second 0) -AcceptAll -AutoReboot -Verbose
 
VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...
VERBOSE: Found[1] Updates in pre search criteria
VERBOSE: Found[1] Updates in post search criteria
VERBOSE: Choosed pre Search Criteria: (UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200)

X ComputerName Result     KB          Size Title
- ------------ ------     --          ---- -----
1 MG-PC        Accepted   KB4023307   13MB Microsoft Silverlight(KB4023307)
VERBOSE: Accepted[1] Updates ready to Download
VERBOSE: Invoke-WUJob: MG-PC(31.08.2017 18:00:00):
VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200)\" -AcceptAll -AutoReboot -Download -Install -MicrosoftUpdate -Verbose *>&1 | Out-File $Env:TEMP\PSWindowsUpdate.log"
```

### Example 7
```
Install updates on remote computer. After all send a report from the installation process.

Install-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -AcceptAll -AutoReboot -SendReport -PSWUSettings @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} -Verbose
or use global PSWUSettings
@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} | Export-Clixml -Path 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate\PSWUSettings.xml'
Install-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -AcceptAll -AutoReboot -SendReport -Verbose
 
VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...
VERBOSE: Found[4] Updates in pre search criteria
VERBOSE: Found[4] Updates in post search criteria

Confirm
Are you sure you want to perform this action?
Performing the operation "Microsoft Silverlight (KB4023307)[13MB]" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

Confirm
Are you sure you want to perform this action?
Performing the operation "2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB3150513)[2MB]" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y

Confirm
Are you sure you want to perform this action?
Performing the operation "Aktualizacja pakietu językowego usługi Microsoft Dynamics 365 2.1[47MB]" on target "MG-PC".
[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): L

X ComputerName Result     KB          Size Title
- ------------ ------     --          ---- -----
1 MG-PC Accepted KB4023307   13MB Microsoft Silverlight (KB4023307)
1 MG-PC Accepted KB3150513    2MB 2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na arc...
1 MG-PC Rejected KB4013759   47MB Aktualizacja pakietu językowego usługi Microsoft Dynamics 365 2.1
1 MG-PC Rejected KB3186568   67MB Program Microsoft .NET Framework 4.7 w syst. Windows 10 Version 1607 i Wi...
VERBOSE: Accepted [2]
Updates ready to Download
VERBOSE: Invoke-WUJob: MG-PC (Now):
VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200) or (UpdateID = '151c4402-513c-4f39-8da1-f84d0956b5e3' and RevisionNumber = 200)\" -AcceptAll -Download -Install -AutoReboot -MicrosoftUpdate -SendReport -ProofOfLife -Verbose *>&1 | Out-File $Env:TEMP\PSWindowsUpdate.log"
```

### Example 8
```
Schedule Job to install all available updates and automatically reboot system if needed. Also send report after installation (but before reboot if needed) and send second instalation history report after reboot.

@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} | Export-Clixml -Path 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate\PSWUSettings.xml'
Install-WindowsUpdate -MicrosoftUpdate -SendReport -SendHistory -AcceptAll -AutoReboot -ScheduleJob (Get-Date -Hour 18 -Minute 30 -Second 0) -ComputerName MG-PC -Verbose
 
VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...
VERBOSE: Found[4] Updates in pre search criteria
VERBOSE: Found[4] Updates in post search criteria

X ComputerName Result     KB          Size Title
- ------------ ------     --          ---- -----
1 MG-PC        Accepted   KB3038936    5MB Aktualizacja systemu Windows 8.1 dla komputerów z procesorami x64(KB3038...
1 MG-PC        Accepted   KB3186606    4MB Pakiety językowe programu Microsoft.NET Framework 4.7 w syst. Windows 8....
1 MG-PC        Accepted   KB4035038   53MB Sierpień 2017: wersja zapozn. pak.zb.aktual.jakości dla pr. .NET Frame...
1 MG-PC        Accepted   KB2267602  309MB Aktualizacja definicji dla: Windows Defender - KB2267602 (Definicja 1.251...
VERBOSE: Accepted[4] Updates ready to Download
VERBOSE: Invoke-WUJob: MG-PC (02.09.2017 08:30:00):
VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'e69c9679-7ce8-489a-a21c-62fb920be67a' and RevisionNumber = 201) or(UpdateID = 'de44604d-ec38-4a7f-ac63-28b3edfdb382' and RevisionNumber = 207) or(UpdateID = '9cf1d8c9-a7c3-4603-90e8-f22131ff6d7e' and RevisionNumber = 201) or(UpdateID = 'b51935f9-0e40-4624-9c26-b29bff92dcf9' and RevisionNumber = 200)\" -AcceptAll -Install -AutoReboot -MicrosoftUpdate -SendReport -SendHistory -Verbose *>&1 | Out-File $Env:TEMP\PSWindowsUpdate.log"
VERBOSE: Send report
```

## PARAMETERS

### -AcceptAll
Do not ask confirmation for updates.
Download or Install all available updates.

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

### -AutoSelectOnWebSites
Pre search criteria - native for WUAPI.
Finds updates where the AutoSelectOnWebSites property has the specified value.

"AutoSelectOnWebSites=1" finds updates that are flagged to be automatically selected by Windows Update.

"AutoSelectOnWebSites=0" finds updates that are not flagged for Automatic Updates.

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

### -AutoSelectOnly
Install only the updates that have status AutoSelectOnWebsites on true.

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

### -BrowseOnly
Pre search criteria - native for WUAPI.
"BrowseOnly=1" finds updates that are considered optional.
"BrowseOnly=0" finds updates that are not considered optional.

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

### -Category
Post search criteria.
Finds updates that contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...

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

### -CategoryIDs
Pre search criteria - native for WUAPI.
Finds updates that belong to a specified category (or sets of UUIDs), such as '0fa1201d-4330-4fa8-8ae9-b877473b6441'.

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

### -ComputerName
Specify one or more computer names for remote connection.
Interactive remote connection works only for checking updates.
For download or install cmdlet creates an Invoke-WUJob task.

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

### -Criteria
Pre search criteria - native for WUAPI.
Set own string that specifies the search criteria.
https://docs.microsoft.com/pl-pl/windows/desktop/api/wuapi/nf-wuapi-iupdatesearcher-search

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

### -Debuger
Debuger return original exceptions.
For additional debug information use $DebugPreference = "Continue"

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

### -DeploymentAction
Pre search criteria - native for WUAPI.
Finds updates that are deployed for a specific action, such as an installation or uninstallation that the administrator of a server specifies.
"DeploymentAction='Installation'" finds updates that are deployed for installation on a destination computer.
"DeploymentAction='Uninstallation'" depends on the other query criteria.

"DeploymentAction='Uninstallation'" finds updates that are deployed for uninstallation on a destination computer.
"DeploymentAction='Uninstallation'" depends on the other query criteria.

If this criterion is not explicitly specified, each group of criteria that is joined to an AND operator implies "DeploymentAction='Installation'".

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Installation, Uninstallation

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Download
Get list of updates and download approved updates, but do not install it.

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

### -ForceDownload
Forces the download of updates that are already installed or that cannot be installed.
Works only with -Download.

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

### -ForceInstall
A forced installation is an installation in which an update is installed even if the metadata indicates that the update is already installed.
Before you use ForceInstall to force an installation, determine whether the update is installed and available.
If an update is not installed, a forced installation fails.
Works only with -Install.

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

### -Hide
Get list of updates and hide/unhide approved updates.

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

### -IgnoreRebootRequired
Post search criteria.
Finds updates that specifies the restart behavior that not occurs when you install or uninstall the update.

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

### -IgnoreUserInput
Post search criteria.
Finds updates that the installation or uninstallation of an update can't prompt for user input.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: Silent

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Install
Get list of updates and install approved updates.

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

### -IsAssigned
Pre search criteria - native for WUAPI.
Finds updates that are intended for deployment by Automatic Updates.
"IsAssigned=1" finds updates that are intended for deployment by Automatic Updates, which depends on the other query criteria.At most, one assigned Windows-based driver update is returned for each local device on a destination computer.

"IsAssigned=0" finds updates that are not intended to be deployed by Automatic Updates.

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

### -IsHidden
Pre search criteria - native for WUAPI.
Finds updates that are marked as hidden on the destination computer.
Default search criteria is only not hidden upadates.

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

### -IsInstalled
Pre search criteria - native for WUAPI.
Finds updates that are installed on the destination computer.

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

### -IsPresent
Pre search criteria - native for WUAPI.
When set to 1, finds updates that are present on a computer.

"IsPresent=1" finds updates that are present on a destination computer.If the update is valid for one or more products, the update is considered present if it is installed for one or more of the products.

"IsPresent=0" finds updates that are not installed for any product on a destination computer.

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
Post search criteria.
Finds updates that contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.

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

### -MaxSize
Post search criteria.
Finds updates that have MaxDownloadSize less or equal.
Size is in Bytes.

```yaml
Type: Int64
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -MicrosoftUpdate
Use Windows Update Service Manager - '9482f4b4-e343-43b6-b170-9a65bc822c77'

```yaml
Type: SwitchParameter
Parameter Sets: MicrosoftUpdate
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinSize
Post search criteria.
Finds updates that have MaxDownloadSize greater or equal.
Size is in Bytes.

```yaml
Type: Int64
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotCategory
Post search criteria.
Finds updates that not contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...

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

### -NotKBArticleID
Post search criteria.
Finds updates that not contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.

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

### -NotSeverity
Post search criteria.
Finds updates that not match part of severity.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:
Accepted values: Critical, Important, Moderate, Low, Unspecified

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotTitle
Post search criteria.
Finds updates that not match part of title (case sensitive).

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

### -NotUpdateID
Pre search criteria - native for WUAPI.
Finds updates without a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.

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

### -RecurseCycle
Specify number of cycles for check updates after successful update installation or system startup.
First run is always main cycle (-RecurseCycle 1 or none).
Second (-RecurseCycle 2) and n (-RecurseCycle n) cycle are recursive.

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

### -RevisionNumber
Pre search criteria - native for WUAPI.
Finds updates with a specific RevisionNumber, such as '100'.
This criterion must be combined with the UpdateID param.

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

### -RootCategories
Post search criteria.
Finds updates that contain a specified root category name 'Critical Updates', 'Definition Updates', 'Drivers', 'Feature Packs', 'Security Updates', 'Service Packs', 'Tools', 'Update Rollups', 'Updates', 'Upgrades', 'Microsoft'.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:
Accepted values: Critical Updates, Definition Updates, Drivers, Feature Packs, Security Updates, Service Packs, Tools, Update Rollups, Updates, Upgrades, Microsoft

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScheduleJob
Specify time when job will start.

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

### -SendHistory
Send install history (Get-WUHistory) report after successful update installation or system startup.

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

### -SendReport
Send report email to specific recipients.

Requires the parameter -PSWUSettings or declare the PSWUSettings.xml file (more preferred) in ModuleBase path.

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
Use specific Service Manager if it's available.

Examples Of ServiceID: \r\n \r\n -- Windows Update 9482f4b4-e343-43b6-b170-9a65bc822c77 \r\n -- Microsoft Update 7971f918-a847-4430-9279-4a52d1efe18d \r\n -- Windows Store 117cab2d-82b1-4b5a-a08c-4d62dbee7782 \r\n -- Windows Server Update Service 3da21691-e39d-4da6-8a4b-b43877bcb1b7

```yaml
Type: String
Parameter Sets: ServiceID
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Severity
Post search criteria.
Finds updates that match part of severity, such as 'Important', 'Critical', 'Moderate', etc...

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:
Accepted values: Critical, Important, Moderate, Low, Unspecified

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowPreSearchCriteria
Show choosen search criteria.
Only works for pre search criteria.

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

### -Title
Post search criteria.
Finds updates that match part of title (case sensitive), such as '.NET Framework 4'.

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

### -UpdateID
Pre search criteria - native for WUAPI.
Finds updates with a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.

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

### -UpdateType
Pre search criteria - native for WUAPI.
Finds updates with a specific type, such as 'Driver' and 'Software'.
Default value contains all updates.

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Driver, Software

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

### -WindowsUpdate
Use Microsoft Update Service Manager - '7971f918-a847-4430-9279-4a52d1efe18d'

```yaml
Type: SwitchParameter
Parameter Sets: WindowsUpdate
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -WithHidden
Pre search criteria - native for WUAPI.
Finds updates that are both hidden and not on the destination computer.
Overwrite IsHidden param.
Default search criteria is only not hidden upadates.

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

### System.String[]

## OUTPUTS

### PSWindowsUpdate.WindowsUpdate

## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

