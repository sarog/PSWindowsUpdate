---
external help file: PSWindowsUpdate.dll-Help.xml
Module Name: PSWindowsUpdate
online version: https://commandlinegeeks.wordpress.com/
schema: 2.0.0
---

# Reset-WUComponents

## SYNOPSIS
Reset Windows Update components.

## SYNTAX

```
Reset-WUComponents [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Reset-WUComponents cmdlet to reset all Windows Update components to default.

## EXAMPLES

### Example 1
```
Reset Windows Update components to default.

 Reset-WUComponents -Verbose
 Step 1: Stop Windows Update services
VERBOSE: Background Intelligent Transfer Service(BITS)
VERBOSE: Windows Update(wuauserv)
VERBOSE: Application Identity(appidsvc)
VERBOSE: Cryptographic Services(cryptsvc)
Step 2: Delete the qmgr*.dat files
Step 3: Backup softare distribution folders
VERBOSE: Renaming Software Distribution folder to C:\WINDOWS\SoftwareDistribution.bak5
VERBOSE: Renaming CatRoot  folder to C:\WINDOWS\System32\Catroot2.bak1
Step 4: Remove old Windows Update logs
VERBOSE: Deleting the C:\WINDOWS\WindowsUpdate.log files.
Step 5: Reset Windows Update services
VERBOSE: Reset BITS service
VERBOSE: Reset Windows Update service
Step 6: Reregister dll's
VERBOSE: regsvr32.exe / s atl.dll
VERBOSE: regsvr32.exe / s urlmon.dll
VERBOSE: regsvr32.exe / s mshtml.dll
VERBOSE: regsvr32.exe / s shdocvw.dll
VERBOSE: regsvr32.exe / s browseui.dll
VERBOSE: regsvr32.exe / s jscript.dll
VERBOSE: regsvr32.exe / s vbscript.dll
VERBOSE: regsvr32.exe / s scrrun.dll
VERBOSE: regsvr32.exe / s msxml.dll
VERBOSE: regsvr32.exe / s msxml3.dll
VERBOSE: regsvr32.exe / s msxml6.dll
VERBOSE: regsvr32.exe / s actxprxy.dll
VERBOSE: regsvr32.exe / s softpub.dll
VERBOSE: regsvr32.exe / s wintrust.dll
VERBOSE: regsvr32.exe / s dssenh.dll
VERBOSE: regsvr32.exe / s rsaenh.dll
VERBOSE: regsvr32.exe / s gpkcsp.dll
VERBOSE: regsvr32.exe / s sccbase.dll
VERBOSE: regsvr32.exe / s slbcsp.dll
VERBOSE: regsvr32.exe / s cryptdlg.dll
VERBOSE: regsvr32.exe / s oleaut32.dll
VERBOSE: regsvr32.exe / s ole32.dll
VERBOSE: regsvr32.exe / s shell32.dll
VERBOSE: regsvr32.exe / s initpki.dll
VERBOSE: regsvr32.exe / s wuapi.dll
VERBOSE: regsvr32.exe / s wuaueng.dll
VERBOSE: regsvr32.exe / s wuaueng1.dll
VERBOSE: regsvr32.exe / s wucltui.dll
VERBOSE: regsvr32.exe / s wups.dll
VERBOSE: regsvr32.exe / s wups2.dll
VERBOSE: regsvr32.exe / s wuweb.dll
VERBOSE: regsvr32.exe / s qmgr.dll
VERBOSE: regsvr32.exe / s qmgrprxy.dll
VERBOSE: regsvr32.exe / s wucltux.dll
VERBOSE: regsvr32.exe / s muweb.dll
VERBOSE: regsvr32.exe / s wuwebv.dll
Step 7: Reset WinSock
VERBOSE: netsh winsock reset
Step 8: Reset Proxy
VERBOSE: netsh winhttp reset proxy
Step 9: Start Windows Update services
VERBOSE: Cryptographic Services (cryptsvc)
VERBOSE: Application Identity (appidsvc)
VERBOSE: Windows Update (wuauserv)
VERBOSE: Background Intelligent Transfer Service (BITS)
Step 10: Start Windows Update services
VERBOSE: wuauclt /resetauthorization /detectnow
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

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS

[Author Blog](https://commandlinegeeks.wordpress.com/)

