using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;
using WUApiLib;

namespace PSWindowsUpdate {
    /// <summary>
    /// <para type="synopsis">Get list of available updates meeting the criteria.</para>
    /// <para type="description">Use Get-WindowsUpdate (aka Get-WUList) cmdlet to get list of available or installed updates meeting specific criteria.</para>
    /// <para type="description">Use Download-WindowsUpdate alias to get list of updates and download it. Equivalent Get-WindowsUpdate -Download.</para>
    /// <para type="description">Use Install-WindowsUpdate (aka Get-WUInstall) alias to get list of updates and install it. Equivalent Get-WindowsUpdate -Install.</para>
    /// <para type="description">Use Hide-WindowsUpdate alias to get list of updates and hide it. Equivalent Get-WindowsUpdate -Hide.</para>
    /// <para type="description">Use Show-WindowsUpdate (aka UnHide-WindowsUpdate) alias to get list of updates and unhide it. Equivalent Get-WindowsUpdate -Hide:$false.</para>
    /// <para type="description">There are two types of filtering update: Pre search criteria, Post search criteria.</para>
    /// <para type="description">- Pre search works on server side, like example: (IsInstalled = 0 and IsHidden = 0 and CategoryIds contains '0fa1201d-4330-4fa8-8ae9-b877473b6441' )</para>
    /// <para type="description">- Post search work on client side after get the pre-filtered list of updates, like example $KBArticleID -match $Update.KBArticleIDs</para>
    /// <para type="description">Status info list:\r\n[A|R]DIMHUB\r\nA-IsAccetped\r\nR-IsRejected\r\n D-IsDownloaded\r\n F-DownloadFailed\r\n ?-IsInvoked\r\n  I-IsInstalled\r\n  F-InstallFailed\r\n  ?-IsInvoked\r\n  R-RebootRequired\r\n   M-IsMandatory\r\n    H-IsHidden\r\n     U-IsUninstallable\r\n      B-IsBeta</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Get windows updates available from default service manager.</para>
    ///
    /// Get-WindowsUpdate -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...</para>
    /// <para>VERBOSE: Found[4] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[4] Updates in post search criteria</para>
    /// <para></para>
    /// <para>ComputerName Status     KB          Size Title</para>
    /// <para>------------ ------     --          ---- -----</para>
    /// <para>MG-PC        -------    KB890830    44MB Narzędzie Windows do usuwania złośliwego oprogramowania dla systemów Window...</para>
    /// <para>MG-PC        -------    KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...</para>
    /// <para>MG-PC        -------    KB4034662   21MB 2017-08 Aktualizacja zabezpieczeń Adobe Flash Player w Windows 10 Version 1...</para>
    /// <para>MG-PC        -------    KB4035631   11MB 2017-08 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Get all installed drivers that are available at Windows Update. Additionaly show pre search criteria.</para>
    ///
    /// Get-WindowsUpdate -WindowsUpdate -UpdateType Driver -IsInstalled -ShowPreSearchCriteria -Verbose
    ///
    /// <para>PreSearchCriteria: IsInstalled = 0 and Type = 'Driver' and IsHidden = 0</para>
    /// <para>VERBOSE: MG-PC: Connecting to Windows Update server.Please wait...</para>
    /// <para>VERBOSE: Found[1] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[1] Updates in post search criteria</para>
    /// <para></para>
    /// <para>ComputerName Status     KB          Size Title</para>
    /// <para>------------ ------     --          ---- -----</para>
    /// <para>MGAJDALAP3   -DI----                 3MB Intel - Other hardware - Intel(R) Watchdog Timer Driver (Intel(R) WDT)</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Get all available update on remote machine MG-PC, that contains in Title this two words 'Aktualizacja' and 'Windows 10' (as regular expression).</para>
    ///
    /// Get-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -Title "Aktualizacja.*Windows 10" -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...</para>
    /// <para>VERBOSE: Found[14] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[5] Updates in post search criteria</para>
    /// <para></para>
    /// <para>ComputerName Status     KB          Size Title</para>
    /// <para>------------ ------     --          ---- -----</para>
    /// <para>MG-PC        -------    KB3150513    2MB 2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...</para>
    /// <para>MG-PC        -------    KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...</para>
    /// <para>MG-PC        -------    KB4034662   21MB 2017-08 Aktualizacja zabezpieczeń Adobe Flash Player w Windows 10 Version 1...</para>
    /// <para>MG-PC        -------    KB4035631   11MB 2017-08 Aktualizacja Windows 10 Version 1607 dla systemów opartych na archi...</para>
    /// <para>MG-PC        -------    KB4033637    4MB Aktualizacja systemu Windows 10 Version 1607 dla komputerów z procesorami x...</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Hide update with KBArticleID: KB4034658.</para>
    ///
    /// Get-WindowsUpdate -KBArticleID KB4034658 -Hide -Verbose
    /// <para>or use alias</para>
    /// Hide-WindowsUpdate -KBArticleID KB4034658 -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...</para>
    /// <para>VERBOSE: Found[4] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[1] Updates in post search criteria</para>
    /// <para></para>
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Hide 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB4034658)[1GB]" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// <para></para>
    /// <para>ComputerName Status     KB          Size Title</para>
    /// <para>------------ ------     --          ---- -----</para>
    /// <para>MG-PC        ---H--     KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Unhide update with KBArticleID: KB4034658.</para>
    ///
    /// Get-WindowsUpdate -KBArticleID KB4034658 -WithHidden -Hide:$false -Verbose
    /// <para>or use alias</para>
    /// Show-WindowsUpdate -KBArticleID KB4034658 -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Windows Server Update Service server. Please wait...</para>
    /// <para>VERBOSE: Found[4] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[1] Updates in post search criteria</para>
    /// <para></para>
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Show 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB4034658)[1GB]" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// <para></para>
    /// <para>ComputerName Status     KB          Size Title</para>
    /// <para>------------ ------     --          ---- -----</para>
    /// <para>MG-PC        ------     KB4034658    1GB 2017-08 Aktualizacja zbiorcza dla systemu Windows 10 Version 1607 dla syste...</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Schedule job at 6:00 PM to install update with UpdateId='ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber=200. Update will be automaticaly accepted and after all serwer will be automaticaly restarted if needed. </para>
    ///
    /// Get-WindowsUpdate -MicrosoftUpdate -UpdateID ddb74579-7a1f-4d1f-80c8-e8647055314e -RevisionNumber 200 -ScheduleJob (Get-Date -Hour 18 -Minute 0 -Second 0) -Install -AcceptAll -AutoReboot -Verbose
    /// <para>or use alias</para>
    /// Install-WindowsUpdate -MicrosoftUpdate -UpdateID ddb74579-7a1f-4d1f-80c8-e8647055314e -RevisionNumber 200 -ScheduleJob (Get-Date -Hour 18 -Minute 0 -Second 0) -AcceptAll -AutoReboot -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...</para>
    /// <para>VERBOSE: Found[1] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[1] Updates in post search criteria</para>
    /// <para>VERBOSE: Choosed pre Search Criteria: (UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200)</para>
    /// <para></para>
    /// <para>X ComputerName Result     KB          Size Title</para>
    /// <para>- ------------ ------     --          ---- -----</para>
    /// <para>1 MG-PC        Accepted   KB4023307   13MB Microsoft Silverlight(KB4023307)</para>
    /// <para>VERBOSE: Accepted[1] Updates ready to Download</para>
    /// <para>VERBOSE: Invoke-WUJob: MG-PC(31.08.2017 18:00:00):</para>
    /// <para>VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200)\" -AcceptAll -AutoReboot -Download -Install -MicrosoftUpdate -Verbose *&gt;&amp;1 | Out-File $Env:TEMP\PSWindowsUpdate.log"</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Install updates on remote computer. After all send a report from the installation process.</para>
    ///
    /// Install-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -AcceptAll -AutoReboot -SendReport -PSWUSettings @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} -Verbose
    /// <para>or use global PSWUSettings</para>
    /// @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} | Export-Clixml -Path 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate\PSWUSettings.xml'
    /// Install-WindowsUpdate -ComputerName MG-PC -MicrosoftUpdate -AcceptAll -AutoReboot -SendReport -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...</para>
    /// <para>VERBOSE: Found[4] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[4] Updates in post search criteria</para>
    /// <para></para>
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Microsoft Silverlight (KB4023307)[13MB]" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// <para></para>
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na architekturze x64 (KB3150513)[2MB]" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// <para></para>
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Aktualizacja pakietu językowego usługi Microsoft Dynamics 365 2.1[47MB]" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): L</para>
    /// <para></para>
    /// <para>X ComputerName Result     KB          Size Title</para>
    /// <para>- ------------ ------     --          ---- -----</para>
    /// <para>1 MG-PC Accepted KB4023307   13MB Microsoft Silverlight (KB4023307)</para>
    /// <para>1 MG-PC Accepted KB3150513    2MB 2017-06 Aktualizacja Windows 10 Version 1607 dla systemów opartych na arc...</para>
    /// <para>1 MG-PC Rejected KB4013759   47MB Aktualizacja pakietu językowego usługi Microsoft Dynamics 365 2.1</para>
    /// <para>1 MG-PC Rejected KB3186568   67MB Program Microsoft .NET Framework 4.7 w syst. Windows 10 Version 1607 i Wi...</para>
    /// <para>VERBOSE: Accepted [2]</para>
    /// <para>Updates ready to Download</para>
    /// <para>VERBOSE: Invoke-WUJob: MG-PC (Now):</para>
    /// <para>VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'ddb74579-7a1f-4d1f-80c8-e8647055314e' and RevisionNumber = 200) or (UpdateID = '151c4402-513c-4f39-8da1-f84d0956b5e3' and RevisionNumber = 200)\" -AcceptAll -Download -Install -AutoReboot -MicrosoftUpdate -SendReport -ProofOfLife -Verbose *&gt;&amp;1 | Out-File $Env:TEMP\PSWindowsUpdate.log"</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Schedule Job to install all available updates and automatically reboot system if needed. Also send report after installation (but before reboot if needed) and send second instalation history report after reboot.</para>
    ///
    /// Set-PSWUSettings -SmtpServer smtp.office365.com -SmtpPort 587 -SmtpEnableSsl $true -SmtpSubject "PSWindowsUpdate Report" -SmtpTo mgajda@psmvp.pl -SmtpFrom mgajda@psmvp.pl -SmtpCredential (Get-Credential mgajda@psmvp.pl)
    /// Install-WindowsUpdate -MicrosoftUpdate -SendReport -SendHistory -AcceptAll -AutoReboot -ScheduleJob (Get-Date -Hour 18 -Minute 30 -Second 0) -ComputerName MG-PC -Verbose
    ///
    /// <para>VERBOSE: MG-PC: Connecting to Microsoft Update server. Please wait...</para>
    /// <para>VERBOSE: Found[4] Updates in pre search criteria</para>
    /// <para>VERBOSE: Found[4] Updates in post search criteria</para>
    /// <para></para>
    /// <para>X ComputerName Result     KB          Size Title</para>
    /// <para>- ------------ ------     --          ---- -----</para>
    /// <para>1 MG-PC        Accepted   KB3038936    5MB Aktualizacja systemu Windows 8.1 dla komputerów z procesorami x64(KB3038...</para>
    /// <para>1 MG-PC        Accepted   KB3186606    4MB Pakiety językowe programu Microsoft.NET Framework 4.7 w syst. Windows 8....</para>
    /// <para>1 MG-PC        Accepted   KB4035038   53MB Sierpień 2017: wersja zapozn. pak.zb.aktual.jakości dla pr. .NET Frame...</para>
    /// <para>1 MG-PC        Accepted   KB2267602  309MB Aktualizacja definicji dla: Windows Defender — KB2267602 (Definicja 1.251...</para>
    /// <para>VERBOSE: Accepted[4] Updates ready to Download</para>
    /// <para>VERBOSE: Invoke-WUJob: MG-PC (02.09.2017 08:30:00):</para>
    /// <para>VERBOSE: powershell.exe -Command "Get-WindowsUpdate -Criteria \"(UpdateID = 'e69c9679-7ce8-489a-a21c-62fb920be67a' and RevisionNumber = 201) or(UpdateID = 'de44604d-ec38-4a7f-ac63-28b3edfdb382' and RevisionNumber = 207) or(UpdateID = '9cf1d8c9-a7c3-4603-90e8-f22131ff6d7e' and RevisionNumber = 201) or(UpdateID = 'b51935f9-0e40-4624-9c26-b29bff92dcf9' and RevisionNumber = 200)\" -AcceptAll -Install -AutoReboot -MicrosoftUpdate -SendReport -SendHistory -Verbose *&gt;&amp;1 | Out-File $Env:TEMP\PSWindowsUpdate.log"</para>
    /// <para>VERBOSE: Send report</para>
    /// </code>
    /// </example>
    [Cmdlet("Get", "WindowsUpdate", SupportsShouldProcess = true, DefaultParameterSetName = "Default", ConfirmImpact = ConfirmImpact.High)]
    [OutputType(typeof(WindowsUpdate))]
    public class GetWindowsUpdate : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection. Interactive remote connection works only for checking updates. For download or install cmdlet creates an Invoke-WUJob task.</para>
        /// </summary>
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        /// <summary>
        /// <para type="description">Specify alternative credential.</para>
        /// </summary>
        [Parameter]
        private PSCredential Credential { get; set; }

        /// <summary>
        /// <para type="description">Send report email to specific recipients.</para>
        /// <para type="description">Requires the parameter -PSWUSettings or declare the PSWUSettings.xml file (more preferred) in ModuleBase path.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SendReport { get; set; }

        /// <summary>
        /// <para type="description">Required parameter for -SendReport.</para>
        /// <para type="description">Passes the parameters (as hashtable) necessary to send the report:
        /// \r\n@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25];[Subject="Alternative Subject"];[Properties="Alternative object properties"];[Style="Table|List"]}</para>
        /// <para type="description">Send parameters can also be saved to a PSWUSettings.xml file in ModuleBase path:
        /// \r\nExport-Clixml @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25]}"</para>
        /// </summary>
        [Parameter]
        public Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Send install history (Get-WUHistory) report after successful update installation or system startup.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SendHistory { get; set; }

        /// <summary>
        /// <para type="description">Specify time when job will start.</para>
        /// </summary>
        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleJob { get; set; }

        /// <summary>
        /// <para type="description">Do not ask confirmation for updates. Download or Install all available updates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AcceptAll { get; set; }

        /// <summary>
        /// <para type="description">Specify number of cycles for check updates after successful update installation or system startup. First run is always main cycle (-RecurseCycle 1 or none). Second (-RecurseCycle 2) and n (-RecurseCycle n) cycle are recursive.</para>
        /// </summary>
        [Parameter]
        [ValidateRecurseCycle]
        public int RecurseCycle { get; set; }

        /// <summary>
        /// <para type="description">Get list of updates and hide/unhide approved updates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Hide { get; set; }

        /// <summary>
        /// <para type="description">Get list of updates and download approved updates, but do not install it.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Download { get; set; }

        /// <summary>
        /// <para type="description">Forces the download of updates that are already installed or that cannot be installed. Works only with -Download.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter ForceDownload { get; set; }

        /// <summary>
        /// <para type="description">Get list of updates and install approved updates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Install { get; set; }

        /// <summary>
        /// <para type="description">A forced installation is an installation in which an update is installed even if the metadata indicates that the update is already installed. Before you use ForceInstall to force an installation, determine whether the update is installed and available. If an update is not installed, a forced installation fails. Works only with -Install.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter ForceInstall { get; set; }

        /// <summary>
        /// <para type="description">Do not ask for reboot if it needed.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AutoReboot { get; set; }

        /// <summary>
        /// <para type="description">Do not ask for reboot if it needed, but do not reboot automaticaly.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IgnoreReboot { get; set; }

        /// <summary>
        /// <para type="description">Specify time when system will be rebooted.</para>
        /// </summary>
        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleReboot { get; set; }

        /// <summary>
        /// <para type="description">Use specific Service Manager if it's available. </para>
        /// <para type="description">Examples Of ServiceID:
        /// \r\n
        /// \r\n -- Windows Update 				    9482f4b4-e343-43b6-b170-9a65bc822c77
        /// \r\n -- Microsoft Update 				7971f918-a847-4430-9279-4a52d1efe18d
        /// \r\n -- Windows Store 					117cab2d-82b1-4b5a-a08c-4d62dbee7782
        /// \r\n -- Windows Server Update Service 	3da21691-e39d-4da6-8a4b-b43877bcb1b7</para>
        /// </summary>
        [Parameter(ParameterSetName = "ServiceID")]
        public string ServiceID { get; set; }

        /// <summary>
        /// <para type="description">Use Microsoft Update Service Manager - '7971f918-a847-4430-9279-4a52d1efe18d'</para>
        /// </summary>
        [Parameter(ParameterSetName = "WindowsUpdate")]
        public SwitchParameter WindowsUpdate { get; set; }

        /// <summary>
        /// <para type="description">Use Windows Update Service Manager - '9482f4b4-e343-43b6-b170-9a65bc822c77'</para>
        /// </summary>
        [Parameter(ParameterSetName = "MicrosoftUpdate")]
        public SwitchParameter MicrosoftUpdate { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Set own string that specifies the search criteria. https://docs.microsoft.com/pl-pl/windows/desktop/api/wuapi/nf-wuapi-iupdatesearcher-search </para>
        /// </summary>
        [Parameter]
        public string Criteria { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates with a specific type, such as 'Driver' and 'Software'. Default value contains all updates.</para>
        /// </summary>
        [ValidateSet("Driver", "Software", IgnoreCase = true)]
        [Parameter]
        public string UpdateType { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that are deployed for a specific action, such as an installation or uninstallation that the administrator of a server specifies. "DeploymentAction='Installation'" finds updates that are deployed for installation on a destination computer. "DeploymentAction='Uninstallation'" depends on the other query criteria.</para>
        /// <para type="description">"DeploymentAction='Uninstallation'" finds updates that are deployed for uninstallation on a destination computer. "DeploymentAction='Uninstallation'" depends on the other query criteria.</para>
        /// <para type="description">If this criterion is not explicitly specified, each group of criteria that is joined to an AND operator implies "DeploymentAction='Installation'".</para>
        /// </summary>
        [ValidateSet("Installation", "Uninstallation", IgnoreCase = true)]
        [Parameter]
        public string DeploymentAction { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that are intended for deployment by Automatic Updates. "IsAssigned=1" finds updates that are intended for deployment by Automatic Updates, which depends on the other query criteria.At most, one assigned Windows-based driver update is returned for each local device on a destination computer.</para>
        /// <para type="description">"IsAssigned=0" finds updates that are not intended to be deployed by Automatic Updates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IsAssigned { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. When set to 1, finds updates that are present on a computer.</para>
        /// <para type="description">"IsPresent=1" finds updates that are present on a destination computer.If the update is valid for one or more products, the update is considered present if it is installed for one or more of the products.</para>
        /// <para type="description">"IsPresent=0" finds updates that are not installed for any product on a destination computer.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IsPresent { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. "BrowseOnly=1" finds updates that are considered optional. "BrowseOnly=0" finds updates that are not considered optional.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter BrowseOnly { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates where the AutoSelectOnWebSites property has the specified value.</para>
        /// <para type="description">"AutoSelectOnWebSites=1" finds updates that are flagged to be automatically selected by Windows Update.</para>
        /// <para type="description">"AutoSelectOnWebSites=0" finds updates that are not flagged for Automatic Updates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AutoSelectOnWebSites { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates with a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.</para>
        /// </summary>
        [Parameter]
        public string[] UpdateID { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates without a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.</para>
        /// </summary>
        [Parameter]
        public string[] NotUpdateID { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates with a specific RevisionNumber, such as '100'. This criterion must be combined with the UpdateID param.</para>
        /// </summary>
        [Parameter]
        public int RevisionNumber { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that belong to a specified category (or sets of UUIDs), such as '0fa1201d-4330-4fa8-8ae9-b877473b6441'.</para>
        /// </summary>
        [Parameter]
        public string[] CategoryIDs { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that are installed on the destination computer.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IsInstalled { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that are marked as hidden on the destination computer. Default search criteria is only not hidden upadates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IsHidden { get; set; }

        /// <summary>
        /// <para type="description">Pre search criteria - native for WUAPI. Finds updates that are both hidden and not on the destination computer. Overwrite IsHidden param. Default search criteria is only not hidden upadates.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter WithHidden { get; set; }

        /// <summary>
        /// <para type="description">Show choosen search criteria. Only works for pre search criteria.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter ShowPreSearchCriteria { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that contain a specified root category name 'Critical Updates', 'Definition Updates', 'Drivers', 'Feature Packs', 'Security Updates', 'Service Packs', 'Tools', 'Update Rollups', 'Updates', 'Upgrades', 'Microsoft'.</para>
        /// </summary>
        [ValidateSet("Critical Updates", "Definition Updates", "Drivers", "Feature Packs", "Security Updates", "Service Packs", "Tools", "Update Rollups", "Updates", "Upgrades",
            "Microsoft")]
        [Parameter]
        public string[] RootCategories { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...</para>
        /// </summary>
        [Parameter]
        public string[] Category { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.</para>
        /// </summary>
        [Parameter]
        public string[] KBArticleID { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that match part of title (case sensitive), such as '.NET Framework 4'.</para>
        /// </summary>
        [Parameter]
        public string Title { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that match part of severity, such as 'Important', 'Critical', 'Moderate', etc...</para>
        /// </summary>
        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified")]
        [Parameter]
        public string[] Severity { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that not contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...</para>
        /// </summary>
        [Parameter]
        public string[] NotCategory { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that not contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.</para>
        /// </summary>
        [Parameter]
        public string[] NotKBArticleID { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that not match part of title (case sensitive).</para>
        /// </summary>
        [Parameter]
        public string NotTitle { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that not match part of severity.</para>
        /// </summary>
        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified")]
        [Parameter]
        public string[] NotSeverity { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that the installation or uninstallation of an update can't prompt for user input.</para>
        /// </summary>
        [Alias("Silent")]
        [Parameter]
        public SwitchParameter IgnoreUserInput { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that specifies the restart behavior that not occurs when you install or uninstall the update.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IgnoreRebootRequired { get; set; }

        /// <summary>
        /// <para type="description">Install only the updates that have status AutoSelectOnWebsites on true.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AutoSelectOnly { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that have MaxDownloadSize less or equal. Size is in Bytes.</para>
        /// </summary>
        [Parameter]
        public long MaxSize { get; set; }

        /// <summary>
        /// <para type="description">Post search criteria. Finds updates that have MaxDownloadSize greater or equal. Size is in Bytes.</para>
        /// </summary>
        [Parameter]
        public long MinSize { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions. For additional debug information use $DebugPreference = "Continue"</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static List<PSObject> OutputObj { get; set; }

        private static UpdateSession UpdateSessionObj { get; set; }

        private static IUpdateSearcher SearcherObj { get; set; }

        private static UpdateServiceManager ServiceManagerObj { get; set; }

        private static bool NeedsReboot { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        /// <summary>Begin</summary>
        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new List<PSObject>();
            if (SendReport) {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new WindowsUpdate() /*new string[5] {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
                        "Title"
                    }*/);
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = false;
                    WriteDebug(DateTime.Now + " Disabling -SendReport");
                }
            }

            if (ComputerName != null) {
                return;
            }

            ComputerName = new string[1] {
                Environment.MachineName
            };
        }

        private void CoreProcessing() {
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " ParameterSetName: " + ParameterSetName);
            if (string.Equals(invocationName, "Install-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
                Install = true;
            }

            if (string.Equals(invocationName, "Download-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
            }

            if (string.Equals(invocationName, "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = true;
            }

            if (string.Equals(invocationName, "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(invocationName, "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = false;
                WithHidden = true;
            }

            string searchCriteria;
            if (Criteria != null) {
                searchCriteria = Criteria;
                WriteDebug(DateTime.Now + " Set pre search criteria: " + searchCriteria);
            } else {
                if (IsInstalled) {
                    searchCriteria = "IsInstalled = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 1");
                } else {
                    searchCriteria = "IsInstalled = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 0");
                }

                if (UpdateType != null) {
                    searchCriteria = searchCriteria + " and Type = '" + UpdateType + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: Type = '" + UpdateType + "'");
                }

                if (IsHidden) {
                    searchCriteria += " and IsHidden = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1");
                } else if (WithHidden) {
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1 or IsHidden = 0");
                } else {
                    searchCriteria += " and IsHidden = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 0");
                }

                if (IgnoreRebootRequired) {
                    searchCriteria += " and RebootRequired = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: RebootRequired = 0");
                }

                if (DeploymentAction != null) {
                    searchCriteria = searchCriteria + " and DeploymentAction = '" + DeploymentAction + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: DeploymentAction = '" + DeploymentAction + "'");
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsAssigned")) {
                    if (IsAssigned) {
                        searchCriteria += " and IsAssigned = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 1");
                    } else {
                        searchCriteria += " and IsAssigned = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsPresent")) {
                    if (IsPresent) {
                        searchCriteria += " and IsPresent = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 1");
                    } else {
                        searchCriteria += " and IsPresent = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("BrowseOnly")) {
                    if (BrowseOnly) {
                        searchCriteria += " and BrowseOnly = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 1");
                    } else {
                        searchCriteria += " and BrowseOnly = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites")) {
                    if (AutoSelectOnWebSites) {
                        searchCriteria += " and AutoSelectOnWebSites = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 1");
                    } else {
                        searchCriteria += " and AutoSelectOnWebSites = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 0");
                    }
                }

                if (NotUpdateID != null) {
                    var notUpdateID = NotUpdateID;
                    foreach (var uID in notUpdateID) {
                        searchCriteria = searchCriteria + " and UpdateID != '" + uID + "'";
                        WriteDebug(DateTime.Now + " Set pre search criteria: NotUpdateID = " + uID);
                    }
                }

                if (UpdateID != null) {
                    var text3 = searchCriteria;
                    searchCriteria = "";
                    var num = 0;
                    var updateID = UpdateID;
                    foreach (var text4 in updateID) {
                        if (num > 0) {
                            searchCriteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        if (RevisionNumber > 0) {
                            searchCriteria = searchCriteria + "(" + text3 + " and UpdateID = '" + text4 + "' and RevisionNumber = " + RevisionNumber + ")";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + text4 + "and RevisionNumber = " + RevisionNumber);
                        } else {
                            searchCriteria = searchCriteria + "(" + text3 + " and UpdateID = '" + text4 + "')";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + text4);
                        }

                        num++;
                    }
                }

                if (CategoryIDs != null) {
                    var text5 = searchCriteria;
                    searchCriteria = "";
                    var num2 = 0;
                    var categoryIDs = CategoryIDs;
                    foreach (var catID in categoryIDs) {
                        if (num2 > 0) {
                            searchCriteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        searchCriteria = searchCriteria + "(" + text5 + " and CategoryIDs contains '" + catID + "')";
                        WriteDebug(DateTime.Now + " Set pre search criteria: CategoryIDs = " + catID);
                        num2++;
                    }
                }
            }

            WriteDebug(DateTime.Now + " Search criteria is: " + searchCriteria);
            if (ShowPreSearchCriteria) {
                Host.UI.WriteLine(ConsoleColor.Green, Host.UI.RawUI.BackgroundColor, "PreSearchCriteria: " + searchCriteria);
            }

            var computerName = ComputerName;
            foreach (var target in computerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (Download) {
                    var errorRecord = WUToolsObj.CheckPSWUModule(target);
                    if (errorRecord != null) {
                        WriteError(errorRecord);
                        continue;
                    }
                }

                var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                WriteDebug(DateTime.Now + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                if (wUApiUpdateSessionObj.Status) {
                    UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                    SearcherObj = UpdateSessionObj.CreateUpdateSearcher();
                    var wUApiServiceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                    WriteDebug(DateTime.Now + " ServiceManagerObj mode: " + wUApiServiceManagerObj.Mode);
                    if (wUApiServiceManagerObj.Status) {
                        ServiceManagerObj = (UpdateServiceManager)wUApiServiceManagerObj.Object;
                        var text8 = "";
                        if (WindowsUpdate) {
                            WriteDebug(DateTime.Now + " Set source of updates to Windows Update");
                            SearcherObj.ServerSelection = ServerSelection.ssWindowsUpdate;
                            text8 = "Windows Update";
                        } else if (MicrosoftUpdate) {
                            WriteDebug(DateTime.Now + " Set source of updates to Microsof Update");
                            ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                            var flag = false;
                            foreach (IUpdateService2 service in ServiceManagerObj.Services) {
                                if (service.ServiceID == ServiceID) {
                                    WriteDebug(DateTime.Now + " Search for ServiceID. Set source of updates to " + service.Name);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = service.Name;
                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag) {
                                try {
                                    Thread.Sleep(500);
                                    var serviceController = new ServiceController("Windows Update", target);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                    ServiceManagerObj.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 2, null);
                                    WriteDebug(DateTime.Now + " Register WUA Microsof Update service");
                                    Thread.Sleep(500);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                } catch (COMException ex) {
                                    var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                    if (wUApiCodeDetails != null) {
                                        switch (wUApiCodeDetails.CodeType) {
                                            case 2:
                                                WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                                break;
                                            case 3:
                                                WriteWarning(wUApiCodeDetails.HResult + ": " + wUApiCodeDetails.Description);
                                                break;
                                        }
                                    } else if (Debuger) {
                                        var errorRecord2 = new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null);
                                        ThrowTerminatingError(errorRecord2);
                                    }
                                }

                                SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                SearcherObj.ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                                text8 = "Microsoft Update";
                            }
                        } else {
                            try {
                                foreach (IUpdateService2 service2 in ServiceManagerObj.Services) {
                                    if (ServiceID != null) {
                                        if (service2.ServiceID == ServiceID) {
                                            WriteDebug(DateTime.Now + " Try ServiceID. Set source of updates to " + service2.Name);
                                            SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                            SearcherObj.ServiceID = ServiceID;
                                            text8 = service2.Name;
                                            break;
                                        }
                                    } else if (service2.IsDefaultAUService) {
                                        WriteDebug(DateTime.Now + " Try Default. Set source of updates to " + service2.Name);
                                        text8 = service2.Name;
                                        break;
                                    }
                                }
                            } catch {
                                if (ServiceID != null) {
                                    WriteDebug(DateTime.Now + " Catch ServiceID. Set source of updates to " + ServiceID);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = ServiceID;
                                } else {
                                    WriteDebug(DateTime.Now + " Catch. Set source of updates to Default");
                                    text8 = "default for " + target + ".";
                                }
                            }
                        }

                        if (text8 == "Windows Server Update Service") {
                            RegistryKey registryKey = null;
                            registryKey = !WUToolsObj.IsLocalHost(target)
                                ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                                : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                            var registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                            var text9 = registryKey2.GetValue("WUServer").ToString();
                            text8 = text8 + " (" + text9 + ")";
                        }

                        WriteVerbose(target + " (" + DateTime.Now + "): Connecting to " + text8 + " server. Please wait...");
                        ISearchResult searchResult;
                        try {
                            searchResult = SearcherObj.Search(searchCriteria);
                        } catch (COMException ex2) {
                            var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex2.ErrorCode);
                            if (wUApiCodeDetails2 != null) {
                                switch (wUApiCodeDetails2.CodeType) {
                                    case 2:
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult, ErrorCategory.CloseError, null));
                                        break;
                                    case 3:
                                        WriteWarning(wUApiCodeDetails2.HResult + ": " + wUApiCodeDetails2.Description);
                                        break;
                                }
                            } else if (Debuger) {
                                var errorRecord3 = new ErrorRecord(ex2, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord3);
                            }

                            WriteDebug(DateTime.Now + " Skip to next computer");
                            continue;
                        }

                        var count = searchResult.Updates.Count;
                        WriteVerbose("Found [" + count + "] Updates in pre search criteria");
                        if (count == 0) {
                            continue;
                        }

                        var updateCollection = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        if (RootCategories != null) {
                            var updateCollection2 =
                                (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                            var rootCategories = RootCategories;
                            foreach (var text10 in rootCategories) {
                                var index = -1;
                                switch (text10) {
                                    case "Critical Updates":
                                        index = 0;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Definition Updates":
                                        index = 1;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Drivers":
                                        index = 2;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Feature Packs":
                                        index = 3;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Security Updates":
                                        index = 4;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Service Packs":
                                        index = 5;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Tools":
                                        index = 6;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Update Rollups":
                                        index = 7;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Updates":
                                        index = 8;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Upgrades":
                                        index = 9;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Microsoft":
                                        index = 10;
                                        WriteDebug(DateTime.Now + " RootCategory: " + text10 + " - " + index);
                                        break;
                                }

                                try {
                                    foreach (IUpdate update6 in searchResult.RootCategories[index].Updates) {
                                        updateCollection2.Add(update6);
                                    }
                                } catch {
                                    WriteWarning("RootCategiries Updates are empty. Use classic filters.");
                                }
                            }

                            updateCollection = updateCollection2;
                        } else {
                            updateCollection = searchResult.Updates;
                        }

                        var num3 = 0;
                        var progressRecord = new ProgressRecord(0, "Post search updates for " + target, "[" + num3 + "/" + count + "]");
                        var collection = new Collection<PSObject>();
                        foreach (IUpdate item in updateCollection) {
                            WriteDebug(DateTime.Now + " " + item.Title);
                            var text11 = WUToolsObj.ConvertSize(item.MaxDownloadSize);
                            progressRecord.StatusDescription = "[" + num3 + "/" + count + "] " + item.Title + " " + text11;
                            progressRecord.PercentComplete = num3 * 100 / count;
                            WriteProgress(progressRecord);
                            num3++;
                            var flag2 = true;
                            if (Category != null || NotCategory != null) {
                                var text12 = "";
                                foreach (ICategory category3 in item.Categories) {
                                    text12 = text12 + category3.Name + "|";
                                }

                                WriteDebug(DateTime.Now + " " + text12);
                                if (Category != null) {
                                    WriteDebug(DateTime.Now + " Set post search criteria: Categories = " + string.Join(", ", Category));
                                    var category2 = Category;
                                    foreach (var pattern in category2) {
                                        if (!Regex.IsMatch(text12, pattern, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            continue;
                                        }

                                        flag2 = true;
                                        break;
                                    }
                                }

                                if (NotCategory != null && flag2) {
                                    WriteDebug(DateTime.Now + " Set post search criteria: NotCategories = " + string.Join(", ", NotCategory));
                                    var notCategory = NotCategory;
                                    foreach (var pattern2 in notCategory) {
                                        if (Regex.IsMatch(text12, pattern2, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                }

                                WriteDebug(DateTime.Now + " UpdateAccess: " + flag2);
                            }

                            if (KBArticleID != null && flag2) {
                                var text13 = string.Join(", ", KBArticleID);
                                WriteDebug(DateTime.Now + " Set post search criteria: KBArticleIDs = " + text13);
                                if (item.KBArticleIDs.Count > 0) {
                                    foreach (string kBArticleID in item.KBArticleIDs) {
                                        if (!Regex.IsMatch(text13, kBArticleID, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                } else {
                                    flag2 = false;
                                }

                                WriteDebug(DateTime.Now + " UpdateAccess: " + flag2);
                            }

                            if (NotKBArticleID != null && flag2) {
                                var text15 = string.Join(", ", NotKBArticleID);
                                WriteDebug(DateTime.Now + " Set post search criteria: NotKBArticleID = " + text15);
                                if (item.KBArticleIDs.Count > 0) {
                                    foreach (string kBArticleID2 in item.KBArticleIDs) {
                                        if (Regex.IsMatch(text15, kBArticleID2, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                }

                                WriteDebug(DateTime.Now + " UpdateAccess: " + flag2);
                            }

                            if (Title != null && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: Title = " + Title);
                                if (!Regex.IsMatch(item.Title, Title, RegexOptions.IgnoreCase)) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess: " + flag2);
                                }
                            }

                            if (NotTitle != null && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: NotTitle = " + NotTitle);
                                if (Regex.IsMatch(item.Title, NotTitle, RegexOptions.IgnoreCase)) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess: " + flag2);
                                }
                            }

                            if (Severity != null && flag2) {
                                var text16 = string.Join(", ", Severity);
                                WriteDebug(DateTime.Now + " Set post search criteria: Severity = " + text16);
                                if (Severity.Contains("Unspecified")) {
                                    Severity = string.Join(",", Severity).Replace("Unspecified", "").Split(',');
                                }

                                if (!Severity.Contains(Convert.ToString((object)item.MsrcSeverity))) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (NotSeverity != null && flag2) {
                                var text17 = string.Join(", ", NotSeverity);
                                WriteDebug(DateTime.Now + " Set post search criteria: NotSeverity = " + text17);
                                if (NotSeverity.Contains("Unspecified")) {
                                    NotSeverity = string.Join(",", NotSeverity).Replace("Unspecified", "").Split(',');
                                }

                                if (NotSeverity.Contains(Convert.ToString((object)item.MsrcSeverity))) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (MaxSize > 0 && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: MaxDownloadSize <= " + MaxSize);
                                if (MaxSize <= item.MaxDownloadSize) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (MinSize > 0 && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: MaxDownloadSize >= " + MinSize);
                                if (MinSize >= item.MaxDownloadSize) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (IgnoreUserInput && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: CanRequestUserInput");
                                if (item.InstallationBehavior.CanRequestUserInput) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (IgnoreRebootRequired && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: RebootBehavior");
                                if (item.InstallationBehavior.RebootBehavior != 0) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (AutoSelectOnly && flag2) {
                                WriteDebug(DateTime.Now + " Set post search criteria: AutoSelectOnWebsites");
                                if (!item.AutoSelectOnWebSites) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now + " UpdateAccess" + flag2);
                                }
                            }

                            if (!flag2) {
                                continue;
                            }

                            WriteDebug(DateTime.Now + " Update was not filtered");
                            var text18 = "";
                            foreach (string kBArticleID3 in item.KBArticleIDs) {
                                text18 = !(text18 == "") ? text18 + ", KB" + kBArticleID3 : text18 + "KB" + kBArticleID3;
                            }

                            var text20 = "-";
                            text20 = !item.IsDownloaded ? text20 + "-" : text20 + "D";
                            text20 = !item.IsInstalled ? text20 + "-" : text20 + "I";
                            text20 = !item.IsMandatory ? text20 + "-" : text20 + "M";
                            text20 = !item.IsHidden ? text20 + "-" : text20 + "H";
                            text20 = !item.IsUninstallable ? text20 + "-" : text20 + "U";
                            text20 = !item.IsBeta ? text20 + "-" : text20 + "B";
                            var pSObject = new PSObject(item);
                            pSObject.Properties.Add(new PSNoteProperty("Size", text11));
                            pSObject.Properties.Add(new PSNoteProperty("Status", text20));
                            pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                            pSObject.Properties.Add(new PSNoteProperty("KB", text18));
                            pSObject.TypeNames.Clear();
                            pSObject.TypeNames.Add("PSWindowsUpdate.WindowsUpdate");
                            collection.Add(pSObject);
                        }

                        progressRecord.RecordType = ProgressRecordType.Completed;
                        WriteProgress(progressRecord);
                        var count2 = collection.Count;
                        WriteVerbose("Found [" + count2 + "] Updates in post search criteria");
                        if (Hide ||
                            string.Equals(MyInvocation.InvocationName, "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(MyInvocation.InvocationName, "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(MyInvocation.InvocationName, "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                            var text21 = !Hide ? "Show" : "Hide";
                            var num5 = 0;
                            var progressRecord2 = new ProgressRecord(1, text21 + " updates for " + target, "[" + num5 + "/" + count2 + "]");
                            foreach (var item2 in collection) {
                                var update2 = (IUpdate)item2.BaseObject;
                                var flag3 = false;
                                progressRecord2.StatusDescription = "[" + num5 + "/" + count2 + "] " + update2.Title + " " + item2.Properties["Size"].Value;
                                progressRecord2.PercentComplete = num5 * 100 / count2;
                                WriteProgress(progressRecord2);
                                num5++;
                                if (AcceptAll) {
                                    flag3 = true;
                                } else if (ShouldProcess(target,
                                               "(" + DateTime.Now + ") " + text21 + " " + update2.Title + "[" + item2.Properties["Size"].Value + "]")) {
                                    flag3 = true;
                                }

                                if (flag3) {
                                    try {
                                        update2.IsHidden = Hide;
                                        item2.Properties.Add(new PSNoteProperty("HideResult", text21));
                                    } catch {
                                        ThrowTerminatingError(new ErrorRecord(new Exception(target + ": You don't have permission to perform this task."), "0x80240044",
                                            ErrorCategory.CloseError, null));
                                    }
                                }

                                var text22 = "";
                                text22 = !update2.IsDownloaded ? text22 + "-" : text22 + "D";
                                text22 = !update2.IsInstalled ? text22 + "-" : text22 + "I";
                                text22 = !update2.IsMandatory ? text22 + "-" : text22 + "M";
                                text22 = !update2.IsHidden ? text22 + "-" : text22 + "H";
                                text22 = !update2.IsUninstallable ? text22 + "-" : text22 + "U";
                                text22 = !update2.IsBeta ? text22 + "-" : text22 + "B";
                                item2.Properties["Status"].Value = text22;
                            }

                            progressRecord2.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord2);
                        }

                        if (!Download && !Install) {
                            WriteDebug(DateTime.Now + " Return update list only");
                            WriteObject(collection);
                            OutputObj.AddRange(collection.ToList());
                            continue;
                        }

                        var totalDownloaded = 0;
                        var totalInstalled = 0;
                        var text23 = "";
                        if (Download || Install) {
                            var num8 = 0;
                            var activityId3 = 1;
                            var activity3 = "Choose updates for " + target;
                            var statusDescription3 = "[" + num8 + "/" + count2 + "]";
                            var progressRecord3 = new ProgressRecord(activityId3, activity3, statusDescription3);
                            var searchCritera = "";
                            foreach (var item3 in collection) {
                                item3.Properties.Add(new PSNoteProperty("X", 1));
                                item3.TypeNames.Clear();
                                item3.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                                var update3 = (IUpdate)item3.BaseObject;
                                progressRecord3.StatusDescription = "[" + num8 + "/" + count2 + "] " + update3.Title + " " + item3.Properties["Size"].Value;
                                progressRecord3.PercentComplete = num8 * 100 / count2;
                                WriteProgress(progressRecord3);
                                num8++;
                                WriteDebug(DateTime.Now + " Show update to accept: " + update3.Title);
                                var flag4 = false;
                                flag4 = AcceptAll || (AutoSelectOnly
                                    ? update3.AutoSelectOnWebSites ? true : false
                                    : ShouldProcess(target, "(" + DateTime.Now + ") " + update3.Title + "[" + item3.Properties["Size"].Value + "]")
                                        ? true
                                        : false);
                                var text25 = "";
                                var text26 = "";
                                if (flag4) {
                                    if (!update3.EulaAccepted) {
                                        WriteDebug(DateTime.Now + " EulaAccepted");
                                        try {
                                            update3.AcceptEula();
                                        } catch (Exception exception) {
                                            if (Debuger) {
                                                var errorRecord5 = new ErrorRecord(exception, "Debug", ErrorCategory.CloseError, null);
                                                WriteError(errorRecord5);
                                            } else {
                                                WriteDebug(DateTime.Now + " Can't accept Eula.");
                                            }
                                        }
                                    }

                                    try {
                                        var pSObject2 = new PSObject(item3.Properties["Identity"].Value);
                                        var text27 = (string)pSObject2.Properties["UpdateID"].Value;
                                        var num9 = (int)pSObject2.Properties["RevisionNumber"].Value;
                                        searchCritera = !(searchCritera == "")
                                            ? searchCritera + " or (UpdateID = '" + text27 + "' and RevisionNumber = " + num9 + ")"
                                            : "(UpdateID = '" + text27 + "' and RevisionNumber = " + num9 + ")";
                                    } catch (Exception ex3) {
                                        flag4 = false;
                                        var errorRecord6 = new ErrorRecord(new Exception("Something goes wrong: " + update3.Title + "; " + ex3.Message), "Debug",
                                            ErrorCategory.CloseError, null);
                                        WriteError(errorRecord6);
                                    }

                                    if (flag4) {
                                        text25 += "A";
                                        text26 = "Accepted";
                                        WriteDebug(DateTime.Now + " " + text26);
                                        item3.Properties.Add(new PSNoteProperty("ChooseResult", text26));
                                        item3.Properties.Add(new PSNoteProperty("Result", text26));
                                    }
                                }

                                if (!flag4) {
                                    try {
                                        var pSObject3 = new PSObject(item3.Properties["Identity"].Value);
                                        var text28 = (string)pSObject3.Properties["UpdateID"].Value;
                                        var num10 = (int)pSObject3.Properties["RevisionNumber"].Value;
                                        text23 = !(text23 == "") ? text23 + ",'" + text28 + "'" : "'" + text28 + "'";
                                    } catch (Exception ex4) {
                                        flag4 = false;
                                        var errorRecord7 = new ErrorRecord(new Exception("Something goes wrong: " + update3.Title + "; " + ex4.Message), "Debug",
                                            ErrorCategory.CloseError, null);
                                        WriteError(errorRecord7);
                                    }

                                    text25 += "R";
                                    text26 = "Rejected";
                                    WriteDebug(DateTime.Now + " " + text26);
                                    item3.Properties.Add(new PSNoteProperty("ChooseResult", text26));
                                    item3.Properties.Add(new PSNoteProperty("Result", text26));
                                }

                                if (!flag4 || (WUToolsObj.IsLocalHost(target) && !(ScheduleJob != DateTime.MinValue))) {
                                    text25 = !update3.IsDownloaded ? text25 + "-" : text25 + "D";
                                    text25 = !update3.IsInstalled ? text25 + "-" : text25 + "I";
                                } else {
                                    if (Download || Install) {
                                        text25 += "?";
                                    }

                                    if (Install) {
                                        text25 += "?";
                                    }
                                }

                                text25 = !update3.IsMandatory ? text25 + "-" : text25 + "M";
                                text25 = !update3.IsHidden ? text25 + "-" : text25 + "H";
                                text25 = !update3.IsUninstallable ? text25 + "-" : text25 + "U";
                                text25 = !update3.IsBeta ? text25 + "-" : text25 + "B";
                                item3.Properties["Status"].Value = text25;
                            }

                            progressRecord3.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord3);
                            if (ShowPreSearchCriteria) {
                                WriteVerbose("Chosen pre-Search Criteria: " + searchCritera);
                            }

                            var totalAccepted = collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                            WriteObject(collection, true);
                            WriteVerbose("Accepted [" + totalAccepted + "] Updates ready to Download");
                            if (totalAccepted > 0 && (!WUToolsObj.IsLocalHost(target) || ScheduleJob != DateTime.MinValue)) {
                                var cmdLine = "";
                                cmdLine = !Debuger ? "Get-WindowsUpdate -AcceptAll" : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";
                                cmdLine = cmdLine + " -Criteria \\\"" + searchCritera + "\\\"";
                                if (RecurseCycle > 1) {
                                    if (text23 != "") {
                                        cmdLine = cmdLine + " -NotUpdateID " + text23;
                                    }

                                    if (IsInstalled) {
                                        cmdLine += " -IsInstalled";
                                    }

                                    if (IsHidden) {
                                        cmdLine += " -IsHidden";
                                    }

                                    if (WithHidden) {
                                        cmdLine += " -WithHidden";
                                    }

                                    if (UpdateType != null) {
                                        cmdLine = cmdLine + " -UpdateType " + UpdateType;
                                    }

                                    if (DeploymentAction != null) {
                                        cmdLine = cmdLine + " -DeploymentAction " + DeploymentAction;
                                    }

                                    if (UpdateID != null) {
                                        cmdLine = cmdLine + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                                    }

                                    if (RevisionNumber > 0) {
                                        cmdLine = cmdLine + " -RevisionNumber " + RevisionNumber;
                                    }

                                    if (CategoryIDs != null) {
                                        cmdLine = cmdLine + " -CategoryIDs " + string.Join(",", CategoryIDs);
                                    }

                                    if (IsAssigned) {
                                        cmdLine = !IsAssigned ? cmdLine + " -IsAssigned:$false" : cmdLine + " -IsAssigned";
                                    }

                                    if (IsPresent) {
                                        cmdLine = !IsPresent ? cmdLine + " -IsPresent:$false" : cmdLine + " -IsPresent";
                                    }

                                    if (AutoSelectOnWebSites) {
                                        cmdLine = !AutoSelectOnWebSites ? cmdLine + " -AutoSelectOnWebSites:$false" : cmdLine + " -AutoSelectOnWebSites";
                                    }

                                    if (RootCategories != null) {
                                        cmdLine = cmdLine + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                                    }

                                    if (Category != null) {
                                        cmdLine = cmdLine + " -Category '" + string.Join("','", Category) + "'";
                                    }

                                    if (KBArticleID != null) {
                                        cmdLine = cmdLine + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                                    }

                                    if (Title != null) {
                                        cmdLine = cmdLine + " -Title '" + Title + "'";
                                    }

                                    if (Severity != null) {
                                        cmdLine = cmdLine + " -Severity '" + string.Join("','", Severity) + "'";
                                    }

                                    if (NotCategory != null) {
                                        cmdLine = cmdLine + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                                    }

                                    if (NotKBArticleID != null) {
                                        cmdLine = cmdLine + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                                    }

                                    if (NotTitle != null) {
                                        cmdLine = cmdLine + " -NotTitle '" + NotTitle + "'";
                                    }

                                    if (NotSeverity != null) {
                                        cmdLine = cmdLine + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                                    }

                                    if (IgnoreUserInput) {
                                        cmdLine += " -IgnoreUserInput";
                                    }

                                    if (IgnoreRebootRequired) {
                                        cmdLine += " -IgnoreRebootRequired";
                                    }

                                    if (AutoSelectOnly) {
                                        cmdLine += " -AutoSelectOnly";
                                    }

                                    if (MaxSize > 0) {
                                        cmdLine = cmdLine + " -MaxSize " + MaxSize;
                                    }

                                    if (MinSize > 0) {
                                        cmdLine = cmdLine + " -MinSize " + MinSize;
                                    }
                                }

                                if (Download) {
                                    cmdLine += " -Download";
                                }

                                if (Install) {
                                    cmdLine += " -Install";
                                }

                                if (IgnoreReboot) {
                                    cmdLine += " -IgnoreReboot";
                                } else if (AutoReboot) {
                                    cmdLine += " -AutoReboot";
                                } else if (ScheduleReboot != DateTime.MinValue) {
                                    cmdLine = cmdLine + " -ScheduleReboot '" + ScheduleReboot + "'";
                                }

                                if (WindowsUpdate) {
                                    cmdLine += " -WindowsUpdate";
                                } else if (MicrosoftUpdate) {
                                    cmdLine += " -MicrosoftUpdate";
                                } else if (ServiceID != null) {
                                    cmdLine = cmdLine + " -ServiceID '" + ServiceID + "'";
                                }

                                if (SendReport) {
                                    cmdLine += " -SendReport";
                                }

                                if (SendHistory) {
                                    cmdLine += " -SendHistory";
                                }

                                if (RecurseCycle > 1) {
                                    cmdLine = cmdLine + " -RecurseCycle " + RecurseCycle;
                                }

                                if (!AutoReboot && ScheduleReboot == DateTime.MinValue && !IgnoreReboot) {
                                    cmdLine += " -IgnoreReboot";
                                }

                                cmdLine += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                                var invokeWUJob = new InvokeWUJob();
                                invokeWUJob.ComputerName = new string[1] { target };
                                if (Credential != null) {
                                    invokeWUJob.Credential = Credential;
                                }

                                invokeWUJob.Script = cmdLine;
                                invokeWUJob.Debuger = true;
                                if (ScheduleJob != DateTime.MinValue) {
                                    if (ScheduleJob.AddSeconds(10.0) <= DateTime.Now) {
                                        WriteVerbose("Execution time is gone. ScheduleJob was set to 1 minute delay from now.");
                                        ScheduleJob = DateTime.Now.AddMinutes(1.0);
                                    }

                                    invokeWUJob.TriggerDate = ScheduleJob;
                                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob + "): ");
                                } else {
                                    invokeWUJob.RunNow = true;
                                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (Now): ");
                                }

                                WriteVerbose("powershell.exe -Command \"" + cmdLine + "\"");
                                var enumerable = invokeWUJob.Invoke();
                                foreach (var item4 in enumerable) {
                                    WriteObject(item4);
                                }

                                WriteDebug(DateTime.Now + " Return invoked update list");
                                OutputObj.AddRange(collection.ToList());
                                continue;
                            }

                            var num12 = 0;
                            var activityId4 = 1;
                            var activity4 = "Download updates for " + target;
                            var statusDescription4 = "[" + num12 + "/" + totalAccepted + "]";
                            var progressRecord4 = new ProgressRecord(activityId4, activity4, statusDescription4);
                            foreach (var item5 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted")) {
                                item5.Properties.Add(new PSNoteProperty("X", 2));
                                var update4 = (IUpdate)item5.BaseObject;
                                var updateCollection3 =
                                    (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                                updateCollection3.Add(update4);
                                progressRecord4.StatusDescription = "[" + num12 + "/" + totalAccepted + "] " + update4.Title + " " + item5.Properties["Size"].Value;
                                progressRecord4.PercentComplete = num12 * 100 / totalAccepted;
                                WriteProgress(progressRecord4);
                                num12++;
                                WriteDebug(DateTime.Now + " Show update to download: " + update4.Title);
                                IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                                updateDownloader.Updates = updateCollection3;
                                if (ForceDownload) {
                                    updateDownloader.IsForced = true;
                                }

                                IDownloadResult downloadResult;
                                try {
                                    downloadResult = updateDownloader.Download();
                                    WriteDebug(DateTime.Now + " Downloaded");
                                } catch (COMException ex5) {
                                    var wUApiCodeDetails3 = WUToolsObj.GetWUApiCodeDetails(ex5.ErrorCode);
                                    var flag5 = false;
                                    if (wUApiCodeDetails3 != null) {
                                        switch (wUApiCodeDetails3.CodeType) {
                                            case 2:
                                                WriteError(new ErrorRecord(new Exception(wUApiCodeDetails3.Description), wUApiCodeDetails3.HResult, ErrorCategory.CloseError,
                                                    null));
                                                flag5 = true;
                                                break;
                                            case 3:
                                                WriteWarning(wUApiCodeDetails3.HResult + ": " + wUApiCodeDetails3.Description);
                                                break;
                                        }

                                        if (flag5) {
                                            WriteDebug(DateTime.Now + " Skip to next computer");
                                            break;
                                        }
                                    } else if (Debuger) {
                                        var errorRecord8 = new ErrorRecord(ex5, "Debug", ErrorCategory.CloseError, null);
                                        ThrowTerminatingError(errorRecord8);
                                    }

                                    WriteDebug(DateTime.Now + " Skip to next update");
                                    continue;
                                }

                                var value2 = "";
                                switch (downloadResult.ResultCode) {
                                    case OperationResultCode.orcNotStarted:
                                        value2 = "NotStarted";
                                        break;
                                    case OperationResultCode.orcInProgress:
                                        value2 = "InProgress";
                                        break;
                                    case OperationResultCode.orcSucceeded:
                                        value2 = "Downloaded";
                                        break;
                                    case OperationResultCode.orcSucceededWithErrors:
                                        value2 = "DownloadedWithErrors";
                                        break;
                                    case OperationResultCode.orcFailed:
                                        value2 = "Failed";
                                        break;
                                    case OperationResultCode.orcAborted:
                                        value2 = "Aborted";
                                        break;
                                }

                                item5.Properties.Add(new PSNoteProperty("DownloadResult", value2));
                                item5.Properties.Add(new PSNoteProperty("Result", value2));
                                var text30 = "";
                                text30 = !(item5.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text30 + "R" : text30 + "A";
                                text30 = !(item5.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text30 + "F" : text30 + "D";
                                text30 = !update4.IsInstalled ? text30 + "-" : text30 + "I";
                                text30 = !update4.IsMandatory ? text30 + "-" : text30 + "M";
                                text30 = !update4.IsHidden ? text30 + "-" : text30 + "H";
                                text30 = !update4.IsUninstallable ? text30 + "-" : text30 + "U";
                                text30 = !update4.IsBeta ? text30 + "-" : text30 + "B";
                                item5.Properties["Status"].Value = text30;
                                WriteObject(item5, true);
                            }

                            progressRecord4.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord4);
                            totalDownloaded = collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                            WriteVerbose("Downloaded [" + totalDownloaded + "] Updates ready to Install");
                            if (!Install) {
                                WriteDebug(DateTime.Now + " Return downloaded update list");
                                OutputObj.AddRange(collection.ToList());
                                continue;
                            }
                        }

                        if (!Install) {
                            continue;
                        }

                        NeedsReboot = false;
                        var num13 = 0;
                        var activityId5 = 1;
                        var activity5 = "Install updates for " + target;
                        var statusDescription5 = "[" + num13 + "/" + totalDownloaded + "]";
                        var progressRecord5 = new ProgressRecord(activityId5, activity5, statusDescription5);
                        foreach (var item6 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded")) {
                            item6.Properties.Add(new PSNoteProperty("X", 3));
                            var update5 = (IUpdate)item6.BaseObject;
                            var updateCollection4 =
                                (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                            updateCollection4.Add(update5);
                            progressRecord5.StatusDescription = "[" + num13 + "/" + totalDownloaded + "] " + update5.Title + " " + item6.Properties["Size"].Value;
                            progressRecord5.PercentComplete = num13 * 100 / totalDownloaded;
                            WriteProgress(progressRecord5);
                            num13++;
                            WriteDebug(DateTime.Now + " Show update to install: " + update5.Title);
                            var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                            updateInstaller.Updates = updateCollection4;
                            if (ForceInstall) {
                                updateInstaller.IsForced = true;
                            }

                            IInstallationResult installationResult;
                            try {
                                installationResult = updateInstaller.Install();
                                WriteDebug(DateTime.Now + " Installed");
                            } catch (COMException ex6) {
                                var wUApiCodeDetails4 = WUToolsObj.GetWUApiCodeDetails(ex6.ErrorCode);
                                var flag6 = false;
                                if (wUApiCodeDetails4 != null) {
                                    switch (wUApiCodeDetails4.CodeType) {
                                        case 2:
                                            WriteError(new ErrorRecord(new Exception(wUApiCodeDetails4.Description), wUApiCodeDetails4.HResult, ErrorCategory.CloseError, null));
                                            flag6 = true;
                                            break;
                                        case 3:
                                            WriteWarning(wUApiCodeDetails4.HResult + ": " + wUApiCodeDetails4.Description);
                                            break;
                                    }

                                    if (flag6) {
                                        WriteDebug(DateTime.Now + " Skip to next computer");
                                        break;
                                    }
                                } else if (Debuger) {
                                    var errorRecord9 = new ErrorRecord(ex6, "Debug", ErrorCategory.CloseError, null);
                                    ThrowTerminatingError(errorRecord9);
                                }

                                WriteDebug(DateTime.Now + " Skip to next update");
                                continue;
                            }

                            if (!NeedsReboot) {
                                NeedsReboot = installationResult.RebootRequired;
                                WriteDebug(DateTime.Now + " Reboot is required");
                            }

                            var value3 = "";
                            switch (installationResult.ResultCode) {
                                case OperationResultCode.orcNotStarted:
                                    value3 = "NotStarted";
                                    break;
                                case OperationResultCode.orcInProgress:
                                    value3 = "InProgress";
                                    break;
                                case OperationResultCode.orcSucceeded:
                                    value3 = "Installed";
                                    break;
                                case OperationResultCode.orcSucceededWithErrors:
                                    value3 = "InstalledWithErrors";
                                    break;
                                case OperationResultCode.orcFailed:
                                    value3 = "Failed";
                                    break;
                                case OperationResultCode.orcAborted:
                                    value3 = "Aborted";
                                    break;
                            }

                            item6.Properties.Add(new PSNoteProperty("InstallResult", value3));
                            item6.Properties.Add(new PSNoteProperty("Result", value3));
                            var text31 = "";
                            text31 = !(item6.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text31 + "R" : text31 + "A";
                            text31 = !(item6.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text31 + "F" : text31 + "D";
                            text31 = !(item6.Properties["InstallResult"].Value.ToString() == "Installed")
                                ? text31 + "F"
                                : !installationResult.RebootRequired
                                    ? text31 + "I"
                                    : text31 + "R";
                            text31 = !update5.IsMandatory ? text31 + "-" : text31 + "M";
                            text31 = !update5.IsHidden ? text31 + "-" : text31 + "H";
                            text31 = !update5.IsUninstallable ? text31 + "-" : text31 + "U";
                            text31 = !update5.IsBeta ? text31 + "-" : text31 + "B";
                            item6.Properties["Status"].Value = text31;
                            WriteObject(item6, true);
                        }

                        progressRecord5.RecordType = ProgressRecordType.Completed;
                        WriteProgress(progressRecord5);
                        totalInstalled = collection.Where(x => x.Properties["Result"].Value.ToString() == "Installed").Count();
                        WriteVerbose("Installed [" + totalInstalled + "] Updates");
                        WriteDebug(DateTime.Now + " Return installed update list");
                        OutputObj.AddRange(collection.ToList());
                        if (totalInstalled > 0 && SendHistory) {
                            var cmdLine2 = "Get-WUHistory -SendReport -Last " + totalInstalled + " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                            var invokeWUJob2 = new InvokeWUJob();
                            invokeWUJob2.ComputerName = new string[1] { target };
                            if (Credential != null) {
                                invokeWUJob2.Credential = Credential;
                            }

                            invokeWUJob2.Script = cmdLine2;
                            invokeWUJob2.TaskName = "PSWindowsUpdate_History";
                            if (NeedsReboot) {
                                invokeWUJob2.TriggerAtStart = true;
                                WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (AtStart): powershell.exe -Command \"" + cmdLine2 + "\"");
                            } else {
                                invokeWUJob2.RunNow = true;
                                WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (Now): powershell.exe -Command \"" + cmdLine2 + "\"");
                            }

                            foreach (var item7 in invokeWUJob2.Invoke()) {
                                WriteObject(item7);
                            }
                        }

                        if (RecurseCycle <= 1) {
                            continue;
                        }

                        RecurseCycle--;

                        var cmdLine3 = !Debuger ? "Get-WindowsUpdate -AcceptAll" : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";

                        if (IsInstalled) {
                            cmdLine3 += " -IsInstalled";
                        }

                        if (IsHidden) {
                            cmdLine3 += " -IsHidden";
                        }

                        if (WithHidden) {
                            cmdLine3 += " -WithHidden";
                        }

                        if (UpdateType != null) {
                            cmdLine3 = cmdLine3 + " -UpdateType " + UpdateType;
                        }

                        if (DeploymentAction != null) {
                            cmdLine3 = cmdLine3 + " -DeploymentAction " + DeploymentAction;
                        }

                        if (UpdateID != null) {
                            cmdLine3 = cmdLine3 + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                        }

                        if (NotUpdateID != null) {
                            cmdLine3 = cmdLine3 + " -NotUpdateID '" + string.Join("','", NotUpdateID) + "'";
                        } else if (text23 != "") {
                            cmdLine3 = cmdLine3 + " -NotUpdateID " + text23;
                        }

                        if (RevisionNumber > 0) {
                            cmdLine3 = cmdLine3 + " -RevisionNumber " + RevisionNumber;
                        }

                        if (CategoryIDs != null) {
                            cmdLine3 = cmdLine3 + " -CategoryIDs '" + string.Join("','", CategoryIDs) + "'";
                        }

                        if (IsAssigned) {
                            cmdLine3 = !IsAssigned ? cmdLine3 + " -IsAssigned:$false" : cmdLine3 + " -IsAssigned";
                        }

                        if (IsPresent) {
                            cmdLine3 = !IsPresent ? cmdLine3 + " -IsPresent:$false" : cmdLine3 + " -IsPresent";
                        }

                        if (AutoSelectOnWebSites) {
                            cmdLine3 = !AutoSelectOnWebSites ? cmdLine3 + " -AutoSelectOnWebSites:$false" : cmdLine3 + " -AutoSelectOnWebSites";
                        }

                        if (RootCategories != null) {
                            cmdLine3 = cmdLine3 + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                        }

                        if (Category != null) {
                            cmdLine3 = cmdLine3 + " -Category '" + string.Join("','", Category) + "'";
                        }

                        if (KBArticleID != null) {
                            cmdLine3 = cmdLine3 + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                        }

                        if (Title != null) {
                            cmdLine3 = cmdLine3 + " -Title '" + Title + "'";
                        }

                        if (Severity != null) {
                            cmdLine3 = cmdLine3 + " -Severity '" + string.Join("','", Severity) + "'";
                        }

                        if (NotCategory != null) {
                            cmdLine3 = cmdLine3 + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                        }

                        if (NotKBArticleID != null) {
                            cmdLine3 = cmdLine3 + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                        }

                        if (NotTitle != null) {
                            cmdLine3 = cmdLine3 + " -NotTitle '" + NotTitle + "'";
                        }

                        if (NotSeverity != null) {
                            cmdLine3 = cmdLine3 + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                        }

                        if (IgnoreUserInput) {
                            cmdLine3 += " -IgnoreUserInput";
                        }

                        if (IgnoreRebootRequired) {
                            cmdLine3 += " -IgnoreRebootRequired";
                        }

                        if (AutoSelectOnly) {
                            cmdLine3 += " -AutoSelectOnly";
                        }

                        if (MaxSize > 0) {
                            cmdLine3 = cmdLine3 + " -MaxSize " + MaxSize;
                        }

                        if (MinSize > 0) {
                            cmdLine3 = cmdLine3 + " -MinSize " + MinSize;
                        }

                        if (Download) {
                            cmdLine3 += " -Download";
                        }

                        if (Install) {
                            cmdLine3 += " -Install";
                        }

                        if (IgnoreReboot) {
                            cmdLine3 += " -IgnoreReboot";
                        } else if (AutoReboot) {
                            cmdLine3 += " -AutoReboot";
                        } else if (ScheduleReboot != DateTime.MinValue) {
                            cmdLine3 = cmdLine3 + " -ScheduleReboot '" + ScheduleReboot + "'";
                        }

                        if (WindowsUpdate) {
                            cmdLine3 += " -WindowsUpdate";
                        } else if (MicrosoftUpdate) {
                            cmdLine3 += " -MicrosoftUpdate";
                        } else if (ServiceID != null) {
                            cmdLine3 = cmdLine3 + " -ServiceID '" + ServiceID + "'";
                        }

                        if (SendReport) {
                            cmdLine3 += " -SendReport";
                        }

                        if (SendHistory) {
                            cmdLine3 += " -SendHistory";
                        }

                        if (RecurseCycle > 1) {
                            cmdLine3 = cmdLine3 + " -RecurseCycle " + RecurseCycle;
                        }

                        cmdLine3 += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                        var invokeWUJob3 = new InvokeWUJob();
                        invokeWUJob3.ComputerName = new string[1] { target };
                        if (Credential != null) {
                            invokeWUJob3.Credential = Credential;
                        }

                        invokeWUJob3.Script = cmdLine3;
                        invokeWUJob3.TaskName = "PSWindowsUpdate_Recurse" + RecurseCycle;
                        invokeWUJob3.Debuger = true;
                        if (NeedsReboot) {
                            invokeWUJob3.TriggerAtStart = true;
                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (AtStart): ");
                        } else {
                            ScheduleJob = DateTime.Now.AddMinutes(5.0);
                            invokeWUJob3.TriggerDate = ScheduleJob;
                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob + "): ");
                        }

                        WriteVerbose("powershell.exe -Command \"" + cmdLine3 + "\"");
                        var enumerable3 = invokeWUJob3.Invoke();
                        foreach (var item8 in enumerable3) {
                            WriteObject(item8);
                        }
                    } else if (Debuger) {
                        var errorRecord10 = new ErrorRecord(wUApiServiceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord10);
                    } else {
                        var error = wUApiServiceManagerObj.Error;
                        WriteError(error);
                    }
                } else if (Debuger) {
                    var errorRecord11 = new ErrorRecord(wUApiUpdateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null);
                    WriteError(errorRecord11);
                } else {
                    var error2 = wUApiUpdateSessionObj.Error;
                    WriteError(error2);
                }
            }
        }

        protected override void ProcessRecord() {
            var flag = false;
            if (Credential != null) {
                var userName = Credential.GetNetworkCredential().UserName;
                var domain = Credential.GetNetworkCredential().Domain;
                var password = Credential.GetNetworkCredential().Password;
                WriteDebug(DateTime.Now + " UserName: " + userName + "; Domain: " + domain + "; Password: " + password.Substring(0, 1) + "*****");
                var windowsPrincipal1 = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                var str1 = "";
                if (windowsPrincipal1.IsInRole(WindowsBuiltInRole.Administrator)) {
                    str1 = "RunAs";
                }

                var strArray1 = new string[5];
                var now = DateTime.Now;
                strArray1[0] = now.ToString();
                strArray1[1] = " Before User: ";
                strArray1[2] = WindowsIdentity.GetCurrent().Name;
                strArray1[3] = " ";
                strArray1[4] = str1;
                WriteDebug(string.Concat(strArray1));
                var logonType = WUImpersonator.LogonSessionType.Interactive;
                var logonProvider = WUImpersonator.LogonProvider.Default;
                if (!WUToolsObj.IsLocalHost(ComputerName[0])) {
                    logonType = WUImpersonator.LogonSessionType.Interactive;
                    logonProvider = WUImpersonator.LogonProvider.Default;
                }

                using (new WUImpersonator(userName, domain, password, logonType, logonProvider)) {
                    if (new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                        var str2 = "RunAs";
                        var strArray2 = new string[9];
                        now = DateTime.Now;
                        strArray2[0] = now.ToString();
                        strArray2[1] = " Impersonated User (";
                        strArray2[2] = logonType.ToString();
                        strArray2[3] = "|";
                        strArray2[4] = logonProvider.ToString();
                        strArray2[5] = "): ";
                        strArray2[6] = WindowsIdentity.GetCurrent().Name;
                        strArray2[7] = " ";
                        strArray2[8] = str2;
                        WriteDebug(string.Concat(strArray2));
                        try {
                            CoreProcessing();
                            flag = false;
                        } catch (Exception ex) {
                            WriteDebug(DateTime.Now + " Something goes wrong: " + ex.Message);
                            flag = true;
                        }
                    } else {
                        var str3 = "Can't RunAs";
                        var strArray3 = new string[9];
                        now = DateTime.Now;
                        strArray3[0] = now.ToString();
                        strArray3[1] = " Impersonated User (";
                        strArray3[2] = logonType.ToString();
                        strArray3[3] = "|";
                        strArray3[4] = logonProvider.ToString();
                        strArray3[5] = "): ";
                        strArray3[6] = WindowsIdentity.GetCurrent().Name;
                        strArray3[7] = " ";
                        strArray3[8] = str3;
                        WriteDebug(string.Concat(strArray3));
                        flag = true;
                    }

                    WriteDebug(DateTime.Now + " Leaving impersonated session");
                }

                var windowsPrincipal2 = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                var str4 = "";
                if (windowsPrincipal2.IsInRole(WindowsBuiltInRole.Administrator)) {
                    str4 = "RunAs";
                }

                WriteDebug(DateTime.Now + " After User: " + WindowsIdentity.GetCurrent().Name + " " + str4);
            } else {
                flag = true;
            }

            if (!flag) {
                return;
            }

            CoreProcessing();
        }

        protected override void EndProcessing() {
            CmdletEnd = DateTime.Now;
            var CmdletInfo = new PSObject();
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletStart", CmdletStart));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletEnd", CmdletEnd));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletLine", MyInvocation.Line));
            if (SendReport && OutputObj.Count > 0) {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new string[5] {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
                        "Title"
                    });
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            if (NeedsReboot) {
                if (ScheduleReboot != DateTime.MinValue) {
                    WriteDebug(DateTime.Now + " Schedule Reboot " + ScheduleReboot);
                    WriteVerbose(WUToolsObj.ScheduleReboot("localhost", ScheduleReboot));
                } else if (AutoReboot) {
                    WriteDebug(DateTime.Now + " Auto Reboot");
                    WriteVerbose(WUToolsObj.RunReboot("localhost"));
                } else if (IgnoreReboot) {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required, but do it manually.");
                } else {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required. Do it now? [Y / N] (default is 'N')");
                    if (Console.ReadLine().ToUpper() == "Y") {
                        WriteDebug(DateTime.Now + " Manually Reboot");
                        WriteVerbose(WUToolsObj.RunReboot("localhost"));
                    }
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}