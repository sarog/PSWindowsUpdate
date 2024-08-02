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

namespace PSWindowsUpdate
{
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
    public class GetWindowsUpdate : PSCmdlet
    {
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
        public Hashtable PSWUSettings
        {
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
        [ValidateSet("Critical Updates", "Definition Updates", "Drivers", "Feature Packs", "Security Updates", "Service Packs", "Tools",
            "Update Rollups", "Updates", "Upgrades",
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
        protected override void BeginProcessing()
        {
            CmdletStart = DateTime.Now;
            WriteDebug(DateTime.Now + " CmdletStart: " + MyInvocation.InvocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                ThrowTerminatingError(new ErrorRecord(
                    new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = [];
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new WindowsUpdate() /*new string[5] {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
                        "Title"
                    }*/);
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = false;
                    WriteDebug(DateTime.Now + " Disabling -SendReport");
                }
            }

            if (ComputerName != null)
            {
                return;
            }

            ComputerName =
            [
                Environment.MachineName
            ];
        }

        private void CoreProcessing()
        {
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " ParameterSetName: " + ParameterSetName);
            if (string.Equals(invocationName, "Install-WindowsUpdate", StringComparison.OrdinalIgnoreCase))
            {
                Download = true;
                Install = true;
            }

            if (string.Equals(invocationName, "Download-WindowsUpdate", StringComparison.OrdinalIgnoreCase))
            {
                Download = true;
            }

            if (string.Equals(invocationName, "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase))
            {
                Hide = true;
            }

            if (string.Equals(invocationName, "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(invocationName, "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase))
            {
                Hide = false;
                WithHidden = true;
            }

            string criteria;
            if (Criteria != null)
            {
                criteria = Criteria;
                WriteDebug(DateTime.Now + " Set pre search criteria: " + criteria);
            }
            else
            {
                if (IsInstalled)
                {
                    criteria = "IsInstalled = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 1");
                }
                else
                {
                    criteria = "IsInstalled = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 0");
                }

                if (UpdateType != null)
                {
                    criteria = criteria + " and Type = '" + UpdateType + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: Type = '" + UpdateType + "'");
                }

                if (IsHidden)
                {
                    criteria += " and IsHidden = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1");
                }
                else if (WithHidden)
                {
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1 or IsHidden = 0");
                }
                else
                {
                    criteria += " and IsHidden = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 0");
                }

                if (IgnoreRebootRequired)
                {
                    criteria += " and RebootRequired = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: RebootRequired = 0");
                }

                if (DeploymentAction != null)
                {
                    criteria = criteria + " and DeploymentAction = '" + DeploymentAction + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: DeploymentAction = '" + DeploymentAction + "'");
                }

                if (IsAssigned)
                {
                    if (IsAssigned)
                    {
                        criteria += " and IsAssigned = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 1");
                    }
                    else
                    {
                        criteria += " and IsAssigned = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 0");
                    }
                }

                if (IsPresent)
                {
                    if (IsPresent)
                    {
                        criteria += " and IsPresent = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 1");
                    }
                    else
                    {
                        criteria += " and IsPresent = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 0");
                    }
                }

                if (BrowseOnly)
                {
                    if (BrowseOnly)
                    {
                        criteria += " and BrowseOnly = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 1");
                    }
                    else
                    {
                        criteria += " and BrowseOnly = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 0");
                    }
                }

                if (AutoSelectOnWebSites)
                {
                    if (AutoSelectOnWebSites)
                    {
                        criteria += " and AutoSelectOnWebSites = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 1");
                    }
                    else
                    {
                        criteria += " and AutoSelectOnWebSites = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 0");
                    }
                }

                if (NotUpdateID != null)
                {
                    foreach (var uid in NotUpdateID)
                    {
                        criteria = criteria + " and UpdateID != '" + uid + "'";
                        WriteDebug(DateTime.Now + " Set pre search criteria: NotUpdateID = " + uid);
                    }
                }

                if (UpdateID != null)
                {
                    var prevCriteria = criteria;
                    criteria = "";
                    var num = 0;
                    foreach (var uid in UpdateID)
                    {
                        if (num > 0)
                        {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        if (RevisionNumber > 0)
                        {
                            criteria = criteria + "(" + prevCriteria + " and UpdateID = '" + uid + "' and RevisionNumber = " + RevisionNumber + ")";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uid + "and RevisionNumber = " + RevisionNumber);
                        }
                        else
                        {
                            criteria = criteria + "(" + prevCriteria + " and UpdateID = '" + uid + "')";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uid);
                        }

                        ++num;
                    }
                }

                if (CategoryIDs != null)
                {
                    var prevCriteria = criteria;
                    criteria = "";
                    var num = 0;
                    foreach (var catID in CategoryIDs)
                    {
                        if (num > 0)
                        {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        criteria = criteria + "(" + prevCriteria + " and CategoryIDs contains '" + catID + "')";
                        WriteDebug(DateTime.Now + " Set pre search criteria: CategoryIDs = " + catID);
                        ++num;
                    }
                }
            }

            WriteDebug(DateTime.Now + " Search criteria is: " + criteria);
            if (ShowPreSearchCriteria)
            {
                Host.UI.WriteLine(ConsoleColor.Green, Host.UI.RawUI.BackgroundColor, "PreSearchCriteria: " + criteria);
            }

            foreach (var target in ComputerName)
            {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try
                {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pswuModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                }
                catch
                {
                }

                if (Download)
                {
                    var errorRecord = WUToolsObj.CheckPSWUModule(target);
                    if (errorRecord != null)
                    {
                        WriteError(errorRecord);
                        continue;
                    }
                }

                var updateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                WriteDebug(DateTime.Now + " UpdateSessionObj mode: " + updateSessionObj.Mode);
                if (updateSessionObj.Status)
                {
                    UpdateSessionObj = (UpdateSession)updateSessionObj.Object;
                    SearcherObj = UpdateSessionObj.CreateUpdateSearcher();
                    var serviceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                    WriteDebug(DateTime.Now + " ServiceManagerObj mode: " + serviceManagerObj.Mode);
                    if (serviceManagerObj.Status)
                    {
                        ServiceManagerObj = (UpdateServiceManager)serviceManagerObj.Object;
                        var text8 = "";
                        if (WindowsUpdate)
                        {
                            WriteDebug(DateTime.Now + " Set source of updates to Windows Update");
                            SearcherObj.ServerSelection = ServerSelection.ssWindowsUpdate;
                            text8 = "Windows Update";
                        }
                        else if (MicrosoftUpdate)
                        {
                            WriteDebug(DateTime.Now + " Set source of updates to Microsof Update");
                            ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                            var flag = false;
                            foreach (IUpdateService2 service in ServiceManagerObj.Services)
                            {
                                if (service.ServiceID == ServiceID)
                                {
                                    WriteDebug(DateTime.Now + " Search for ServiceID. Set source of updates to " + service.Name);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = service.Name;
                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag)
                            {
                                try
                                {
                                    Thread.Sleep(500);
                                    var serviceController = new ServiceController("Windows Update", target);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                    ServiceManagerObj.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 2, null);
                                    WriteDebug(DateTime.Now + " Register WUA Microsoft Update service");
                                    Thread.Sleep(500);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                }
                                catch (COMException ex)
                                {
                                    var wuApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                    if (wuApiCodeDetails != null)
                                    {
                                        switch (wuApiCodeDetails.CodeType)
                                        {
                                            case 2:
                                                WriteError(new ErrorRecord(new Exception(wuApiCodeDetails.Description),
                                                    wuApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                                break;
                                            case 3:
                                                WriteWarning(wuApiCodeDetails.HResult + ": " + wuApiCodeDetails.Description);
                                                break;
                                        }
                                    }
                                    else if (Debuger)
                                    {
                                        ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                                    }
                                }

                                SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                SearcherObj.ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                                text8 = "Microsoft Update";
                            }
                        }
                        else
                        {
                            try
                            {
                                foreach (IUpdateService2 service in ServiceManagerObj.Services)
                                {
                                    if (ServiceID != null)
                                    {
                                        if (service.ServiceID == ServiceID)
                                        {
                                            WriteDebug(DateTime.Now + " Try ServiceID. Set source of updates to " + service.Name);
                                            SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                            SearcherObj.ServiceID = ServiceID;
                                            text8 = service.Name;
                                            break;
                                        }
                                    }
                                    else if (service.IsDefaultAUService)
                                    {
                                        WriteDebug(DateTime.Now + " Try Default. Set source of updates to " + service.Name);
                                        text8 = service.Name;
                                        break;
                                    }
                                }
                            }
                            catch
                            {
                                if (ServiceID != null)
                                {
                                    WriteDebug(DateTime.Now + " Catch ServiceID. Set source of updates to " + ServiceID);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = ServiceID;
                                }
                                else
                                {
                                    WriteDebug(DateTime.Now + " Catch. Set source of updates to Default");
                                    text8 = "default for " + target + ".";
                                }
                            }
                        }

                        if (text8 == "Windows Server Update Service")
                        {
                            var regValue = (!WUToolsObj.IsLocalHost(target)
                                    ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                                    : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default))
                                .OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\")
                                .GetValue("WUServer").ToString();
                            text8 = text8 + " (" + regValue + ")";
                        }

                        WriteVerbose(target + " (" + DateTime.Now + "): Connecting to " + text8 + " server. Please wait...");
                        ISearchResult searchResult;
                        try
                        {
                            searchResult = SearcherObj.Search(criteria);
                        }
                        catch (COMException ex)
                        {
                            var wuApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                            if (wuApiCodeDetails != null)
                            {
                                switch (wuApiCodeDetails.CodeType)
                                {
                                    case 2:
                                        WriteError(new ErrorRecord(new Exception(wuApiCodeDetails.Description), wuApiCodeDetails.HResult,
                                            ErrorCategory.CloseError, null));
                                        break;
                                    case 3:
                                        WriteWarning(wuApiCodeDetails.HResult + ": " + wuApiCodeDetails.Description);
                                        break;
                                }
                            }
                            else if (Debuger)
                            {
                                ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                            }

                            WriteDebug(DateTime.Now + " Skip to next computer");
                            continue;
                        }

                        var count = searchResult.Updates.Count;
                        WriteVerbose("Found [" + count + "] Updates in pre search criteria");
                        if (count != 0)
                        {
                            UpdateCollection updateCollection;

                            if (RootCategories != null)
                            {
                                var updateCollection2 =
                                    (UpdateCollection)Activator.CreateInstance(
                                        Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));

                                foreach (var rootCategory in RootCategories)
                                {
                                    var index = -1;
                                    switch (rootCategory)
                                    {
                                        case "Critical Updates":
                                            index = 0;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Definition Updates":
                                            index = 1;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Drivers":
                                            index = 2;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Feature Packs":
                                            index = 3;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Security Updates":
                                            index = 4;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Service Packs":
                                            index = 5;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Tools":
                                            index = 6;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Update Rollups":
                                            index = 7;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Updates":
                                            index = 8;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Upgrades":
                                            index = 9;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                        case "Microsoft":
                                            index = 10;
                                            WriteDebug(DateTime.Now + " RootCategory: " + rootCategory + " - " + index);
                                            break;
                                    }

                                    try
                                    {
                                        foreach (IUpdate update in searchResult.RootCategories[index].Updates)
                                        {
                                            updateCollection2.Add(update);
                                        }
                                    }
                                    catch
                                    {
                                        WriteWarning("RootCategories Updates are empty. Use classic filters.");
                                    }
                                }

                                updateCollection = updateCollection2;
                            }
                            else
                            {
                                updateCollection = searchResult.Updates;
                            }

                            var num3 = 0;
                            var progressRecord = new ProgressRecord(0, "Post search updates for " + target, "[" + num3 + "/" + count + "]");
                            var collection = new Collection<PSObject>();
                            foreach (IUpdate update in updateCollection)
                            {
                                WriteDebug(DateTime.Now + " " + update.Title);
                                var text11 = WUToolsObj.ConvertSize(update.MaxDownloadSize);
                                progressRecord.StatusDescription = "[" + num3 + "/" + count + "] " + update.Title + " " + text11;
                                progressRecord.PercentComplete = num3 * 100 / count;
                                WriteProgress(progressRecord);
                                ++num3;
                                var foundMatch = true;
                                if (Category != null || NotCategory != null)
                                {
                                    var input = "";
                                    foreach (ICategory category in update.Categories)
                                    {
                                        input = input + category.Name + "|";
                                    }

                                    WriteDebug(DateTime.Now + " " + input);
                                    if (Category != null)
                                    {
                                        WriteDebug(DateTime.Now + " Set post search criteria: Categories = " + string.Join(", ", Category));
                                        foreach (var pattern in Category)
                                        {
                                            if (!Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase))
                                            {
                                                foundMatch = false;
                                            }
                                            else
                                            {
                                                foundMatch = true;
                                                break;
                                            }
                                        }
                                    }

                                    if (NotCategory != null && foundMatch)
                                    {
                                        WriteDebug(
                                            DateTime.Now + " Set post search criteria: NotCategories = " + string.Join(", ", NotCategory));
                                        foreach (var pattern in NotCategory)
                                        {
                                            if (Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase))
                                            {
                                                foundMatch = false;
                                                break;
                                            }
                                        }
                                    }

                                    WriteDebug(DateTime.Now + " UpdateAccess: " + foundMatch);
                                }

                                if (KBArticleID != null && foundMatch)
                                {
                                    var input = string.Join(", ", KBArticleID);
                                    WriteDebug(DateTime.Now + " Set post search criteria: KBArticleIDs = " + input);
                                    if (update.KBArticleIDs.Count > 0)
                                    {
                                        foreach (string kBArticleID in update.KBArticleIDs)
                                        {
                                            if (!Regex.IsMatch(input, kBArticleID, RegexOptions.IgnoreCase))
                                            {
                                                foundMatch = false;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foundMatch = false;
                                    }

                                    WriteDebug(DateTime.Now + " UpdateAccess: " + foundMatch);
                                }

                                if (NotKBArticleID != null && foundMatch)
                                {
                                    var input = string.Join(", ", NotKBArticleID);
                                    WriteDebug(DateTime.Now + " Set post search criteria: NotKBArticleID = " + input);
                                    if (update.KBArticleIDs.Count > 0)
                                    {
                                        foreach (string kBArticleID2 in update.KBArticleIDs)
                                        {
                                            if (Regex.IsMatch(input, kBArticleID2, RegexOptions.IgnoreCase))
                                            {
                                                foundMatch = false;
                                                break;
                                            }
                                        }
                                    }

                                    WriteDebug(DateTime.Now + " UpdateAccess: " + foundMatch);
                                }

                                if (Title != null && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: Title = " + Title);
                                    if (!Regex.IsMatch(update.Title, Title, RegexOptions.IgnoreCase))
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess: " + foundMatch);
                                    }
                                }

                                if (NotTitle != null && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: NotTitle = " + NotTitle);
                                    if (Regex.IsMatch(update.Title, NotTitle, RegexOptions.IgnoreCase))
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess: " + foundMatch);
                                    }
                                }

                                if (Severity != null && foundMatch)
                                {
                                    var text16 = string.Join(", ", Severity);
                                    WriteDebug(DateTime.Now + " Set post search criteria: Severity = " + text16);
                                    if (Severity.Contains("Unspecified"))
                                    {
                                        Severity = string.Join(",", Severity).Replace("Unspecified", "").Split(',');
                                    }

                                    if (!Severity.Contains(Convert.ToString((object)update.MsrcSeverity)))
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (NotSeverity != null && foundMatch)
                                {
                                    var text17 = string.Join(", ", NotSeverity);
                                    WriteDebug(DateTime.Now + " Set post search criteria: NotSeverity = " + text17);
                                    if (NotSeverity.Contains("Unspecified"))
                                    {
                                        NotSeverity = string.Join(",", NotSeverity).Replace("Unspecified", "").Split(',');
                                    }

                                    if (NotSeverity.Contains(Convert.ToString((object)update.MsrcSeverity)))
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (MaxSize > 0 && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: MaxDownloadSize <= " + MaxSize);
                                    if (MaxSize <= update.MaxDownloadSize)
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (MinSize > 0 && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: MaxDownloadSize >= " + MinSize);
                                    if (MinSize >= update.MaxDownloadSize)
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (IgnoreUserInput && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: CanRequestUserInput");
                                    if (update.InstallationBehavior.CanRequestUserInput)
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (IgnoreRebootRequired && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: RebootBehavior");
                                    if (update.InstallationBehavior.RebootBehavior != 0)
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (AutoSelectOnly && foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Set post search criteria: AutoSelectOnWebsites");
                                    if (!update.AutoSelectOnWebSites)
                                    {
                                        foundMatch = false;
                                        WriteDebug(DateTime.Now + " UpdateAccess" + foundMatch);
                                    }
                                }

                                if (foundMatch)
                                {
                                    WriteDebug(DateTime.Now + " Update was not filtered");
                                    var text18 = "";
                                    foreach (string kbArticleId in update.KBArticleIDs)
                                    {
                                        text18 = !(text18 == "") ? text18 + ", KB" + kbArticleId : text18 + "KB" + kbArticleId;
                                    }

                                    var text20 = "-";
                                    text20 = !update.IsDownloaded ? text20 + "-" : text20 + "D";
                                    text20 = !update.IsInstalled ? text20 + "-" : text20 + "I";
                                    text20 = !update.IsMandatory ? text20 + "-" : text20 + "M";
                                    text20 = !update.IsHidden ? text20 + "-" : text20 + "H";
                                    text20 = !update.IsUninstallable ? text20 + "-" : text20 + "U";
                                    text20 = !update.IsBeta ? text20 + "-" : text20 + "B";
                                    var psObject = new PSObject(update);
                                    psObject.Properties.Add(new PSNoteProperty("Size", text11));
                                    psObject.Properties.Add(new PSNoteProperty("Status", text20));
                                    psObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                                    psObject.Properties.Add(new PSNoteProperty("KB", text18));
                                    psObject.TypeNames.Clear();
                                    psObject.TypeNames.Add("PSWindowsUpdate.WindowsUpdate");
                                    collection.Add(psObject);
                                }
                            }

                            progressRecord.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord);
                            var count2 = collection.Count;
                            WriteVerbose("Found [" + count2 + "] Updates in post search criteria");
                            if (Hide ||
                                string.Equals(invocationName, "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(invocationName, "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(invocationName, "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase))
                            {
                                var text21 = !Hide ? "Show" : "Hide";
                                var num5 = 0;
                                var progressRecord2 = new ProgressRecord(1, text21 + " updates for " + target, "[" + num5 + "/" + count2 + "]");
                                foreach (var item in collection)
                                {
                                    var update2 = (IUpdate)item.BaseObject;
                                    var flag = false;
                                    progressRecord2.StatusDescription = "[" + num5 + "/" + count2 + "] " + update2.Title + " " +
                                                                        item.Properties["Size"].Value;
                                    progressRecord2.PercentComplete = num5 * 100 / count2;
                                    WriteProgress(progressRecord2);
                                    ++num5;
                                    if (AcceptAll)
                                    {
                                        flag = true;
                                    }
                                    else if (ShouldProcess(target,
                                                 "(" + DateTime.Now + ") " + text21 + " " + update2.Title + "[" +
                                                 item.Properties["Size"].Value + "]"))
                                    {
                                        flag = true;
                                    }

                                    if (flag)
                                    {
                                        try
                                        {
                                            update2.IsHidden = Hide;
                                            item.Properties.Add(new PSNoteProperty("HideResult", text21));
                                        }
                                        catch
                                        {
                                            ThrowTerminatingError(new ErrorRecord(
                                                new Exception(target + ": You don't have permission to perform this task."), "0x80240044",
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
                                    item.Properties["Status"].Value = text22;
                                }

                                progressRecord2.RecordType = ProgressRecordType.Completed;
                                WriteProgress(progressRecord2);
                            }

                            if (!Download && !Install)
                            {
                                WriteDebug(DateTime.Now + " Return update list only");
                                WriteObject(collection);
                                OutputObj.AddRange(collection.ToList());
                            }
                            else
                            {
                                var totalDownloaded = 0;
                                var totalInstalled = 0;
                                var text23 = "";
                                if (Download || Install)
                                {
                                    var num8 = 0;
                                    var statusDescription3 = "[" + num8 + "/" + count2 + "]";
                                    var progressRecord3 = new ProgressRecord(1, "Choose updates for " + target, statusDescription3);
                                    var searchCritera = "";
                                    foreach (var item in collection)
                                    {
                                        item.Properties.Add(new PSNoteProperty("X", 1));
                                        item.TypeNames.Clear();
                                        item.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                                        var update = (IUpdate)item.BaseObject;
                                        progressRecord3.StatusDescription = "[" + num8 + "/" + count2 + "] " + update.Title + " " +
                                                                            item.Properties["Size"].Value;
                                        progressRecord3.PercentComplete = num8 * 100 / count2;
                                        WriteProgress(progressRecord3);
                                        ++num8;
                                        WriteDebug(DateTime.Now + " Show update to accept: " + update.Title);
                                        var flag = AcceptAll || (AutoSelectOnly
                                            ? update.AutoSelectOnWebSites
                                            : ShouldProcess(target, "(" + DateTime.Now + ") " + update.Title + "[" + item.Properties["Size"].Value + "]"));
                                        var text25 = "";
                                        var status = "";
                                        if (flag)
                                        {
                                            if (!update.EulaAccepted)
                                            {
                                                WriteDebug(DateTime.Now + " EulaAccepted");
                                                try
                                                {
                                                    update.AcceptEula();
                                                }
                                                catch (Exception ex)
                                                {
                                                    if (Debuger)
                                                    {
                                                        WriteError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                                                    }
                                                    else
                                                    {
                                                        WriteDebug(DateTime.Now + " Can't accept Eula.");
                                                    }
                                                }
                                            }

                                            try
                                            {
                                                var identity = new PSObject(item.Properties["Identity"].Value);
                                                var updateId = (string)identity.Properties["UpdateID"].Value;
                                                var revision = (int)identity.Properties["RevisionNumber"].Value;
                                                searchCritera = !(searchCritera == "")
                                                    ? searchCritera + " or (UpdateID = '" + updateId + "' and RevisionNumber = " + revision + ")"
                                                    : "(UpdateID = '" + updateId + "' and RevisionNumber = " + revision + ")";
                                            }
                                            catch (Exception ex)
                                            {
                                                // ignored
                                            }

                                            if (flag)
                                            {
                                                text25 += "A";
                                                status = "Accepted";
                                                WriteDebug(DateTime.Now + " " + status);
                                                item.Properties.Add(new PSNoteProperty("ChooseResult", status));
                                                item.Properties.Add(new PSNoteProperty("Result", status));
                                            }
                                        }

                                        if (!flag)
                                        {
                                            try
                                            {
                                                var identity = new PSObject(item.Properties["Identity"].Value);
                                                var updateId = (string)identity.Properties["UpdateID"].Value;
                                                var revision = (int)identity.Properties["RevisionNumber"].Value;
                                                text23 = !(text23 == "") ? text23 + ",'" + updateId + "'" : "'" + updateId + "'";
                                            }
                                            catch (Exception ex)
                                            {
                                                // ignored
                                            }

                                            text25 += "R";
                                            status = "Rejected";
                                            WriteDebug(DateTime.Now + " " + status);
                                            item.Properties.Add(new PSNoteProperty("ChooseResult", status));
                                            item.Properties.Add(new PSNoteProperty("Result", status));
                                        }

                                        if (!flag || (WUToolsObj.IsLocalHost(target) && !(ScheduleJob != DateTime.MinValue)))
                                        {
                                            text25 = !update.IsDownloaded ? text25 + "-" : text25 + "D";
                                            text25 = !update.IsInstalled ? text25 + "-" : text25 + "I";
                                        }
                                        else
                                        {
                                            if (Download || Install)
                                            {
                                                text25 += "?";
                                            }

                                            if (Install)
                                            {
                                                text25 += "?";
                                            }
                                        }

                                        text25 = !update.IsMandatory ? text25 + "-" : text25 + "M";
                                        text25 = !update.IsHidden ? text25 + "-" : text25 + "H";
                                        text25 = !update.IsUninstallable ? text25 + "-" : text25 + "U";
                                        text25 = !update.IsBeta ? text25 + "-" : text25 + "B";
                                        item.Properties["Status"].Value = text25;
                                    }

                                    progressRecord3.RecordType = ProgressRecordType.Completed;
                                    WriteProgress(progressRecord3);
                                    if (ShowPreSearchCriteria)
                                    {
                                        WriteVerbose("Chosen pre-Search Criteria: " + searchCritera);
                                    }

                                    var totalAccepted = collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                                    WriteObject(collection, true);
                                    WriteVerbose("Accepted [" + totalAccepted + "] Updates ready to Download");
                                    if (totalAccepted > 0 && (!WUToolsObj.IsLocalHost(target) || ScheduleJob != DateTime.MinValue))
                                    {
                                        var cmdLine = "";
                                        cmdLine = !Debuger
                                            ? "Get-WindowsUpdate -AcceptAll"
                                            : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";
                                        cmdLine = cmdLine + " -Criteria \\\"" + searchCritera + "\\\"";
                                        if (RecurseCycle > 1)
                                        {
                                            if (text23 != "")
                                            {
                                                cmdLine = cmdLine + " -NotUpdateID " + text23;
                                            }

                                            if (IsInstalled)
                                            {
                                                cmdLine += " -IsInstalled";
                                            }

                                            if (IsHidden)
                                            {
                                                cmdLine += " -IsHidden";
                                            }

                                            if (WithHidden)
                                            {
                                                cmdLine += " -WithHidden";
                                            }

                                            if (UpdateType != null)
                                            {
                                                cmdLine = cmdLine + " -UpdateType " + UpdateType;
                                            }

                                            if (DeploymentAction != null)
                                            {
                                                cmdLine = cmdLine + " -DeploymentAction " + DeploymentAction;
                                            }

                                            if (UpdateID != null)
                                            {
                                                cmdLine = cmdLine + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                                            }

                                            if (RevisionNumber > 0)
                                            {
                                                cmdLine = cmdLine + " -RevisionNumber " + RevisionNumber;
                                            }

                                            if (CategoryIDs != null)
                                            {
                                                cmdLine = cmdLine + " -CategoryIDs " + string.Join(",", CategoryIDs);
                                            }

                                            if (IsAssigned)
                                            {
                                                cmdLine = !IsAssigned ? cmdLine + " -IsAssigned:$false" : cmdLine + " -IsAssigned";
                                            }

                                            if (IsPresent)
                                            {
                                                cmdLine = !IsPresent ? cmdLine + " -IsPresent:$false" : cmdLine + " -IsPresent";
                                            }

                                            if (AutoSelectOnWebSites)
                                            {
                                                cmdLine = !AutoSelectOnWebSites
                                                    ? cmdLine + " -AutoSelectOnWebSites:$false"
                                                    : cmdLine + " -AutoSelectOnWebSites";
                                            }

                                            if (RootCategories != null)
                                            {
                                                cmdLine = cmdLine + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                                            }

                                            if (Category != null)
                                            {
                                                cmdLine = cmdLine + " -Category '" + string.Join("','", Category) + "'";
                                            }

                                            if (KBArticleID != null)
                                            {
                                                cmdLine = cmdLine + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                                            }

                                            if (Title != null)
                                            {
                                                cmdLine = cmdLine + " -Title '" + Title + "'";
                                            }

                                            if (Severity != null)
                                            {
                                                cmdLine = cmdLine + " -Severity '" + string.Join("','", Severity) + "'";
                                            }

                                            if (NotCategory != null)
                                            {
                                                cmdLine = cmdLine + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                                            }

                                            if (NotKBArticleID != null)
                                            {
                                                cmdLine = cmdLine + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                                            }

                                            if (NotTitle != null)
                                            {
                                                cmdLine = cmdLine + " -NotTitle '" + NotTitle + "'";
                                            }

                                            if (NotSeverity != null)
                                            {
                                                cmdLine = cmdLine + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                                            }

                                            if (IgnoreUserInput)
                                            {
                                                cmdLine += " -IgnoreUserInput";
                                            }

                                            if (IgnoreRebootRequired)
                                            {
                                                cmdLine += " -IgnoreRebootRequired";
                                            }

                                            if (AutoSelectOnly)
                                            {
                                                cmdLine += " -AutoSelectOnly";
                                            }

                                            if (MaxSize > 0)
                                            {
                                                cmdLine = cmdLine + " -MaxSize " + MaxSize;
                                            }

                                            if (MinSize > 0)
                                            {
                                                cmdLine = cmdLine + " -MinSize " + MinSize;
                                            }
                                        }

                                        if (Download)
                                        {
                                            cmdLine += " -Download";
                                        }

                                        if (Install)
                                        {
                                            cmdLine += " -Install";
                                        }

                                        if (IgnoreReboot)
                                        {
                                            cmdLine += " -IgnoreReboot";
                                        }
                                        else if (AutoReboot)
                                        {
                                            cmdLine += " -AutoReboot";
                                        }
                                        else if (ScheduleReboot != DateTime.MinValue)
                                        {
                                            cmdLine = cmdLine + " -ScheduleReboot '" + ScheduleReboot + "'";
                                        }

                                        if (WindowsUpdate)
                                        {
                                            cmdLine += " -WindowsUpdate";
                                        }
                                        else if (MicrosoftUpdate)
                                        {
                                            cmdLine += " -MicrosoftUpdate";
                                        }
                                        else if (ServiceID != null)
                                        {
                                            cmdLine = cmdLine + " -ServiceID '" + ServiceID + "'";
                                        }

                                        if (SendReport)
                                        {
                                            cmdLine += " -SendReport";
                                        }

                                        if (SendHistory)
                                        {
                                            cmdLine += " -SendHistory";
                                        }

                                        if (RecurseCycle > 1)
                                        {
                                            cmdLine = cmdLine + " -RecurseCycle " + RecurseCycle;
                                        }

                                        if (!AutoReboot && ScheduleReboot == DateTime.MinValue && !IgnoreReboot)
                                        {
                                            cmdLine += " -IgnoreReboot";
                                        }

                                        cmdLine += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                                        var invokeWuJob = new InvokeWUJob();
                                        invokeWuJob.ComputerName = [target];
                                        if (Credential != null)
                                        {
                                            invokeWuJob.Credential = Credential;
                                        }

                                        invokeWuJob.Script = cmdLine;
                                        invokeWuJob.Debuger = true;
                                        if (ScheduleJob != DateTime.MinValue)
                                        {
                                            if (ScheduleJob.AddSeconds(10.0) <= DateTime.Now)
                                            {
                                                WriteVerbose("Execution time is gone. ScheduleJob was set to 1 minute delay from now.");
                                                ScheduleJob = DateTime.Now.AddMinutes(1.0);
                                            }

                                            invokeWuJob.TriggerDate = ScheduleJob;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob + "): ");
                                        }
                                        else
                                        {
                                            invokeWuJob.RunNow = true;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (Now): ");
                                        }

                                        WriteVerbose("powershell.exe -Command \"" + cmdLine + "\"");
                                        foreach (var sendToPipeline in invokeWuJob.Invoke())
                                        {
                                            WriteObject(sendToPipeline);
                                        }

                                        WriteDebug(DateTime.Now + " Return invoked update list");
                                        OutputObj.AddRange(collection.ToList());
                                        continue;
                                    }

                                    var num12 = 0;
                                    var progressRecord4 = new ProgressRecord(1, "Download updates for " + target, "[" + num12 + "/" + totalAccepted + "]");
                                    foreach (var sendToPipeline in collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted"))
                                    {
                                        sendToPipeline.Properties.Add(new PSNoteProperty("X", 2));
                                        var update = (IUpdate)sendToPipeline.BaseObject;
                                        var updateCollection3 =
                                            (UpdateCollection)Activator.CreateInstance(
                                                Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                                        updateCollection3.Add(update);
                                        progressRecord4.StatusDescription = "[" + num12 + "/" + totalAccepted + "] " + update.Title + " " +
                                                                            sendToPipeline.Properties["Size"].Value;
                                        progressRecord4.PercentComplete = num12 * 100 / totalAccepted;
                                        WriteProgress(progressRecord4);
                                        ++num12;
                                        WriteDebug(DateTime.Now + " Show update to download: " + update.Title);
                                        IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                                        updateDownloader.Updates = updateCollection3;
                                        if (ForceDownload)
                                        {
                                            updateDownloader.IsForced = true;
                                        }

                                        IDownloadResult downloadResult;
                                        try
                                        {
                                            downloadResult = updateDownloader.Download();
                                            WriteDebug(DateTime.Now + " Downloaded");
                                        }
                                        catch (COMException ex)
                                        {
                                            var wuApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                            var skip = false;
                                            if (wuApiCodeDetails != null)
                                            {
                                                switch (wuApiCodeDetails.CodeType)
                                                {
                                                    case 2:
                                                        WriteError(new ErrorRecord(new Exception(wuApiCodeDetails.Description),
                                                            wuApiCodeDetails.HResult, ErrorCategory.CloseError,
                                                            null));
                                                        skip = true;
                                                        break;
                                                    case 3:
                                                        WriteWarning(wuApiCodeDetails.HResult + ": " + wuApiCodeDetails.Description);
                                                        break;
                                                }

                                                if (skip)
                                                {
                                                    WriteDebug(DateTime.Now + " Skip to next computer");
                                                    break;
                                                }
                                            }
                                            else if (Debuger)
                                            {
                                                ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                                            }

                                            WriteDebug(DateTime.Now + " Skip to next update");
                                            continue;
                                        }

                                        var downloadMsg = "";
                                        switch (downloadResult.ResultCode)
                                        {
                                            case OperationResultCode.orcNotStarted:
                                                downloadMsg = "NotStarted";
                                                break;
                                            case OperationResultCode.orcInProgress:
                                                downloadMsg = "InProgress";
                                                break;
                                            case OperationResultCode.orcSucceeded:
                                                downloadMsg = "Downloaded";
                                                break;
                                            case OperationResultCode.orcSucceededWithErrors:
                                                downloadMsg = "DownloadedWithErrors";
                                                break;
                                            case OperationResultCode.orcFailed:
                                                downloadMsg = "Failed";
                                                break;
                                            case OperationResultCode.orcAborted:
                                                downloadMsg = "Aborted";
                                                break;
                                        }

                                        sendToPipeline.Properties.Add(new PSNoteProperty("DownloadResult", downloadMsg));
                                        sendToPipeline.Properties.Add(new PSNoteProperty("Result", downloadMsg));
                                        var text30 = "";
                                        text30 = !(sendToPipeline.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text30 + "R" : text30 + "A";
                                        text30 = !(sendToPipeline.Properties["DownloadResult"].Value.ToString() == "Downloaded")
                                            ? text30 + "F"
                                            : text30 + "D";
                                        text30 = !update.IsInstalled ? text30 + "-" : text30 + "I";
                                        text30 = !update.IsMandatory ? text30 + "-" : text30 + "M";
                                        text30 = !update.IsHidden ? text30 + "-" : text30 + "H";
                                        text30 = !update.IsUninstallable ? text30 + "-" : text30 + "U";
                                        text30 = !update.IsBeta ? text30 + "-" : text30 + "B";
                                        sendToPipeline.Properties["Status"].Value = text30;
                                        WriteObject(sendToPipeline, true);
                                    }

                                    progressRecord4.RecordType = ProgressRecordType.Completed;
                                    WriteProgress(progressRecord4);
                                    totalDownloaded = collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                                    WriteVerbose("Downloaded [" + totalDownloaded + "] Updates ready to Install");
                                    if (!Install)
                                    {
                                        WriteDebug(DateTime.Now + " Return downloaded update list");
                                        OutputObj.AddRange(collection.ToList());
                                        continue;
                                    }
                                }

                                if (Install)
                                {
                                    NeedsReboot = false;
                                    var num13 = 0;
                                    var progressRecord5 = new ProgressRecord(1, "Install updates for " + target, "[" + num13 + "/" + totalDownloaded + "]");
                                    foreach (var sendToPipeline in collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded"))
                                    {
                                        sendToPipeline.Properties.Add(new PSNoteProperty("X", 3));
                                        var update = (IUpdate)sendToPipeline.BaseObject;
                                        var updateCollection4 =
                                            (UpdateCollection)Activator.CreateInstance(
                                                Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                                        updateCollection4.Add(update);
                                        progressRecord5.StatusDescription = "[" + num13 + "/" + totalDownloaded + "] " + update.Title + " " +
                                                                            sendToPipeline.Properties["Size"].Value;
                                        progressRecord5.PercentComplete = num13 * 100 / totalDownloaded;
                                        WriteProgress(progressRecord5);
                                        ++num13;
                                        WriteDebug(DateTime.Now + " Show update to install: " + update.Title);
                                        var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                                        updateInstaller.Updates = updateCollection4;
                                        if (ForceInstall)
                                        {
                                            updateInstaller.IsForced = true;
                                        }

                                        IInstallationResult installationResult;
                                        try
                                        {
                                            installationResult = updateInstaller.Install();
                                            WriteDebug(DateTime.Now + " Installed");
                                        }
                                        catch (COMException ex)
                                        {
                                            var wuApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                            var skip = false;
                                            if (wuApiCodeDetails != null)
                                            {
                                                switch (wuApiCodeDetails.CodeType)
                                                {
                                                    case 2:
                                                        WriteError(new ErrorRecord(new Exception(wuApiCodeDetails.Description),
                                                            wuApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                                        skip = true;
                                                        break;
                                                    case 3:
                                                        WriteWarning(wuApiCodeDetails.HResult + ": " + wuApiCodeDetails.Description);
                                                        break;
                                                }

                                                if (skip)
                                                {
                                                    WriteDebug(DateTime.Now + " Skip to next computer");
                                                    break;
                                                }
                                            }
                                            else if (Debuger)
                                            {
                                                ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                                            }

                                            WriteDebug(DateTime.Now + " Skip to next update");
                                            continue;
                                        }

                                        if (!NeedsReboot)
                                        {
                                            NeedsReboot = installationResult.RebootRequired;
                                            WriteDebug(DateTime.Now + " Reboot is required");
                                        }

                                        var installMsg = "";
                                        switch (installationResult.ResultCode)
                                        {
                                            case OperationResultCode.orcNotStarted:
                                                installMsg = "NotStarted";
                                                break;
                                            case OperationResultCode.orcInProgress:
                                                installMsg = "InProgress";
                                                break;
                                            case OperationResultCode.orcSucceeded:
                                                installMsg = "Installed";
                                                break;
                                            case OperationResultCode.orcSucceededWithErrors:
                                                installMsg = "InstalledWithErrors";
                                                break;
                                            case OperationResultCode.orcFailed:
                                                installMsg = "Failed";
                                                break;
                                            case OperationResultCode.orcAborted:
                                                installMsg = "Aborted";
                                                break;
                                        }

                                        sendToPipeline.Properties.Add(new PSNoteProperty("InstallResult", installMsg));
                                        sendToPipeline.Properties.Add(new PSNoteProperty("Result", installMsg));
                                        var text31 = "";
                                        text31 = !(sendToPipeline.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text31 + "R" : text31 + "A";
                                        text31 = !(sendToPipeline.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text31 + "F" : text31 + "D";
                                        text31 = !(sendToPipeline.Properties["InstallResult"].Value.ToString() == "Installed")
                                            ? text31 + "F"
                                            : !installationResult.RebootRequired
                                                ? text31 + "I"
                                                : text31 + "R";
                                        text31 = !update.IsMandatory ? text31 + "-" : text31 + "M";
                                        text31 = !update.IsHidden ? text31 + "-" : text31 + "H";
                                        text31 = !update.IsUninstallable ? text31 + "-" : text31 + "U";
                                        text31 = !update.IsBeta ? text31 + "-" : text31 + "B";
                                        sendToPipeline.Properties["Status"].Value = text31;
                                        WriteObject(sendToPipeline, true);
                                    }

                                    progressRecord5.RecordType = ProgressRecordType.Completed;
                                    WriteProgress(progressRecord5);
                                    totalInstalled = collection.Where(x => x.Properties["Result"].Value.ToString() == "Installed").Count();
                                    WriteVerbose("Installed [" + totalInstalled + "] Updates");
                                    WriteDebug(DateTime.Now + " Return installed update list");
                                    OutputObj.AddRange(collection.ToList());
                                    if (totalInstalled > 0 && SendHistory)
                                    {
                                        var cmdLine2 = "Get-WUHistory -SendReport -Last " + totalInstalled +
                                                       " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                                        var invokeWuJob = new InvokeWUJob();
                                        invokeWuJob.ComputerName = [target];
                                        if (Credential != null)
                                        {
                                            invokeWuJob.Credential = Credential;
                                        }

                                        invokeWuJob.Script = cmdLine2;
                                        invokeWuJob.TaskName = "PSWindowsUpdate_History";
                                        if (NeedsReboot)
                                        {
                                            invokeWuJob.TriggerAtStart = true;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (AtStart): powershell.exe -Command \"" +
                                                         cmdLine2 + "\"");
                                        }
                                        else
                                        {
                                            invokeWuJob.RunNow = true;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (Now): powershell.exe -Command \"" +
                                                         cmdLine2 + "\"");
                                        }

                                        foreach (var sendToPipeline in invokeWuJob.Invoke())
                                        {
                                            WriteObject(sendToPipeline);
                                        }
                                    }

                                    if (RecurseCycle > 1)
                                    {
                                        --RecurseCycle;

                                        var cmdLine3 = !Debuger
                                            ? "Get-WindowsUpdate -AcceptAll"
                                            : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";

                                        if (IsInstalled)
                                        {
                                            cmdLine3 += " -IsInstalled";
                                        }

                                        if (IsHidden)
                                        {
                                            cmdLine3 += " -IsHidden";
                                        }

                                        if (WithHidden)
                                        {
                                            cmdLine3 += " -WithHidden";
                                        }

                                        if (UpdateType != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -UpdateType " + UpdateType;
                                        }

                                        if (DeploymentAction != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -DeploymentAction " + DeploymentAction;
                                        }

                                        if (UpdateID != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                                        }

                                        if (NotUpdateID != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -NotUpdateID '" + string.Join("','", NotUpdateID) + "'";
                                        }
                                        else if (text23 != "")
                                        {
                                            cmdLine3 = cmdLine3 + " -NotUpdateID " + text23;
                                        }

                                        if (RevisionNumber > 0)
                                        {
                                            cmdLine3 = cmdLine3 + " -RevisionNumber " + RevisionNumber;
                                        }

                                        if (CategoryIDs != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -CategoryIDs '" + string.Join("','", CategoryIDs) + "'";
                                        }

                                        if (IsAssigned)
                                        {
                                            cmdLine3 = !IsAssigned ? cmdLine3 + " -IsAssigned:$false" : cmdLine3 + " -IsAssigned";
                                        }

                                        if (IsPresent)
                                        {
                                            cmdLine3 = !IsPresent ? cmdLine3 + " -IsPresent:$false" : cmdLine3 + " -IsPresent";
                                        }

                                        if (AutoSelectOnWebSites)
                                        {
                                            cmdLine3 = !AutoSelectOnWebSites
                                                ? cmdLine3 + " -AutoSelectOnWebSites:$false"
                                                : cmdLine3 + " -AutoSelectOnWebSites";
                                        }

                                        if (RootCategories != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                                        }

                                        if (Category != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -Category '" + string.Join("','", Category) + "'";
                                        }

                                        if (KBArticleID != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                                        }

                                        if (Title != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -Title '" + Title + "'";
                                        }

                                        if (Severity != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -Severity '" + string.Join("','", Severity) + "'";
                                        }

                                        if (NotCategory != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                                        }

                                        if (NotKBArticleID != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                                        }

                                        if (NotTitle != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -NotTitle '" + NotTitle + "'";
                                        }

                                        if (NotSeverity != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                                        }

                                        if (IgnoreUserInput)
                                        {
                                            cmdLine3 += " -IgnoreUserInput";
                                        }

                                        if (IgnoreRebootRequired)
                                        {
                                            cmdLine3 += " -IgnoreRebootRequired";
                                        }

                                        if (AutoSelectOnly)
                                        {
                                            cmdLine3 += " -AutoSelectOnly";
                                        }

                                        if (MaxSize > 0)
                                        {
                                            cmdLine3 = cmdLine3 + " -MaxSize " + MaxSize;
                                        }

                                        if (MinSize > 0)
                                        {
                                            cmdLine3 = cmdLine3 + " -MinSize " + MinSize;
                                        }

                                        if (Download)
                                        {
                                            cmdLine3 += " -Download";
                                        }

                                        if (Install)
                                        {
                                            cmdLine3 += " -Install";
                                        }

                                        if (IgnoreReboot)
                                        {
                                            cmdLine3 += " -IgnoreReboot";
                                        }
                                        else if (AutoReboot)
                                        {
                                            cmdLine3 += " -AutoReboot";
                                        }
                                        else if (ScheduleReboot != DateTime.MinValue)
                                        {
                                            cmdLine3 = cmdLine3 + " -ScheduleReboot '" + ScheduleReboot + "'";
                                        }

                                        if (WindowsUpdate)
                                        {
                                            cmdLine3 += " -WindowsUpdate";
                                        }
                                        else if (MicrosoftUpdate)
                                        {
                                            cmdLine3 += " -MicrosoftUpdate";
                                        }
                                        else if (ServiceID != null)
                                        {
                                            cmdLine3 = cmdLine3 + " -ServiceID '" + ServiceID + "'";
                                        }

                                        if (SendReport)
                                        {
                                            cmdLine3 += " -SendReport";
                                        }

                                        if (SendHistory)
                                        {
                                            cmdLine3 += " -SendHistory";
                                        }

                                        if (RecurseCycle > 1)
                                        {
                                            cmdLine3 = cmdLine3 + " -RecurseCycle " + RecurseCycle;
                                        }

                                        cmdLine3 += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                                        var invokeWuJob = new InvokeWUJob();
                                        invokeWuJob.ComputerName = [target];
                                        if (Credential != null)
                                        {
                                            invokeWuJob.Credential = Credential;
                                        }

                                        invokeWuJob.Script = cmdLine3;
                                        invokeWuJob.TaskName = "PSWindowsUpdate_Recurse" + RecurseCycle;
                                        invokeWuJob.Debuger = true;
                                        if (NeedsReboot)
                                        {
                                            invokeWuJob.TriggerAtStart = true;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (AtStart): ");
                                        }
                                        else
                                        {
                                            ScheduleJob = DateTime.Now.AddMinutes(5.0);
                                            invokeWuJob.TriggerDate = ScheduleJob;
                                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob + "): ");
                                        }

                                        WriteVerbose("powershell.exe -Command \"" + cmdLine3 + "\"");
                                        foreach (var sendToPipeline in invokeWuJob.Invoke())
                                        {
                                            WriteObject(sendToPipeline);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (Debuger)
                    {
                        WriteError(new ErrorRecord(serviceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null));
                    }
                    else
                    {
                        WriteError(serviceManagerObj.Error);
                    }
                }
                else if (Debuger)
                {
                    WriteError(new ErrorRecord(updateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null));
                }
                else
                {
                    WriteError(updateSessionObj.Error);
                }
            }
        }

        /// <summary>Process</summary>
        protected override void ProcessRecord()
        {
            var flag = false;
            if (Credential != null)
            {
                var userName = Credential.GetNetworkCredential().UserName;
                var domain = Credential.GetNetworkCredential().Domain;
                var password = Credential.GetNetworkCredential().Password;
                WriteDebug(DateTime.Now + " UserName: " + userName + "; Domain: " + domain + "; Password: " + password.Substring(0, 1) +
                           "*****");
                var windowsPrincipal1 = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                var str1 = "";
                if (windowsPrincipal1.IsInRole(WindowsBuiltInRole.Administrator))
                {
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
                if (!WUToolsObj.IsLocalHost(ComputerName[0]))
                {
                    logonType = WUImpersonator.LogonSessionType.Interactive;
                    logonProvider = WUImpersonator.LogonProvider.Default;
                }

                using (new WUImpersonator(userName, domain, password, logonType, logonProvider))
                {
                    if (new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
                    {
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
                        try
                        {
                            CoreProcessing();
                            flag = false;
                        }
                        catch (Exception ex)
                        {
                            WriteDebug(DateTime.Now + " Something goes wrong: " + ex.Message);
                            flag = true;
                        }
                    }
                    else
                    {
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
                if (windowsPrincipal2.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    str4 = "RunAs";
                }

                WriteDebug(DateTime.Now + " After User: " + WindowsIdentity.GetCurrent().Name + " " + str4);
            }
            else
            {
                flag = true;
            }

            if (!flag)
            {
                return;
            }

            CoreProcessing();
        }

        /// <summary>End</summary>
        protected override void EndProcessing()
        {
            CmdletEnd = DateTime.Now;
            var CmdletInfo = new PSObject();
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletStart", CmdletStart));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletEnd", CmdletEnd));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletLine", MyInvocation.Line));
            if (SendReport && OutputObj.Count > 0)
            {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new string[5]
                    {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
                        "Title"
                    });
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            if (NeedsReboot)
            {
                if (ScheduleReboot != DateTime.MinValue)
                {
                    WriteDebug(DateTime.Now + " Schedule Reboot " + ScheduleReboot);
                    WriteVerbose(WUToolsObj.ScheduleReboot("localhost", ScheduleReboot));
                }
                else if (AutoReboot)
                {
                    WriteDebug(DateTime.Now + " Auto Reboot");
                    WriteVerbose(WUToolsObj.RunReboot("localhost"));
                }
                else if (IgnoreReboot)
                {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required, but do it manually.");
                }
                else
                {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor,
                        "Reboot is required. Do it now? [Y / N] (default is 'N')");
                    if (Console.ReadLine().ToUpper() == "Y")
                    {
                        WriteDebug(DateTime.Now + " Manually Reboot");
                        WriteVerbose(WUToolsObj.RunReboot("localhost"));
                    }
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}