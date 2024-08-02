using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using Microsoft.Management.Infrastructure;
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
    /// @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";Port=25} | Export-Clixml -Path 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate\PSWUSettings.xml'
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
    [Cmdlet("Get", "WindowsUpdate_v2", SupportsShouldProcess = true, DefaultParameterSetName = "Default",
        ConfirmImpact = ConfirmImpact.High)]
    [OutputType(typeof(WindowsUpdate))]
    public class GetWindowsUpdate_v2 : PSCmdlet
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

        private static Collection<PSObject> OutputObj { get; set; }

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
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                ThrowTerminatingError(new ErrorRecord(
                    new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new WindowsUpdate());
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

            ComputerName = new string[1]
            {
                Environment.MachineName
            };
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

                if (MyInvocation.BoundParameters.ContainsKey("IsAssigned"))
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

                if (MyInvocation.BoundParameters.ContainsKey("IsPresent"))
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

                if (MyInvocation.BoundParameters.ContainsKey("BrowseOnly"))
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

                if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites"))
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
                    var notUpdateID = NotUpdateID;
                    foreach (var text2 in notUpdateID)
                    {
                        criteria = criteria + " and UpdateID != '" + text2 + "'";
                        WriteDebug(DateTime.Now + " Set pre search criteria: NotUpdateID = " + text2);
                    }
                }

                if (UpdateID != null)
                {
                    var text3 = criteria;
                    criteria = "";
                    var num = 0;
                    var updateID = UpdateID;
                    foreach (var uID in updateID)
                    {
                        if (num > 0)
                        {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        if (RevisionNumber > 0)
                        {
                            criteria = criteria + "(" + text3 + " and UpdateID = '" + uID + "' and RevisionNumber = " + RevisionNumber +
                                       ")";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uID + "and RevisionNumber = " +
                                       RevisionNumber);
                        }
                        else
                        {
                            criteria = criteria + "(" + text3 + " and UpdateID = '" + uID + "')";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uID);
                        }

                        num++;
                    }
                }

                if (CategoryIDs != null)
                {
                    var text5 = criteria;
                    criteria = "";
                    var num2 = 0;
                    var categoryIDs = CategoryIDs;
                    foreach (var catIDs in categoryIDs)
                    {
                        if (num2 > 0)
                        {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        criteria = criteria + "(" + text5 + " and CategoryIDs contains '" + catIDs + "')";
                        WriteDebug(DateTime.Now + " Set pre search criteria: CategoryIDs = " + catIDs);
                        num2++;
                    }
                }
            }

            WriteDebug(DateTime.Now + " Search criteria is: " + criteria);
            if (ShowPreSearchCriteria)
            {
                Host.UI.WriteLine(ConsoleColor.Green, Host.UI.RawUI.BackgroundColor, "PreSearchCriteria: " + criteria);
            }

            var computerName = ComputerName;
            foreach (var target in computerName)
            {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try
                {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
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

                var wUCimSession = new WUCimSession(target);
                CimInstance[] cimInstances;
                try
                {
                    cimInstances = wUCimSession.ScanForUpdates(criteria);
                }
                catch (COMException exception)
                {
                    if (Debuger)
                    {
                        var errorRecord2 = new ErrorRecord(exception, "Debug", ErrorCategory.CloseError, null);
                        ThrowTerminatingError(errorRecord2);
                    }

                    WriteDebug(DateTime.Now + " Skip to next computer");
                    continue;
                }

                var cimInstancesLength = cimInstances.Length;
                WriteVerbose("Found [" + cimInstancesLength + "] Updates in pre search criteria");
                if (cimInstancesLength == 0)
                {
                    continue;
                }

                var num4 = 0;
                var activityId = 0;
                var activity = "Post search updates for " + target;
                var statusDescription = "[" + num4 + "/" + cimInstancesLength + "]";
                var progressRecord = new ProgressRecord(activityId, activity, statusDescription);
                var collection = new Collection<PSObject>();
                foreach (var cimInstance in cimInstances)
                {
                    WriteDebug(DateTime.Now + " " + cimInstance.CimInstanceProperties["Title"]);
                    progressRecord.StatusDescription =
                        "[" + num4 + "/" + cimInstancesLength + "] " + cimInstance.CimInstanceProperties["Title"];
                    progressRecord.PercentComplete = num4 * 100 / cimInstancesLength;
                    WriteProgress(progressRecord);
                    num4++;
                    var updateAccess = true;
                    if (KBArticleID != null && updateAccess)
                    {
                        WriteDebug(DateTime.Now + " Set post search criteria: KBArticleID = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), Title, RegexOptions.IgnoreCase))
                        {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (NotKBArticleID != null && updateAccess)
                    {
                        WriteDebug(DateTime.Now + " Set post search criteria: NotKBArticleID = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), NotTitle, RegexOptions.IgnoreCase))
                        {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (Title != null && updateAccess)
                    {
                        WriteDebug(DateTime.Now + " Set post search criteria: Title = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), Title, RegexOptions.IgnoreCase))
                        {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (NotTitle != null && updateAccess)
                    {
                        WriteDebug(DateTime.Now + " Set post search criteria: NotTitle = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), NotTitle, RegexOptions.IgnoreCase))
                        {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (updateAccess)
                    {
                        WriteDebug(DateTime.Now + " Update was not filtered");
                        var pSObject = new PSObject(cimInstance);
                        pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                        pSObject.Properties.Add(new PSNoteProperty("BaseObject", cimInstance));
                        pSObject.TypeNames.Clear();
                        pSObject.TypeNames.Add("PSWindowsUpdate.WindowsUpdate");
                        collection.Add(pSObject);
                    }
                }

                progressRecord.RecordType = ProgressRecordType.Completed;
                WriteProgress(progressRecord);
                var count = collection.Count;
                WriteVerbose("Found [" + count + "] Updates in post search criteria");
                if (!Download && !Install)
                {
                    WriteDebug(DateTime.Now + " Return update list only");
                    WriteObject(collection);
                    OutputObj = new Collection<PSObject>(collection);
                    continue;
                }

                var totalDownloaded = 0;
                var totalInstalled = 0;
                var text8 = "";
                if (Download || Install)
                {
                    var num7 = 0;
                    var activityId2 = 1;
                    var activity2 = "Choose updates for " + target;
                    var statusDescription2 = "[" + num7 + "/" + count + "]";
                    var progressRecord2 = new ProgressRecord(activityId2, activity2, statusDescription2);
                    var text9 = "";
                    foreach (var item in collection)
                    {
                        item.Properties.Add(new PSNoteProperty("X", 1));
                        item.TypeNames.Clear();
                        item.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                        var cimInstance2 = (CimInstance)item.BaseObject;
                        progressRecord2.StatusDescription = "[" + num7 + "/" + count + "] " + cimInstance2.CimInstanceProperties["Title"];
                        progressRecord2.PercentComplete = num7 * 100 / count;
                        WriteProgress(progressRecord2);
                        num7++;
                        WriteDebug(DateTime.Now + " Show update to accept: " + cimInstance2.CimInstanceProperties["Title"]);
                        var flag2 = false;
                        flag2 = AcceptAll || (ShouldProcess(target, "(" + DateTime.Now + ") " + cimInstance2.CimInstanceProperties["Title"])
                            ? true
                            : false);
                        var text10 = "";
                        var text11 = "";
                        if (flag2)
                        {
                            try
                            {
                                var pSObject2 = new PSObject(item.Properties["Identity"].Value);
                                var text12 = (string)pSObject2.Properties["UpdateID"].Value;
                                var num8 = (int)pSObject2.Properties["RevisionNumber"].Value;
                                text9 = !(text9 == "")
                                    ? text9 + " or (UpdateID = '" + text12 + "' and RevisionNumber = " + num8 + ")"
                                    : "(UpdateID = '" + text12 + "' and RevisionNumber = " + num8 + ")";
                            }
                            catch (Exception ex)
                            {
                                flag2 = false;
                                var errorRecord3 =
                                    new ErrorRecord(
                                        new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"] + "; " +
                                                      ex.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord3);
                            }

                            if (flag2)
                            {
                                text10 += "A";
                                text11 = "Accepted";
                                WriteDebug(DateTime.Now + " " + text11);
                                item.Properties.Add(new PSNoteProperty("ChooseResult", text11));
                                item.Properties.Add(new PSNoteProperty("Result", text11));
                            }
                        }

                        if (!flag2)
                        {
                            try
                            {
                                var pSObject3 = new PSObject(item.Properties["Identity"].Value);
                                var text13 = (string)pSObject3.Properties["UpdateID"].Value;
                                var num9 = (int)pSObject3.Properties["RevisionNumber"].Value;
                                text8 = !(text8 == "") ? text8 + ",'" + text13 + "'" : "'" + text13 + "'";
                            }
                            catch (Exception ex2)
                            {
                                flag2 = false;
                                var errorRecord4 =
                                    new ErrorRecord(
                                        new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"] + "; " +
                                                      ex2.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord4);
                            }

                            text10 += "R";
                            text11 = "Rejected";
                            WriteDebug(DateTime.Now + " " + text11);
                            item.Properties.Add(new PSNoteProperty("ChooseResult", text11));
                            item.Properties.Add(new PSNoteProperty("Result", text11));
                        }
                    }

                    progressRecord2.RecordType = ProgressRecordType.Completed;
                    WriteProgress(progressRecord2);
                    if (ShowPreSearchCriteria)
                    {
                        WriteVerbose("Choosed pre Search Criteria: " + text9);
                    }

                    var totalAccepted = collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                    WriteObject(collection, true);
                    WriteVerbose("Accepted [" + totalAccepted + "] Updates ready to Download");
                    var num11 = 0;
                    var activityId3 = 1;
                    var activity3 = "Download updates for " + target;
                    var statusDescription3 = "[" + num11 + "/" + totalAccepted + "]";
                    var progressRecord3 = new ProgressRecord(activityId3, activity3, statusDescription3);
                    foreach (var item2 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted"))
                    {
                        item2.Properties.Add(new PSNoteProperty("X", 2));
                        var update = (IUpdate)item2.BaseObject;
                        var updateCollection =
                            (UpdateCollection)Activator.CreateInstance(
                                Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        updateCollection.Add(update);
                        progressRecord3.StatusDescription = "[" + num11 + "/" + totalAccepted + "] " + update.Title + " " +
                                                            item2.Properties["Size"].Value;
                        progressRecord3.PercentComplete = num11 * 100 / totalAccepted;
                        WriteProgress(progressRecord3);
                        num11++;
                        WriteDebug(DateTime.Now + " Show update to download: " + update.Title);
                        IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                        updateDownloader.Updates = updateCollection;
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
                        catch (COMException ex3)
                        {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex3.ErrorCode);
                            var flag3 = false;
                            if (wUApiCodeDetails != null)
                            {
                                switch (wUApiCodeDetails.CodeType)
                                {
                                    case 2: // WUTools.CodeType.Error
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult,
                                            ErrorCategory.CloseError, null));
                                        flag3 = true;
                                        break;
                                    case 3:
                                        WriteWarning(wUApiCodeDetails.HResult + ": " + wUApiCodeDetails.Description);
                                        break;
                                }

                                if (flag3)
                                {
                                    WriteDebug(DateTime.Now + " Skip to next computer");
                                    break;
                                }
                            }
                            else if (Debuger)
                            {
                                // Debuger
                                var errorRecord5 = new ErrorRecord(ex3, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord5);
                            }

                            WriteDebug(DateTime.Now + " Skip to next update");
                            continue;
                        }

                        var value = "";
                        switch (downloadResult.ResultCode)
                        {
                            case OperationResultCode.orcNotStarted:
                                value = "NotStarted";
                                break;
                            case OperationResultCode.orcInProgress:
                                value = "InProgress";
                                break;
                            case OperationResultCode.orcSucceeded:
                                value = "Downloaded";
                                break;
                            case OperationResultCode.orcSucceededWithErrors:
                                value = "DownloadedWithErrors";
                                break;
                            case OperationResultCode.orcFailed:
                                value = "Failed";
                                break;
                            case OperationResultCode.orcAborted:
                                value = "Aborted";
                                break;
                        }

                        item2.Properties.Add(new PSNoteProperty("DownloadResult", value));
                        item2.Properties.Add(new PSNoteProperty("Result", value));
                        var text14 = "";
                        text14 = !(item2.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text14 + "R" : text14 + "A";
                        text14 = !(item2.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text14 + "F" : text14 + "D";
                        text14 = !update.IsInstalled ? text14 + "-" : text14 + "I";
                        text14 = !update.IsMandatory ? text14 + "-" : text14 + "M";
                        text14 = !update.IsHidden ? text14 + "-" : text14 + "H";
                        text14 = !update.IsUninstallable ? text14 + "-" : text14 + "U";
                        text14 = !update.IsBeta ? text14 + "-" : text14 + "B";
                        item2.Properties["Status"].Value = text14;
                        WriteObject(item2, true);
                    }

                    progressRecord3.RecordType = ProgressRecordType.Completed;
                    WriteProgress(progressRecord3);
                    totalDownloaded = collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                    WriteVerbose("Downloaded [" + totalDownloaded + "] Updates ready to Install");
                    if (!Install)
                    {
                        WriteDebug(DateTime.Now + " Return downloaded update list");
                        OutputObj = new Collection<PSObject>(collection.ToList());
                        continue;
                    }
                }

                if (!Install)
                {
                    continue;
                }

                NeedsReboot = false;
                var num12 = 0;
                var activityId4 = 1;
                var activity4 = "Install updates for " + target;
                var statusDescription4 = "[" + num12 + "/" + totalDownloaded + "]";
                var progressRecord4 = new ProgressRecord(activityId4, activity4, statusDescription4);
                foreach (var item3 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded"))
                {
                    item3.Properties.Add(new PSNoteProperty("X", 3));
                    var update2 = (IUpdate)item3.BaseObject;
                    var updateCollection2 =
                        (UpdateCollection)Activator.CreateInstance(
                            Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                    updateCollection2.Add(update2);
                    progressRecord4.StatusDescription = "[" + num12 + "/" + totalDownloaded + "] " + update2.Title + " " +
                                                        item3.Properties["Size"].Value;
                    progressRecord4.PercentComplete = num12 * 100 / totalDownloaded;
                    WriteProgress(progressRecord4);
                    num12++;
                    WriteDebug(DateTime.Now + " Show update to install: " + update2.Title);
                    var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                    updateInstaller.Updates = updateCollection2;
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
                    catch (COMException ex4)
                    {
                        var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex4.ErrorCode);
                        var skip = false;
                        if (wUApiCodeDetails2 != null)
                        {
                            switch (wUApiCodeDetails2.CodeType)
                            {
                                case 2:
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult,
                                        ErrorCategory.CloseError, null));
                                    skip = true;
                                    break;
                                case 3:
                                    WriteWarning(wUApiCodeDetails2.HResult + ": " + wUApiCodeDetails2.Description);
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
                            // Debuger
                            var errorRecord6 = new ErrorRecord(ex4, "Debug", ErrorCategory.CloseError, null);
                            ThrowTerminatingError(errorRecord6);
                        }

                        WriteDebug(DateTime.Now + " Skip to next update");
                        continue;
                    }

                    if (!NeedsReboot)
                    {
                        NeedsReboot = installationResult.RebootRequired;
                        WriteDebug(DateTime.Now + " Reboot is required");
                    }

                    var installResult = "";
                    switch (installationResult.ResultCode)
                    {
                        case OperationResultCode.orcNotStarted:
                            installResult = "NotStarted";
                            break;
                        case OperationResultCode.orcInProgress:
                            installResult = "InProgress";
                            break;
                        case OperationResultCode.orcSucceeded:
                            installResult = "Installed";
                            break;
                        case OperationResultCode.orcSucceededWithErrors:
                            installResult = "InstalledWithErrors";
                            break;
                        case OperationResultCode.orcFailed:
                            installResult = "Failed";
                            break;
                        case OperationResultCode.orcAborted:
                            installResult = "Aborted";
                            break;
                    }

                    item3.Properties.Add(new PSNoteProperty("InstallResult", installResult));
                    item3.Properties.Add(new PSNoteProperty("Result", installResult));
                    var text15 = "";
                    text15 = !(item3.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text15 + "R" : text15 + "A";
                    text15 = !(item3.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text15 + "F" : text15 + "D";
                    text15 = !(item3.Properties["InstallResult"].Value.ToString() == "Installed")
                        ? text15 + "F"
                        : !installationResult.RebootRequired
                            ? text15 + "I"
                            : text15 + "R";
                    text15 = !update2.IsMandatory ? text15 + "-" : text15 + "M";
                    text15 = !update2.IsHidden ? text15 + "-" : text15 + "H";
                    text15 = !update2.IsUninstallable ? text15 + "-" : text15 + "U";
                    text15 = !update2.IsBeta ? text15 + "-" : text15 + "B";
                    item3.Properties["Status"].Value = text15;
                    WriteObject(item3, true);
                }

                progressRecord4.RecordType = ProgressRecordType.Completed;
                WriteProgress(progressRecord4);
                totalInstalled = collection.Where(x => x.Properties["Result"].Value.ToString() == "Installed").Count();
                WriteVerbose("Installed [" + totalInstalled + "] Updates");
                WriteDebug(DateTime.Now + " Return installed update list");
                OutputObj = new Collection<PSObject>(collection.ToList());
                if (totalInstalled <= 0 || !SendHistory)
                {
                    continue;
                }

                var pCommand = "Get-WUHistory -SendReport -Last " + totalInstalled +
                               " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { target };
                if (Credential != null)
                {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = pCommand;
                invokeWUJob.TaskName = "PSWindowsUpdate_History";
                if (NeedsReboot)
                {
                    invokeWUJob.TriggerAtStart = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (AtStart): powershell.exe -Command \"" + pCommand +
                                 "\"");
                }
                else
                {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (Now): powershell.exe -Command \"" + pCommand +
                                 "\"");
                }

                var enumerable = invokeWUJob.Invoke();
                foreach (var item4 in enumerable)
                {
                    WriteObject(item4);
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

                    now = DateTime.Now;
                    WriteDebug(now + " Leaving impersonated session");
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
                    PSWUSettings.Add("Properties", new WindowsUpdate());
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