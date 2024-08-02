using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;
using Microsoft.Win32;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Set Windows Update Client settings.</para>
    /// <para type="description">Use Set-WUSettings cmdlet to Set configuration of Windows Update Client.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Enable IncludeRecommendedUpdates attributes.</para>
    /// 
    /// Set-WUSettings -IncludeRecommendedUpdates
    /// 
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Set Windows Update settings" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// <para></para>
    /// <para>ComputerName IncludeRecommendedUpdates</para>
    /// <para>------------ -------------------------</para>
    /// <para>MG-PC        True</para>
    /// </code>
    /// </example>
    [Cmdlet("Set", "WUSettings", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    [OutputType(typeof(WUSettings))]
    public class SetWUSettings : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
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
        /// <para type="description">Requires the parameter -PSWUSettings or declare the PSWUSettings.xml file in ModuleBase path.</para>
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
        /// <para type="description">Enabled - The WSUS server distributes available signed non-Microsoft updates.</para>
        /// <para type="description">Disabled - The WSUS server does not distribute available signed non-Microsoft updates.</para>
        /// <para type="description">Supported on: At least Windows Server 2003 operating systems or Windows XP Professional with SP1, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AcceptTrustedPublisherCerts { get; set; }

        /// <summary>
        /// <para type="description">Disables/Enables access to Windows Update.</para>
        /// <para type="description">Supported on: At least Windows Server 2003 operating systems, Windows XP Professional Service Pack 1, or Windows 2000 Service Pack 3</para>
        /// </summary>
        [Parameter]
        public SwitchParameter DisableWindowsUpdateAccess { get; set; }

        /// <summary>
        /// <para type="description">This policy setting allows you to control whether non-administrative users will receive update notifications based on the "Configure Automatic Updates" policy setting.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or At least Windows 2000 Service Pack 3 through Windows 8.1 or Windows Server 2012 R2 with most current service pack. Not supported on Windows 10 and above..</para>
        /// </summary>
        [Parameter]
        public SwitchParameter ElevateNonAdmins { get; set; }

        /// <summary>
        /// <para type="description">Name of the computer group to which the computer belongs.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public string TargetGroup { get; set; }

        /// <summary>
        /// <para type="description">Use/Do not use client-side targeting.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter TargetGroupEnabled { get; set; }

        /// <summary>
        /// <para type="description">HTTP(S) URL of the WSUS server that is used by Automatic Updates and API callers (by default).</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public string WUServer { get; set; }

        /// <summary>
        /// <para type="description">Notify mode: "Not configured", "Disabled", "Notify before download", "Notify before installation", "Scheduled installation", "Users configure"</para>
        /// <para type="description">Supported on: Windows XP Professional Service Pack 1 or At least Windows 2000 Service Pack 3 Option 7 only supported on servers of at least Windows Server 2016 edition</para>
        /// </summary>
        [ValidateSet("Not configured", "Disabled", "Notify before download", "Notify before installation", "Scheduled installation",
            "Users configure")]
        [Alias("AUOptions")]
        [Parameter]
        public string NotificationLevel { get; set; }

        /// <summary>
        /// <para type="description">Silently install minor updates.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or At least Windows 2000 Service Pack 3 through Windows 8.1 or Windows Server 2012 R2 with most current service pack. Not supported on Windows 10 and above.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AutoInstallMinorUpdates { get; set; }

        /// <summary>
        /// <para type="description">Time between detection cycles. Time in hours (1–22).</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [ValidateRange(1, 22)]
        [Parameter]
        public int DetectionFrequency { get; set; }

        /// <summary>
        /// <para type="description">Enable/Disable detection frequency.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter DetectionFrequencyEnabled { get; set; }

        /// <summary>
        /// <para type="description">Enable/Disable recommended updates.</para>
        /// <para type="description">Supported on: At least Windows Vista through Windows 8.1 or Windows Server 2012 R2 with most current service pack. Not supported on Windows 10 and above.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IncludeRecommendedUpdates { get; set; }

        /// <summary>
        /// <para type="description">Logged-on user can decide whether to restart the client computer.</para>
        /// <para type="description">Supported on: Windows XP Professional Service Pack 1 or At least Windows 2000 Service Pack 3</para>
        /// </summary>
        [Parameter]
        public SwitchParameter NoAutoRebootWithLoggedOnUsers { get; set; }

        /// <summary>
        /// <para type="description">Enable/Disable Automatic Updates.</para>
        /// <para type="description">Supported on: Windows XP Professional Service Pack 1 or At least Windows 2000 Service Pack 3 Option 7 only supported on servers of at least Windows Server 2016 edition</para>
        /// </summary>
        [Parameter]
        public SwitchParameter NoAutoUpdate { get; set; }

        /// <summary>
        /// <para type="description">Time between prompts for a scheduled restart. Time in minutes (1–1440).</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [ValidateRange(1, 1440)]
        [Parameter]
        public int RebootRelaunchTimeout { get; set; }

        /// <summary>
        /// <para type="description">Specifies the amount of time for Automatic Updates to wait before prompting again with a scheduled restart.</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [Parameter]
        public SwitchParameter RebootRelaunchTimeoutEnabled { get; set; }

        /// <summary>
        /// <para type="description">Length, in minutes, of the restart warning countdown after updates have been installed that have a deadline or scheduled updates. Time in minutes (1–30).</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [Parameter]
        public int RebootWarningTimeout { get; set; }

        /// <summary>
        /// <para type="description">Enable/Disable RebootWarningTimeout.</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [Parameter]
        public SwitchParameter RebootWarningTimeoutEnabled { get; set; }

        /// <summary>
        /// <para type="description">Scheduled day of install: "Every day", "Every Sunday", "Every Monday", "Every Tuesday", "Every Wednesday", "Every Thursday", ""Every Friday", "EverySaturday". Only valid if NotificationLevel (AUOptions) = "Scheduled installation"</para>
        /// <para type="description">Starting with Windows 8 and Windows Server 2012, ScheduledInstallationDay are not supported and will return unreliable values.If you try to modify these properties, the operation will appear to succeed but will have no effect.</para>
        /// </summary>
        [ValidateSet("Every Day", "Every Sunday", "Every Monday", "Every Tuesday", "Every Wednesday", "Every Thursday", "Every Friday",
            "EverySaturday")]
        [Parameter]
        public string ScheduledInstallDay { get; set; }

        /// <summary>
        /// <para type="description">Scheduled time of install in 24-hour format(0–23).</para>
        /// <para type="description">Starting with Windows 8 and Windows Server 2012, ScheduledInstallTime are not supported and will return unreliable values.If you try to modify these properties, the operation will appear to succeed but will have no effect.</para>
        /// /// </summary>
        [ValidateRange(0, 23)]
        [Parameter]
        public int ScheduledInstallTime { get; set; }

        /// <summary>
        /// <para type="description">The computer gets its updates from a WSUS server or from Microsoft Update.</para>
        /// <para type="description">Supported on: At least Windows XP Professional Service Pack 1 or Windows 2000 Service Pack 3, excluding Windows RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter UseWUServer { get; set; }

        /// <summary>
        /// <para type="description">Enable settings for control maximum version of Feature updates.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter TargetReleaseVersion { get; set; }

        /// <summary>
        /// <para type="description">TargetReleaseVersionInfo allow to control maximum version of Feature updates. For example "22H2".</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public string TargetReleaseVersionInfo { get; set; }

        /// <summary>
        /// <para type="description">ProductVersion allow to control maximum version of Feature updates. For example "Windows 11" or "11" or "Windows 10".</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public string ProductVersion { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify when to receive Feature Updates.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 365)]
        [Parameter]
        public int DeferFeatureUpdatesPeriodInDays { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify when to receive quality updates.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 30)]
        [Parameter]
        public int DeferQualityUpdatesPeriodInDays { get; set; }

        /// <summary>
        /// <para type="description">This policy setting specifies that a Windows Update for Business device should skip safeguards.</para>
        /// <para type="description">Supported on: At least Windows Server, Windows 10 Version 1903</para>
        /// </summary>
        [Parameter]
        public SwitchParameter DisableWUfBSafeguards { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to manage which updates you receive prior to the update being released to the world.</para>
        /// <para type="description">0	Disable Preview builds.</para>
        /// <para type="description">1	Disable Preview builds once the next release is public.</para>
        /// <para type="description">2	Enable Preview builds.</para>
        /// <para type="description">3 (Default) Preview builds is left to user selection.</para>
        /// <para type="description">Supported on: At least Windows Server, Windows 10 Version 1709</para>
        /// </summary>
        [ValidateRange(0, 3)]
        [Parameter]
        public int ManagePreviewBuilds { get; set; }

        /// <summary>
        /// <para type="description">Enabling this policy will automatically download updates, even over metered data connections (charges may apply).</para>
        /// <para type="description">Supported on: At least Windows 10 Server, Windows 10 or Windows 10 RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AllowAutoWindowsUpdateDownloadOverMeteredNetwork { get; set; }

        /// <summary>
        /// <para type="description">If you enable this policy, a restart timer will always begin immediately after Windows Update installs important updates, instead of first notifying users on the login screen for at least two days..</para>
        /// <para type="description">Supported on: At least Windows Server 2012, Windows 8 or Windows RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AlwaysAutoRebootAtScheduledTime { get; set; }

        /// <summary>
        /// <para type="description">The restart timer can be configured to start with any value from 15 to 180 minutes. When the timer runs out, the restart will proceed even if the PC has signed-in users.</para>
        /// <para type="description">Supported on: At least Windows Server 2012, Windows 8 or Windows RT</para>
        /// </summary>
        [ValidateRange(15, 180)]
        [Parameter]
        public int AlwaysAutoRebootAtScheduledTimeMinutes { get; set; }

        /// <summary>
        /// <para type="description">Specifies the type of notification used when notifying the user about a scheduled restart to install updates.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetAutoRestartNotificationConfig { get; set; }

        /// <summary>
        /// <para type="description">Specifies the type of notification used when notifying the user about a scheduled restart to install updates.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [ValidateSet("15", "30", "60", "120", "240")]
        [Parameter]
        public int AutoRestartNotificationSchedule { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify the method by which the auto-restart required notification is dismissed. When a restart is required to install updates, the auto-restart required notification is displayed. By default, the notification is automatically dismissed after 25 seconds.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetAutoRestartRequiredNotificationDismissal { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify the method by which the auto-restart required notification is dismissed. When a restart is required to install updates, the auto-restart required notification is displayed. By default, the notification is automatically dismissed after 25 seconds.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [ValidateSet("Auto", "User Action")]
        [Parameter]
        public string AutoRestartRequiredNotificationDismissal { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to control when notifications are displayed to warn users about a scheduled restart for the update installation deadline.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetRestartWarningSchd { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to control when notifications are displayed to warn users about a scheduled restart for the update installation deadline. 2 Hours, 4 Hours (Default), 8 Hours, 12 Hours, 24 Hours.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [ValidateSet("2", "4", "8", "12", "24")]
        [Parameter]
        public int ScheduleRestartWarning { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to control when notifications are displayed to warn users about a scheduled restart for the update installation deadline. 15 Minutes (Default), 30 Minutes or 60 Minutes.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [ValidateSet("15", "30", "60")]
        [Parameter]
        public int ScheduleImminentRestartWarning { get; set; }

        /// <summary>
        /// <para type="description">You can also choose to delay updates for up to one month. If you do not delay updates, your PC will remain up to date with security updates as they become available.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter DeferUpgrade { get; set; }

        /// <summary>
        /// <para type="description">Allows IT Admins to specify additional upgrade delays for up to 8 months. Supported values are 0-8, which refers to the number of months to defer upgrades.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 8)]
        [Parameter]
        public int DeferUpgradePeriod { get; set; }

        /// <summary>
        /// <para type="description">Allows IT Admins to specify update delays for up to four weeks. Supported values are 0-4, which refers to the number of weeks to defer updates.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 4)]
        [Parameter]
        public int DeferUpdatePeriod { get; set; }

        /// <summary>
        /// <para type="description">Allows IT Admins to pause updates and upgrades for up to 5 weeks.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter PauseDeferrals { get; set; }

        /// <summary>
        /// <para type="description">This policy allows you to define what Windows Update notifications users see. This policy doesn't control how and when updates are downloaded and installed.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetUpdateNotificationLevel { get; set; }

        /// <summary>
        /// <para type="description">This policy allows you to define what Windows Update notifications users see. This policy doesn't control how and when updates are downloaded and installed.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateSet("DefaultOS", "DisableAllExcludingRestart", "DisableAllIncludingRestart")]
        [Parameter]
        public string UpdateNotificationLevel { get; set; }

        /// <summary>
        /// <para type="description">This policy setting allows you to manage whether the 'Install Updates and Shut Down' option is displayed in the Shut Down Windows dialog box.</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows XP SP2</para>
        /// </summary>
        [Parameter]
        public SwitchParameter NoAUShutdownOption { get; set; }

        /// <summary>
        /// <para type="description">Specifies whether the Windows Update will use the Windows Power Management features to automatically wake up the system from sleep, if there are updates scheduled for installation.</para>
        /// <para type="description">Supported on: Windows Server 2008, Windows 7, Windows Vista, and Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AUPowerManagement { get; set; }

        /// <summary>
        /// <para type="description">This setting allows to remove access to "Pause updates" feature. Once enabled user access to pause updates is removed.</para>
        /// <para type="description">Supported on: At least Windows Server, Windows 10 Version 1809</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetDisablePauseUXAccess { get; set; }

        /// <summary>
        /// <para type="description">This setting allows you to remove access to scan Windows Update. If you enable this setting, Windows Update scan access is removed.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetDisableUXWUAccess { get; set; }

        /// <summary>
        /// <para type="description">Specifies the amount of time for Automatic Updates to wait, following system startup, before proceeding with a scheduled installation that was missed previously.</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [Parameter]
        public SwitchParameter RescheduleWaitTimeEnabled { get; set; }

        /// <summary>
        /// <para type="description">Specifies the amount of time for Automatic Updates to wait, following system startup, before proceeding with a scheduled installation that was missed previously.</para>
        /// <para type="description">Supported on: Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2003, Windows XP SP2, Windows XP SP1 , Windows 2000 SP4, Windows 2000 SP3</para>
        /// </summary>
        [ValidateRange(1, 60)]
        [Parameter]
        public int RescheduleWaitTime { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify the maximum number of hours from the start time that users can set their active hours. The max active hours range can be set between 8 and 18 hours.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetActiveHoursMaxRange { get; set; }

        /// <summary>
        /// <para type="description">Enable this policy to specify the maximum number of hours from the start time that users can set their active hours. The max active hours range can be set between 8 and 18 hours.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(8, 18)]
        [Parameter]
        public int ActiveHoursMaxRange { get; set; }

        /// <summary>
        /// <para type="description">When this policy is enabled, devices will receive Windows updates for the classes listed from the specified update source: either Windows Update or Windows Server Update Service.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter UseUpdateClassPolicySource { get; set; }

        /// <summary>
        /// <para type="description">Configure this policy to specify whether to receive Windows Feature Updates from Windows Update endpoint, managed by Windows Update for Business policies, or through your configured Windows Server Update Service (WSUS) server.</para>
        /// <para type="description">Supported on: At least Windows Server 2022, or Windows 10 Version 2004</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetPolicyDrivenUpdateSourceForFeatureUpdates { get; set; }

        /// <summary>
        /// <para type="description">Configure this policy to specify whether to receive Windows Quality Updates from Windows Update endpoint, managed by Windows Update for Business policies, or through your configured Windows Server Update Service (WSUS) server.</para>
        /// <para type="description">Supported on: At least Windows Server 2022, or Windows 10 Version 2004</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetPolicyDrivenUpdateSourceForQualityUpdates { get; set; }

        /// <summary>
        /// <para type="description">Configure this policy to specify whether to receive Windows Driver Updates from Windows Update endpoint, managed by Windows Update for Business policies, or through your configured Windows Server Update Service (WSUS) server.</para>
        /// <para type="description">Supported on: At least Windows Server 2022, or Windows 10 Version 2004</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetPolicyDrivenUpdateSourceForDriverUpdates { get; set; }

        /// <summary>
        /// <para type="description">Configure this policy to specify whether to receive Other Updates from Windows Update endpoint, managed by Windows Update for Business policies, or through your configured Windows Server Update Service (WSUS) server.</para>
        /// <para type="description">Supported on: At least Windows Server 2022, or Windows 10 Version 2004</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetPolicyDrivenUpdateSourceForOtherUpdates { get; set; }

        /// <summary>
        /// <para type="description">If you enable this policy, the PC will not automatically restart after updates during active hours. The PC will attempt to restart outside of active hours.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetActiveHours { get; set; }

        /// <summary>
        /// <para type="description">If you enable this policy, the PC will not automatically restart after updates during active hours. The PC will attempt to restart outside of active hours.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 23)]
        [Parameter]
        public int ActiveHoursStart { get; set; }

        /// <summary>
        /// <para type="description">If you enable this policy, the PC will not automatically restart after updates during active hours. The PC will attempt to restart outside of active hours.</para>
        /// <para type="description">Supported on: At least Windows 10 Server or Windows 10</para>
        /// </summary>
        [ValidateRange(0, 23)]
        [Parameter]
        public int ActiveHoursEnd { get; set; }

        /// <summary>
        /// <para type="description">This policy setting allows you to control whether users receive notifications for auto restarts for update installations including reminder and warning notifications.</para>
        /// <para type="description">Supported on: Windows Server 2016 through Windows Server 2022, or Windows 10</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetAutoRestartNotificationDisable { get; set; }

        /// <summary>
        /// <para type="description">This policy setting allows you to control whether users see detailed enhanced notification messages about featured software from the Microsoft Update service.</para>
        /// <para type="description">Supported on: Windows Server 2008, Windows 7, and Windows Vista</para>
        /// </summary>
        [Parameter]
        public SwitchParameter EnableFeaturedSoftware { get; set; }

        /// <summary>
        /// <para type="description">Enabling this policy for EDU devices that remain on Carts overnight will skip power checks to ensure update reboots will happen at the scheduled install time.</para>
        /// <para type="description">Supported on: At least Windows 10 Server, Windows 10 or Windows 10 RT</para>
        /// </summary>
        [Parameter]
        public SwitchParameter SetEDURestart { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

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
                    PSWUSettings.Add("Properties", new string[1]
                    {
                        "*"
                    });
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
            Hashtable hashtable1 = new Hashtable();
            hashtable1.Add("DefaultOS", 0);
            hashtable1.Add("DisableAllExcludingRestart", 1);
            hashtable1.Add("DisableAllIncludingRestart", 2);

            Hashtable hashtable2 = new Hashtable();
            hashtable2.Add("Auto", 1);
            hashtable2.Add("User Action", 2);

            Hashtable hashtable3 = new Hashtable();
            hashtable3.Add("Not configured", 0);
            hashtable3.Add("Disabled", 1);
            hashtable3.Add("Notify before download", 2);
            hashtable3.Add("Notify before installation", 3);
            hashtable3.Add("Scheduled installation", 4);
            hashtable3.Add("Users configure", 5);

            Hashtable hashtable4 = new Hashtable();
            hashtable4.Add("Every Day", 0);
            hashtable4.Add("Every Sunday", 1);
            hashtable4.Add("Every Monday", 2);
            hashtable4.Add("Every Tuesday", 3);
            hashtable4.Add("Every Wednesday", 4);
            hashtable4.Add("Every Thursday", 5);
            hashtable4.Add("Every Friday", 6);
            hashtable4.Add("Every Saturday", 7);

            /*var notifyTypes = new Hashtable {
                { 0, "Not configured" },
                { 1, "Disabled" },
                { 2, "Notify before download" },
                { 3, "Notify before installation" },
                { 4, "Scheduled installation" },
                { 5, "Users configure" }
            };

            var scheduleDay = new Hashtable {
                { 0, "Every Day" },
                { 1, "Every Sunday" },
                { 2, "Every Monday" },
                { 3, "Every Tuesday" },
                { 4, "Every Wednesday" },
                { 5, "Every Thursday" },
                { 6, "Every Friday" },
                { 7, "Every Saturday" }
            };*/

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

                if (ShouldProcess(target, "(" + DateTime.Now + ") Set Windows Update settings"))
                {
                    var registryKey1 = !WUToolsObj.IsLocalHost(target)
                        ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                        : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                    var registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                    var sendToPipeline = new PSObject();
                    sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                    if (registryKey2 != null)
                    {
                        WriteVerbose("Some settings are managed by your system administrator. Changes may not be applied.");
                    }

                    if (registryKey2 == null)
                    {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                        registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                        sendToPipeline.Properties.Add(new PSNoteProperty("RegistrySubKey1", "Created"));
                    }

                    if (registryKey3 == null)
                    {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\");
                        registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                        sendToPipeline.Properties.Add(new PSNoteProperty("RegistrySubKey2", "Created"));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AcceptTrustedPublisherCerts"))
                    {
                        bool flag1 = AcceptTrustedPublisherCerts.ToBool();
                        bool flag2 = false;
                        try
                        {
                            flag2 = Convert.ToBoolean(registryKey2.GetValue("AcceptTrustedPublisherCerts"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AcceptTrustedPublisherCerts", flag1));
                        if (flag1 != flag2)
                        {
                            try
                            {
                                WriteVerbose(target + " AcceptTrustedPublisherCerts: " + flag2 + " => " + flag1);
                                registryKey2.SetValue("AcceptTrustedPublisherCerts", Convert.ToInt32(flag1), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AcceptTrustedPublisherCerts",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                        {
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AcceptTrustedPublisherCerts' have been modified.");
                        }
                    }

                    if (DisableWindowsUpdateAccess)
                    {
                        bool flag3 = DisableWindowsUpdateAccess.ToBool();
                        bool flag4 = false;
                        try
                        {
                            flag4 = Convert.ToBoolean(registryKey2.GetValue("DisableWindowsUpdateAccess"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DisableWindowsUpdateAccess", flag3));
                        if (flag3 != flag4)
                        {
                            try
                            {
                                WriteVerbose(target + " DisableWindowsUpdateAccess: " + flag4 + " => " + flag3);
                                registryKey2.SetValue("DisableWindowsUpdateAccess", Convert.ToInt32(flag3), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DisableWindowsUpdateAccess",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                        {
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DisableWindowsUpdateAccess' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ElevateNonAdmins"))
                    {
                        bool flag5 = ElevateNonAdmins.ToBool();
                        bool flag6 = false;
                        try
                        {
                            flag6 = Convert.ToBoolean(registryKey2.GetValue("ElevateNonAdmins"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ElevateNonAdmins", flag5));
                        if (flag5 != flag6)
                        {
                            try
                            {
                                WriteVerbose(target + " ElevateNonAdmins: " + flag6 + " => " + flag5);
                                registryKey2.SetValue("ElevateNonAdmins", Convert.ToInt32(flag5), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ElevateNonAdmins",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ElevateNonAdmins' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetGroup"))
                    {
                        string targetGroup = TargetGroup;
                        string empty = string.Empty;
                        try
                        {
                            empty = (string)registryKey2.GetValue("\r\n");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetGroup", targetGroup));
                        if (targetGroup != empty)
                        {
                            try
                            {
                                WriteVerbose(target + " TargetGroup: " + empty + " => " + targetGroup);
                                registryKey2.SetValue("TargetGroup", targetGroup, RegistryValueKind.String);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetGroup",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(
                                target + " The command completed successfully but no settings of 'TargetGroup' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetGroupEnabled"))
                    {
                        bool flag7 = TargetGroupEnabled.ToBool();
                        bool flag8 = false;
                        try
                        {
                            flag8 = Convert.ToBoolean(registryKey2.GetValue("TargetGroupEnabled"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetGroupEnabled", flag7));
                        if (flag7 != flag8)
                        {
                            try
                            {
                                WriteVerbose(target + " TargetGroupEnabled: " + flag8 + " => " + flag7);
                                registryKey2.SetValue("TargetGroupEnabled", Convert.ToInt32(flag7), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetGroupEnabled",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'TargetGroupEnabled' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("WUServer"))
                    {
                        string wuServer = WUServer;
                        string empty = string.Empty;
                        try
                        {
                            empty = (string)registryKey2.GetValue("WUServer");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("WUServer", wuServer));
                        if (wuServer != empty)
                        {
                            try
                            {
                                WriteVerbose(target + " WUServer: " + empty + " => " + wuServer);
                                registryKey2.SetValue("WUServer", wuServer, RegistryValueKind.String);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "WUServer", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target + " The command completed successfully but no settings of 'WUServer' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NoAutoUpdate"))
                    {
                        bool flag9 = TargetGroupEnabled.ToBool();
                        bool flag10 = false;
                        try
                        {
                            flag10 = Convert.ToBoolean(registryKey3.GetValue("NoAutoUpdate"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("NoAutoUpdate", flag9));
                        if (flag9 != flag10)
                        {
                            try
                            {
                                WriteVerbose(target + " NoAutoUpdateNoAutoUpdate: " + flag10 + " => " + flag9);
                                registryKey3.SetValue("TargetGroupEnabled", Convert.ToInt32(flag9), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "NoAutoUpdate",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'NoAutoUpdate' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NotificationLevel"))
                    {
                        int num1 = (int)hashtable3[NotificationLevel];
                        int num2 = 0;
                        try
                        {
                            num2 = (int)registryKey3.GetValue("AUOptions");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AUOptions", num1));
                        if (num1 != num2)
                        {
                            try
                            {
                                WriteVerbose(target + " AUOptions: " + num2 + " => " + num1);
                                registryKey3.SetValue("AUOptions", num1, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AUOptions",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target + " The command completed successfully but no settings of 'AUOptions' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AutoInstallMinorUpdates"))
                    {
                        bool flag11 = AutoInstallMinorUpdates.ToBool();
                        bool flag12 = false;
                        try
                        {
                            flag12 = Convert.ToBoolean(registryKey3.GetValue("AutoInstallMinorUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AutoInstallMinorUpdates", flag11));
                        if (flag11 != flag12)
                        {
                            try
                            {
                                WriteVerbose(target + " AutoInstallMinorUpdates: " + flag12 + " => " + flag11);
                                registryKey3.SetValue("AutoInstallMinorUpdates", Convert.ToInt32(flag11), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AutoInstallMinorUpdates",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AutoInstallMinorUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DetectionFrequency"))
                    {
                        int detectionFrequency = DetectionFrequency;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("DetectionFrequency");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DetectionFrequency", detectionFrequency));
                        if (detectionFrequency != num)
                        {
                            try
                            {
                                WriteVerbose(target + " DetectionFrequency: " + num + " => " + detectionFrequency);
                                registryKey3.SetValue("DetectionFrequency", detectionFrequency, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DetectionFrequency",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DetectionFrequency' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DetectionFrequencyEnabled"))
                    {
                        bool flag13 = DetectionFrequencyEnabled.ToBool();
                        bool flag14 = false;
                        try
                        {
                            flag14 = Convert.ToBoolean(registryKey3.GetValue("DetectionFrequencyEnabled"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DetectionFrequencyEnabled", flag13));
                        if (flag13 != flag14)
                        {
                            try
                            {
                                WriteVerbose(target + " DetectionFrequencyEnabled: " + flag14 + " => " + flag13);
                                registryKey3.SetValue("DetectionFrequencyEnabled", Convert.ToInt32(flag13), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DetectionFrequencyEnabled",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DetectionFrequencyEnabled' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("IncludeRecommendedUpdates"))
                    {
                        bool flag15 = IncludeRecommendedUpdates.ToBool();
                        bool flag16 = false;
                        try
                        {
                            flag16 = Convert.ToBoolean(registryKey3.GetValue("IncludeRecommendedUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("IncludeRecommendedUpdates", flag15));
                        if (flag15 != flag16)
                        {
                            try
                            {
                                WriteVerbose(target + " IncludeRecommendedUpdates: " + flag16 + " => " + flag15);
                                registryKey3.SetValue("IncludeRecommendedUpdates", Convert.ToInt32(flag15), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "IncludeRecommendedUpdates",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'IncludeRecommendedUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NoAutoRebootWithLoggedOnUsers"))
                    {
                        bool flag17 = NoAutoRebootWithLoggedOnUsers.ToBool();
                        bool flag18 = false;
                        try
                        {
                            flag18 = Convert.ToBoolean(registryKey3.GetValue("NoAutoRebootWithLoggedOnUsers"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("NoAutoRebootWithLoggedOnUsers", flag17));
                        if (flag17 != flag18)
                        {
                            try
                            {
                                WriteVerbose(target + " NoAutoRebootWithLoggedOnUsers: " + flag18 + " => " + flag17);
                                registryKey3.SetValue("NoAutoRebootWithLoggedOnUsers", Convert.ToInt32(flag17), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(
                                    new ErrorRecord(new Exception(target + ": Access denied."), "NoAutoRebootWithLoggedOnUsers",
                                        ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'NoAutoRebootWithLoggedOnUsers' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootRelaunchTimeout"))
                    {
                        int rebootRelaunchTimeout = RebootRelaunchTimeout;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("RebootRelaunchTimeout");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootRelaunchTimeout", rebootRelaunchTimeout));
                        if (rebootRelaunchTimeout != num)
                        {
                            try
                            {
                                WriteVerbose(target + " RebootRelaunchTimeout: " + num + " => " + rebootRelaunchTimeout);
                                registryKey3.SetValue("RebootRelaunchTimeout", rebootRelaunchTimeout, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootRelaunchTimeout",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RebootRelaunchTimeout' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootRelaunchTimeoutEnabled"))
                    {
                        bool flag19 = RebootRelaunchTimeoutEnabled.ToBool();
                        bool flag20 = false;
                        try
                        {
                            flag20 = Convert.ToBoolean(registryKey3.GetValue("RebootRelaunchTimeoutEnabled"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootRelaunchTimeoutEnabled", flag19));
                        if (flag19 != flag20)
                        {
                            try
                            {
                                WriteVerbose(target + " RebootRelaunchTimeoutEnabled: " + flag20 + " => " + flag19);
                                registryKey3.SetValue("RebootRelaunchTimeoutEnabled", Convert.ToInt32(flag19), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootRelaunchTimeoutEnabled",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RebootRelaunchTimeoutEnabled' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootWarningTimeout"))
                    {
                        int rebootWarningTimeout = RebootWarningTimeout;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("RebootWarningTimeout");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootWarningTimeout", rebootWarningTimeout));
                        if (rebootWarningTimeout != num)
                        {
                            try
                            {
                                WriteVerbose(target + " RebootWarningTimeout: " + num + " => " + rebootWarningTimeout);
                                registryKey3.SetValue("RebootWarningTimeout", rebootWarningTimeout, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootWarningTimeout",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RebootWarningTimeout' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootWarningTimeoutEnabled"))
                    {
                        bool flag21 = RebootWarningTimeoutEnabled.ToBool();
                        bool flag22 = false;
                        try
                        {
                            flag22 = Convert.ToBoolean(registryKey3.GetValue("RebootWarningTimeoutEnabled"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootWarningTimeoutEnabled", flag21));
                        if (flag21 != flag22)
                        {
                            try
                            {
                                WriteVerbose(target + " RebootWarningTimeoutEnabled: " + flag22 + " => " + flag21);
                                registryKey3.SetValue("RebootWarningTimeoutEnabled", Convert.ToInt32(flag21), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootWarningTimeoutEnabled",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RebootWarningTimeoutEnabled' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ScheduledInstallDay"))
                    {
                        int num3 = (int)hashtable4[ScheduledInstallDay];
                        int num4 = 0;
                        try
                        {
                            num4 = (int)registryKey3.GetValue("ScheduledInstallDay");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduledInstallDay", num3));
                        if (num3 != num4)
                        {
                            try
                            {
                                WriteVerbose(target + " ScheduledInstallDay: " + num4 + " => " + num3);
                                registryKey3.SetValue("ScheduledInstallDay", num3, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduledInstallDay",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ScheduledInstallDay' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ScheduledInstallTime"))
                    {
                        int scheduledInstallTime = ScheduledInstallTime;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("ScheduledInstallTime");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduledInstallTime", scheduledInstallTime));
                        if (scheduledInstallTime != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ScheduledInstallTime: " + num + " => " + scheduledInstallTime);
                                registryKey3.SetValue("ScheduledInstallTime", scheduledInstallTime, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduledInstallTime",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ScheduledInstallTime' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("UseWUServer"))
                    {
                        bool flag23 = UseWUServer.ToBool();
                        bool flag24 = false;
                        try
                        {
                            flag24 = Convert.ToBoolean(registryKey3.GetValue("UseWUServer"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("UseWUServer", flag23));
                        if (flag23 != flag24)
                        {
                            try
                            {
                                WriteVerbose(target + " UseWUServer: " + flag24 + " => " + flag23);
                                registryKey3.SetValue("UseWUServer", Convert.ToInt32(flag23), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "UseWUServer",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(
                                target + " The command completed successfully but no settings of 'UseWUServer' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetReleaseVersion"))
                    {
                        bool flag25 = TargetReleaseVersion.ToBool();
                        bool flag26 = false;
                        try
                        {
                            flag26 = Convert.ToBoolean(registryKey2.GetValue("TargetReleaseVersion"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetReleaseVersion", flag25));
                        if (flag25 != flag26)
                        {
                            try
                            {
                                WriteVerbose(target + " TargetReleaseVersion: " + flag26 + " => " + flag25);
                                registryKey2.SetValue("TargetReleaseVersion", Convert.ToInt32(flag25), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetReleaseVersion",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'TargetReleaseVersion' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetReleaseVersionInfo"))
                    {
                        string releaseVersionInfo = TargetReleaseVersionInfo;
                        string empty = string.Empty;
                        try
                        {
                            empty = (string)registryKey2.GetValue("TargetReleaseVersionInfo");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetReleaseVersionInfo", releaseVersionInfo));
                        if (releaseVersionInfo != empty)
                        {
                            try
                            {
                                WriteVerbose(target + " TargetRelTargetReleaseVersionInfoeaseVersion: " + empty + " => " +
                                             releaseVersionInfo);
                                registryKey2.SetValue("TargetReleaseVersionInfo", releaseVersionInfo, RegistryValueKind.String);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetReleaseVersionInfo",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'TargetReleaseVersionInfo' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ProductVersion"))
                    {
                        string productVersion = ProductVersion;
                        string empty = string.Empty;
                        try
                        {
                            empty = (string)registryKey2.GetValue("ProductVersion");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ProductVersion", productVersion));
                        if (productVersion != empty)
                        {
                            try
                            {
                                WriteVerbose(target + " ProductVersion: " + empty + " => " + productVersion);
                                registryKey2.SetValue("ProductVersion", productVersion, RegistryValueKind.String);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ProductVersion",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ProductVersion' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DisableWUfBSafeguards"))
                    {
                        bool flag27 = DisableWUfBSafeguards.ToBool();
                        bool flag28 = false;
                        try
                        {
                            flag28 = Convert.ToBoolean(registryKey2.GetValue("DisableWUfBSafeguards"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DisableWUfBSafeguards", flag27));
                        if (flag27 != flag28)
                        {
                            try
                            {
                                WriteVerbose(target + " DisableWUfBSafeguards: " + flag28 + " => " + flag27);
                                registryKey2.SetValue("DisableWUfBSafeguards", Convert.ToInt32(flag27), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DisableWUfBSafeguards",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DisableWUfBSafeguards' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ManagePreviewBuilds"))
                    {
                        int managePreviewBuilds = ManagePreviewBuilds;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ManagePreviewBuildsPolicyValue");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ManagePreviewBuilds", 1));
                        sendToPipeline.Properties.Add(new PSNoteProperty("ManagePreviewBuildsPolicyValue", managePreviewBuilds));
                        if (managePreviewBuilds != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ManagePreviewBuildsPolicyValue: " + num + " => " + managePreviewBuilds);
                                registryKey2.SetValue("ManagePreviewBuilds", 1, RegistryValueKind.DWord);
                                registryKey2.SetValue("ManagePreviewBuildsPolicyValue", managePreviewBuilds, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ManagePreviewBuilds",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ManagePreviewBuilds' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DeferFeatureUpdatesPeriodInDays"))
                    {
                        int updatesPeriodInDays = DeferFeatureUpdatesPeriodInDays;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("DeferFeatureUpdatesPeriodInDays");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferFeatureUpdates", 1));
                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferFeatureUpdatesPeriodInDays", updatesPeriodInDays));
                        if (updatesPeriodInDays != num)
                        {
                            try
                            {
                                WriteVerbose(target + " DeferFeatureUpdatesPeriodInDays: " + num + " => " + updatesPeriodInDays);
                                registryKey2.SetValue("DeferFeatureUpdates", 1, RegistryValueKind.DWord);
                                registryKey2.SetValue("DeferFeatureUpdatesPeriodInDays", updatesPeriodInDays, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DeferFeatureUpdates",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DeferFeatureUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DeferQualityUpdatesPeriodInDays"))
                    {
                        int updatesPeriodInDays = DeferQualityUpdatesPeriodInDays;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("DeferQualityUpdatesPeriodInDays");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferQualityUpdates", 1));
                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferQualityUpdatesPeriodInDays", updatesPeriodInDays));
                        if (updatesPeriodInDays != num)
                        {
                            try
                            {
                                WriteVerbose(target + " DeferQualityUpdatesPeriodInDays: " + num + " => " + updatesPeriodInDays);
                                registryKey2.SetValue("DeferQualityUpdates", 1, RegistryValueKind.DWord);
                                registryKey2.SetValue("DeferQualityUpdatesPeriodInDays", updatesPeriodInDays, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DeferQualityUpdates",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DeferQualityUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AllowAutoWindowsUpdateDownloadOverMeteredNetwork"))
                    {
                        bool flag29 = AllowAutoWindowsUpdateDownloadOverMeteredNetwork.ToBool();
                        bool flag30 = false;
                        try
                        {
                            flag30 = Convert.ToBoolean(registryKey2.GetValue("AllowAutoWindowsUpdateDownloadOverMeteredNetwork"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AllowAutoWindowsUpdateDownloadOverMeteredNetwork", flag29));
                        if (flag29 != flag30)
                        {
                            try
                            {
                                WriteVerbose(target + " AllowAutoWindowsUpdateDownloadOverMeteredNetwork: " + flag30 + " => " + flag29);
                                registryKey2.SetValue("AllowAutoWindowsUpdateDownloadOverMeteredNetwork", Convert.ToInt32(flag29),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "AllowAutoWindowsUpdateDownloadOverMeteredNetwork",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AllowAutoWindowsUpdateDownloadOverMeteredNetwork' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AlwaysAutoRebootAtScheduledTime"))
                    {
                        bool flag31 = AlwaysAutoRebootAtScheduledTime.ToBool();
                        bool flag32 = false;
                        try
                        {
                            flag32 = Convert.ToBoolean(registryKey3.GetValue("AlwaysAutoRebootAtScheduledTime"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AlwaysAutoRebootAtScheduledTime", flag31));
                        if (flag31 != flag32)
                        {
                            try
                            {
                                WriteVerbose(target + " AlwaysAutoRebootAtScheduledTime: " + flag32 + " => " + flag31);
                                registryKey3.SetValue("AlwaysAutoRebootAtScheduledTime", Convert.ToInt32(flag31), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AlwaysAutoRebootAtScheduledTime",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AlwaysAutoRebootAtScheduledTime' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AlwaysAutoRebootAtScheduledTimeMinutes"))
                    {
                        int scheduledTimeMinutes = AlwaysAutoRebootAtScheduledTimeMinutes;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("AlwaysAutoRebootAtScheduledTimeMinutes");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AlwaysAutoRebootAtScheduledTimeMinutes", scheduledTimeMinutes));
                        if (scheduledTimeMinutes != num)
                        {
                            try
                            {
                                WriteVerbose(target + " AlwaysAutoRebootAtScheduledTimeMinutes: " + num + " => " + scheduledTimeMinutes);
                                registryKey3.SetValue("AlwaysAutoRebootAtScheduledTimeMinutes", scheduledTimeMinutes,
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "AlwaysAutoRebootAtScheduledTimeMinutes", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AlwaysAutoRebootAtScheduledTimeMinutes' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetAutoRestartNotificationConfig"))
                    {
                        bool flag33 = SetAutoRestartNotificationConfig.ToBool();
                        bool flag34 = false;
                        try
                        {
                            flag34 = Convert.ToBoolean(registryKey2.GetValue("SetAutoRestartNotificationConfig"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetAutoRestartNotificationConfig", flag33));
                        if (flag33 != flag34)
                        {
                            try
                            {
                                WriteVerbose(target + " SetAutoRestartNotificationConfig: " + flag34 + " => " + flag33);
                                registryKey2.SetValue("SetAutoRestartNotificationConfig", Convert.ToInt32(flag33), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(
                                    new ErrorRecord(new Exception(target + ": Access denied."), "SetAutoRestartNotificationConfig",
                                        ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetAutoRestartNotificationConfig' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AutoRestartNotificationSchedule"))
                    {
                        int notificationSchedule = AutoRestartNotificationSchedule;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("AutoRestartNotificationSchedule");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AutoRestartNotificationSchedule", notificationSchedule));
                        if (notificationSchedule != num)
                        {
                            try
                            {
                                WriteVerbose(target + " AutoRestartNotificationSchedule: " + num + " => " + notificationSchedule);
                                registryKey3.SetValue("AutoRestartNotificationSchedule", notificationSchedule, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AutoRestartNotificationSchedule",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AutoRestartNotificationSchedule' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetAutoRestartRequiredNotificationDismissal"))
                    {
                        bool flag35 = SetAutoRestartRequiredNotificationDismissal.ToBool();
                        bool flag36 = false;
                        try
                        {
                            flag36 = Convert.ToBoolean(registryKey2.GetValue("SetAutoRestartRequiredNotificationDismissal"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetAutoRestartRequiredNotificationDismissal", flag35));
                        if (flag35 != flag36)
                        {
                            try
                            {
                                WriteVerbose(target + " SetAutoRestartRequiredNotificationDismissal: " + flag36 + " => " + flag35);
                                registryKey2.SetValue("SetAutoRestartRequiredNotificationDismissal", Convert.ToInt32(flag35),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "SetAutoRestartRequiredNotificationDismissal", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetAutoRestartRequiredNotificationDismissal' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AutoRestartRequiredNotificationDismissal"))
                    {
                        int num5 = (int)hashtable2[AutoRestartRequiredNotificationDismissal];
                        int num6 = 0;
                        try
                        {
                            num6 = (int)registryKey2.GetValue("AutoRestartRequiredNotificationDismissal");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AutoRestartRequiredNotificationDismissal", num5));
                        if (num5 != num6)
                        {
                            try
                            {
                                WriteVerbose(target + " AutoRestartRequiredNotificationDismissal: " + num6 + " => " + num5);
                                registryKey2.SetValue("AutoRestartRequiredNotificationDismissal", num5, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "AutoRestartRequiredNotificationDismissal", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AutoRestartRequiredNotificationDismissal' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetRestartWarningSchd"))
                    {
                        bool flag37 = SetRestartWarningSchd.ToBool();
                        bool flag38 = false;
                        try
                        {
                            flag38 = Convert.ToBoolean(registryKey2.GetValue("SetRestartWarningSchd"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetRestartWarningSchd", flag37));
                        if (flag37 != flag38)
                        {
                            try
                            {
                                WriteVerbose(target + " SetRestartWarningSchd: " + flag38 + " => " + flag37);
                                registryKey2.SetValue("SetRestartWarningSchd", Convert.ToInt32(flag37), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetRestartWarningSchd",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetRestartWarningSchd' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ScheduleRestartWarning"))
                    {
                        int scheduleRestartWarning = ScheduleRestartWarning;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ScheduleRestartWarning");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduleRestartWarning", scheduleRestartWarning));
                        if (scheduleRestartWarning != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ScheduleRestartWarning: " + num + " => " + scheduleRestartWarning);
                                registryKey2.SetValue("ScheduleRestartWarning", scheduleRestartWarning, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduleRestartWarning",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ScheduleRestartWarning' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ScheduleImminentRestartWarning"))
                    {
                        int imminentRestartWarning = ScheduleImminentRestartWarning;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ScheduleImminentRestartWarning");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduleImminentRestartWarning", imminentRestartWarning));
                        if (imminentRestartWarning != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ScheduleImminentRestartWarning: " + num + " => " + imminentRestartWarning);
                                registryKey2.SetValue("ScheduleImminentRestartWarning", imminentRestartWarning, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduleImminentRestartWarning",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ScheduleImminentRestartWarning' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DeferUpgrade"))
                    {
                        bool flag39 = DeferUpgrade.ToBool();
                        bool flag40 = false;
                        try
                        {
                            flag40 = Convert.ToBoolean(registryKey2.GetValue("DeferUpgrade"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferUpgrade", flag39));
                        if (flag39 != flag40)
                        {
                            try
                            {
                                WriteVerbose(target + " DeferUpgrade: " + flag40 + " => " + flag39);
                                registryKey2.SetValue("DeferUpgrade", Convert.ToInt32(flag39), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DeferUpgrade",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DeferUpgrade' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DeferUpgradePeriod"))
                    {
                        int deferUpgradePeriod = DeferUpgradePeriod;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("DeferUpgradePeriod");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferUpgradePeriod", deferUpgradePeriod));
                        if (deferUpgradePeriod != num)
                        {
                            try
                            {
                                WriteVerbose(target + " DeferUpgradePeriod: " + num + " => " + deferUpgradePeriod);
                                registryKey2.SetValue("DeferUpgradePeriod", deferUpgradePeriod, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DeferUpgradePeriod",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DeferUpgradePeriod' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DeferUpdatePeriod"))
                    {
                        int deferUpdatePeriod = DeferUpdatePeriod;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("DeferUpdatePeriod");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("DeferUpdatePeriod", deferUpdatePeriod));
                        if (deferUpdatePeriod != num)
                        {
                            try
                            {
                                WriteVerbose(target + " DeferUpdatePeriod: " + num + " => " + deferUpdatePeriod);
                                registryKey2.SetValue("DeferUpdatePeriod", deferUpdatePeriod, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DeferUpdatePeriod",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'DeferUpdatePeriod' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("PauseDeferrals"))
                    {
                        bool flag41 = PauseDeferrals.ToBool();
                        bool flag42 = false;
                        try
                        {
                            flag42 = Convert.ToBoolean(registryKey2.GetValue("PauseDeferrals"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("PauseDeferrals", flag41));
                        if (flag41 != flag42)
                        {
                            try
                            {
                                WriteVerbose(target + " PauseDeferrals: " + flag42 + " => " + flag41);
                                registryKey2.SetValue("PauseDeferrals", Convert.ToInt32(flag41), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "PauseDeferrals",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'PauseDeferrals' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetUpdateNotificationLevel"))
                    {
                        bool flag43 = SetUpdateNotificationLevel.ToBool();
                        bool flag44 = false;
                        try
                        {
                            flag44 = Convert.ToBoolean(registryKey2.GetValue("SetUpdateNotificationLevel"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetUpdateNotificationLevel", flag43));
                        if (flag43 != flag44)
                        {
                            try
                            {
                                WriteVerbose(target + " SetUpdateNotificationLevel: " + flag44 + " => " + flag43);
                                registryKey2.SetValue("SetUpdateNotificationLevel", Convert.ToInt32(flag43), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetUpdateNotificationLevel",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetUpdateNotificationLevel' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("UpdateNotificationLevel"))
                    {
                        int num7 = (int)hashtable1[UpdateNotificationLevel];
                        int num8 = 0;
                        try
                        {
                            num8 = (int)registryKey2.GetValue("UpdateNotificationLevel");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("UpdateNotificationLevel", num7));
                        if (num7 != num8)
                        {
                            try
                            {
                                WriteVerbose(target + " UpdateNotificationLevel: " + num8 + " => " + num7);
                                registryKey2.SetValue("UpdateNotificationLevel", num7, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "UpdateNotificationLevel",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'UpdateNotificationLevel' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NoAUShutdownOption"))
                    {
                        bool flag45 = NoAUShutdownOption.ToBool();
                        bool flag46 = false;
                        try
                        {
                            flag46 = Convert.ToBoolean(registryKey2.GetValue("NoAUShutdownOption"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("NoAUShutdownOption", flag45));
                        if (flag45 != flag46)
                        {
                            try
                            {
                                WriteVerbose(target + " NoAUShutdownOption: " + flag46 + " => " + flag45);
                                registryKey2.SetValue("NoAUShutdownOption", Convert.ToInt32(flag45), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "NoAUShutdownOption",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'NoAUShutdownOption' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AUPowerManagement"))
                    {
                        bool flag47 = AUPowerManagement.ToBool();
                        bool flag48 = false;
                        try
                        {
                            flag48 = Convert.ToBoolean(registryKey2.GetValue("AUPowerManagement"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("AUPowerManagement", flag47));
                        if (flag47 != flag48)
                        {
                            try
                            {
                                WriteVerbose(target + " AUPowerManagement: " + flag48 + " => " + flag47);
                                registryKey2.SetValue("AUPowerManagement", Convert.ToInt32(flag47), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AUPowerManagement",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'AUPowerManagement' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetDisablePauseUXAccess"))
                    {
                        bool flag49 = SetDisablePauseUXAccess.ToBool();
                        bool flag50 = false;
                        try
                        {
                            flag50 = Convert.ToBoolean(registryKey2.GetValue("SetDisablePauseUXAccess"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetDisablePauseUXAccess", flag49));
                        if (flag49 != flag50)
                        {
                            try
                            {
                                WriteVerbose(target + " SetDisablePauseUXAccess: " + flag50 + " => " + flag49);
                                registryKey2.SetValue("SetDisablePauseUXAccess", Convert.ToInt32(flag49), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetDisablePauseUXAccess",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetDisablePauseUXAccess' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetDisableUXWUAccess"))
                    {
                        bool flag51 = SetDisableUXWUAccess.ToBool();
                        bool flag52 = false;
                        try
                        {
                            flag52 = Convert.ToBoolean(registryKey2.GetValue("SetDisableUXWUAccess"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetDisableUXWUAccess", flag51));
                        if (flag51 != flag52)
                        {
                            try
                            {
                                WriteVerbose(target + " SetDisablePauseUXAccess: " + flag52 + " => " + flag51);
                                registryKey2.SetValue("SetDisablePauseUXAccess", Convert.ToInt32(flag51), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetDisableUXWUAccess",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetDisableUXWUAccess' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RescheduleWaitTimeEnabled"))
                    {
                        bool flag53 = RescheduleWaitTimeEnabled.ToBool();
                        bool flag54 = false;
                        try
                        {
                            flag54 = Convert.ToBoolean(registryKey3.GetValue("RescheduleWaitTimeEnabled"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RescheduleWaitTimeEnabledRescheduleWaitTimeEnabled", flag53));
                        if (flag53 != flag54)
                        {
                            try
                            {
                                WriteVerbose(target + " RescheduleWaitTimeEnabled: " + flag54 + " => " + flag53);
                                registryKey3.SetValue("RescheduleWaitTimeEnabled", Convert.ToInt32(flag53), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RescheduleWaitTimeEnabled",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RescheduleWaitTimeEnabled' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RescheduleWaitTime"))
                    {
                        int rescheduleWaitTime = RescheduleWaitTime;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey3.GetValue("RescheduleWaitTime");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("RescheduleWaitTime", rescheduleWaitTime));
                        if (rescheduleWaitTime != num)
                        {
                            try
                            {
                                WriteVerbose(target + " RescheduleWaitTime: " + num + " => " + rescheduleWaitTime);
                                registryKey3.SetValue("RescheduleWaitTime", rescheduleWaitTime, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RescheduleWaitTime",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'RescheduleWaitTime' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetActiveHoursMaxRange"))
                    {
                        bool flag55 = SetActiveHoursMaxRange.ToBool();
                        bool flag56 = false;
                        try
                        {
                            flag56 = Convert.ToBoolean(registryKey2.GetValue("SetActiveHoursMaxRange"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetActiveHoursMaxRange", flag55));
                        if (flag55 != flag56)
                        {
                            try
                            {
                                WriteVerbose(target + " SetActiveHoursMaxRange: " + flag56 + " => " + flag55);
                                registryKey2.SetValue("SetActiveHoursMaxRange", Convert.ToInt32(flag55), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetActiveHoursMaxRange",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetActiveHoursMaxRange' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ActiveHoursMaxRange"))
                    {
                        int activeHoursMaxRange = ActiveHoursMaxRange;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ActiveHoursMaxRange");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ActiveHoursMaxRange", activeHoursMaxRange));
                        if (activeHoursMaxRange != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ActiveHoursMaxRange: " + num + " => " + activeHoursMaxRange);
                                registryKey2.SetValue("ActiveHoursMaxRange", activeHoursMaxRange, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ActiveHoursMaxRange",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ActiveHoursMaxRange' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("UseUpdateClassPolicySource"))
                    {
                        bool flag57 = UseUpdateClassPolicySource.ToBool();
                        bool flag58 = false;
                        try
                        {
                            flag58 = Convert.ToBoolean(registryKey3.GetValue("UseUpdateClassPolicySource"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("UseUpdateClassPolicySource", flag57));
                        if (flag57 != flag58)
                        {
                            try
                            {
                                WriteVerbose(target + " UseUpdateClassPolicySource: " + flag58 + " => " + flag57);
                                registryKey3.SetValue("UseUpdateClassPolicySource", Convert.ToInt32(flag57), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "UseUpdateClassPolicySource",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'UseUpdateClassPolicySource' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetPolicyDrivenUpdateSourceForFeatureUpdates"))
                    {
                        bool flag59 = SetPolicyDrivenUpdateSourceForFeatureUpdates.ToBool();
                        bool flag60 = false;
                        try
                        {
                            flag60 = Convert.ToBoolean(registryKey2.GetValue("SetPolicyDrivenUpdateSourceForFeatureUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetPolicyDrivenUpdateSourceForFeatureUpdates", flag59));
                        if (flag59 != flag60)
                        {
                            try
                            {
                                WriteVerbose(target + " SetPolicyDrivenUpdateSourceForFeatureUpdates: " + flag60 + " => " + flag59);
                                registryKey2.SetValue("SetPolicyDrivenUpdateSourceForFeatureUpdates", Convert.ToInt32(flag59),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "SetPolicyDrivenUpdateSourceForFeatureUpdates",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetPolicyDrivenUpdateSourceForFeatureUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetPolicyDrivenUpdateSourceForQualityUpdates"))
                    {
                        bool flag61 = SetPolicyDrivenUpdateSourceForQualityUpdates.ToBool();
                        bool flag62 = false;
                        try
                        {
                            flag62 = Convert.ToBoolean(registryKey2.GetValue("SetPolicyDrivenUpdateSourceForQualityUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetPolicyDrivenUpdateSourceForQualityUpdates", flag61));
                        if (flag61 != flag62)
                        {
                            try
                            {
                                WriteVerbose(target + " SetPolicyDrivenUpdateSourceForQualityUpdates: " + flag62 + " => " + flag61);
                                registryKey2.SetValue("SetPolicyDrivenUpdateSourceForQualityUpdates", Convert.ToInt32(flag61),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "SetPolicyDrivenUpdateSourceForQualityUpdates",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetPolicyDrivenUpdateSourceForQualityUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetPolicyDrivenUpdateSourceForDriverUpdates"))
                    {
                        bool flag63 = SetPolicyDrivenUpdateSourceForDriverUpdates.ToBool();
                        bool flag64 = false;
                        try
                        {
                            flag64 = Convert.ToBoolean(registryKey2.GetValue("SetPolicyDrivenUpdateSourceForDriverUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetPolicyDrivenUpdateSourceForDriverUpdates", flag63));
                        if (flag63 != flag64)
                        {
                            try
                            {
                                WriteVerbose(target + " SetPolicyDrivenUpdateSourceForDriverUpdates: " + flag64 + " => " + flag63);
                                registryKey2.SetValue("SetPolicyDrivenUpdateSourceForDriverUpdates", Convert.ToInt32(flag63),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "SetPolicyDrivenUpdateSourceForDriverUpdates", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetPolicyDrivenUpdateSourceForDriverUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetPolicyDrivenUpdateSourceForOtherUpdates"))
                    {
                        bool flag65 = SetPolicyDrivenUpdateSourceForOtherUpdates.ToBool();
                        bool flag66 = false;
                        try
                        {
                            flag66 = Convert.ToBoolean(registryKey2.GetValue("SetPolicyDrivenUpdateSourceForOtherUpdates"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetPolicyDrivenUpdateSourceForOtherUpdates", flag65));
                        if (flag65 != flag66)
                        {
                            try
                            {
                                WriteVerbose(target + " SetPolicyDrivenUpdateSourceForOtherUpdates: " + flag66 + " => " + flag65);
                                registryKey2.SetValue("SetPolicyDrivenUpdateSourceForOtherUpdates", Convert.ToInt32(flag65),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."),
                                    "SetPolicyDrivenUpdateSourceForOtherUpdates", ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetPolicyDrivenUpdateSourceForOtherUpdates' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetActiveHours"))
                    {
                        bool flag67 = SetActiveHours.ToBool();
                        bool flag68 = false;
                        try
                        {
                            flag68 = Convert.ToBoolean(registryKey2.GetValue("SetActiveHours"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetActiveHours", flag67));
                        if (flag67 != flag68)
                        {
                            try
                            {
                                WriteVerbose(target + " SetActiveHours: " + flag68 + " => " + flag67);
                                registryKey2.SetValue("SetActiveHours", Convert.ToInt32(flag67), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetActiveHours",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetActiveHours' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ActiveHoursStart"))
                    {
                        int activeHoursStart = ActiveHoursStart;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ActiveHoursStart");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ActiveHoursStart", activeHoursStart));
                        if (activeHoursStart != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ActiveHoursStart: " + num + " => " + activeHoursStart);
                                registryKey2.SetValue("ActiveHoursStart", activeHoursStart, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ActiveHoursStart",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ActiveHoursStart' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("ActiveHoursEnd"))
                    {
                        int activeHoursEnd = ActiveHoursEnd;
                        int num = 0;
                        try
                        {
                            num = (int)registryKey2.GetValue("ActiveHoursEnd");
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("ActiveHoursEnd", activeHoursEnd));
                        if (activeHoursEnd != num)
                        {
                            try
                            {
                                WriteVerbose(target + " ActiveHoursEnd: " + num + " => " + activeHoursEnd);
                                registryKey2.SetValue("ActiveHoursEnd", activeHoursEnd, RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ActiveHoursEnd",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'ActiveHoursEnd' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetAutoRestartNotificationDisable"))
                    {
                        bool flag69 = SetAutoRestartNotificationDisable.ToBool();
                        bool flag70 = false;
                        try
                        {
                            flag70 = Convert.ToBoolean(registryKey2.GetValue("SetAutoRestartNotificationDisable"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetAutoRestartNotificationDisable", flag69));
                        if (flag69 != flag70)
                        {
                            try
                            {
                                WriteVerbose(target + " SetAutoRestartNotificationDisable: " + flag70 + " => " + flag69);
                                registryKey2.SetValue("SetAutoRestartNotificationDisable", Convert.ToInt32(flag69),
                                    RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetAutoRestartNotificationDisable",
                                    ErrorCategory.CloseError,
                                    null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetAutoRestartNotificationDisable' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("EnableFeaturedSoftware"))
                    {
                        bool flag71 = EnableFeaturedSoftware.ToBool();
                        bool flag72 = false;
                        try
                        {
                            flag72 = Convert.ToBoolean(registryKey3.GetValue("EnableFeaturedSoftware"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("EnableFeaturedSoftware", flag71));
                        if (flag71 != flag72)
                        {
                            try
                            {
                                WriteVerbose(target + " EnableFeaturedSoftware: " + flag72 + " => " + flag71);
                                registryKey3.SetValue("EnableFeaturedSoftware", Convert.ToInt32(flag71), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "EnableFeaturedSoftware",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'EnableFeaturedSoftware' have been modified.");
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SetEDURestart"))
                    {
                        bool flag73 = SetEDURestart.ToBool();
                        bool flag74 = false;
                        try
                        {
                            flag74 = Convert.ToBoolean(registryKey2.GetValue("SetEDURestart"));
                        }
                        catch
                        {
                        }

                        sendToPipeline.Properties.Add(new PSNoteProperty("SetEDURestart", flag73));
                        if (flag73 != flag74)
                        {
                            try
                            {
                                WriteVerbose(target + " SetEDURestart: " + flag74 + " => " + flag73);
                                registryKey2.SetValue("SetEDURestart", Convert.ToInt32(flag73), RegistryValueKind.DWord);
                            }
                            catch
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "SetEDURestart",
                                    ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        }
                        else
                            WriteVerbose(target +
                                         " The command completed successfully but no settings of 'SetEDURestart' have been modified.");
                    }

                    WriteObject(sendToPipeline, true);
                    OutputObj.Add(sendToPipeline);
                }
            }
        }

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
                    logonType = WUImpersonator.LogonSessionType.NewCredentials;
                    logonProvider = WUImpersonator.LogonProvider.WinNT50;
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

        protected override void EndProcessing()
        {
            CmdletEnd = DateTime.Now;
            var CmdletInfo = new PSObject();
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletStart", CmdletStart));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletEnd", CmdletEnd));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletLine", MyInvocation.Line));
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", "*");
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}