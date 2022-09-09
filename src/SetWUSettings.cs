using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;
using Microsoft.Win32;

namespace PSWindowsUpdate {
    
    [Cmdlet("Set", "WUSettings", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    [OutputType(typeof(WUSettings))]
    public class SetWUSettings : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        private PSCredential Credential { get; set; }

        [Parameter]
        public SwitchParameter SendReport { get; set; }

        [Parameter]
        public Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter]
        public SwitchParameter AcceptTrustedPublisherCerts { get; set; }

        [Parameter]
        public SwitchParameter DisableWindowsUpdateAccess { get; set; }

        [Parameter]
        public SwitchParameter NonAdministratorsElevated { get; set; }

        [Parameter]
        public string TargetGroup { get; set; }

        [Parameter]
        public SwitchParameter TargetGroupEnabled { get; set; }

        [Parameter]
        public string WUServer { get; set; }

        [ValidateSet("Not configured", "Disabled", "Notify before download", "Notify before installation", "Scheduled installation", "Users configure")]
        [Alias("AUOptions")]
        [Parameter]
        public string NotificationLevel { get; set; }

        [Parameter]
        public SwitchParameter AutoInstallMinorUpdates { get; set; }

        [ValidateRange(1, 22)]
        [Parameter]
        public int DetectionFrequency { get; set; }

        [Parameter]
        public SwitchParameter DetectionFrequencyEnabled { get; set; }

        [Parameter]
        public SwitchParameter IncludeRecommendedUpdates { get; set; }

        [Parameter]
        public SwitchParameter NoAutoRebootWithLoggedOnUsers { get; set; }

        [Parameter]
        public SwitchParameter NoAutoUpdate { get; set; }

        [ValidateRange(1, 1440)]
        [Parameter]
        public int RebootRelaunchTimeout { get; set; }

        [Parameter]
        public SwitchParameter RebootRelaunchTimeoutEnabled { get; set; }

        [Parameter]
        public int RebootWarningTimeout { get; set; }

        [Parameter]
        public SwitchParameter RebootWarningTimeoutEnabled { get; set; }

        [ValidateRange(1, 60)]
        [Parameter]
        public int RescheduleWaitTime { get; set; }

        [ValidateSet("Every Day", "Every Sunday", "Every Monday", "Every Tuesday", "Every Wednesday", "Every Thursday", "Every Friday", "EverySaturday")]
        [Parameter]
        public string ScheduledInstallDay { get; set; }

        [ValidateRange(0, 23)]
        [Parameter]
        public int ScheduledInstallTime { get; set; }

        [Parameter]
        public SwitchParameter UseWUServer { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (SendReport) {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new string[1] {
                        "*"
                    });
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
            var notifyTypes = new Hashtable {
                { 0, "0 - Not configured" },
                { 1, "1 - Disabled" },
                { 2, "2 - Notify before download" },
                { 3, "3 - Notify before installation" },
                { 4, "4 - Scheduled installation" },
                { 5, "5 - Users configure" }
            };
            
            var scheduleDay = new Hashtable {
                { 0, "0 - Every Day" },
                { 1, "1 - Every Sunday" },
                { 2, "2 - Every Monday" },
                { 3, "3 - Every Tuesday" },
                { 4, "4 - Every Wednesday" },
                { 5, "5 - Every Thursday" },
                { 6, "6 - Every Friday" },
                { 7, "7 - Every Saturday" }
            };

            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pswuModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (ShouldProcess(target, "(" + DateTime.Now + ") Set Windows Update settings")) {
                    var registryKey1 = !WUToolsObj.IsLocalHost(target)
                        ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                        : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                    var registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                    var sendToPipeline = new PSObject();
                    sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                    if (registryKey2 != null) {
                        WriteVerbose("Some settings are managed by your system administrator. Changes may not be applied.");
                    }

                    if (registryKey2 == null) {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                        registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                        sendToPipeline.Properties.Add(new PSNoteProperty("RegistrySubKey1", "Created"));
                    }

                    if (registryKey3 == null) {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\");
                        registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                        sendToPipeline.Properties.Add(new PSNoteProperty("RegistrySubKey2", "Created"));
                    }

                    SwitchParameter switchParameter;
                    if (MyInvocation.BoundParameters.ContainsKey("AcceptTrustedPublisherCerts")) {
                        switchParameter = AcceptTrustedPublisherCerts;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("AcceptTrustedPublisherCerts"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("AcceptTrustedPublisherCerts", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " AcceptTrustedPublisherCerts: " + boolean + " => " + flag);
                                registryKey2.SetValue("AcceptTrustedPublisherCerts", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AcceptTrustedPublisherCerts", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'AcceptTrustedPublisherCerts' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DisableWindowsUpdateAccess")) {
                        switchParameter = DisableWindowsUpdateAccess;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("DisableWindowsUpdateAccess"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("DisableWindowsUpdateAccess", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " DisableWindowsUpdateAccess: " + boolean + " => " + flag);
                                registryKey2.SetValue("DisableWindowsUpdateAccess", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DisableWindowsUpdateAccess", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'DisableWindowsUpdateAccess' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NonAdministratorsElevated")) {
                        switchParameter = NonAdministratorsElevated;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("NonAdministratorsElevated"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("NonAdministratorsElevated", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " NonAdministratorsElevated: " + boolean + " => " + flag);
                                registryKey2.SetValue("NonAdministratorsElevated", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "NonAdministratorsElevated", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'NonAdministratorsElevated' have been modified.");
                        }
                    }

                    if (TargetGroup != null) {
                        var targetGroup = TargetGroup;
                        var str2 = (string)registryKey2.GetValue("TargetGroup");
                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetGroup", targetGroup));
                        if (targetGroup != str2) {
                            try {
                                WriteVerbose(target + " TargetGroup: " + str2 + " => " + targetGroup);
                                registryKey2.SetValue("TargetGroup", targetGroup, RegistryValueKind.String);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetGroup", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'TargetGroup' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetGroupEnabled")) {
                        switchParameter = TargetGroupEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("TargetGroupEnabled"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("TargetGroupEnabled", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " TargetGroupEnabled: " + boolean + " => " + flag);
                                registryKey2.SetValue("TargetGroupEnabled", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "TargetGroupEnabled", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'TargetGroupEnabled' have been modified.");
                        }
                    }

                    if (WUServer != null) {
                        var wuServer = WUServer;
                        var str3 = (string)registryKey2.GetValue("WUServer");
                        sendToPipeline.Properties.Add(new PSNoteProperty("WUServer", wuServer));
                        if (wuServer != str3) {
                            try {
                                WriteVerbose(target + " WUServer: " + str3 + " => " + wuServer);
                                registryKey2.SetValue("WUServer", wuServer, RegistryValueKind.String);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "WUServer", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'WUServer' have been modified.");
                        }
                    }

                    if (NotificationLevel != null) {
                        var num1 = (int)notifyTypes[NotificationLevel];
                        var num2 = (int)registryKey3.GetValue("AUOptions");
                        sendToPipeline.Properties.Add(new PSNoteProperty("AUOptions", num1));
                        if (num1 != num2) {
                            try {
                                WriteVerbose(target + " AUOptions: " + num2 + " => " + num1);
                                registryKey3.SetValue("AUOptions", num1, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AUOptions", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'AUOptions' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AutoInstallMinorUpdates")) {
                        switchParameter = AutoInstallMinorUpdates;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("AutoInstallMinorUpdates"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("AutoInstallMinorUpdates", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " AutoInstallMinorUpdates: " + boolean + " => " + flag);
                                registryKey3.SetValue("AutoInstallMinorUpdates", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "AutoInstallMinorUpdates", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'AutoInstallMinorUpdates' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(DetectionFrequency.ToString())) {
                        var detectionFrequency = DetectionFrequency;
                        var num = (int)registryKey3.GetValue("DetectionFrequency");
                        sendToPipeline.Properties.Add(new PSNoteProperty("DetectionFrequency", detectionFrequency));
                        if (detectionFrequency != num) {
                            try {
                                WriteVerbose(target + " DetectionFrequency: " + num + " => " + detectionFrequency);
                                registryKey3.SetValue("DetectionFrequency", detectionFrequency, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DetectionFrequency", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'DetectionFrequency' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DetectionFrequencyEnabled")) {
                        switchParameter = DetectionFrequencyEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("DetectionFrequencyEnabled"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("DetectionFrequencyEnabled", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " DetectionFrequencyEnabled: " + boolean + " => " + flag);
                                registryKey3.SetValue("DetectionFrequencyEnabled", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "DetectionFrequencyEnabled", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'DetectionFrequencyEnabled' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("IncludeRecommendedUpdates")) {
                        switchParameter = IncludeRecommendedUpdates;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("IncludeRecommendedUpdates"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("IncludeRecommendedUpdates", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " IncludeRecommendedUpdates: " + boolean + " => " + flag);
                                registryKey3.SetValue("IncludeRecommendedUpdates", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "IncludeRecommendedUpdates", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'IncludeRecommendedUpdates' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NoAutoRebootWithLoggedOnUsers")) {
                        switchParameter = NoAutoRebootWithLoggedOnUsers;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("NoAutoRebootWithLoggedOnUsers"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("NoAutoRebootWithLoggedOnUsers", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " NoAutoRebootWithLoggedOnUsers: " + boolean + " => " + flag);
                                registryKey3.SetValue("NoAutoRebootWithLoggedOnUsers", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "NoAutoRebootWithLoggedOnUsers", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'NoAutoRebootWithLoggedOnUsers' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RebootRelaunchTimeout.ToString())) {
                        var rebootRelaunchTimeout = RebootRelaunchTimeout;
                        var num = (int)registryKey3.GetValue("RebootRelaunchTimeout");
                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootRelaunchTimeout", rebootRelaunchTimeout));
                        if (rebootRelaunchTimeout != num) {
                            try {
                                WriteVerbose(target + " RebootRelaunchTimeout: " + num + " => " + rebootRelaunchTimeout);
                                registryKey3.SetValue("RebootRelaunchTimeout", rebootRelaunchTimeout, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootRelaunchTimeout", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'RebootRelaunchTimeout' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootRelaunchTimeoutEnabled")) {
                        switchParameter = RebootRelaunchTimeoutEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("RebootRelaunchTimeoutEnabled"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootRelaunchTimeoutEnabled", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " RebootRelaunchTimeoutEnabled: " + boolean + " => " + flag);
                                registryKey3.SetValue("RebootRelaunchTimeoutEnabled", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootRelaunchTimeoutEnabled", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'RebootRelaunchTimeoutEnabled' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RebootWarningTimeout.ToString())) {
                        var rebootWarningTimeout = RebootWarningTimeout;
                        var num = (int)registryKey3.GetValue("RebootWarningTimeout");
                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootWarningTimeout", rebootWarningTimeout));
                        if (rebootWarningTimeout != num) {
                            try {
                                WriteVerbose(target + " RebootWarningTimeout: " + num + " => " + rebootWarningTimeout);
                                registryKey3.SetValue("RebootWarningTimeout", rebootWarningTimeout, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootWarningTimeout", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'RebootWarningTimeout' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootWarningTimeoutEnabled")) {
                        switchParameter = RebootWarningTimeoutEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("RebootWarningTimeoutEnabled"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("RebootWarningTimeoutEnabled", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " RebootWarningTimeoutEnabled: " + boolean + " => " + flag);
                                registryKey3.SetValue("RebootWarningTimeoutEnabled", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RebootWarningTimeoutEnabled", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'RebootWarningTimeoutEnabled' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RescheduleWaitTime.ToString())) {
                        var rescheduleWaitTime = RescheduleWaitTime;
                        var num = (int)registryKey3.GetValue("RescheduleWaitTime");
                        sendToPipeline.Properties.Add(new PSNoteProperty("RescheduleWaitTime", rescheduleWaitTime));
                        if (rescheduleWaitTime != num) {
                            try {
                                WriteVerbose(target + " RescheduleWaitTime: " + num + " => " + rescheduleWaitTime);
                                registryKey3.SetValue("RescheduleWaitTime", rescheduleWaitTime, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "RescheduleWaitTime", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'RescheduleWaitTime' have been modified.");
                        }
                    }

                    if (ScheduledInstallDay != null) {
                        var num3 = (int)scheduleDay[ScheduledInstallDay];
                        var num4 = (int)registryKey3.GetValue("ScheduledInstallDay");
                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduledInstallDay", num3));
                        if (num3 != num4) {
                            try {
                                WriteVerbose(target + " ScheduledInstallDay: " + num4 + " => " + num3);
                                registryKey3.SetValue("ScheduledInstallDay", num3, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduledInstallDay", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'ScheduledInstallDay' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(ScheduledInstallTime.ToString())) {
                        var scheduledInstallTime = ScheduledInstallTime;
                        var num = (int)registryKey3.GetValue("ScheduledInstallTime");
                        sendToPipeline.Properties.Add(new PSNoteProperty("ScheduledInstallTime", scheduledInstallTime));
                        if (scheduledInstallTime != num) {
                            try {
                                WriteVerbose(target + " ScheduledInstallTime: " + num + " => " + scheduledInstallTime);
                                registryKey3.SetValue("ScheduledInstallTime", scheduledInstallTime, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "ScheduledInstallTime", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'ScheduledInstallTime' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("UseWUServer")) {
                        switchParameter = UseWUServer;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("UseWUServer"));
                        sendToPipeline.Properties.Add(new PSNoteProperty("UseWUServer", flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(target + " UseWUServer: " + boolean + " => " + flag);
                                registryKey3.SetValue("UseWUServer", Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(target + ": Access denied."), "UseWUServer", ErrorCategory.CloseError, null));
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(target + " The command completed successfully but no settings of 'UseWUServer' have been modified.");
                        }
                    }

                    WriteObject(sendToPipeline, true);
                    OutputObj.Add(sendToPipeline);
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
                    logonType = WUImpersonator.LogonSessionType.NewCredentials;
                    logonProvider = WUImpersonator.LogonProvider.WinNT50;
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

                    now = DateTime.Now;
                    WriteDebug(now + " Leaving impersonated session");
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
            if (SendReport) {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", "*");
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}