using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;

namespace PSWindowsUpdate {
    [Cmdlet("Set", "WUSettings", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(WUSettings) })]
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

        [ValidateSet(new string[] { "Not configured", "Disabled", "Notify before download", "Notify before installation", "Scheduled installation", "Users configure" })]
        [Alias(new string[] { "AUOptions" })]
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

        [ValidateSet(new string[] { "Every Day", "Every Sunday", "Every Monday", "Every Tuesday", "Every Wednesday", "Every Thursday", "Every Friday", "EverySaturday" })]
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
            WriteDebug(DateTime.Now.ToString() + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, (object)null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if ((bool)SendReport) {
                WriteDebug(DateTime.Now.ToString() + " Test smtp settings");
                if (!PSWUSettings.ContainsKey((object)"Properties")) {
                    PSWUSettings.Add((object)"Properties", (object)new string[1] {
                        "*"
                    });
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = (SwitchParameter)false;
                    WriteDebug(DateTime.Now.ToString() + " Disabling -SendReport");
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
            var hashtable1 = new Hashtable();
            hashtable1.Add((object)0, (object)"0 - Not configured");
            hashtable1.Add((object)1, (object)"1 - Disabled");
            hashtable1.Add((object)2, (object)"2 - Notify before download");
            hashtable1.Add((object)3, (object)"3 - Notify before installation");
            hashtable1.Add((object)4, (object)"4 - Scheduled installation");
            hashtable1.Add((object)5, (object)"5 - Users configure");
            var hashtable2 = new Hashtable();
            hashtable2.Add((object)0, (object)"0 - Every Day");
            hashtable2.Add((object)1, (object)"1 - Every Sunday");
            hashtable2.Add((object)2, (object)"2 - Every Monday");
            hashtable2.Add((object)3, (object)"3 - Every Tuesday");
            hashtable2.Add((object)4, (object)"4 - Every Wednesday");
            hashtable2.Add((object)5, (object)"5 - Every Thursday");
            hashtable2.Add((object)6, (object)"6 - Every Friday");
            hashtable2.Add((object)7, (object)"7 - Every Saturday");
            foreach (var str1 in ComputerName) {
                WriteDebug(DateTime.Now.ToString() + " " + str1 + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(str1);
                    WriteDebug(DateTime.Now.ToString() + " Module version: " + pswuModule.Properties["Version"].Value?.ToString());
                    WriteDebug(DateTime.Now.ToString() + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value?.ToString());
                } catch { }

                if (ShouldProcess(str1, "(" + DateTime.Now.ToString() + ") Set Windows Update settings")) {
                    var registryKey1 = !WUToolsObj.IsLocalHost(str1)
                        ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, str1)
                        : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                    var registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                    var sendToPipeline = new PSObject();
                    sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("ComputerName", (object)str1));
                    if (registryKey2 != null) {
                        WriteVerbose("Some settings are managed by your system administrator. Changes may don't be applied.");
                    }

                    if (registryKey2 == null) {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                        registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\", true);
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RegistrySubKey1", (object)"Created"));
                    }

                    if (registryKey3 == null) {
                        registryKey1.CreateSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\");
                        registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\", true);
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RegistrySubKey2", (object)"Created"));
                    }

                    SwitchParameter switchParameter;
                    if (MyInvocation.BoundParameters.ContainsKey("AcceptTrustedPublisherCerts")) {
                        switchParameter = AcceptTrustedPublisherCerts;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("AcceptTrustedPublisherCerts"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("AcceptTrustedPublisherCerts", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " AcceptTrustedPublisherCerts: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey2.SetValue("AcceptTrustedPublisherCerts", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "AcceptTrustedPublisherCerts", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'AcceptTrustedPublisherCerts' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DisableWindowsUpdateAccess")) {
                        switchParameter = DisableWindowsUpdateAccess;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("DisableWindowsUpdateAccess"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("DisableWindowsUpdateAccess", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " DisableWindowsUpdateAccess: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey2.SetValue("DisableWindowsUpdateAccess", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "DisableWindowsUpdateAccess", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'DisableWindowsUpdateAccess' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NonAdministratorsElevated")) {
                        switchParameter = NonAdministratorsElevated;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("NonAdministratorsElevated"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("NonAdministratorsElevated", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " NonAdministratorsElevated: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey2.SetValue("NonAdministratorsElevated", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "NonAdministratorsElevated", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'NonAdministratorsElevated' have been modified.");
                        }
                    }

                    if (TargetGroup != null) {
                        var targetGroup = TargetGroup;
                        var str2 = (string)registryKey2.GetValue("TargetGroup");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("TargetGroup", (object)targetGroup));
                        if (targetGroup != str2) {
                            try {
                                WriteVerbose(str1 + " TargetGroup: " + str2 + " => " + targetGroup);
                                registryKey2.SetValue("TargetGroup", (object)targetGroup, RegistryValueKind.String);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "TargetGroup", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'TargetGroup' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("TargetGroupEnabled")) {
                        switchParameter = TargetGroupEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey2.GetValue("TargetGroupEnabled"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("TargetGroupEnabled", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " TargetGroupEnabled: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey2.SetValue("TargetGroupEnabled", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "TargetGroupEnabled", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'TargetGroupEnabled' have been modified.");
                        }
                    }

                    if (WUServer != null) {
                        var wuServer = WUServer;
                        var str3 = (string)registryKey2.GetValue("WUServer");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("WUServer", (object)wuServer));
                        if (wuServer != str3) {
                            try {
                                WriteVerbose(str1 + " WUServer: " + str3 + " => " + wuServer);
                                registryKey2.SetValue("WUServer", (object)wuServer, RegistryValueKind.String);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "WUServer", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'WUServer' have been modified.");
                        }
                    }

                    if (NotificationLevel != null) {
                        var num1 = (int)hashtable1[(object)NotificationLevel];
                        var num2 = (int)registryKey3.GetValue("AUOptions");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("AUOptions", (object)num1));
                        if (num1 != num2) {
                            try {
                                WriteVerbose(str1 + " AUOptions: " + num2.ToString() + " => " + num1.ToString());
                                registryKey3.SetValue("AUOptions", (object)num1, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "AUOptions", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'AUOptions' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("AutoInstallMinorUpdates")) {
                        switchParameter = AutoInstallMinorUpdates;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("AutoInstallMinorUpdates"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("AutoInstallMinorUpdates", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " AutoInstallMinorUpdates: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("AutoInstallMinorUpdates", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "AutoInstallMinorUpdates", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'AutoInstallMinorUpdates' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(DetectionFrequency.ToString())) {
                        var detectionFrequency = DetectionFrequency;
                        var num = (int)registryKey3.GetValue("DetectionFrequency");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("DetectionFrequency", (object)detectionFrequency));
                        if (detectionFrequency != num) {
                            try {
                                WriteVerbose(str1 + " DetectionFrequency: " + num.ToString() + " => " + detectionFrequency.ToString());
                                registryKey3.SetValue("DetectionFrequency", (object)detectionFrequency, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "DetectionFrequency", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'DetectionFrequency' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("DetectionFrequencyEnabled")) {
                        switchParameter = DetectionFrequencyEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("DetectionFrequencyEnabled"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("DetectionFrequencyEnabled", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " DetectionFrequencyEnabled: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("DetectionFrequencyEnabled", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "DetectionFrequencyEnabled", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'DetectionFrequencyEnabled' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("IncludeRecommendedUpdates")) {
                        switchParameter = IncludeRecommendedUpdates;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("IncludeRecommendedUpdates"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("IncludeRecommendedUpdates", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " IncludeRecommendedUpdates: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("IncludeRecommendedUpdates", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "IncludeRecommendedUpdates", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'IncludeRecommendedUpdates' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("NoAutoRebootWithLoggedOnUsers")) {
                        switchParameter = NoAutoRebootWithLoggedOnUsers;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("NoAutoRebootWithLoggedOnUsers"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("NoAutoRebootWithLoggedOnUsers", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " NoAutoRebootWithLoggedOnUsers: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("NoAutoRebootWithLoggedOnUsers", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "NoAutoRebootWithLoggedOnUsers", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'NoAutoRebootWithLoggedOnUsers' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RebootRelaunchTimeout.ToString())) {
                        var rebootRelaunchTimeout = RebootRelaunchTimeout;
                        var num = (int)registryKey3.GetValue("RebootRelaunchTimeout");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootRelaunchTimeout", (object)rebootRelaunchTimeout));
                        if (rebootRelaunchTimeout != num) {
                            try {
                                WriteVerbose(str1 + " RebootRelaunchTimeout: " + num.ToString() + " => " + rebootRelaunchTimeout.ToString());
                                registryKey3.SetValue("RebootRelaunchTimeout", (object)rebootRelaunchTimeout, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "RebootRelaunchTimeout", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'RebootRelaunchTimeout' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootRelaunchTimeoutEnabled")) {
                        switchParameter = RebootRelaunchTimeoutEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("RebootRelaunchTimeoutEnabled"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootRelaunchTimeoutEnabled", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " RebootRelaunchTimeoutEnabled: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("RebootRelaunchTimeoutEnabled", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "RebootRelaunchTimeoutEnabled", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'RebootRelaunchTimeoutEnabled' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RebootWarningTimeout.ToString())) {
                        var rebootWarningTimeout = RebootWarningTimeout;
                        var num = (int)registryKey3.GetValue("RebootWarningTimeout");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootWarningTimeout", (object)rebootWarningTimeout));
                        if (rebootWarningTimeout != num) {
                            try {
                                WriteVerbose(str1 + " RebootWarningTimeout: " + num.ToString() + " => " + rebootWarningTimeout.ToString());
                                registryKey3.SetValue("RebootWarningTimeout", (object)rebootWarningTimeout, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "RebootWarningTimeout", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'RebootWarningTimeout' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("RebootWarningTimeoutEnabled")) {
                        switchParameter = RebootWarningTimeoutEnabled;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("RebootWarningTimeoutEnabled"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootWarningTimeoutEnabled", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " RebootWarningTimeoutEnabled: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("RebootWarningTimeoutEnabled", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "RebootWarningTimeoutEnabled", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'RebootWarningTimeoutEnabled' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(RescheduleWaitTime.ToString())) {
                        var rescheduleWaitTime = RescheduleWaitTime;
                        var num = (int)registryKey3.GetValue("RescheduleWaitTime");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("RescheduleWaitTime", (object)rescheduleWaitTime));
                        if (rescheduleWaitTime != num) {
                            try {
                                WriteVerbose(str1 + " RescheduleWaitTime: " + num.ToString() + " => " + rescheduleWaitTime.ToString());
                                registryKey3.SetValue("RescheduleWaitTime", (object)rescheduleWaitTime, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "RescheduleWaitTime", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'RescheduleWaitTime' have been modified.");
                        }
                    }

                    if (ScheduledInstallDay != null) {
                        var num3 = (int)hashtable2[(object)ScheduledInstallDay];
                        var num4 = (int)registryKey3.GetValue("ScheduledInstallDay");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("ScheduledInstallDay", (object)num3));
                        if (num3 != num4) {
                            try {
                                WriteVerbose(str1 + " ScheduledInstallDay: " + num4.ToString() + " => " + num3.ToString());
                                registryKey3.SetValue("ScheduledInstallDay", (object)num3, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "ScheduledInstallDay", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'ScheduledInstallDay' have been modified.");
                        }
                    }

                    if (string.IsNullOrEmpty(ScheduledInstallTime.ToString())) {
                        var scheduledInstallTime = ScheduledInstallTime;
                        var num = (int)registryKey3.GetValue("ScheduledInstallTime");
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("ScheduledInstallTime", (object)scheduledInstallTime));
                        if (scheduledInstallTime != num) {
                            try {
                                WriteVerbose(str1 + " ScheduledInstallTime: " + num.ToString() + " => " + scheduledInstallTime.ToString());
                                registryKey3.SetValue("ScheduledInstallTime", (object)scheduledInstallTime, RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "ScheduledInstallTime", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'ScheduledInstallTime' have been modified.");
                        }
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("UseWUServer")) {
                        switchParameter = UseWUServer;
                        var flag = switchParameter.ToBool();
                        var boolean = Convert.ToBoolean(registryKey3.GetValue("UseWUServer"));
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("UseWUServer", (object)flag));
                        if (flag != boolean) {
                            try {
                                WriteVerbose(str1 + " UseWUServer: " + boolean.ToString() + " => " + flag.ToString());
                                registryKey3.SetValue("UseWUServer", (object)Convert.ToInt32(flag), RegistryValueKind.DWord);
                            } catch {
                                WriteError(new ErrorRecord(new Exception(str1 + ": Access denied."), "UseWUServer", ErrorCategory.CloseError, (object)null));
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else {
                            WriteVerbose(str1 + " The command completed successfully but no settings of 'UseWUServer' have been modified.");
                        }
                    }

                    WriteObject((object)sendToPipeline, true);
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
                WriteDebug(DateTime.Now.ToString() + " UserName: " + userName + "; Domain: " + domain + "; Password: " + password.Substring(0, 1) + "*****");
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
                            WriteDebug(DateTime.Now.ToString() + " Something goes wrong: " + ex.Message);
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
                    WriteDebug(now.ToString() + " Leaving impersonated session");
                }

                var windowsPrincipal2 = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                var str4 = "";
                if (windowsPrincipal2.IsInRole(WindowsBuiltInRole.Administrator)) {
                    str4 = "RunAs";
                }

                WriteDebug(DateTime.Now.ToString() + " After User: " + WindowsIdentity.GetCurrent().Name + " " + str4);
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
            CmdletInfo.Properties.Add((PSPropertyInfo)new PSNoteProperty("CmdletStart", (object)CmdletStart));
            CmdletInfo.Properties.Add((PSPropertyInfo)new PSNoteProperty("CmdletEnd", (object)CmdletEnd));
            CmdletInfo.Properties.Add((PSPropertyInfo)new PSNoteProperty("CmdletLine", (object)MyInvocation.Line));
            if ((bool)SendReport) {
                WriteDebug(DateTime.Now.ToString() + " Send report");
                if (!PSWUSettings.ContainsKey((object)"Properties")) {
                    PSWUSettings.Add((object)"Properties", (object)"*");
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            WriteDebug(DateTime.Now.ToString() + " CmdletEnd");
        }

        protected override void StopProcessing() {
            base.StopProcessing();
        }
    }
}