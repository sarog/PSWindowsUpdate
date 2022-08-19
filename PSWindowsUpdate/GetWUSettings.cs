using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUSettings", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(WUSettings) })]
    public class GetWUSettings : PSCmdlet {
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
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
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

                if (ShouldProcess(str1, "(" + DateTime.Now.ToString() + ") Get Windows Update settings")) {
                    var registryKey1 = !WUToolsObj.IsLocalHost(str1)
                        ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, str1)
                        : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                    var registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\");
                    if (registryKey2 != null) {
                        WriteVerbose("Some settings are managed by your system administrator.");
                        var sendToPipeline = new PSObject();
                        sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty("ComputerName", (object)str1));
                        foreach (var valueName in registryKey2.GetValueNames()) {
                            var obj = registryKey2.GetValue(valueName);
                            sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty(valueName, obj));
                        }

                        foreach (var valueName in registryKey3.GetValueNames()) {
                            var key = registryKey3.GetValue(valueName);
                            var str2 = valueName;
                            if (!(str2 == "AUOptions")) {
                                if (str2 == "ScheduledInstallDay") {
                                    key = hashtable2[key];
                                }
                            } else {
                                key = hashtable1[key];
                            }

                            sendToPipeline.Properties.Add((PSPropertyInfo)new PSNoteProperty(valueName, key));
                        }

                        WriteObject((object)sendToPipeline, true);
                        OutputObj.Add(sendToPipeline);
                    } else {
                        WriteVerbose("Settings not found.");
                    }
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