using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WURebootStatus", ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = "ManualReboot", SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(RebootStatus) })]
    public class GetWURebootStatus : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter]
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

        [Parameter(ParameterSetName = "AutoReboot", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "ScheduleReboot", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "ManualReboot", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "CancelReboot", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Silent { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "AutoReboot")]
        public SwitchParameter AutoReboot { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "ScheduleReboot")]
        public DateTime ScheduleReboot { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "CancelReboot")]
        public SwitchParameter CancelReboot { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static SystemInformation SystemInfoObj { get; set; }

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
                    PSWUSettings.Add((object)"Properties", (object)new string[2] {
                        "ComputerName",
                        "RebootRequired"
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
            WriteDebug(DateTime.Now.ToString() + " Connect to local machine: " + ComputerName[0]);
        }

        private void CoreProcessing() {
            var invocationName = MyInvocation.InvocationName;
            foreach (var str in ComputerName) {
                var dateTime = DateTime.Now;
                WriteDebug(dateTime.ToString() + " " + str + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(str);
                    dateTime = DateTime.Now;
                    WriteDebug(dateTime.ToString() + " Module version: " + pswuModule.Properties["Version"].Value?.ToString());
                    dateTime = DateTime.Now;
                    WriteDebug(dateTime.ToString() + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value?.ToString());
                } catch { }

                var target = str;
                dateTime = DateTime.Now;
                var action = "(" + dateTime.ToString() + ") Get Windows Update reboot status";
                if (ShouldProcess(target, action)) {
                    bool sendToPipeline1;
                    if (WUToolsObj.IsLocalHost(str)) {
                        var apiSystemInfoObj = WUToolsObj.GetWUApiSystemInfoObj(str);
                        dateTime = DateTime.Now;
                        WriteDebug(dateTime.ToString() + " SystemInfoObj mode: " + apiSystemInfoObj.Mode);
                        if (apiSystemInfoObj.Status) {
                            SystemInfoObj = (SystemInformation)apiSystemInfoObj.Object;
                            sendToPipeline1 = SystemInfoObj.RebootRequired;
                        } else {
                            if ((bool)Debuger) {
                                WriteError(new ErrorRecord(apiSystemInfoObj.Exception, "Debug", ErrorCategory.CloseError, (object)null));
                                continue;
                            }

                            WriteError(apiSystemInfoObj.Error);
                            continue;
                        }
                    } else {
                        sendToPipeline1 = ((IEnumerable<string>)RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, str)
                            .OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\Auto Update\\").GetSubKeyNames()).Contains<string>("RebootRequired");
                    }

                    var sendToPipeline2 = new PSObject();
                    sendToPipeline2.Properties.Add((PSPropertyInfo)new PSNoteProperty("ComputerName", (object)str));
                    sendToPipeline2.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootRequired", (object)sendToPipeline1));
                    if (ScheduleReboot != DateTime.MinValue) {
                        sendToPipeline2.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootScheduled", (object)ScheduleReboot));
                    } else {
                        sendToPipeline2.Properties.Add((PSPropertyInfo)new PSNoteProperty("RebootScheduled", (object)null));
                    }

                    if ((bool)Silent) {
                        WriteObject((object)sendToPipeline1, true);
                    } else {
                        WriteObject((object)sendToPipeline2, true);
                    }

                    OutputObj.Add(sendToPipeline2);
                    if (sendToPipeline1) {
                        if (ScheduleReboot != DateTime.MinValue) {
                            var strArray = new string[5];
                            dateTime = DateTime.Now;
                            strArray[0] = dateTime.ToString();
                            strArray[1] = " ";
                            strArray[2] = str;
                            strArray[3] = ": Schedule Reboot ";
                            dateTime = ScheduleReboot;
                            strArray[4] = dateTime.ToString();
                            WriteDebug(string.Concat(strArray));
                            WUToolsObj.CancelReboot(str);
                            WriteVerbose(WUToolsObj.ScheduleReboot(str, ScheduleReboot));
                        } else if ((bool)AutoReboot) {
                            dateTime = DateTime.Now;
                            WriteDebug(dateTime.ToString() + " " + str + ": Auto Reboot");
                            WUToolsObj.CancelReboot(str);
                            WriteVerbose(WUToolsObj.RunReboot(str));
                        } else if (!(bool)Silent) {
                            Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required. Do it now? [Y / N] (default is 'N')");
                            if (Console.ReadLine().ToUpper() == "Y") {
                                dateTime = DateTime.Now;
                                WriteDebug(dateTime.ToString() + " " + str + ": Manually Reboot");
                                WUToolsObj.CancelReboot(str);
                                WriteVerbose(WUToolsObj.RunReboot(str));
                            }
                        }
                    }

                    if ((bool)CancelReboot) {
                        dateTime = DateTime.Now;
                        WriteDebug(dateTime.ToString() + " " + str + ": Cancel Reboot");
                        WriteVerbose(WUToolsObj.CancelReboot(str));
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
                    PSWUSettings.Add((object)"Properties", (object)new string[2] {
                        "ComputerName",
                        "RebootRequired"
                    });
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