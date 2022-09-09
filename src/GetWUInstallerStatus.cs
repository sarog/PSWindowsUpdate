﻿using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUInstallerStatus", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(InstallerStatus))]
    public class GetWUInstallerStatus : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        private string[] ComputerName { get; set; }

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
        public SwitchParameter Silent { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateInstaller InstallerObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (SendReport) {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new InstallerStatus());
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
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pswuModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (ShouldProcess(target, "(" + DateTime.Now + ") Get Windows Update installer status")) {
                    var wuApiInstallerObj = WUToolsObj.GetWUApiInstallerObj(target);
                    WriteDebug(DateTime.Now + " InstallerObj mode: " + wuApiInstallerObj.Mode);
                    if (wuApiInstallerObj.Status) {
                        InstallerObj = (UpdateInstaller)wuApiInstallerObj.Object;
                        var isBusy = InstallerObj.IsBusy;
                        var sendToPipeline = new PSObject();
                        sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                        sendToPipeline.Properties.Add(new PSNoteProperty("IsBusy", isBusy));
                        OutputObj.Add(sendToPipeline);
                        if (!Silent) {
                            WriteObject(sendToPipeline, true);
                        } else {
                            WriteObject(isBusy, true);
                        }
                    } else if (Debuger) {
                        WriteError(new ErrorRecord(wuApiInstallerObj.Exception, "Debug", ErrorCategory.CloseError, null));
                    } else {
                        WriteError(wuApiInstallerObj.Error);
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
                    PSWUSettings.Add("Properties", new InstallerStatus());
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