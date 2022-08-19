using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;

namespace PSWindowsUpdate {
    [Cmdlet("Update", "WUModule", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "Online", SupportsShouldProcess = true)]
    public class UpdateWUModule : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(Mandatory = true, ParameterSetName = "Local", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "Online", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        public PSCredential Credential { get; set; }

        [Parameter]
        private SwitchParameter SendReport { get; set; }

        [Parameter]
        private Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter(Mandatory = true, ParameterSetName = "Online")]
        public SwitchParameter Online { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "Local")]
        public SwitchParameter Local { get; set; }

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
            if ((bool)Local) {
                var pswuModule = WUToolsObj.GetPSWUModule(Environment.MachineName);
                try {
                    var version1 = (Version)pswuModule.Properties["Version"].Value;
                    var version2 = (Version)pswuModule.Properties["PSWUDllVersion"].Value;
                    WriteVerbose(Environment.MachineName + ": Local module: " + version1.ToString() + " (" + version2.ToString() + ")");
                } catch { }
            }

            foreach (var str in ComputerName) {
                var now = DateTime.Now;
                WriteDebug(now.ToString() + " " + str + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(str);
                    now = DateTime.Now;
                    WriteDebug(now.ToString() + " Module version: " + pswuModule.Properties["Version"].Value?.ToString());
                    now = DateTime.Now;
                    WriteDebug(now.ToString() + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value?.ToString());
                } catch { }

                var pswuModule1 = WUToolsObj.GetPSWUModule(str);
                try {
                    var version3 = (Version)pswuModule1.Properties["Version"].Value;
                    var version4 = (Version)pswuModule1.Properties["PSWUDllVersion"].Value;
                    WriteVerbose(str + ": Remote module: " + version3.ToString() + " (" + version4.ToString() + ")");
                } catch {
                    WriteDebug(DateTime.Now.ToString() + " Can't get PSWU module from " + str);
                }

                if (ShouldProcess(str, "(" + DateTime.Now.ToString() + ") Update PSWindowsUpdate module")) {
                    string script;
                    if ((bool)Online) {
                        WriteVerbose("Installing module...");
                        script =
                            "Invoke-Command -ComputerName $Computer -Command { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force; Install-Module -Name PSWindowsUpdate -Force -Verbose } -Authentication Negotiate"
                                .Replace("$Computer", str);
                        WriteDebug(DateTime.Now.ToString() + " Online mode: " + script);
                    } else {
                        WriteVerbose("Cloning module...");
                        script =
                            "$ModuleBase = (Get-Module PSWindowsUpdate -ListAvailable | Sort Version -Descending | Select -First 1).ModuleBase\r\n$Destination = '\\\\$Computer\\' + $($ModuleBase -replace ':','$') \r\n\r\nCopy-Item $ModuleBase -Destination $Destination -Force\r\nCopy-Item ($ModuleBase + '\\*') -Destination $Destination -Force\r\n\r\nInvoke-Command -ComputerName $Computer -Command { ipmo PSWindowsUpdate; Enable-WURemoting -Verbose } -Authentication Negotiate"
                                .Replace("$Computer", str);
                        WriteDebug(DateTime.Now.ToString() + " Local mode: " + script);
                    }

                    var errorRecord = (ErrorRecord)null;
                    using (var powerShell = PowerShell.Create()) {
                        powerShell.AddScript(script);
                        powerShell.Invoke();
                        WriteObject((object)powerShell.Streams.Verbose);
                        WriteObject((object)powerShell.Streams.Error);
                        if (powerShell.Streams.Error.Count > 0) {
                            errorRecord = powerShell.Streams.Error[0];
                        }
                    }

                    var pswuModule2 = WUToolsObj.GetPSWUModule(str);
                    try {
                        var version5 = (Version)pswuModule2.Properties["Version"].Value;
                        var version6 = (Version)pswuModule2.Properties["PSWUDllVersion"].Value;
                        WriteVerbose(str + ": Remote module: " + version5.ToString() + " (" + version6.ToString() + ")");
                    } catch {
                        WriteDebug(DateTime.Now.ToString() + " Can't get PSWU module from " + str);
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