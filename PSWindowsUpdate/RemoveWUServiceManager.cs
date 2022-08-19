using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Remove", "WUServiceManager", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class RemoveWUServiceManager : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        private PSCredential Credential { get; set; }

        [Parameter]
        private SwitchParameter SendReport { get; set; }

        [Parameter]
        private Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter]
        public DateTime ScheduleJob { get; set; }

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string ServiceID { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateServiceManager ServiceManagerObj { get; set; }

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
            var computerName = ComputerName;
            foreach (var text in computerName) {
                WriteDebug(DateTime.Now.ToString() + " " + text + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(text);
                    WriteDebug(DateTime.Now.ToString() + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now.ToString() + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (!ShouldProcess(text, "(" + DateTime.Now.ToString() + ") Unregister Windows Update Service Manager: " + ServiceID)) {
                    continue;
                }

                if (WUToolsObj.IsLocalHost(text) && ScheduleJob == DateTime.MinValue) {
                    WriteVerbose("Run WUApi");
                    var wUApiServiceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(text);
                    WriteDebug(DateTime.Now.ToString() + " ServiceManagerObj mode: " + wUApiServiceManagerObj.Mode);
                    if (wUApiServiceManagerObj.Status) {
                        ServiceManagerObj = (UpdateServiceManager)wUApiServiceManagerObj.Object;
                        try {
                            WUToolsObj.RestartService(text);
                            ServiceManagerObj.RemoveService(ServiceID);
                            WUToolsObj.RestartService(text);
                        } catch (COMException ex) {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                            if (wUApiCodeDetails != null) {
                                var codeType = wUApiCodeDetails.CodeType;
                                var num = codeType;
                                if (num == 2) {
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                }
                            } else if (MyInvocation.BoundParameters.ContainsKey("Debuger")) {
                                var errorRecord = new ErrorRecord(ex, "Debug: " + ex.ErrorCode, ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord);
                            } else {
                                WriteError(new ErrorRecord(new Exception(text + ": Unknown failure. Try debug."), "Unknown", ErrorCategory.CloseError, null));
                            }
                        }
                    } else if ((bool)Debuger) {
                        var errorRecord2 = new ErrorRecord(wUApiServiceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord2);
                    } else {
                        var error = wUApiServiceManagerObj.Error;
                        WriteError(error);
                    }

                    continue;
                }

                var errorRecord3 = WUToolsObj.CheckPSWUModule(text);
                if (errorRecord3 != null) {
                    WriteError(errorRecord3);
                    continue;
                }

                var text2 = "Remove-WUServiceManager -ServiceID " + ServiceID + " -Verbose -Confirm:$false *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { text };
                if (Credential != null) {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = text2;
                if (ScheduleJob != DateTime.MinValue) {
                    invokeWUJob.TriggerDate = ScheduleJob;
                    WriteVerbose("Invoke-WUJob: " + text + " (" + ScheduleJob.ToString() + "): ");
                } else {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: " + text + " (Now): ");
                }

                WriteVerbose("powershell.exe -Command \"" + text2 + "\"");
                var enumerable = invokeWUJob.Invoke();
                foreach (var item in enumerable) {
                    WriteObject(item);
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
            WriteDebug(DateTime.Now.ToString() + " CmdletEnd");
        }

        protected override void StopProcessing() {
            base.StopProcessing();
        }
    }
}