using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using System.Text.RegularExpressions;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUHistory", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(History) })]
    public class GetWUHistory : PSCmdlet {
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
        public DateTime MaxDate { get; set; }

        [Parameter]
        [ValidateIntGt0]
        public int Last { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateSession UpdateSessionObj { get; set; }

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
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "OperationName",
                        "Result",
                        "Date",
                        "Title"
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
            var computerName = ComputerName;
            foreach (var text in computerName) {
                WriteDebug(DateTime.Now.ToString() + " " + text + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(text);
                    WriteDebug(DateTime.Now.ToString() + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now.ToString() + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (!ShouldProcess(text, "(" + DateTime.Now.ToString() + ") Get Windows Update History")) {
                    continue;
                }

                var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(text);
                WriteDebug(DateTime.Now.ToString() + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                if (wUApiUpdateSessionObj.Status) {
                    UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                    var updateSearcher = UpdateSessionObj.CreateUpdateSearcher();
                    var collection = new Collection<PSObject>();
                    var num = 0;
                    var flag = false;
                    do {
                        try {
                            var updateHistoryEntryCollection = updateSearcher.QueryHistory(num, 1);
                            foreach (IUpdateHistoryEntry item in updateHistoryEntryCollection) {
                                if (item.Date.Year > 1900 && item.Date > MaxDate) {
                                    var pSObject = new PSObject(item);
                                    var value = "";
                                    switch (item.Operation) {
                                        case tagUpdateOperation.uoInstallation:
                                            value = "Installation";
                                            break;
                                        case tagUpdateOperation.uoUninstallation:
                                            value = "Uninstallation";
                                            break;
                                    }

                                    var value2 = "";
                                    switch (item.ResultCode) {
                                        case OperationResultCode.orcNotStarted:
                                            value2 = "NotStarted";
                                            break;
                                        case OperationResultCode.orcInProgress:
                                            value2 = "InProgress";
                                            break;
                                        case OperationResultCode.orcSucceeded:
                                            value2 = "Succeeded";
                                            break;
                                        case OperationResultCode.orcSucceededWithErrors:
                                            value2 = "SucceededWithErrors";
                                            break;
                                        case OperationResultCode.orcFailed:
                                            value2 = "Failed";
                                            break;
                                        case OperationResultCode.orcAborted:
                                            value2 = "Aborted ";
                                            break;
                                    }

                                    var value3 = "";
                                    var match = Regex.Match(item.Title, "KB(\\d+)");
                                    if (match.Success) {
                                        value3 = match.Value;
                                    }

                                    pSObject.Properties.Add(new PSNoteProperty("ComputerName", text));
                                    pSObject.Properties.Add(new PSNoteProperty("OperationName", value));
                                    pSObject.Properties.Add(new PSNoteProperty("Date", item.Date.ToLocalTime()));
                                    pSObject.Properties.Add(new PSNoteProperty("Title", item.Title));
                                    pSObject.Properties.Add(new PSNoteProperty("Result", value2));
                                    pSObject.Properties.Add(new PSNoteProperty("KB", value3));
                                    pSObject.TypeNames.Clear();
                                    pSObject.TypeNames.Add("PSWindowsUpdate.History");
                                    collection.Add(pSObject);
                                    WriteObject(pSObject, true);
                                    flag = true;
                                } else {
                                    flag = false;
                                }
                            }
                        } catch {
                            break;
                        }

                        num++;
                        if (Last > 0 && num >= Last) {
                            flag = false;
                        }
                    } while (flag);

                    OutputObj = new Collection<PSObject>(OutputObj.Concat(collection).ToList());
                } else if ((bool)Debuger) {
                    var errorRecord = new ErrorRecord(wUApiUpdateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null);
                    WriteError(errorRecord);
                } else {
                    var error = wUApiUpdateSessionObj.Error;
                    WriteError(error);
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
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "OperationName",
                        "Result",
                        "Date",
                        "Title"
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