using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using System.Text.RegularExpressions;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUHistory", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(History))]
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
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (SendReport) {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new History());
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
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (!ShouldProcess(target, "(" + DateTime.Now + ") Get Windows Update History")) {
                    continue;
                }

                var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                WriteDebug(DateTime.Now + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                if (wUApiUpdateSessionObj.Status) {
                    UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                    var updateSearcher = UpdateSessionObj.CreateUpdateSearcher();
                    var collection = new Collection<PSObject>();
                    var startIndex = 0;
                    var flag = false;
                    do {
                        try {
                            var historyEntryCollection = updateSearcher.QueryHistory(startIndex, 1);
                            foreach (IUpdateHistoryEntry updateHistoryEntry in historyEntryCollection) {
                                if (updateHistoryEntry.Date.Year > 1900 && updateHistoryEntry.Date > MaxDate) {
                                    var pSObject = new PSObject(updateHistoryEntry);
                                    var opName = "";
                                    switch (updateHistoryEntry.Operation) {
                                        case tagUpdateOperation.uoInstallation:
                                            opName = "Installation";
                                            break;
                                        case tagUpdateOperation.uoUninstallation:
                                            opName = "Uninstallation";
                                            break;
                                    }

                                    var result = "";
                                    switch (updateHistoryEntry.ResultCode) {
                                        case OperationResultCode.orcNotStarted:
                                            result = "NotStarted";
                                            break;
                                        case OperationResultCode.orcInProgress:
                                            result = "InProgress";
                                            break;
                                        case OperationResultCode.orcSucceeded:
                                            result = "Succeeded";
                                            break;
                                        case OperationResultCode.orcSucceededWithErrors:
                                            result = "SucceededWithErrors";
                                            break;
                                        case OperationResultCode.orcFailed:
                                            result = "Failed";
                                            break;
                                        case OperationResultCode.orcAborted:
                                            result = "Aborted ";
                                            break;
                                    }

                                    var kbId = "";
                                    var match = Regex.Match(updateHistoryEntry.Title, "KB(\\d+)");
                                    if (match.Success) {
                                        kbId = match.Value;
                                    }

                                    pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                                    pSObject.Properties.Add(new PSNoteProperty("OperationName", opName));
                                    pSObject.Properties.Add(new PSNoteProperty("Date", updateHistoryEntry.Date.ToLocalTime()));
                                    pSObject.Properties.Add(new PSNoteProperty("Title", updateHistoryEntry.Title));
                                    pSObject.Properties.Add(new PSNoteProperty("Result", result));
                                    pSObject.Properties.Add(new PSNoteProperty("KB", kbId));
                                    pSObject.TypeNames.Clear();
                                    pSObject.TypeNames.Add("PSWindowsUpdate.History");
                                    collection.Add(pSObject);
                                    WriteObject(pSObject, true);
                                    flag = true;
                                } else {
                                    flag = false;
                                }
                                if (historyEntryCollection.Count == 0) {
                                    flag = false;
                                }
                            }
                        } catch {
                            break;
                        }

                        startIndex++;
                        if (Last > 0 && startIndex >= Last) {
                            flag = false;
                        }
                    } while (flag);

                    OutputObj = new Collection<PSObject>(OutputObj.Concat(collection).ToList());
                } else if (Debuger) {
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
                    PSWUSettings.Add("Properties", new History());
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