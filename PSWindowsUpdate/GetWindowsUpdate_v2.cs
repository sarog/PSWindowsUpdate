using Microsoft.Management.Infrastructure;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WindowsUpdate_v2", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "Default", SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(WindowsUpdate) })]
    public class GetWindowsUpdate_v2 : PSCmdlet {
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
        public SwitchParameter SendHistory { get; set; }

        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleJob { get; set; }

        [Parameter]
        public SwitchParameter AcceptAll { get; set; }

        [Parameter]
        [ValidateRecurseCycle]
        public int RecurseCycle { get; set; }

        [Parameter]
        public SwitchParameter Hide { get; set; }

        [Parameter]
        public SwitchParameter Download { get; set; }

        [Parameter]
        public SwitchParameter ForceDownload { get; set; }

        [Parameter]
        public SwitchParameter Install { get; set; }

        [Parameter]
        public SwitchParameter ForceInstall { get; set; }

        [Parameter]
        public SwitchParameter AutoReboot { get; set; }

        [Parameter]
        public SwitchParameter IgnoreReboot { get; set; }

        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleReboot { get; set; }

        [Parameter(ParameterSetName = "ServiceID")]
        public string ServiceID { get; set; }

        [Parameter(ParameterSetName = "WindowsUpdate")]
        public SwitchParameter WindowsUpdate { get; set; }

        [Parameter(ParameterSetName = "MicrosoftUpdate")]
        public SwitchParameter MicrosoftUpdate { get; set; }

        [Parameter]
        public string Criteria { get; set; }

        [ValidateSet(new string[] { "Driver", "Software" }, IgnoreCase = true)]
        [Parameter]
        public string UpdateType { get; set; }

        [ValidateSet(new string[] { "Installation", "Uninstallation" }, IgnoreCase = true)]
        [Parameter]
        public string DeploymentAction { get; set; }

        [Parameter]
        public SwitchParameter IsAssigned { get; set; }

        [Parameter]
        public SwitchParameter IsPresent { get; set; }

        [Parameter]
        public SwitchParameter BrowseOnly { get; set; }

        [Parameter]
        public SwitchParameter AutoSelectOnWebSites { get; set; }

        [Parameter]
        public string[] UpdateID { get; set; }

        [Parameter]
        public string[] NotUpdateID { get; set; }

        [Parameter]
        public int RevisionNumber { get; set; }

        [Parameter]
        public string[] CategoryIDs { get; set; }

        [Parameter]
        public SwitchParameter IsInstalled { get; set; }

        [Parameter]
        public SwitchParameter IsHidden { get; set; }

        [Parameter]
        public SwitchParameter WithHidden { get; set; }

        [Parameter]
        public SwitchParameter ShowPreSearchCriteria { get; set; }

        [Parameter]
        public string[] Category { get; set; }

        [Parameter]
        public string[] KBArticleID { get; set; }

        [Parameter]
        public string Title { get; set; }

        [ValidateSet(new string[] { "Critical", "Important", "Moderate", "Low", "Unspecified" })]
        [Parameter]
        public string[] Severity { get; set; }

        [Parameter]
        public string[] NotCategory { get; set; }

        [Parameter]
        public string[] NotKBArticleID { get; set; }

        [Parameter]
        public string NotTitle { get; set; }

        [ValidateSet(new string[] { "Critical", "Important", "Moderate", "Low", "Unspecified" })]
        [Parameter]
        public string[] NotSeverity { get; set; }

        [Alias(new string[] { "Silent" })]
        [Parameter]
        public SwitchParameter IgnoreUserInput { get; set; }

        [Parameter]
        public SwitchParameter IgnoreRebootRequired { get; set; }

        [Parameter]
        public SwitchParameter AutoSelectOnly { get; set; }

        [Parameter]
        public long MaxSize { get; set; }

        [Parameter]
        public long MinSize { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateSession UpdateSessionObj { get; set; }

        private static IUpdateSearcher SearcherObj { get; set; }

        private static UpdateServiceManager ServiceManagerObj { get; set; }

        private static bool NeedsReboot { get; set; }

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
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
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
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now.ToString() + " ParameterSetName: " + ParameterSetName);
            if (string.Equals(invocationName.ToString(), "Install-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
                Install = true;
            }

            if (string.Equals(invocationName.ToString(), "Download-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
            }

            if (string.Equals(invocationName.ToString(), "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = true;
            }

            if (string.Equals(invocationName.ToString(), "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(invocationName.ToString(), "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = false;
                WithHidden = true;
            }

            string text;
            if (Criteria != null) {
                text = Criteria;
                WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: " + text);
            } else {
                if ((bool)IsInstalled) {
                    text = "IsInstalled = 1";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsInstalled = 1");
                } else {
                    text = "IsInstalled = 0";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsInstalled = 0");
                }

                if (UpdateType != null) {
                    text = text + " and Type = '" + UpdateType + "'";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: Type = '" + UpdateType + "'");
                }

                if ((bool)IsHidden) {
                    text += " and IsHidden = 1";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsHidden = 1");
                } else if ((bool)WithHidden) {
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsHidden = 1 or IsHidden = 0");
                } else {
                    text += " and IsHidden = 0";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsHidden = 0");
                }

                if ((bool)IgnoreRebootRequired) {
                    text += " and RebootRequired = 0";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: RebootRequired = 0");
                }

                if (DeploymentAction != null) {
                    text = text + " and DeploymentAction = '" + DeploymentAction + "'";
                    WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: DeploymentAction = '" + DeploymentAction + "'");
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsAssigned")) {
                    if ((bool)IsAssigned) {
                        text += " and IsAssigned = 1";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsAssigned = 1");
                    } else {
                        text += " and IsAssigned = 0";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsAssigned = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsPresent")) {
                    if ((bool)IsPresent) {
                        text += " and IsPresent = 1";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsPresent = 1");
                    } else {
                        text += " and IsPresent = 0";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: IsPresent = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("BrowseOnly")) {
                    if ((bool)BrowseOnly) {
                        text += " and BrowseOnly = 1";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: BrowseOnly = 1");
                    } else {
                        text += " and BrowseOnly = 0";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: BrowseOnly = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites")) {
                    if ((bool)AutoSelectOnWebSites) {
                        text += " and AutoSelectOnWebSites = 1";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: AutoSelectOnWebSites = 1");
                    } else {
                        text += " and AutoSelectOnWebSites = 0";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: AutoSelectOnWebSites = 0");
                    }
                }

                if (NotUpdateID != null) {
                    var notUpdateID = NotUpdateID;
                    foreach (var text2 in notUpdateID) {
                        text = text + " and UpdateID != '" + text2 + "'";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: NotUpdateID = " + text2);
                    }
                }

                if (UpdateID != null) {
                    var text3 = text;
                    text = "";
                    var num = 0;
                    var updateID = UpdateID;
                    foreach (var text4 in updateID) {
                        if (num > 0) {
                            text += " or ";
                            WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: or ");
                        }

                        if (RevisionNumber > 0) {
                            text = text + "(" + text3 + " and UpdateID = '" + text4 + "' and RevisionNumber = " + RevisionNumber + ")";
                            WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: UpdateID = " + text4 + "and RevisionNumber = " + RevisionNumber);
                        } else {
                            text = text + "(" + text3 + " and UpdateID = '" + text4 + "')";
                            WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: UpdateID = " + text4);
                        }

                        num++;
                    }
                }

                if (CategoryIDs != null) {
                    var text5 = text;
                    text = "";
                    var num2 = 0;
                    var categoryIDs = CategoryIDs;
                    foreach (var text6 in categoryIDs) {
                        if (num2 > 0) {
                            text += " or ";
                            WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: or ");
                        }

                        text = text + "(" + text5 + " and CategoryIDs contains '" + text6 + "')";
                        WriteDebug(DateTime.Now.ToString() + " Set pre search criteria: CategoryIDs = " + text6);
                        num2++;
                    }
                }
            }

            WriteDebug(DateTime.Now.ToString() + " Search criteria is: " + text);
            if ((bool)ShowPreSearchCriteria) {
                Host.UI.WriteLine(ConsoleColor.Green, Host.UI.RawUI.BackgroundColor, "PreSearchCriteria: " + text);
            }

            var computerName = ComputerName;
            foreach (var text7 in computerName) {
                WriteDebug(DateTime.Now.ToString() + " " + text7 + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(text7);
                    WriteDebug(DateTime.Now.ToString() + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now.ToString() + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if ((bool)Download) {
                    var errorRecord = WUToolsObj.CheckPSWUModule(text7);
                    if (errorRecord != null) {
                        WriteError(errorRecord);
                        continue;
                    }
                }

                var wUCimSession = new WUCimSession(text7);
                CimInstance[] array;
                try {
                    array = wUCimSession.ScanForUpdates(text);
                } catch (COMException exception) {
                    if (MyInvocation.BoundParameters.ContainsKey("Debuger")) {
                        var errorRecord2 = new ErrorRecord(exception, "Debug", ErrorCategory.CloseError, null);
                        ThrowTerminatingError(errorRecord2);
                    }

                    WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                    continue;
                }

                var num3 = array.Length;
                WriteVerbose("Found [" + num3 + "] Updates in pre search criteria");
                if (num3 == 0) {
                    continue;
                }

                var num4 = 0;
                var activityId = 0;
                var activity = "Post search updates for " + text7;
                var statusDescription = "[" + num4 + "/" + num3 + "]";
                var progressRecord = new ProgressRecord(activityId, activity, statusDescription);
                var collection = new Collection<PSObject>();
                var array2 = array;
                foreach (var cimInstance in array2) {
                    WriteDebug(DateTime.Now.ToString() + " " + cimInstance.CimInstanceProperties["Title"].ToString());
                    progressRecord.StatusDescription = "[" + num4 + "/" + num3 + "] " + cimInstance.CimInstanceProperties["Title"].ToString();
                    progressRecord.PercentComplete = num4 * 100 / num3;
                    WriteProgress(progressRecord);
                    num4++;
                    var flag = true;
                    if (KBArticleID != null && flag) {
                        WriteDebug(DateTime.Now.ToString() + " Set post search criteria: KBArticleID = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), Title, RegexOptions.IgnoreCase)) {
                            flag = false;
                            WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag);
                        }
                    }

                    if (NotKBArticleID != null && flag) {
                        WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotKBArticleID = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), NotTitle, RegexOptions.IgnoreCase)) {
                            flag = false;
                            WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag);
                        }
                    }

                    if (Title != null && flag) {
                        WriteDebug(DateTime.Now.ToString() + " Set post search criteria: Title = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), Title, RegexOptions.IgnoreCase)) {
                            flag = false;
                            WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag);
                        }
                    }

                    if (NotTitle != null && flag) {
                        WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotTitle = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), NotTitle, RegexOptions.IgnoreCase)) {
                            flag = false;
                            WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag);
                        }
                    }

                    if (flag) {
                        WriteDebug(DateTime.Now.ToString() + " Update was not filtered");
                        var pSObject = new PSObject(cimInstance);
                        pSObject.Properties.Add(new PSNoteProperty("ComputerName", text7));
                        pSObject.Properties.Add(new PSNoteProperty("BaseObject", cimInstance));
                        pSObject.TypeNames.Clear();
                        pSObject.TypeNames.Add("PSWindowsUpdate.WindowsUpdate");
                        collection.Add(pSObject);
                    }
                }

                progressRecord.RecordType = ProgressRecordType.Completed;
                WriteProgress(progressRecord);
                var count = collection.Count;
                WriteVerbose("Found [" + count + "] Updates in post search criteria");
                if (!Download && !Install) {
                    WriteDebug(DateTime.Now.ToString() + " Return update list only");
                    WriteObject(collection);
                    OutputObj = new Collection<PSObject>(collection);
                    continue;
                }

                var num5 = 0;
                var num6 = 0;
                var text8 = "";
                if ((bool)Download || (bool)Install) {
                    var num7 = 0;
                    var activityId2 = 1;
                    var activity2 = "Choose updates for " + text7;
                    var statusDescription2 = "[" + num7 + "/" + count + "]";
                    var progressRecord2 = new ProgressRecord(activityId2, activity2, statusDescription2);
                    var text9 = "";
                    foreach (var item in collection) {
                        item.Properties.Add(new PSNoteProperty("X", 1));
                        item.TypeNames.Clear();
                        item.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                        var cimInstance2 = (CimInstance)item.BaseObject;
                        progressRecord2.StatusDescription = "[" + num7 + "/" + count + "] " + cimInstance2.CimInstanceProperties["Title"].ToString();
                        progressRecord2.PercentComplete = num7 * 100 / count;
                        WriteProgress(progressRecord2);
                        num7++;
                        WriteDebug(DateTime.Now.ToString() + " Show update to accept: " + cimInstance2.CimInstanceProperties["Title"].ToString());
                        var flag2 = false;
                        flag2 = (bool)AcceptAll || (ShouldProcess(text7, "(" + DateTime.Now.ToString() + ") " + cimInstance2.CimInstanceProperties["Title"].ToString())
                            ? true
                            : false);
                        var text10 = "";
                        var text11 = "";
                        if (flag2) {
                            try {
                                var pSObject2 = new PSObject(item.Properties["Identity"].Value);
                                var text12 = (string)pSObject2.Properties["UpdateID"].Value;
                                var num8 = (int)pSObject2.Properties["RevisionNumber"].Value;
                                text9 = !(text9 == "")
                                    ? text9 + " or (UpdateID = '" + text12 + "' and RevisionNumber = " + num8 + ")"
                                    : "(UpdateID = '" + text12 + "' and RevisionNumber = " + num8 + ")";
                            } catch (Exception ex) {
                                flag2 = false;
                                var errorRecord3 =
                                    new ErrorRecord(new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"].ToString() + "; " + ex.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord3);
                            }

                            if (flag2) {
                                text10 += "A";
                                text11 = "Accepted";
                                WriteDebug(DateTime.Now.ToString() + " " + text11);
                                item.Properties.Add(new PSNoteProperty("ChooseResult", text11));
                                item.Properties.Add(new PSNoteProperty("Result", text11));
                            }
                        }

                        if (!flag2) {
                            try {
                                var pSObject3 = new PSObject(item.Properties["Identity"].Value);
                                var text13 = (string)pSObject3.Properties["UpdateID"].Value;
                                var num9 = (int)pSObject3.Properties["RevisionNumber"].Value;
                                text8 = !(text8 == "") ? text8 + ",'" + text13 + "'" : "'" + text13 + "'";
                            } catch (Exception ex2) {
                                flag2 = false;
                                var errorRecord4 =
                                    new ErrorRecord(new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"].ToString() + "; " + ex2.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord4);
                            }

                            text10 += "R";
                            text11 = "Rejected";
                            WriteDebug(DateTime.Now.ToString() + " " + text11);
                            item.Properties.Add(new PSNoteProperty("ChooseResult", text11));
                            item.Properties.Add(new PSNoteProperty("Result", text11));
                        }
                    }

                    progressRecord2.RecordType = ProgressRecordType.Completed;
                    WriteProgress(progressRecord2);
                    if ((bool)ShowPreSearchCriteria) {
                        WriteVerbose("Choosed pre Search Criteria: " + text9);
                    }

                    var num10 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                    WriteObject(collection, true);
                    WriteVerbose("Accepted [" + num10 + "] Updates ready to Download");
                    var num11 = 0;
                    var activityId3 = 1;
                    var activity3 = "Download updates for " + text7;
                    var statusDescription3 = "[" + num11 + "/" + num10 + "]";
                    var progressRecord3 = new ProgressRecord(activityId3, activity3, statusDescription3);
                    foreach (var item2 in collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Accepted")) {
                        item2.Properties.Add(new PSNoteProperty("X", 2));
                        var update = (IUpdate)item2.BaseObject;
                        var updateCollection = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        updateCollection.Add(update);
                        progressRecord3.StatusDescription = "[" + num11 + "/" + num10 + "] " + update.Title + " " + item2.Properties["Size"].Value.ToString();
                        progressRecord3.PercentComplete = num11 * 100 / num10;
                        WriteProgress(progressRecord3);
                        num11++;
                        WriteDebug(DateTime.Now.ToString() + " Show update to download: " + update.Title);
                        IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                        updateDownloader.Updates = updateCollection;
                        if ((bool)ForceDownload) {
                            updateDownloader.IsForced = true;
                        }

                        IDownloadResult downloadResult;
                        try {
                            downloadResult = updateDownloader.Download();
                            WriteDebug(DateTime.Now.ToString() + " Downloaded");
                        } catch (COMException ex3) {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex3.ErrorCode);
                            var flag3 = false;
                            if (wUApiCodeDetails != null) {
                                switch (wUApiCodeDetails.CodeType) {
                                    case 2:
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                        flag3 = true;
                                        break;
                                    case 3:
                                        WriteWarning(wUApiCodeDetails.HResult + ": " + wUApiCodeDetails.Description);
                                        break;
                                }

                                if (flag3) {
                                    WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                    break;
                                }
                            } else if (MyInvocation.BoundParameters.ContainsKey("Debuger")) {
                                var errorRecord5 = new ErrorRecord(ex3, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord5);
                            }

                            WriteDebug(DateTime.Now.ToString() + " Skip to next update");
                            continue;
                        }

                        var value = "";
                        switch (downloadResult.ResultCode) {
                            case OperationResultCode.orcNotStarted:
                                value = "NotStarted";
                                break;
                            case OperationResultCode.orcInProgress:
                                value = "InProgress";
                                break;
                            case OperationResultCode.orcSucceeded:
                                value = "Downloaded";
                                break;
                            case OperationResultCode.orcSucceededWithErrors:
                                value = "DownloadedWithErrors";
                                break;
                            case OperationResultCode.orcFailed:
                                value = "Failed";
                                break;
                            case OperationResultCode.orcAborted:
                                value = "Aborted";
                                break;
                        }

                        item2.Properties.Add(new PSNoteProperty("DownloadResult", value));
                        item2.Properties.Add(new PSNoteProperty("Result", value));
                        var text14 = "";
                        text14 = !(item2.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text14 + "R" : text14 + "A";
                        text14 = !(item2.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text14 + "F" : text14 + "D";
                        text14 = !update.IsInstalled ? text14 + "-" : text14 + "I";
                        text14 = !update.IsMandatory ? text14 + "-" : text14 + "M";
                        text14 = !update.IsHidden ? text14 + "-" : text14 + "H";
                        text14 = !update.IsUninstallable ? text14 + "-" : text14 + "U";
                        text14 = !update.IsBeta ? text14 + "-" : text14 + "B";
                        item2.Properties["Status"].Value = text14;
                        WriteObject(item2, true);
                    }

                    progressRecord3.RecordType = ProgressRecordType.Completed;
                    WriteProgress(progressRecord3);
                    num5 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                    WriteVerbose("Downloaded [" + num5 + "] Updates ready to Install");
                    if (!Install) {
                        WriteDebug(DateTime.Now.ToString() + " Return downloaded update list");
                        OutputObj = new Collection<PSObject>(collection.ToList());
                        continue;
                    }
                }

                if (!Install) {
                    continue;
                }

                NeedsReboot = false;
                var num12 = 0;
                var activityId4 = 1;
                var activity4 = "Install updates for " + text7;
                var statusDescription4 = "[" + num12 + "/" + num5 + "]";
                var progressRecord4 = new ProgressRecord(activityId4, activity4, statusDescription4);
                foreach (var item3 in collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Downloaded")) {
                    item3.Properties.Add(new PSNoteProperty("X", 3));
                    var update2 = (IUpdate)item3.BaseObject;
                    var updateCollection2 = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                    updateCollection2.Add(update2);
                    progressRecord4.StatusDescription = "[" + num12 + "/" + num5 + "] " + update2.Title + " " + item3.Properties["Size"].Value.ToString();
                    progressRecord4.PercentComplete = num12 * 100 / num5;
                    WriteProgress(progressRecord4);
                    num12++;
                    WriteDebug(DateTime.Now.ToString() + " Show update to install: " + update2.Title);
                    var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                    updateInstaller.Updates = updateCollection2;
                    if ((bool)ForceInstall) {
                        updateInstaller.IsForced = true;
                    }

                    IInstallationResult installationResult;
                    try {
                        installationResult = updateInstaller.Install();
                        WriteDebug(DateTime.Now.ToString() + " Installed");
                    } catch (COMException ex4) {
                        var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex4.ErrorCode);
                        var flag4 = false;
                        if (wUApiCodeDetails2 != null) {
                            switch (wUApiCodeDetails2.CodeType) {
                                case 2:
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult, ErrorCategory.CloseError, null));
                                    flag4 = true;
                                    break;
                                case 3:
                                    WriteWarning(wUApiCodeDetails2.HResult + ": " + wUApiCodeDetails2.Description);
                                    break;
                            }

                            if (flag4) {
                                WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                break;
                            }
                        } else if (MyInvocation.BoundParameters.ContainsKey("Debuger")) {
                            var errorRecord6 = new ErrorRecord(ex4, "Debug", ErrorCategory.CloseError, null);
                            ThrowTerminatingError(errorRecord6);
                        }

                        WriteDebug(DateTime.Now.ToString() + " Skip to next update");
                        continue;
                    }

                    if (!NeedsReboot) {
                        NeedsReboot = installationResult.RebootRequired;
                        WriteDebug(DateTime.Now.ToString() + " Reboot is required");
                    }

                    var value2 = "";
                    switch (installationResult.ResultCode) {
                        case OperationResultCode.orcNotStarted:
                            value2 = "NotStarted";
                            break;
                        case OperationResultCode.orcInProgress:
                            value2 = "InProgress";
                            break;
                        case OperationResultCode.orcSucceeded:
                            value2 = "Installed";
                            break;
                        case OperationResultCode.orcSucceededWithErrors:
                            value2 = "InstalledWithErrors";
                            break;
                        case OperationResultCode.orcFailed:
                            value2 = "Failed";
                            break;
                        case OperationResultCode.orcAborted:
                            value2 = "Aborted";
                            break;
                    }

                    item3.Properties.Add(new PSNoteProperty("InstallResult", value2));
                    item3.Properties.Add(new PSNoteProperty("Result", value2));
                    var text15 = "";
                    text15 = !(item3.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text15 + "R" : text15 + "A";
                    text15 = !(item3.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text15 + "F" : text15 + "D";
                    text15 = !(item3.Properties["InstallResult"].Value.ToString() == "Installed")
                        ? text15 + "F"
                        : !installationResult.RebootRequired
                            ? text15 + "I"
                            : text15 + "R";
                    text15 = !update2.IsMandatory ? text15 + "-" : text15 + "M";
                    text15 = !update2.IsHidden ? text15 + "-" : text15 + "H";
                    text15 = !update2.IsUninstallable ? text15 + "-" : text15 + "U";
                    text15 = !update2.IsBeta ? text15 + "-" : text15 + "B";
                    item3.Properties["Status"].Value = text15;
                    WriteObject(item3, true);
                }

                progressRecord4.RecordType = ProgressRecordType.Completed;
                WriteProgress(progressRecord4);
                num6 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Installed").Count();
                WriteVerbose("Installed [" + num6 + "] Updates");
                WriteDebug(DateTime.Now.ToString() + " Return installed update list");
                OutputObj = new Collection<PSObject>(collection.ToList());
                if (num6 <= 0 || !SendHistory) {
                    continue;
                }

                var text16 = "Get-WUHistory -SendReport -Last " + num6 + " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { text7 };
                if (Credential != null) {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = text16;
                invokeWUJob.TaskName = "PSWindowsUpdate_History";
                if (NeedsReboot) {
                    invokeWUJob.TriggerAtStart = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + text7 + " (AtStart): powershell.exe -Command \"" + text16 + "\"");
                } else {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + text7 + " (Now): powershell.exe -Command \"" + text16 + "\"");
                }

                var enumerable = invokeWUJob.Invoke();
                foreach (var item4 in enumerable) {
                    WriteObject(item4);
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
                    logonType = WUImpersonator.LogonSessionType.Interactive;
                    logonProvider = WUImpersonator.LogonProvider.Default;
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
            if ((bool)SendReport && OutputObj.Count > 0) {
                WriteDebug(DateTime.Now.ToString() + " Send report");
                if (!PSWUSettings.ContainsKey((object)"Properties")) {
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "Status",
                        "KB",
                        "Size",
                        "Title"
                    });
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            if (NeedsReboot) {
                if (ScheduleReboot != DateTime.MinValue) {
                    WriteDebug(DateTime.Now.ToString() + " Schedule Reboot " + ScheduleReboot.ToString());
                    WriteVerbose(WUToolsObj.ScheduleReboot("localhost", ScheduleReboot));
                } else if ((bool)AutoReboot) {
                    WriteDebug(DateTime.Now.ToString() + " Auto Reboot");
                    WriteVerbose(WUToolsObj.RunReboot("localhost"));
                } else if ((bool)IgnoreReboot) {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required, but do it manually.");
                } else {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required. Do it now? [Y / N] (default is 'N')");
                    if (Console.ReadLine().ToUpper() == "Y") {
                        WriteDebug(DateTime.Now.ToString() + " Manually Reboot");
                        WriteVerbose(WUToolsObj.RunReboot("localhost"));
                    }
                }
            }

            WriteDebug(DateTime.Now.ToString() + " CmdletEnd");
        }

        protected override void StopProcessing() {
            base.StopProcessing();
        }
    }
}