using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using Microsoft.Management.Infrastructure;
using WUApiLib;

namespace PSWindowsUpdate {
    
    [Cmdlet("Get", "WindowsUpdate_v2", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "Default", SupportsShouldProcess = true)]
    [OutputType(typeof(WindowsUpdate))]
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

        [ValidateSet("Driver", "Software", IgnoreCase = true)]
        [Parameter]
        public string UpdateType { get; set; }

        [ValidateSet("Installation", "Uninstallation", IgnoreCase = true)]
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

        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified")]
        [Parameter]
        public string[] Severity { get; set; }

        [Parameter]
        public string[] NotCategory { get; set; }

        [Parameter]
        public string[] NotKBArticleID { get; set; }

        [Parameter]
        public string NotTitle { get; set; }

        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified")]
        [Parameter]
        public string[] NotSeverity { get; set; }

        [Alias("Silent")]
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
                    PSWUSettings.Add("Properties", new WindowsUpdate());
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
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " ParameterSetName: " + ParameterSetName);
            if (string.Equals(invocationName, "Install-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
                Install = true;
            }

            if (string.Equals(invocationName, "Download-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Download = true;
            }

            if (string.Equals(invocationName, "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = true;
            }

            if (string.Equals(invocationName, "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(invocationName, "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                Hide = false;
                WithHidden = true;
            }

            string criteria;
            if (Criteria != null) {
                criteria = Criteria;
                WriteDebug(DateTime.Now + " Set pre search criteria: " + criteria);
            } else {
                if (IsInstalled) {
                    criteria = "IsInstalled = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 1");
                } else {
                    criteria = "IsInstalled = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsInstalled = 0");
                }

                if (UpdateType != null) {
                    criteria = criteria + " and Type = '" + UpdateType + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: Type = '" + UpdateType + "'");
                }

                if (IsHidden) {
                    criteria += " and IsHidden = 1";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1");
                } else if (WithHidden) {
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 1 or IsHidden = 0");
                } else {
                    criteria += " and IsHidden = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: IsHidden = 0");
                }

                if (IgnoreRebootRequired) {
                    criteria += " and RebootRequired = 0";
                    WriteDebug(DateTime.Now + " Set pre search criteria: RebootRequired = 0");
                }

                if (DeploymentAction != null) {
                    criteria = criteria + " and DeploymentAction = '" + DeploymentAction + "'";
                    WriteDebug(DateTime.Now + " Set pre search criteria: DeploymentAction = '" + DeploymentAction + "'");
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsAssigned")) {
                    if (IsAssigned) {
                        criteria += " and IsAssigned = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 1");
                    } else {
                        criteria += " and IsAssigned = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsAssigned = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("IsPresent")) {
                    if (IsPresent) {
                        criteria += " and IsPresent = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 1");
                    } else {
                        criteria += " and IsPresent = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: IsPresent = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("BrowseOnly")) {
                    if (BrowseOnly) {
                        criteria += " and BrowseOnly = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 1");
                    } else {
                        criteria += " and BrowseOnly = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: BrowseOnly = 0");
                    }
                }

                if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites")) {
                    if (AutoSelectOnWebSites) {
                        criteria += " and AutoSelectOnWebSites = 1";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 1");
                    } else {
                        criteria += " and AutoSelectOnWebSites = 0";
                        WriteDebug(DateTime.Now + " Set pre search criteria: AutoSelectOnWebSites = 0");
                    }
                }

                if (NotUpdateID != null) {
                    var notUpdateID = NotUpdateID;
                    foreach (var text2 in notUpdateID) {
                        criteria = criteria + " and UpdateID != '" + text2 + "'";
                        WriteDebug(DateTime.Now + " Set pre search criteria: NotUpdateID = " + text2);
                    }
                }

                if (UpdateID != null) {
                    var text3 = criteria;
                    criteria = "";
                    var num = 0;
                    var updateID = UpdateID;
                    foreach (var uID in updateID) {
                        if (num > 0) {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        if (RevisionNumber > 0) {
                            criteria = criteria + "(" + text3 + " and UpdateID = '" + uID + "' and RevisionNumber = " + RevisionNumber + ")";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uID + "and RevisionNumber = " + RevisionNumber);
                        } else {
                            criteria = criteria + "(" + text3 + " and UpdateID = '" + uID + "')";
                            WriteDebug(DateTime.Now + " Set pre search criteria: UpdateID = " + uID);
                        }

                        num++;
                    }
                }

                if (CategoryIDs != null) {
                    var text5 = criteria;
                    criteria = "";
                    var num2 = 0;
                    var categoryIDs = CategoryIDs;
                    foreach (var catIDs in categoryIDs) {
                        if (num2 > 0) {
                            criteria += " or ";
                            WriteDebug(DateTime.Now + " Set pre search criteria: or ");
                        }

                        criteria = criteria + "(" + text5 + " and CategoryIDs contains '" + catIDs + "')";
                        WriteDebug(DateTime.Now + " Set pre search criteria: CategoryIDs = " + catIDs);
                        num2++;
                    }
                }
            }

            WriteDebug(DateTime.Now + " Search criteria is: " + criteria);
            if (ShowPreSearchCriteria) {
                Host.UI.WriteLine(ConsoleColor.Green, Host.UI.RawUI.BackgroundColor, "PreSearchCriteria: " + criteria);
            }

            var computerName = ComputerName;
            foreach (var target in computerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (Download) {
                    var errorRecord = WUToolsObj.CheckPSWUModule(target);
                    if (errorRecord != null) {
                        WriteError(errorRecord);
                        continue;
                    }
                }

                var wUCimSession = new WUCimSession(target);
                CimInstance[] cimInstances;
                try {
                    cimInstances = wUCimSession.ScanForUpdates(criteria);
                } catch (COMException exception) {
                    if (Debuger) {
                        var errorRecord2 = new ErrorRecord(exception, "Debug", ErrorCategory.CloseError, null);
                        ThrowTerminatingError(errorRecord2);
                    }

                    WriteDebug(DateTime.Now + " Skip to next computer");
                    continue;
                }

                var cimInstancesLength = cimInstances.Length;
                WriteVerbose("Found [" + cimInstancesLength + "] Updates in pre search criteria");
                if (cimInstancesLength == 0) {
                    continue;
                }

                var num4 = 0;
                var activityId = 0;
                var activity = "Post search updates for " + target;
                var statusDescription = "[" + num4 + "/" + cimInstancesLength + "]";
                var progressRecord = new ProgressRecord(activityId, activity, statusDescription);
                var collection = new Collection<PSObject>();
                foreach (var cimInstance in cimInstances) {
                    WriteDebug(DateTime.Now + " " + cimInstance.CimInstanceProperties["Title"]);
                    progressRecord.StatusDescription = "[" + num4 + "/" + cimInstancesLength + "] " + cimInstance.CimInstanceProperties["Title"];
                    progressRecord.PercentComplete = num4 * 100 / cimInstancesLength;
                    WriteProgress(progressRecord);
                    num4++;
                    var updateAccess = true;
                    if (KBArticleID != null && updateAccess) {
                        WriteDebug(DateTime.Now + " Set post search criteria: KBArticleID = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), Title, RegexOptions.IgnoreCase)) {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (NotKBArticleID != null && updateAccess) {
                        WriteDebug(DateTime.Now + " Set post search criteria: NotKBArticleID = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["KBArticleID"].ToString(), NotTitle, RegexOptions.IgnoreCase)) {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (Title != null && updateAccess) {
                        WriteDebug(DateTime.Now + " Set post search criteria: Title = " + Title);
                        if (!Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), Title, RegexOptions.IgnoreCase)) {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (NotTitle != null && updateAccess) {
                        WriteDebug(DateTime.Now + " Set post search criteria: NotTitle = " + NotTitle);
                        if (Regex.IsMatch(cimInstance.CimInstanceProperties["Title"].ToString(), NotTitle, RegexOptions.IgnoreCase)) {
                            updateAccess = false;
                            WriteDebug(DateTime.Now + " UpdateAccess: " + updateAccess);
                        }
                    }

                    if (updateAccess) {
                        WriteDebug(DateTime.Now + " Update was not filtered");
                        var pSObject = new PSObject(cimInstance);
                        pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
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
                    WriteDebug(DateTime.Now + " Return update list only");
                    WriteObject(collection);
                    OutputObj = new Collection<PSObject>(collection);
                    continue;
                }

                var totalDownloaded = 0;
                var totalInstalled = 0;
                var text8 = "";
                if (Download || Install) {
                    var num7 = 0;
                    var activityId2 = 1;
                    var activity2 = "Choose updates for " + target;
                    var statusDescription2 = "[" + num7 + "/" + count + "]";
                    var progressRecord2 = new ProgressRecord(activityId2, activity2, statusDescription2);
                    var text9 = "";
                    foreach (var item in collection) {
                        item.Properties.Add(new PSNoteProperty("X", 1));
                        item.TypeNames.Clear();
                        item.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                        var cimInstance2 = (CimInstance)item.BaseObject;
                        progressRecord2.StatusDescription = "[" + num7 + "/" + count + "] " + cimInstance2.CimInstanceProperties["Title"];
                        progressRecord2.PercentComplete = num7 * 100 / count;
                        WriteProgress(progressRecord2);
                        num7++;
                        WriteDebug(DateTime.Now + " Show update to accept: " + cimInstance2.CimInstanceProperties["Title"]);
                        var flag2 = false;
                        flag2 = AcceptAll || (ShouldProcess(target, "(" + DateTime.Now + ") " + cimInstance2.CimInstanceProperties["Title"])
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
                                    new ErrorRecord(new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"] + "; " + ex.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord3);
                            }

                            if (flag2) {
                                text10 += "A";
                                text11 = "Accepted";
                                WriteDebug(DateTime.Now + " " + text11);
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
                                    new ErrorRecord(new Exception("Something goes wrong: " + cimInstance2.CimInstanceProperties["Title"] + "; " + ex2.Message), "Debug",
                                        ErrorCategory.CloseError, null);
                                WriteError(errorRecord4);
                            }

                            text10 += "R";
                            text11 = "Rejected";
                            WriteDebug(DateTime.Now + " " + text11);
                            item.Properties.Add(new PSNoteProperty("ChooseResult", text11));
                            item.Properties.Add(new PSNoteProperty("Result", text11));
                        }
                    }

                    progressRecord2.RecordType = ProgressRecordType.Completed;
                    WriteProgress(progressRecord2);
                    if (ShowPreSearchCriteria) {
                        WriteVerbose("Choosed pre Search Criteria: " + text9);
                    }

                    var totalAccepted = collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                    WriteObject(collection, true);
                    WriteVerbose("Accepted [" + totalAccepted + "] Updates ready to Download");
                    var num11 = 0;
                    var activityId3 = 1;
                    var activity3 = "Download updates for " + target;
                    var statusDescription3 = "[" + num11 + "/" + totalAccepted + "]";
                    var progressRecord3 = new ProgressRecord(activityId3, activity3, statusDescription3);
                    foreach (var item2 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted")) {
                        item2.Properties.Add(new PSNoteProperty("X", 2));
                        var update = (IUpdate)item2.BaseObject;
                        var updateCollection = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        updateCollection.Add(update);
                        progressRecord3.StatusDescription = "[" + num11 + "/" + totalAccepted + "] " + update.Title + " " + item2.Properties["Size"].Value;
                        progressRecord3.PercentComplete = num11 * 100 / totalAccepted;
                        WriteProgress(progressRecord3);
                        num11++;
                        WriteDebug(DateTime.Now + " Show update to download: " + update.Title);
                        IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                        updateDownloader.Updates = updateCollection;
                        if (ForceDownload) {
                            updateDownloader.IsForced = true;
                        }

                        IDownloadResult downloadResult;
                        try {
                            downloadResult = updateDownloader.Download();
                            WriteDebug(DateTime.Now + " Downloaded");
                        } catch (COMException ex3) {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex3.ErrorCode);
                            var flag3 = false;
                            if (wUApiCodeDetails != null) {
                                switch (wUApiCodeDetails.CodeType) {
                                    case 2: // WUTools.CodeType.Error
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                        flag3 = true;
                                        break;
                                    case 3:
                                        WriteWarning(wUApiCodeDetails.HResult + ": " + wUApiCodeDetails.Description);
                                        break;
                                }

                                if (flag3) {
                                    WriteDebug(DateTime.Now + " Skip to next computer");
                                    break;
                                }
                            } else if (Debuger) { // Debuger
                                var errorRecord5 = new ErrorRecord(ex3, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord5);
                            }

                            WriteDebug(DateTime.Now + " Skip to next update");
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
                    totalDownloaded = collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                    WriteVerbose("Downloaded [" + totalDownloaded + "] Updates ready to Install");
                    if (!Install) {
                        WriteDebug(DateTime.Now + " Return downloaded update list");
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
                var activity4 = "Install updates for " + target;
                var statusDescription4 = "[" + num12 + "/" + totalDownloaded + "]";
                var progressRecord4 = new ProgressRecord(activityId4, activity4, statusDescription4);
                foreach (var item3 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded")) {
                    item3.Properties.Add(new PSNoteProperty("X", 3));
                    var update2 = (IUpdate)item3.BaseObject;
                    var updateCollection2 = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                    updateCollection2.Add(update2);
                    progressRecord4.StatusDescription = "[" + num12 + "/" + totalDownloaded + "] " + update2.Title + " " + item3.Properties["Size"].Value;
                    progressRecord4.PercentComplete = num12 * 100 / totalDownloaded;
                    WriteProgress(progressRecord4);
                    num12++;
                    WriteDebug(DateTime.Now + " Show update to install: " + update2.Title);
                    var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                    updateInstaller.Updates = updateCollection2;
                    if (ForceInstall) {
                        updateInstaller.IsForced = true;
                    }

                    IInstallationResult installationResult;
                    try {
                        installationResult = updateInstaller.Install();
                        WriteDebug(DateTime.Now + " Installed");
                    } catch (COMException ex4) {
                        var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex4.ErrorCode);
                        var skip = false;
                        if (wUApiCodeDetails2 != null) {
                            switch (wUApiCodeDetails2.CodeType) {
                                case 2:
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult, ErrorCategory.CloseError, null));
                                    skip = true;
                                    break;
                                case 3:
                                    WriteWarning(wUApiCodeDetails2.HResult + ": " + wUApiCodeDetails2.Description);
                                    break;
                            }

                            if (skip) {
                                WriteDebug(DateTime.Now + " Skip to next computer");
                                break;
                            }
                        } else if (Debuger) { // Debuger
                            var errorRecord6 = new ErrorRecord(ex4, "Debug", ErrorCategory.CloseError, null);
                            ThrowTerminatingError(errorRecord6);
                        }

                        WriteDebug(DateTime.Now + " Skip to next update");
                        continue;
                    }

                    if (!NeedsReboot) {
                        NeedsReboot = installationResult.RebootRequired;
                        WriteDebug(DateTime.Now + " Reboot is required");
                    }

                    var installResult = "";
                    switch (installationResult.ResultCode) {
                        case OperationResultCode.orcNotStarted:
                            installResult = "NotStarted";
                            break;
                        case OperationResultCode.orcInProgress:
                            installResult = "InProgress";
                            break;
                        case OperationResultCode.orcSucceeded:
                            installResult = "Installed";
                            break;
                        case OperationResultCode.orcSucceededWithErrors:
                            installResult = "InstalledWithErrors";
                            break;
                        case OperationResultCode.orcFailed:
                            installResult = "Failed";
                            break;
                        case OperationResultCode.orcAborted:
                            installResult = "Aborted";
                            break;
                    }

                    item3.Properties.Add(new PSNoteProperty("InstallResult", installResult));
                    item3.Properties.Add(new PSNoteProperty("Result", installResult));
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
                totalInstalled = collection.Where(x => x.Properties["Result"].Value.ToString() == "Installed").Count();
                WriteVerbose("Installed [" + totalInstalled + "] Updates");
                WriteDebug(DateTime.Now + " Return installed update list");
                OutputObj = new Collection<PSObject>(collection.ToList());
                if (totalInstalled <= 0 || !SendHistory) {
                    continue;
                }

                var pCommand = "Get-WUHistory -SendReport -Last " + totalInstalled + " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { target };
                if (Credential != null) {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = pCommand;
                invokeWUJob.TaskName = "PSWindowsUpdate_History";
                if (NeedsReboot) {
                    invokeWUJob.TriggerAtStart = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (AtStart): powershell.exe -Command \"" + pCommand + "\"");
                } else {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (Now): powershell.exe -Command \"" + pCommand + "\"");
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
            if (SendReport && OutputObj.Count > 0) {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new WindowsUpdate());
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            if (NeedsReboot) {
                if (ScheduleReboot != DateTime.MinValue) {
                    WriteDebug(DateTime.Now + " Schedule Reboot " + ScheduleReboot);
                    WriteVerbose(WUToolsObj.ScheduleReboot("localhost", ScheduleReboot));
                } else if (AutoReboot) {
                    WriteDebug(DateTime.Now + " Auto Reboot");
                    WriteVerbose(WUToolsObj.RunReboot("localhost"));
                } else if (IgnoreReboot) {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required, but do it manually.");
                } else {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required. Do it now? [Y / N] (default is 'N')");
                    if (Console.ReadLine().ToUpper() == "Y") {
                        WriteDebug(DateTime.Now + " Manually Reboot");
                        WriteVerbose(WUToolsObj.RunReboot("localhost"));
                    }
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}