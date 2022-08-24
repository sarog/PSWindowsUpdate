using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WindowsUpdate", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "Default", SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(WindowsUpdate) })]
    public class GetWindowsUpdate : PSCmdlet {
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

        [ValidateSet(new string[] {
            "Critical Updates", "Definition Updates", "Drivers", "Feature Packs", "Security Updates", "Service Packs", "Tools", "Update Rollups", "Updates", "Upgrades", "Microsoft"
        })]
        [Parameter]
        public string[] RootCategories { get; set; }

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
            foreach (var target in computerName) {
                WriteDebug(DateTime.Now.ToString() + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now.ToString() + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now.ToString() + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if ((bool)Download) {
                    var errorRecord = WUToolsObj.CheckPSWUModule(target);
                    if (errorRecord != null) {
                        WriteError(errorRecord);
                        continue;
                    }
                }

                var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                WriteDebug(DateTime.Now.ToString() + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                if (wUApiUpdateSessionObj.Status) {
                    UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                    SearcherObj = UpdateSessionObj.CreateUpdateSearcher();
                    var wUApiServiceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                    WriteDebug(DateTime.Now.ToString() + " ServiceManagerObj mode: " + wUApiServiceManagerObj.Mode);
                    if (wUApiServiceManagerObj.Status) {
                        ServiceManagerObj = (UpdateServiceManager)wUApiServiceManagerObj.Object;
                        var text8 = "";
                        if ((bool)WindowsUpdate) {
                            WriteDebug(DateTime.Now.ToString() + " Set source of updates to Windows Update");
                            SearcherObj.ServerSelection = ServerSelection.ssWindowsUpdate;
                            text8 = "Windows Update";
                        } else if ((bool)MicrosoftUpdate) {
                            WriteDebug(DateTime.Now.ToString() + " Set source of updates to Microsof Update");
                            ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                            var flag = false;
                            foreach (IUpdateService2 service in ServiceManagerObj.Services) {
                                if (service.ServiceID == ServiceID) {
                                    WriteDebug(DateTime.Now.ToString() + " Search for ServiceID. Set source of updates to " + service.Name);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = service.Name;
                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag) {
                                try {
                                    Thread.Sleep(500);
                                    var serviceController = new ServiceController("Windows Update", target);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                    ServiceManagerObj.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 2, null);
                                    WriteDebug(DateTime.Now.ToString() + " Register WUA Microsof Update service");
                                    Thread.Sleep(500);
                                    serviceController.Stop();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                                    Thread.Sleep(500);
                                    serviceController.Start();
                                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                                } catch (COMException ex) {
                                    var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                    if (wUApiCodeDetails != null) {
                                        switch (wUApiCodeDetails.CodeType) {
                                            case 2:
                                                WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                                break;
                                            case 3:
                                                WriteWarning(wUApiCodeDetails.HResult + ": " + wUApiCodeDetails.Description);
                                                break;
                                        }
                                    } else if (Debuger) {
                                        var errorRecord2 = new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null);
                                        ThrowTerminatingError(errorRecord2);
                                    }
                                }

                                SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                SearcherObj.ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
                                text8 = "Microsoft Update";
                            }
                        } else {
                            try {
                                foreach (IUpdateService2 service2 in ServiceManagerObj.Services) {
                                    if (ServiceID != null) {
                                        if (service2.ServiceID == ServiceID) {
                                            WriteDebug(DateTime.Now.ToString() + " Try ServiceID. Set source of updates to " + service2.Name);
                                            SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                            SearcherObj.ServiceID = ServiceID;
                                            text8 = service2.Name;
                                            break;
                                        }
                                    } else if (service2.IsDefaultAUService) {
                                        WriteDebug(DateTime.Now.ToString() + " Try Default. Set source of updates to " + service2.Name);
                                        text8 = service2.Name;
                                        break;
                                    }
                                }
                            } catch {
                                if (ServiceID != null) {
                                    WriteDebug(DateTime.Now.ToString() + " Catch ServiceID. Set source of updates to " + ServiceID);
                                    SearcherObj.ServerSelection = ServerSelection.ssOthers;
                                    SearcherObj.ServiceID = ServiceID;
                                    text8 = ServiceID;
                                } else {
                                    WriteDebug(DateTime.Now.ToString() + " Catch. Set source of updates to Default");
                                    text8 = "default for " + target + ".";
                                }
                            }
                        }

                        if (text8 == "Windows Server Update Service") {
                            RegistryKey registryKey = null;
                            registryKey = !WUToolsObj.IsLocalHost(target)
                                ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                                : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                            var registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                            var text9 = registryKey2.GetValue("WUServer").ToString();
                            text8 = text8 + " (" + text9 + ")";
                        }

                        WriteVerbose(target + " (" + DateTime.Now.ToString() + "): Connecting to " + text8 + " server. Please wait...");
                        ISearchResult searchResult;
                        try {
                            searchResult = SearcherObj.Search(text);
                        } catch (COMException ex2) {
                            var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex2.ErrorCode);
                            if (wUApiCodeDetails2 != null) {
                                switch (wUApiCodeDetails2.CodeType) {
                                    case 2:
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult, ErrorCategory.CloseError, null));
                                        break;
                                    case 3:
                                        WriteWarning(wUApiCodeDetails2.HResult + ": " + wUApiCodeDetails2.Description);
                                        break;
                                }
                            } else if (Debuger) {
                                var errorRecord3 = new ErrorRecord(ex2, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord3);
                            }

                            WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                            continue;
                        }

                        var count = searchResult.Updates.Count;
                        WriteVerbose("Found [" + count + "] Updates in pre search criteria");
                        if (count == 0) {
                            continue;
                        }

                        var updateCollection = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        if (RootCategories != null) {
                            var updateCollection2 =
                                (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                            var rootCategories = RootCategories;
                            foreach (var text10 in rootCategories) {
                                var index = -1;
                                switch (text10) {
                                    case "Critical Updates":
                                        index = 0;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Definition Updates":
                                        index = 1;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Drivers":
                                        index = 2;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Feature Packs":
                                        index = 3;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Security Updates":
                                        index = 4;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Service Packs":
                                        index = 5;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Tools":
                                        index = 6;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Update Rollups":
                                        index = 7;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Updates":
                                        index = 8;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Upgrades":
                                        index = 9;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                    case "Microsoft":
                                        index = 10;
                                        WriteDebug(DateTime.Now.ToString() + " RootCategory: " + text10 + " - " + index);
                                        break;
                                }

                                try {
                                    foreach (IUpdate update6 in searchResult.RootCategories[index].Updates) {
                                        updateCollection2.Add(update6);
                                    }
                                } catch {
                                    WriteWarning("RootCategiries Updates are empty. Use classic filters.");
                                }
                            }

                            updateCollection = updateCollection2;
                        } else {
                            updateCollection = searchResult.Updates;
                        }

                        var num3 = 0;
                        var activityId = 0;
                        var activity = "Post search updates for " + target;
                        var statusDescription = "[" + num3 + "/" + count + "]";
                        var progressRecord = new ProgressRecord(activityId, activity, statusDescription);
                        var collection = new Collection<PSObject>();
                        foreach (IUpdate item in updateCollection) {
                            WriteDebug(DateTime.Now.ToString() + " " + item.Title);
                            var text11 = WUToolsObj.ConvertSize(item.MaxDownloadSize);
                            progressRecord.StatusDescription = "[" + num3 + "/" + count + "] " + item.Title + " " + text11;
                            progressRecord.PercentComplete = num3 * 100 / count;
                            WriteProgress(progressRecord);
                            num3++;
                            var flag2 = true;
                            if (Category != null || NotCategory != null) {
                                var text12 = "";
                                foreach (ICategory category3 in item.Categories) {
                                    text12 = text12 + category3.Name + "|";
                                }

                                WriteDebug(DateTime.Now.ToString() + " " + text12);
                                if (Category != null) {
                                    WriteDebug(DateTime.Now.ToString() + " Set post search criteria: Categories = " + string.Join(", ", Category));
                                    var category2 = Category;
                                    foreach (var pattern in category2) {
                                        if (!Regex.IsMatch(text12, pattern, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            continue;
                                        }

                                        flag2 = true;
                                        break;
                                    }
                                }

                                if (NotCategory != null && flag2) {
                                    WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotCategories = " + string.Join(", ", NotCategory));
                                    var notCategory = NotCategory;
                                    foreach (var pattern2 in notCategory) {
                                        if (Regex.IsMatch(text12, pattern2, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                }

                                WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag2);
                            }

                            if (KBArticleID != null && flag2) {
                                var text13 = string.Join(", ", KBArticleID);
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: KBArticleIDs = " + text13);
                                if (item.KBArticleIDs.Count > 0) {
                                    foreach (string kBArticleID in item.KBArticleIDs) {
                                        if (!Regex.IsMatch(text13, kBArticleID.ToString(), RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                } else {
                                    flag2 = false;
                                }

                                WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag2);
                            }

                            if (NotKBArticleID != null && flag2) {
                                var text15 = string.Join(", ", NotKBArticleID);
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotKBArticleID = " + text15);
                                if (item.KBArticleIDs.Count > 0) {
                                    foreach (string kBArticleID2 in item.KBArticleIDs) {
                                        if (Regex.IsMatch(text15, kBArticleID2, RegexOptions.IgnoreCase)) {
                                            flag2 = false;
                                            break;
                                        }
                                    }
                                }

                                WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag2);
                            }

                            if (Title != null && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: Title = " + Title);
                                if (!Regex.IsMatch(item.Title.ToString(), Title, RegexOptions.IgnoreCase)) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag2);
                                }
                            }

                            if (NotTitle != null && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotTitle = " + NotTitle);
                                if (Regex.IsMatch(item.Title.ToString(), NotTitle, RegexOptions.IgnoreCase)) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess: " + flag2);
                                }
                            }

                            if (Severity != null && flag2) {
                                var text16 = string.Join(", ", Severity);
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: Severity = " + text16);
                                if (Severity.Contains("Unspecified")) {
                                    Severity = string.Join(",", Severity).Replace("Unspecified", "").Split(',');
                                }

                                if (!Severity.Contains(Convert.ToString((object)item.MsrcSeverity))) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if (NotSeverity != null && flag2) {
                                var text17 = string.Join(", ", NotSeverity);
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: NotSeverity = " + text17);
                                if (NotSeverity.Contains("Unspecified")) {
                                    NotSeverity = string.Join(",", NotSeverity).Replace("Unspecified", "").Split(',');
                                }

                                if (NotSeverity.Contains(Convert.ToString((object)item.MsrcSeverity))) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if (MaxSize > 0 && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: MaxDownloadSize <= " + MaxSize);
                                if ((decimal)MaxSize <= item.MaxDownloadSize) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if (MinSize > 0 && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: MaxDownloadSize >= " + MinSize);
                                if ((decimal)MinSize >= item.MaxDownloadSize) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if ((bool)IgnoreUserInput && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: CanRequestUserInput");
                                if (item.InstallationBehavior.CanRequestUserInput) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if ((bool)IgnoreRebootRequired && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: RebootBehavior");
                                if (item.InstallationBehavior.RebootBehavior != 0) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if ((bool)AutoSelectOnly && flag2) {
                                WriteDebug(DateTime.Now.ToString() + " Set post search criteria: AutoSelectOnWebsites");
                                if (!item.AutoSelectOnWebSites) {
                                    flag2 = false;
                                    WriteDebug(DateTime.Now.ToString() + " UpdateAccess" + flag2);
                                }
                            }

                            if (!flag2) {
                                continue;
                            }

                            WriteDebug(DateTime.Now.ToString() + " Update was not filtered");
                            var text18 = "";
                            foreach (string kBArticleID3 in item.KBArticleIDs) {
                                text18 = !(text18 == "") ? text18 + ", KB" + kBArticleID3 : text18 + "KB" + kBArticleID3;
                            }

                            var text20 = "-";
                            text20 = !item.IsDownloaded ? text20 + "-" : text20 + "D";
                            text20 = !item.IsInstalled ? text20 + "-" : text20 + "I";
                            text20 = !item.IsMandatory ? text20 + "-" : text20 + "M";
                            text20 = !item.IsHidden ? text20 + "-" : text20 + "H";
                            text20 = !item.IsUninstallable ? text20 + "-" : text20 + "U";
                            text20 = !item.IsBeta ? text20 + "-" : text20 + "B";
                            var pSObject = new PSObject(item);
                            pSObject.Properties.Add(new PSNoteProperty("Size", text11));
                            pSObject.Properties.Add(new PSNoteProperty("Status", text20));
                            pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                            pSObject.Properties.Add(new PSNoteProperty("KB", text18));
                            pSObject.TypeNames.Clear();
                            pSObject.TypeNames.Add("PSWindowsUpdate.WindowsUpdate");
                            collection.Add(pSObject);
                        }

                        progressRecord.RecordType = ProgressRecordType.Completed;
                        WriteProgress(progressRecord);
                        var count2 = collection.Count;
                        WriteVerbose("Found [" + count2 + "] Updates in post search criteria");
                        if (MyInvocation.BoundParameters.ContainsKey("Hide") ||
                            string.Equals(MyInvocation.InvocationName.ToString(), "Hide-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(MyInvocation.InvocationName.ToString(), "Show-WindowsUpdate", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(MyInvocation.InvocationName.ToString(), "UnHide-WindowsUpdate", StringComparison.OrdinalIgnoreCase)) {
                            var text21 = "";
                            text21 = !Hide ? "Show" : "Hide";
                            var num5 = 0;
                            var activityId2 = 1;
                            var activity2 = text21 + " updates for " + target;
                            var statusDescription2 = "[" + num5 + "/" + count2 + "]";
                            var progressRecord2 = new ProgressRecord(activityId2, activity2, statusDescription2);
                            foreach (var item2 in collection) {
                                var update2 = (IUpdate)item2.BaseObject;
                                var flag3 = false;
                                progressRecord2.StatusDescription = "[" + num5 + "/" + count2 + "] " + update2.Title + " " + item2.Properties["Size"].Value.ToString();
                                progressRecord2.PercentComplete = num5 * 100 / count2;
                                WriteProgress(progressRecord2);
                                num5++;
                                if ((bool)AcceptAll) {
                                    flag3 = true;
                                } else if (ShouldProcess(target,
                                               "(" + DateTime.Now.ToString() + ") " + text21 + " " + update2.Title + "[" + item2.Properties["Size"].Value.ToString() + "]")) {
                                    flag3 = true;
                                }

                                if (flag3) {
                                    try {
                                        update2.IsHidden = Hide;
                                        item2.Properties.Add(new PSNoteProperty("HideResult", text21));
                                    } catch {
                                        var errorRecord4 = new ErrorRecord(new Exception(target + ": You don't have permission to perform this task."), "0x80240044",
                                            ErrorCategory.CloseError, null);
                                        ThrowTerminatingError(errorRecord4);
                                    }
                                }

                                var text22 = "";
                                text22 = !update2.IsDownloaded ? text22 + "-" : text22 + "D";
                                text22 = !update2.IsInstalled ? text22 + "-" : text22 + "I";
                                text22 = !update2.IsMandatory ? text22 + "-" : text22 + "M";
                                text22 = !update2.IsHidden ? text22 + "-" : text22 + "H";
                                text22 = !update2.IsUninstallable ? text22 + "-" : text22 + "U";
                                text22 = !update2.IsBeta ? text22 + "-" : text22 + "B";
                                item2.Properties["Status"].Value = text22;
                            }

                            progressRecord2.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord2);
                        }

                        if (!Download && !Install) {
                            WriteDebug(DateTime.Now.ToString() + " Return update list only");
                            WriteObject(collection);
                            OutputObj = new Collection<PSObject>(collection);
                            continue;
                        }

                        var num6 = 0;
                        var num7 = 0;
                        var text23 = "";
                        if ((bool)Download || (bool)Install) {
                            var num8 = 0;
                            var activityId3 = 1;
                            var activity3 = "Choose updates for " + target;
                            var statusDescription3 = "[" + num8 + "/" + count2 + "]";
                            var progressRecord3 = new ProgressRecord(activityId3, activity3, statusDescription3);
                            var text24 = "";
                            foreach (var item3 in collection) {
                                item3.Properties.Add(new PSNoteProperty("X", 1));
                                item3.TypeNames.Clear();
                                item3.TypeNames.Add("PSWindowsUpdate.WindowsUpdateJob");
                                var update3 = (IUpdate)item3.BaseObject;
                                progressRecord3.StatusDescription = "[" + num8 + "/" + count2 + "] " + update3.Title + " " + item3.Properties["Size"].Value.ToString();
                                progressRecord3.PercentComplete = num8 * 100 / count2;
                                WriteProgress(progressRecord3);
                                num8++;
                                WriteDebug(DateTime.Now.ToString() + " Show update to accept: " + update3.Title);
                                var flag4 = false;
                                flag4 = (bool)AcceptAll || (AutoSelectOnly
                                    ? update3.AutoSelectOnWebSites ? true : false
                                    : ShouldProcess(target, "(" + DateTime.Now.ToString() + ") " + update3.Title + "[" + item3.Properties["Size"].Value.ToString() + "]")
                                        ? true
                                        : false);
                                var text25 = "";
                                var text26 = "";
                                if (flag4) {
                                    if (!update3.EulaAccepted) {
                                        WriteDebug(DateTime.Now.ToString() + " EulaAccepted");
                                        try {
                                            update3.AcceptEula();
                                        } catch (Exception exception) {
                                            if ((bool)Debuger) {
                                                var errorRecord5 = new ErrorRecord(exception, "Debug", ErrorCategory.CloseError, null);
                                                WriteError(errorRecord5);
                                            } else {
                                                WriteDebug(DateTime.Now.ToString() + " Can't accept Eula.");
                                            }
                                        }
                                    }

                                    try {
                                        var pSObject2 = new PSObject(item3.Properties["Identity"].Value);
                                        var text27 = (string)pSObject2.Properties["UpdateID"].Value;
                                        var num9 = (int)pSObject2.Properties["RevisionNumber"].Value;
                                        text24 = !(text24 == "")
                                            ? text24 + " or (UpdateID = '" + text27 + "' and RevisionNumber = " + num9 + ")"
                                            : "(UpdateID = '" + text27 + "' and RevisionNumber = " + num9 + ")";
                                    } catch (Exception ex3) {
                                        flag4 = false;
                                        var errorRecord6 = new ErrorRecord(new Exception("Something goes wrong: " + update3.Title + "; " + ex3.Message), "Debug",
                                            ErrorCategory.CloseError, null);
                                        WriteError(errorRecord6);
                                    }

                                    if (flag4) {
                                        text25 += "A";
                                        text26 = "Accepted";
                                        WriteDebug(DateTime.Now.ToString() + " " + text26);
                                        item3.Properties.Add(new PSNoteProperty("ChooseResult", text26));
                                        item3.Properties.Add(new PSNoteProperty("Result", text26));
                                    }
                                }

                                if (!flag4) {
                                    try {
                                        var pSObject3 = new PSObject(item3.Properties["Identity"].Value);
                                        var text28 = (string)pSObject3.Properties["UpdateID"].Value;
                                        var num10 = (int)pSObject3.Properties["RevisionNumber"].Value;
                                        text23 = !(text23 == "") ? text23 + ",'" + text28 + "'" : "'" + text28 + "'";
                                    } catch (Exception ex4) {
                                        flag4 = false;
                                        var errorRecord7 = new ErrorRecord(new Exception("Something goes wrong: " + update3.Title + "; " + ex4.Message), "Debug",
                                            ErrorCategory.CloseError, null);
                                        WriteError(errorRecord7);
                                    }

                                    text25 += "R";
                                    text26 = "Rejected";
                                    WriteDebug(DateTime.Now.ToString() + " " + text26);
                                    item3.Properties.Add(new PSNoteProperty("ChooseResult", text26));
                                    item3.Properties.Add(new PSNoteProperty("Result", text26));
                                }

                                if (!flag4 || (WUToolsObj.IsLocalHost(target) && !(ScheduleJob != DateTime.MinValue))) {
                                    text25 = !update3.IsDownloaded ? text25 + "-" : text25 + "D";
                                    text25 = !update3.IsInstalled ? text25 + "-" : text25 + "I";
                                } else {
                                    if ((bool)Download || (bool)Install) {
                                        text25 += "?";
                                    }

                                    if ((bool)Install) {
                                        text25 += "?";
                                    }
                                }

                                text25 = !update3.IsMandatory ? text25 + "-" : text25 + "M";
                                text25 = !update3.IsHidden ? text25 + "-" : text25 + "H";
                                text25 = !update3.IsUninstallable ? text25 + "-" : text25 + "U";
                                text25 = !update3.IsBeta ? text25 + "-" : text25 + "B";
                                item3.Properties["Status"].Value = text25;
                            }

                            progressRecord3.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord3);
                            if ((bool)ShowPreSearchCriteria) {
                                WriteVerbose("Choosed pre Search Criteria: " + text24);
                            }

                            var num11 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                            WriteObject(collection, true);
                            WriteVerbose("Accepted [" + num11 + "] Updates ready to Download");
                            if (num11 > 0 && (!WUToolsObj.IsLocalHost(target) || ScheduleJob != DateTime.MinValue)) {
                                var text29 = "";
                                text29 = !Debuger ? "Get-WindowsUpdate -AcceptAll" : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";
                                text29 = text29 + " -Criteria \\\"" + text24 + "\\\"";
                                if (RecurseCycle > 1) {
                                    if (text23 != "") {
                                        text29 = text29 + " -NotUpdateID " + text23;
                                    }

                                    if ((bool)IsInstalled) {
                                        text29 += " -IsInstalled";
                                    }

                                    if ((bool)IsHidden) {
                                        text29 += " -IsHidden";
                                    }

                                    if ((bool)WithHidden) {
                                        text29 += " -WithHidden";
                                    }

                                    if (UpdateType != null) {
                                        text29 = text29 + " -UpdateType " + UpdateType;
                                    }

                                    if (DeploymentAction != null) {
                                        text29 = text29 + " -DeploymentAction " + DeploymentAction;
                                    }

                                    if (UpdateID != null) {
                                        text29 = text29 + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                                    }

                                    if (RevisionNumber > 0) {
                                        text29 = text29 + " -RevisionNumber " + RevisionNumber;
                                    }

                                    if (CategoryIDs != null) {
                                        text29 = text29 + " -CategoryIDs " + string.Join(",", CategoryIDs);
                                    }

                                    if (MyInvocation.BoundParameters.ContainsKey("IsAssigned")) {
                                        text29 = !IsAssigned ? text29 + " -IsAssigned:$false" : text29 + " -IsAssigned";
                                    }

                                    if (MyInvocation.BoundParameters.ContainsKey("IsPresent")) {
                                        text29 = !IsPresent ? text29 + " -IsPresent:$false" : text29 + " -IsPresent";
                                    }

                                    if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites")) {
                                        text29 = !AutoSelectOnWebSites ? text29 + " -AutoSelectOnWebSites:$false" : text29 + " -AutoSelectOnWebSites";
                                    }

                                    if (RootCategories != null) {
                                        text29 = text29 + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                                    }

                                    if (Category != null) {
                                        text29 = text29 + " -Category '" + string.Join("','", Category) + "'";
                                    }

                                    if (KBArticleID != null) {
                                        text29 = text29 + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                                    }

                                    if (Title != null) {
                                        text29 = text29 + " -Title '" + Title + "'";
                                    }

                                    if (Severity != null) {
                                        text29 = text29 + " -Severity '" + string.Join("','", Severity) + "'";
                                    }

                                    if (NotCategory != null) {
                                        text29 = text29 + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                                    }

                                    if (NotKBArticleID != null) {
                                        text29 = text29 + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                                    }

                                    if (NotTitle != null) {
                                        text29 = text29 + " -NotTitle '" + NotTitle + "'";
                                    }

                                    if (NotSeverity != null) {
                                        text29 = text29 + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                                    }

                                    if ((bool)IgnoreUserInput) {
                                        text29 += " -IgnoreUserInput";
                                    }

                                    if ((bool)IgnoreRebootRequired) {
                                        text29 += " -IgnoreRebootRequired";
                                    }

                                    if ((bool)AutoSelectOnly) {
                                        text29 += " -AutoSelectOnly";
                                    }

                                    if (MaxSize > 0) {
                                        text29 = text29 + " -MaxSize " + MaxSize;
                                    }

                                    if (MinSize > 0) {
                                        text29 = text29 + " -MinSize " + MinSize;
                                    }
                                }

                                if ((bool)Download) {
                                    text29 += " -Download";
                                }

                                if ((bool)Install) {
                                    text29 += " -Install";
                                }

                                if ((bool)IgnoreReboot) {
                                    text29 += " -IgnoreReboot";
                                } else if ((bool)AutoReboot) {
                                    text29 += " -AutoReboot";
                                } else if (ScheduleReboot != DateTime.MinValue) {
                                    text29 = text29 + " -ScheduleReboot '" + ScheduleReboot.ToString() + "'";
                                }

                                if ((bool)WindowsUpdate) {
                                    text29 += " -WindowsUpdate";
                                } else if ((bool)MicrosoftUpdate) {
                                    text29 += " -MicrosoftUpdate";
                                } else if (ServiceID != null) {
                                    text29 = text29 + " -ServiceID '" + ServiceID + "'";
                                }

                                if ((bool)SendReport) {
                                    text29 += " -SendReport";
                                }

                                if ((bool)SendHistory) {
                                    text29 += " -SendHistory";
                                }

                                if (RecurseCycle > 1) {
                                    text29 = text29 + " -RecurseCycle " + RecurseCycle;
                                }

                                if (!AutoReboot && ScheduleReboot == DateTime.MinValue && !IgnoreReboot) {
                                    text29 += " -IgnoreReboot";
                                }

                                text29 += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                                var invokeWUJob = new InvokeWUJob();
                                invokeWUJob.ComputerName = new string[1] { target };
                                if (Credential != null) {
                                    invokeWUJob.Credential = Credential;
                                }

                                invokeWUJob.Script = text29;
                                invokeWUJob.Debuger = true;
                                if (ScheduleJob != DateTime.MinValue) {
                                    if (ScheduleJob.AddSeconds(10.0) <= DateTime.Now) {
                                        WriteVerbose("Execution time is gone. ScheduleJob was set to 1 minute delay from now.");
                                        ScheduleJob = DateTime.Now.AddMinutes(1.0);
                                    }

                                    invokeWUJob.TriggerDate = ScheduleJob;
                                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob.ToString() + "): ");
                                } else {
                                    invokeWUJob.RunNow = true;
                                    WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (Now): ");
                                }

                                WriteVerbose("powershell.exe -Command \"" + text29 + "\"");
                                var enumerable = invokeWUJob.Invoke();
                                foreach (var item4 in enumerable) {
                                    WriteObject(item4);
                                }

                                WriteDebug(DateTime.Now.ToString() + " Return invoked update list");
                                OutputObj = new Collection<PSObject>(collection.ToList());
                                continue;
                            }

                            var num12 = 0;
                            var activityId4 = 1;
                            var activity4 = "Download updates for " + target;
                            var statusDescription4 = "[" + num12 + "/" + num11 + "]";
                            var progressRecord4 = new ProgressRecord(activityId4, activity4, statusDescription4);
                            foreach (var item5 in collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Accepted")) {
                                item5.Properties.Add(new PSNoteProperty("X", 2));
                                var update4 = (IUpdate)item5.BaseObject;
                                var updateCollection3 =
                                    (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                                updateCollection3.Add(update4);
                                progressRecord4.StatusDescription = "[" + num12 + "/" + num11 + "] " + update4.Title + " " + item5.Properties["Size"].Value.ToString();
                                progressRecord4.PercentComplete = num12 * 100 / num11;
                                WriteProgress(progressRecord4);
                                num12++;
                                WriteDebug(DateTime.Now.ToString() + " Show update to download: " + update4.Title);
                                IUpdateDownloader updateDownloader = UpdateSessionObj.CreateUpdateDownloader();
                                updateDownloader.Updates = updateCollection3;
                                if ((bool)ForceDownload) {
                                    updateDownloader.IsForced = true;
                                }

                                IDownloadResult downloadResult;
                                try {
                                    downloadResult = updateDownloader.Download();
                                    WriteDebug(DateTime.Now.ToString() + " Downloaded");
                                } catch (COMException ex5) {
                                    var wUApiCodeDetails3 = WUToolsObj.GetWUApiCodeDetails(ex5.ErrorCode);
                                    var flag5 = false;
                                    if (wUApiCodeDetails3 != null) {
                                        switch (wUApiCodeDetails3.CodeType) {
                                            case 2:
                                                WriteError(new ErrorRecord(new Exception(wUApiCodeDetails3.Description), wUApiCodeDetails3.HResult, ErrorCategory.CloseError,
                                                    null));
                                                flag5 = true;
                                                break;
                                            case 3:
                                                WriteWarning(wUApiCodeDetails3.HResult + ": " + wUApiCodeDetails3.Description);
                                                break;
                                        }

                                        if (flag5) {
                                            WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                            break;
                                        }
                                    } else if (Debuger) {
                                        var errorRecord8 = new ErrorRecord(ex5, "Debug", ErrorCategory.CloseError, null);
                                        ThrowTerminatingError(errorRecord8);
                                    }

                                    WriteDebug(DateTime.Now.ToString() + " Skip to next update");
                                    continue;
                                }

                                var value2 = "";
                                switch (downloadResult.ResultCode) {
                                    case OperationResultCode.orcNotStarted:
                                        value2 = "NotStarted";
                                        break;
                                    case OperationResultCode.orcInProgress:
                                        value2 = "InProgress";
                                        break;
                                    case OperationResultCode.orcSucceeded:
                                        value2 = "Downloaded";
                                        break;
                                    case OperationResultCode.orcSucceededWithErrors:
                                        value2 = "DownloadedWithErrors";
                                        break;
                                    case OperationResultCode.orcFailed:
                                        value2 = "Failed";
                                        break;
                                    case OperationResultCode.orcAborted:
                                        value2 = "Aborted";
                                        break;
                                }

                                item5.Properties.Add(new PSNoteProperty("DownloadResult", value2));
                                item5.Properties.Add(new PSNoteProperty("Result", value2));
                                var text30 = "";
                                text30 = !(item5.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text30 + "R" : text30 + "A";
                                text30 = !(item5.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text30 + "F" : text30 + "D";
                                text30 = !update4.IsInstalled ? text30 + "-" : text30 + "I";
                                text30 = !update4.IsMandatory ? text30 + "-" : text30 + "M";
                                text30 = !update4.IsHidden ? text30 + "-" : text30 + "H";
                                text30 = !update4.IsUninstallable ? text30 + "-" : text30 + "U";
                                text30 = !update4.IsBeta ? text30 + "-" : text30 + "B";
                                item5.Properties["Status"].Value = text30;
                                WriteObject(item5, true);
                            }

                            progressRecord4.RecordType = ProgressRecordType.Completed;
                            WriteProgress(progressRecord4);
                            num6 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                            WriteVerbose("Downloaded [" + num6 + "] Updates ready to Install");
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
                        var num13 = 0;
                        var activityId5 = 1;
                        var activity5 = "Install updates for " + target;
                        var statusDescription5 = "[" + num13 + "/" + num6 + "]";
                        var progressRecord5 = new ProgressRecord(activityId5, activity5, statusDescription5);
                        foreach (var item6 in collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Downloaded")) {
                            item6.Properties.Add(new PSNoteProperty("X", 3));
                            var update5 = (IUpdate)item6.BaseObject;
                            var updateCollection4 =
                                (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                            updateCollection4.Add(update5);
                            progressRecord5.StatusDescription = "[" + num13 + "/" + num6 + "] " + update5.Title + " " + item6.Properties["Size"].Value.ToString();
                            progressRecord5.PercentComplete = num13 * 100 / num6;
                            WriteProgress(progressRecord5);
                            num13++;
                            WriteDebug(DateTime.Now.ToString() + " Show update to install: " + update5.Title);
                            var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                            updateInstaller.Updates = updateCollection4;
                            if ((bool)ForceInstall) {
                                updateInstaller.IsForced = true;
                            }

                            IInstallationResult installationResult;
                            try {
                                installationResult = updateInstaller.Install();
                                WriteDebug(DateTime.Now.ToString() + " Installed");
                            } catch (COMException ex6) {
                                var wUApiCodeDetails4 = WUToolsObj.GetWUApiCodeDetails(ex6.ErrorCode);
                                var flag6 = false;
                                if (wUApiCodeDetails4 != null) {
                                    switch (wUApiCodeDetails4.CodeType) {
                                        case 2:
                                            WriteError(new ErrorRecord(new Exception(wUApiCodeDetails4.Description), wUApiCodeDetails4.HResult, ErrorCategory.CloseError, null));
                                            flag6 = true;
                                            break;
                                        case 3:
                                            WriteWarning(wUApiCodeDetails4.HResult + ": " + wUApiCodeDetails4.Description);
                                            break;
                                    }

                                    if (flag6) {
                                        WriteDebug(DateTime.Now.ToString() + " Skip to next computer");
                                        break;
                                    }
                                } else if (Debuger) {
                                    var errorRecord9 = new ErrorRecord(ex6, "Debug", ErrorCategory.CloseError, null);
                                    ThrowTerminatingError(errorRecord9);
                                }

                                WriteDebug(DateTime.Now.ToString() + " Skip to next update");
                                continue;
                            }

                            if (!NeedsReboot) {
                                NeedsReboot = installationResult.RebootRequired;
                                WriteDebug(DateTime.Now.ToString() + " Reboot is required");
                            }

                            var value3 = "";
                            switch (installationResult.ResultCode) {
                                case OperationResultCode.orcNotStarted:
                                    value3 = "NotStarted";
                                    break;
                                case OperationResultCode.orcInProgress:
                                    value3 = "InProgress";
                                    break;
                                case OperationResultCode.orcSucceeded:
                                    value3 = "Installed";
                                    break;
                                case OperationResultCode.orcSucceededWithErrors:
                                    value3 = "InstalledWithErrors";
                                    break;
                                case OperationResultCode.orcFailed:
                                    value3 = "Failed";
                                    break;
                                case OperationResultCode.orcAborted:
                                    value3 = "Aborted";
                                    break;
                            }

                            item6.Properties.Add(new PSNoteProperty("InstallResult", value3));
                            item6.Properties.Add(new PSNoteProperty("Result", value3));
                            var text31 = "";
                            text31 = !(item6.Properties["ChooseResult"].Value.ToString() == "Accepted") ? text31 + "R" : text31 + "A";
                            text31 = !(item6.Properties["DownloadResult"].Value.ToString() == "Downloaded") ? text31 + "F" : text31 + "D";
                            text31 = !(item6.Properties["InstallResult"].Value.ToString() == "Installed")
                                ? text31 + "F"
                                : !installationResult.RebootRequired
                                    ? text31 + "I"
                                    : text31 + "R";
                            text31 = !update5.IsMandatory ? text31 + "-" : text31 + "M";
                            text31 = !update5.IsHidden ? text31 + "-" : text31 + "H";
                            text31 = !update5.IsUninstallable ? text31 + "-" : text31 + "U";
                            text31 = !update5.IsBeta ? text31 + "-" : text31 + "B";
                            item6.Properties["Status"].Value = text31;
                            WriteObject(item6, true);
                        }

                        progressRecord5.RecordType = ProgressRecordType.Completed;
                        WriteProgress(progressRecord5);
                        num7 = collection.Where((PSObject x) => x.Properties["Result"].Value.ToString() == "Installed").Count();
                        WriteVerbose("Installed [" + num7 + "] Updates");
                        WriteDebug(DateTime.Now.ToString() + " Return installed update list");
                        OutputObj = new Collection<PSObject>(collection.ToList());
                        if (num7 > 0 && (bool)SendHistory) {
                            var text32 = "Get-WUHistory -SendReport -Last " + num7 + " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                            var invokeWUJob2 = new InvokeWUJob();
                            invokeWUJob2.ComputerName = new string[1] { target };
                            if (Credential != null) {
                                invokeWUJob2.Credential = Credential;
                            }

                            invokeWUJob2.Script = text32;
                            invokeWUJob2.TaskName = "PSWindowsUpdate_History";
                            if (NeedsReboot) {
                                invokeWUJob2.TriggerAtStart = true;
                                WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (AtStart): powershell.exe -Command \"" + text32 + "\"");
                            } else {
                                invokeWUJob2.RunNow = true;
                                WriteVerbose("Invoke-WUJob: PSWindowsUpdate_History " + target + " (Now): powershell.exe -Command \"" + text32 + "\"");
                            }

                            var enumerable2 = invokeWUJob2.Invoke();
                            foreach (var item7 in enumerable2) {
                                WriteObject(item7);
                            }
                        }

                        if (RecurseCycle <= 1) {
                            continue;
                        }

                        RecurseCycle--;
                        var text33 = "";
                        text33 = !Debuger ? "Get-WindowsUpdate -AcceptAll" : "$DebugPreference = 'Continue'; Get-WindowsUpdate -AcceptAll";
                        if ((bool)IsInstalled) {
                            text33 += " -IsInstalled";
                        }

                        if ((bool)IsHidden) {
                            text33 += " -IsHidden";
                        }

                        if ((bool)WithHidden) {
                            text33 += " -WithHidden";
                        }

                        if (UpdateType != null) {
                            text33 = text33 + " -UpdateType " + UpdateType;
                        }

                        if (DeploymentAction != null) {
                            text33 = text33 + " -DeploymentAction " + DeploymentAction;
                        }

                        if (UpdateID != null) {
                            text33 = text33 + " -UpdateID '" + string.Join("','", UpdateID) + "'";
                        }

                        if (NotUpdateID != null) {
                            text33 = text33 + " -NotUpdateID '" + string.Join("','", NotUpdateID) + "'";
                        } else if (text23 != "") {
                            text33 = text33 + " -NotUpdateID " + text23;
                        }

                        if (RevisionNumber > 0) {
                            text33 = text33 + " -RevisionNumber " + RevisionNumber;
                        }

                        if (CategoryIDs != null) {
                            text33 = text33 + " -CategoryIDs '" + string.Join("','", CategoryIDs) + "'";
                        }

                        if (MyInvocation.BoundParameters.ContainsKey("IsAssigned")) {
                            text33 = !IsAssigned ? text33 + " -IsAssigned:$false" : text33 + " -IsAssigned";
                        }

                        if (MyInvocation.BoundParameters.ContainsKey("IsPresent")) {
                            text33 = !IsPresent ? text33 + " -IsPresent:$false" : text33 + " -IsPresent";
                        }

                        if (MyInvocation.BoundParameters.ContainsKey("AutoSelectOnWebSites")) {
                            text33 = !AutoSelectOnWebSites ? text33 + " -AutoSelectOnWebSites:$false" : text33 + " -AutoSelectOnWebSites";
                        }

                        if (RootCategories != null) {
                            text33 = text33 + " -RootCategories '" + string.Join("','", RootCategories) + "'";
                        }

                        if (Category != null) {
                            text33 = text33 + " -Category '" + string.Join("','", Category) + "'";
                        }

                        if (KBArticleID != null) {
                            text33 = text33 + " -KBArticleID '" + string.Join("','", KBArticleID) + "'";
                        }

                        if (Title != null) {
                            text33 = text33 + " -Title '" + Title + "'";
                        }

                        if (Severity != null) {
                            text33 = text33 + " -Severity '" + string.Join("','", Severity) + "'";
                        }

                        if (NotCategory != null) {
                            text33 = text33 + " -NotCategory '" + string.Join("','", NotCategory) + "'";
                        }

                        if (NotKBArticleID != null) {
                            text33 = text33 + " -NotKBArticleID '" + string.Join("','", NotKBArticleID) + "'";
                        }

                        if (NotTitle != null) {
                            text33 = text33 + " -NotTitle '" + NotTitle + "'";
                        }

                        if (NotSeverity != null) {
                            text33 = text33 + " -NotSeverity '" + string.Join("','", NotSeverity) + "'";
                        }

                        if ((bool)IgnoreUserInput) {
                            text33 += " -IgnoreUserInput";
                        }

                        if ((bool)IgnoreRebootRequired) {
                            text33 += " -IgnoreRebootRequired";
                        }

                        if ((bool)AutoSelectOnly) {
                            text33 += " -AutoSelectOnly";
                        }

                        if (MaxSize > 0) {
                            text33 = text33 + " -MaxSize " + MaxSize;
                        }

                        if (MinSize > 0) {
                            text33 = text33 + " -MinSize " + MinSize;
                        }

                        if ((bool)Download) {
                            text33 += " -Download";
                        }

                        if ((bool)Install) {
                            text33 += " -Install";
                        }

                        if ((bool)IgnoreReboot) {
                            text33 += " -IgnoreReboot";
                        } else if ((bool)AutoReboot) {
                            text33 += " -AutoReboot";
                        } else if (ScheduleReboot != DateTime.MinValue) {
                            text33 = text33 + " -ScheduleReboot '" + ScheduleReboot.ToString() + "'";
                        }

                        if ((bool)WindowsUpdate) {
                            text33 += " -WindowsUpdate";
                        } else if ((bool)MicrosoftUpdate) {
                            text33 += " -MicrosoftUpdate";
                        } else if (ServiceID != null) {
                            text33 = text33 + " -ServiceID '" + ServiceID + "'";
                        }

                        if ((bool)SendReport) {
                            text33 += " -SendReport";
                        }

                        if ((bool)SendHistory) {
                            text33 += " -SendHistory";
                        }

                        if (RecurseCycle > 1) {
                            text33 = text33 + " -RecurseCycle " + RecurseCycle;
                        }

                        text33 += " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log -Append";
                        var invokeWUJob3 = new InvokeWUJob();
                        invokeWUJob3.ComputerName = new string[1] { target };
                        if (Credential != null) {
                            invokeWUJob3.Credential = Credential;
                        }

                        invokeWUJob3.Script = text33;
                        invokeWUJob3.TaskName = "PSWindowsUpdate_Recurse" + RecurseCycle;
                        invokeWUJob3.Debuger = true;
                        if (NeedsReboot) {
                            invokeWUJob3.TriggerAtStart = true;
                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (AtStart): ");
                        } else {
                            ScheduleJob = DateTime.Now.AddMinutes(5.0);
                            invokeWUJob3.TriggerDate = ScheduleJob;
                            WriteVerbose("Invoke-WUJob: PSWindowsUpdate " + target + " (" + ScheduleJob.ToString() + "): ");
                        }

                        WriteVerbose("powershell.exe -Command \"" + text33 + "\"");
                        var enumerable3 = invokeWUJob3.Invoke();
                        foreach (var item8 in enumerable3) {
                            WriteObject(item8);
                        }
                    } else if ((bool)Debuger) {
                        var errorRecord10 = new ErrorRecord(wUApiServiceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord10);
                    } else {
                        var error = wUApiServiceManagerObj.Error;
                        WriteError(error);
                    }
                } else if ((bool)Debuger) {
                    var errorRecord11 = new ErrorRecord(wUApiUpdateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null);
                    WriteError(errorRecord11);
                } else {
                    var error2 = wUApiUpdateSessionObj.Error;
                    WriteError(error2);
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


    }
}