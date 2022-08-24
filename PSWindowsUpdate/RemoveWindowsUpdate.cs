using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Remove", "WindowsUpdate", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class RemoveWindowsUpdate : PSCmdlet {
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

        [Parameter(Mandatory = true, ParameterSetName = "WUSAMode")]
        [Alias(new string[] { "HotFixID" })]
        public string KBArticleID { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "WUApiMode")]
        public string UpdateID { get; set; }

        [Parameter]
        public SwitchParameter AutoReboot { get; set; }

        [Parameter]
        public SwitchParameter IgnoreReboot { get; set; }

        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleReboot { get; set; }

        [Parameter]
        private SwitchParameter NoRestart { get; set; }

        [Parameter]
        public SwitchParameter WUSAMode { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateSession UpdateSessionObj { get; set; }

        private static IUpdateSearcher SearcherObj { get; set; }

        private static UpdateInstaller InstallerObj { get; set; }

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
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (ScheduleJob == DateTime.MinValue && WUToolsObj.IsLocalHost(target)) {
                    if (WUSAMode) {
                        // 2022-08-19: possible fix for issue described by MizardX (https://github.com/mgajda83/PSWindowsUpdate/pull/9)
                        var articleId = KBArticleID.Replace("KB", "");
                        WriteVerbose(target + ": Try to uninstall KB" + KBArticleID);
                        var processStartInfo = new ProcessStartInfo();
                        processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        processStartInfo.FileName = "cmd";
                        if (WUToolsObj.IsLocalHost(target)) {
                            processStartInfo.Arguments = "/C wusa.exe /uninstall /KB:" + articleId + " /quiet /norestart";
                        } else {
                            processStartInfo.Arguments = "/C winrs.exe -r:" + target + " wusa.exe /uninstall /KB:" + articleId + " /quiet /norestart";
                        }

                        WriteDebug("Process: " + processStartInfo.FileName + " " + processStartInfo.Arguments);
                        if (ShouldProcess(target, "(" + DateTime.Now + ") Remove (wusa.exe) Windows Update: " + KBArticleID)) {
                            Process.Start(processStartInfo);
                            WriteDebug("Process started.");
                        }

                        continue;
                    }

                    var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                    WriteDebug(DateTime.Now + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                    if (wUApiUpdateSessionObj.Status) {
                        UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                        SearcherObj = UpdateSessionObj.CreateUpdateSearcher();
                        var text3 = "IsInstalled = 1";
                        if (UpdateID != null) {
                            text3 = text3 + " and UpdateID='" + UpdateID + "'";
                        }

                        WriteDebug(DateTime.Now + " Pre search criteria: " + text3);
                        ISearchResult searchResult;
                        try {
                            searchResult = SearcherObj.Search(text3);
                        } catch (COMException ex) {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                            if (wUApiCodeDetails != null) {
                                var codeType = wUApiCodeDetails.CodeType;
                                var num = codeType;
                                if (num == 2) {
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                }
                            } else if (Debuger) {
                                var errorRecord = new ErrorRecord(ex, "Debug: " + ex.ErrorCode, ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord);
                            } else {
                                WriteError(new ErrorRecord(new Exception(target + ": Unknown failure. Try debug."), "Unknown", ErrorCategory.CloseError, null));
                            }

                            continue;
                        }

                        var count = searchResult.Updates.Count;
                        WriteDebug("Found [" + count + "] Updates in pre search criteria");
                        if (count == 0) {
                            continue;
                        }

                        var updateCollection = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        var updateCollection2 = (UpdateCollection)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        updateCollection2 = searchResult.Updates;
                        if (KBArticleID != null) {
                            foreach (IUpdate item in updateCollection2) {
                                var flag = false;
                                foreach (string kBArticleID in item.KBArticleIDs) {
                                    if (Regex.IsMatch(KBArticleID, kBArticleID, RegexOptions.IgnoreCase)) {
                                        flag = true;
                                        break;
                                    }
                                }

                                if (flag) {
                                    updateCollection.Add(item);
                                    break;
                                }
                            }
                        } else {
                            updateCollection = updateCollection2;
                        }

                        var count2 = updateCollection.Count;
                        WriteDebug("Found [" + count2 + "] Updates in post search criteria");
                        if (count2 == 0 || !ShouldProcess(target, "(" + DateTime.Now + ") Remove (WUApi) Windows Update: " + updateCollection[0].Title)) {
                            continue;
                        }

                        var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                        updateInstaller.Updates = updateCollection;
                        IInstallationResult installationResult = null;
                        try {
                            installationResult = updateInstaller.Uninstall();
                            WriteObject(installationResult);
                            var text5 = "";
                            switch (installationResult.ResultCode) {
                                case OperationResultCode.orcNotStarted:
                                    text5 = "NotStarted";
                                    break;
                                case OperationResultCode.orcInProgress:
                                    text5 = "InProgress";
                                    break;
                                case OperationResultCode.orcSucceeded:
                                    text5 = "Uninstalled";
                                    break;
                                case OperationResultCode.orcSucceededWithErrors:
                                    text5 = "UninstalledWithErrors";
                                    break;
                                case OperationResultCode.orcFailed:
                                    text5 = "Failed";
                                    break;
                                case OperationResultCode.orcAborted:
                                    text5 = "Aborted";
                                    break;
                            }

                            WriteDebug(DateTime.Now + " Uninstallation status: " + text5);
                            if (installationResult.RebootRequired) {
                                NeedsReboot = installationResult.RebootRequired;
                                WriteDebug(DateTime.Now + " Reboot is required");
                            }
                        } catch (Exception ex2) {
                            if (Regex.IsMatch(ex2.Message, "HRESULT: 0x80240028")) {
                                WriteWarning(ex2.Message + ". The update could not be uninstalled because the request did not originate from a WSUS server.");
                            }

                            if (Regex.IsMatch(ex2.Message, "HRESULT: 0x80240044")) {
                                WriteWarning(ex2.Message + ". Your security policy don't allow a non-administator identity to perform this task.");
                            }
                        }
                    } else if (Debuger) {
                        var errorRecord2 = new ErrorRecord(wUApiUpdateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord2);
                    } else {
                        var error = wUApiUpdateSessionObj.Error;
                        WriteError(error);
                    }

                    continue;
                }

                var errorRecord3 = WUToolsObj.CheckPSWUModule(target);
                if (errorRecord3 != null) {
                    WriteError(errorRecord3);
                    continue;
                }

                var text6 = "";
                text6 = !Debuger ? "Remove-WindowsUpdate" : "$DebugPreference = 'Continue'; Remove-WindowsUpdate";
                if (KBArticleID != null) {
                    text6 = text6 + " -KBArticleID " + KBArticleID;
                }

                if (UpdateID != null) {
                    text6 = text6 + " -UpdateID '" + UpdateID + "'";
                }

                if (AutoReboot) {
                    text6 += " -AutoReboot";
                }

                if (ScheduleReboot != DateTime.MinValue) {
                    text6 = text6 + " -ScheduleReboot " + ScheduleReboot;
                }

                if (IgnoreReboot) {
                    text6 += " -IgnoreReboot";
                }

                text6 += " -Verbose -Confirm:$false *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { target };
                if (Credential != null) {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = text6;
                if (ScheduleJob != DateTime.MinValue) {
                    invokeWUJob.TriggerDate = ScheduleJob;
                    WriteVerbose("Invoke-WUJob: " + target + " (" + ScheduleJob + ") " + text6);
                } else {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: " + target + " (Now) " + text6);
                }

                WriteVerbose("powershell.exe -Command \"" + text6 + "\"");
                var enumerable = invokeWUJob.Invoke();
                foreach (var item2 in enumerable) {
                    WriteObject(item2);
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