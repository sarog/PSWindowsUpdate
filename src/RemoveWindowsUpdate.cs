using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using WUApiLib;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Uninstall update.</para>
    /// <para type="description">Use Remove-WindowsUpdate to uninstall update.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Try to uninstall update with specific KBArticleID = KB958830.</para>
    ///
    /// Get-WUUninstall -KBArticleID KB958830
    ///
    /// </code>
    /// </example>
    [Cmdlet("Remove", "WindowsUpdate", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class RemoveWindowsUpdate : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        /// <summary>
        /// <para type="description">Specify alternative credential.</para>
        /// </summary>
        [Parameter]
        private PSCredential Credential { get; set; }

        /// <summary>
        /// <para type="description">Send report email to specific recipients.</para>
        /// <para type="description">Requires the parameter -PSWUSettings or declare the PSWUSettings.xml file in ModuleBase path.</para>
        /// </summary>
        [Parameter]
        private SwitchParameter SendReport { get; set; }

        /// <summary>
        /// <para type="description">Required parameter for -SendReport.</para>
        /// <para type="description">Passes the parameters (as hashtable) necessary to send the report:
        /// \r\n@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25];[Subject="Alternative Subject"];[Properties="Alternative object properties"];[Style="Table|List"]}</para>
        /// <para type="description">Send parameters can also be saved to a PSWUSettings.xml file in ModuleBase path:
        /// \r\nExport-Clixml @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25]}"</para>
        /// </summary>
        [Parameter]
        private Hashtable PSWUSettings
        {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Specify schedule time job.</para>
        /// </summary>
        [Parameter]
        public DateTime ScheduleJob { get; set; }

        /// <summary>
        /// <para type="description">KBArticleID that will be uninstalled.</para>
        /// </summary>
        [Parameter(ParameterSetName = "WUSAMode", Mandatory = true)]
        [Alias("HotFixID")]
        public string KBArticleID { get; set; }

        /// <summary>
        /// <para type="description">Update ID that will be uninstalled.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "WUApiMode")]
        public string UpdateID { get; set; }

        /// <summary>
        /// <para type="description">Do not ask for reboot if it needed.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter AutoReboot { get; set; }

        /// <summary>
        /// <para type="description">Do not ask for reboot if it needed, but do not reboot automaticaly.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter IgnoreReboot { get; set; }

        /// <summary>
        /// <para type="description">Specify time when system will be rebooted.</para>
        /// </summary>
        [Parameter]
        [ValidateDateTime]
        public DateTime ScheduleReboot { get; set; }

        /// <summary>
        /// <para type="description">Don't use automatic reboot.</para>
        /// </summary>
        [Parameter]
        private SwitchParameter NoRestart { get; set; }

        /// <summary>
        /// <para type="description">Wse wusa.exe instead of WU Api.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter WUSAMode { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
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

        /// <summary>Begin</summary>
        protected override void BeginProcessing()
        {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                ThrowTerminatingError(new ErrorRecord(
                    new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = [];
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new string[1]
                    {
                        "*"
                    });
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = false;
                    WriteDebug(DateTime.Now + " Disabling -SendReport");
                }
            }

            if (ComputerName != null)
            {
                return;
            }

            ComputerName =
            [
                Environment.MachineName
            ];
        }

        private void CoreProcessing()
        {
            var articleId = string.Empty;
            // 2023-12-13: fixes issue https://github.com/mgajda83/PSWindowsUpdate/pull/9
            if (MyInvocation.BoundParameters.ContainsKey("KBArticleID"))
            {
                articleId = KBArticleID.Replace("KB", "");
            }

            foreach (var target in ComputerName)
            {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try
                {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                }
                catch
                {
                }

                if (ScheduleJob == DateTime.MinValue && WUToolsObj.IsLocalHost(target))
                {
                    if (WUSAMode)
                    {
                        WriteVerbose(target + ": Try to uninstall KB" + KBArticleID);
                        var processStartInfo = new ProcessStartInfo();
                        processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        processStartInfo.FileName = "cmd";
                        if (WUToolsObj.IsLocalHost(target))
                        {
                            processStartInfo.Arguments = "/C wusa.exe /uninstall /KB:" + articleId + " /quiet /norestart";
                        }
                        else
                        {
                            processStartInfo.Arguments = "/C winrs.exe -r:" + target + " wusa.exe /uninstall /KB:" + articleId +
                                                         " /quiet /norestart";
                        }

                        WriteDebug("Process: " + processStartInfo.FileName + " " + processStartInfo.Arguments);
                        if (ShouldProcess(target, "(" + DateTime.Now + ") Remove (wusa.exe) Windows Update: " + KBArticleID))
                        {
                            Process.Start(processStartInfo);
                            WriteDebug("Process started.");
                        }

                        continue;
                    }

                    var wUApiUpdateSessionObj = WUToolsObj.GetWUApiUpdateSessionObj(target);
                    WriteDebug(DateTime.Now + " UpdateSessionObj mode: " + wUApiUpdateSessionObj.Mode);
                    if (wUApiUpdateSessionObj.Status)
                    {
                        UpdateSessionObj = (UpdateSession)wUApiUpdateSessionObj.Object;
                        SearcherObj = UpdateSessionObj.CreateUpdateSearcher();
                        var text3 = "IsInstalled = 1";
                        if (UpdateID != null)
                        {
                            text3 = text3 + " and UpdateID='" + UpdateID + "'";
                        }

                        WriteDebug(DateTime.Now + " Pre search criteria: " + text3);
                        ISearchResult searchResult;
                        try
                        {
                            searchResult = SearcherObj.Search(text3);
                        }
                        catch (COMException ex)
                        {
                            var wuApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                            if (wuApiCodeDetails != null)
                            {
                                if (wuApiCodeDetails.CodeType == 2)
                                {
                                    WriteError(new ErrorRecord(new Exception(wuApiCodeDetails.Description), wuApiCodeDetails.HResult,
                                        ErrorCategory.CloseError, null));
                                }
                            }
                            else if (Debuger)
                            {
                                var errorRecord = new ErrorRecord(ex, "Debug: " + ex.ErrorCode, ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord);
                            }
                            else
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Unknown failure. Try debug."), "Unknown",
                                    ErrorCategory.CloseError, null));
                            }

                            continue;
                        }

                        var count = searchResult.Updates.Count;
                        WriteDebug("Found [" + count + "] Updates in pre search criteria");
                        if (count == 0)
                        {
                            continue;
                        }

                        var updateCollection =
                            (UpdateCollection)Activator.CreateInstance(
                                Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        var updateCollection2 =
                            (UpdateCollection)Activator.CreateInstance(
                                Marshal.GetTypeFromCLSID(new Guid("13639463-00DB-4646-803D-528026140D88")));
                        updateCollection2 = searchResult.Updates;
                        if (KBArticleID != null)
                        {
                            foreach (IUpdate item in updateCollection2)
                            {
                                var flag = false;
                                foreach (string kBArticleID in item.KBArticleIDs)
                                {
                                    if (Regex.IsMatch(KBArticleID, kBArticleID, RegexOptions.IgnoreCase))
                                    {
                                        flag = true;
                                        break;
                                    }
                                }

                                if (flag)
                                {
                                    updateCollection.Add(item);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            updateCollection = updateCollection2;
                        }

                        var count2 = updateCollection.Count;
                        WriteDebug("Found [" + count2 + "] Updates in post search criteria");
                        if (count2 == 0 || !ShouldProcess(target,
                                "(" + DateTime.Now + ") Remove (WUApi) Windows Update: " + updateCollection[0].Title))
                        {
                            continue;
                        }

                        var updateInstaller = UpdateSessionObj.CreateUpdateInstaller();
                        updateInstaller.Updates = updateCollection;
                        IInstallationResult installationResult = null;
                        try
                        {
                            installationResult = updateInstaller.Uninstall();
                            WriteObject(installationResult);
                            var text5 = "";
                            switch (installationResult.ResultCode)
                            {
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
                            if (installationResult.RebootRequired)
                            {
                                NeedsReboot = installationResult.RebootRequired;
                                WriteDebug(DateTime.Now + " Reboot is required");
                            }
                        }
                        catch (Exception ex2)
                        {
                            if (Regex.IsMatch(ex2.Message, "HRESULT: 0x80240028"))
                            {
                                WriteWarning(ex2.Message +
                                             ". The update could not be uninstalled because the request did not originate from a WSUS server.");
                            }

                            if (Regex.IsMatch(ex2.Message, "HRESULT: 0x80240044"))
                            {
                                WriteWarning(ex2.Message +
                                             ". Your security policy don't allow a non-administator identity to perform this task.");
                            }
                        }
                    }
                    else if (Debuger)
                    {
                        var errorRecord2 = new ErrorRecord(wUApiUpdateSessionObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord2);
                    }
                    else
                    {
                        var error = wUApiUpdateSessionObj.Error;
                        WriteError(error);
                    }

                    continue;
                }

                var errorRecord3 = WUToolsObj.CheckPSWUModule(target);
                if (errorRecord3 != null)
                {
                    WriteError(errorRecord3);
                    continue;
                }

                var cmdLine = "";
                cmdLine = !Debuger ? "Remove-WindowsUpdate" : "$DebugPreference = 'Continue'; Remove-WindowsUpdate";
                if (KBArticleID != null)
                {
                    cmdLine = cmdLine + " -KBArticleID " + KBArticleID;
                }

                if (UpdateID != null)
                {
                    cmdLine = cmdLine + " -UpdateID '" + UpdateID + "'";
                }

                if (AutoReboot)
                {
                    cmdLine += " -AutoReboot";
                }

                if (ScheduleReboot != DateTime.MinValue)
                {
                    cmdLine = cmdLine + " -ScheduleReboot " + ScheduleReboot;
                }

                if (IgnoreReboot)
                {
                    cmdLine += " -IgnoreReboot";
                }

                cmdLine += " -Verbose -Confirm:$false *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = [target];
                if (Credential != null)
                {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = cmdLine;
                if (ScheduleJob != DateTime.MinValue)
                {
                    invokeWUJob.TriggerDate = ScheduleJob;
                    WriteVerbose("Invoke-WUJob: " + target + " (" + ScheduleJob + ") " + cmdLine);
                }
                else
                {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: " + target + " (Now) " + cmdLine);
                }

                WriteVerbose("powershell.exe -Command \"" + cmdLine + "\"");
                var enumerable = invokeWUJob.Invoke();
                foreach (var item2 in enumerable)
                {
                    WriteObject(item2);
                }
            }
        }

        /// <summary>Process</summary>
        protected override void ProcessRecord()
        {
            var flag = false;
            if (Credential != null)
            {
                var userName = Credential.GetNetworkCredential().UserName;
                var domain = Credential.GetNetworkCredential().Domain;
                var password = Credential.GetNetworkCredential().Password;
                WriteDebug(DateTime.Now + " UserName: " + userName + "; Domain: " + domain + "; Password: " + password.Substring(0, 1) +
                           "*****");
                var windowsPrincipal1 = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                var str1 = "";
                if (windowsPrincipal1.IsInRole(WindowsBuiltInRole.Administrator))
                {
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
                if (!WUToolsObj.IsLocalHost(ComputerName[0]))
                {
                    logonType = WUImpersonator.LogonSessionType.NewCredentials;
                    logonProvider = WUImpersonator.LogonProvider.WinNT50;
                }

                using (new WUImpersonator(userName, domain, password, logonType, logonProvider))
                {
                    if (new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
                    {
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
                        try
                        {
                            CoreProcessing();
                            flag = false;
                        }
                        catch (Exception ex)
                        {
                            WriteDebug(DateTime.Now + " Something goes wrong: " + ex.Message);
                            flag = true;
                        }
                    }
                    else
                    {
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
                if (windowsPrincipal2.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    str4 = "RunAs";
                }

                WriteDebug(DateTime.Now + " After User: " + WindowsIdentity.GetCurrent().Name + " " + str4);
            }
            else
            {
                flag = true;
            }

            if (!flag)
            {
                return;
            }

            CoreProcessing();
        }

        /// <summary>End</summary>
        protected override void EndProcessing()
        {
            CmdletEnd = DateTime.Now;
            var CmdletInfo = new PSObject();
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletStart", CmdletStart));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletEnd", CmdletEnd));
            CmdletInfo.Properties.Add(new PSNoteProperty("CmdletLine", MyInvocation.Line));
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Send report");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", "*");
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            if (NeedsReboot)
            {
                if (ScheduleReboot != DateTime.MinValue)
                {
                    WriteDebug(DateTime.Now + " Schedule Reboot " + ScheduleReboot);
                    WriteVerbose(WUToolsObj.ScheduleReboot("localhost", ScheduleReboot));
                }
                else if (AutoReboot)
                {
                    WriteDebug(DateTime.Now + " Auto Reboot");
                    WriteVerbose(WUToolsObj.RunReboot("localhost"));
                }
                else if (IgnoreReboot)
                {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor, "Reboot is required, but do it manually.");
                }
                else
                {
                    Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor,
                        "Reboot is required. Do it now? [Y / N] (default is 'N')");
                    if (Console.ReadLine().ToUpper() == "Y")
                    {
                        WriteDebug(DateTime.Now + " Manually Reboot");
                        WriteVerbose(WUToolsObj.RunReboot("localhost"));
                    }
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}