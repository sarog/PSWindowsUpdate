using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using Microsoft.Win32;
using WUApiLib;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Get Windows Update reboot status.</para>
    /// <para type="description">Use Get-WURebootStatus cmdlet to check if reboot is needed.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Check if restart is necessary. If yes, ask to do this or don't.</para>
    ///
    /// Get-WURebootStatus
    ///
    /// <para>Reboot is required. Do it now ? [Y/N] (default is 'N')</para>
    /// <para>ComputerName RebootRequired RebootScheduled</para>
    /// <para>------------ -------------- ---------------</para>
    /// <para>MG-PC        True</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Check if restart is necessary. If yes, then shedule it.</para>
    ///
    /// Get-WURebootStatus -ScheduleReboot (Get-Date -Hour 18 -Minute 0 -Second 0)
    ///
    /// <para>ComputerName RebootRequired RebootScheduled</para>
    /// <para>------------ -------------- ---------------</para>
    /// <para>MG-PC        True           31.08.2017 18:00:00</para>
    /// </code>
    /// </example>
    [Cmdlet("Get", "WURebootStatus", ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = "ManualReboot",
        SupportsShouldProcess = true)]
    [OutputType(typeof(RebootStatus))]
    public class GetWURebootStatus : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter]
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
        public SwitchParameter SendReport { get; set; }

        /// <summary>
        /// <para type="description">Required parameter for -SendReport.</para>
        /// <para type="description">Passes the parameters (as hashtable) necessary to send the report:
        /// \r\n@{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25];[Subject="Alternative Subject"];[Properties="Alternative object properties"];[Style="Table|List"]}</para>
        /// <para type="description">Send parameters can also be saved to a PSWUSettings.xml file in ModuleBase path:
        /// \r\nExport-Clixml @{SmtpServer="your.smtp.server";From="sender@email.address";To="recipient@email.address";[Port=25]}"</para>
        /// </summary>
        [Parameter]
        public Hashtable PSWUSettings
        {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Return true/false only.</para>
        /// </summary>
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "AutoReboot")]
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "ScheduleReboot")]
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "ManualReboot")]
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "CancelReboot")]
        public SwitchParameter Silent { get; set; }

        /// <summary>
        /// <para type="description">Do not ask for reboot if it needed, but do it now.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "AutoReboot")]
        public SwitchParameter AutoReboot { get; set; }

        /// <summary>
        /// <para type="description">Specify schedule time for reboot.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "ScheduleReboot")]
        public DateTime ScheduleReboot { get; set; }

        /// <summary>
        /// <para type="description">Cancel scheduled reboot.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "CancelReboot")]
        public SwitchParameter CancelReboot { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static SystemInformation SystemInfoObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        /// <summary>Begin</summary>
        protected override void BeginProcessing()
        {
            CmdletStart = DateTime.Now;
            WriteDebug(DateTime.Now + " CmdletStart: " + MyInvocation.InvocationName);
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
                    PSWUSettings.Add("Properties", new RebootStatus());
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
            WriteDebug(DateTime.Now + " Connect to local machine: " + ComputerName[0]);
        }

        private void CoreProcessing()
        {
            foreach (var target in ComputerName)
            {
                var dateTime = DateTime.Now;
                WriteDebug(dateTime + " " + target + ": Connecting...");
                try
                {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    dateTime = DateTime.Now;
                    WriteDebug(dateTime + " Module version: " + pswuModule.Properties["Version"].Value);
                    dateTime = DateTime.Now;
                    WriteDebug(dateTime + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                }
                catch
                {
                }

                // var target = target;
                dateTime = DateTime.Now;
                var action = "(" + dateTime + ") Get Windows Update reboot status";
                if (ShouldProcess(target, action))
                {
                    bool rebootRequired;
                    if (WUToolsObj.IsLocalHost(target))
                    {
                        var apiSystemInfoObj = WUToolsObj.GetWUApiSystemInfoObj(target);
                        dateTime = DateTime.Now;
                        WriteDebug(dateTime + " SystemInfoObj mode: " + apiSystemInfoObj.Mode);
                        if (apiSystemInfoObj.Status)
                        {
                            SystemInfoObj = (SystemInformation)apiSystemInfoObj.Object;
                            rebootRequired = SystemInfoObj.RebootRequired;
                        }
                        else
                        {
                            if (Debuger)
                            {
                                WriteError(new ErrorRecord(apiSystemInfoObj.Exception, "Debug", ErrorCategory.CloseError, null));
                                continue;
                            }

                            WriteError(apiSystemInfoObj.Error);
                            continue;
                        }
                    }
                    else
                    {
                        rebootRequired = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                            .OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\").GetSubKeyNames()
                            .Contains("RebootRequired");
                    }

                    var rebootStatus = new PSObject();
                    rebootStatus.Properties.Add(new PSNoteProperty("ComputerName", target));
                    rebootStatus.Properties.Add(new PSNoteProperty("RebootRequired", rebootRequired));

                    if (ScheduleReboot != DateTime.MinValue)
                    {
                        rebootStatus.Properties.Add(new PSNoteProperty("RebootScheduled", ScheduleReboot));
                    }
                    else
                    {
                        rebootStatus.Properties.Add(new PSNoteProperty("RebootScheduled", null));
                    }

                    if (Silent)
                    {
                        WriteObject(rebootRequired, true);
                    }
                    else
                    {
                        WriteObject(rebootStatus, true);
                    }

                    OutputObj.Add(rebootStatus);

                    if (rebootRequired)
                    {
                        if (ScheduleReboot != DateTime.MinValue)
                        {
                            var strArray = new string[5];
                            dateTime = DateTime.Now;
                            strArray[0] = dateTime.ToString();
                            strArray[1] = " ";
                            strArray[2] = target;
                            strArray[3] = ": Schedule Reboot ";
                            dateTime = ScheduleReboot;
                            strArray[4] = dateTime.ToString();
                            WriteDebug(string.Concat(strArray));
                            WUToolsObj.CancelReboot(target);
                            WriteVerbose(WUToolsObj.ScheduleReboot(target, ScheduleReboot));
                        }
                        else if (AutoReboot)
                        {
                            dateTime = DateTime.Now;
                            WriteDebug(dateTime + " " + target + ": Auto Reboot");
                            WUToolsObj.CancelReboot(target);
                            WriteVerbose(WUToolsObj.RunReboot(target));
                        }
                        else if (!Silent)
                        {
                            Host.UI.WriteLine(ConsoleColor.White, Host.UI.RawUI.BackgroundColor,
                                "Reboot is required. Do it now? [Y / N] (default is 'N')");
                            if (Console.ReadLine().ToUpper() == "Y")
                            {
                                dateTime = DateTime.Now;
                                WriteDebug(dateTime + " " + target + ": Manually Reboot");
                                WUToolsObj.CancelReboot(target);
                                WriteVerbose(WUToolsObj.RunReboot(target));
                            }
                        }
                    }

                    if (CancelReboot)
                    {
                        dateTime = DateTime.Now;
                        WriteDebug(dateTime + " " + target + ": Cancel Reboot");
                        WriteVerbose(WUToolsObj.CancelReboot(target));
                    }
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
                    PSWUSettings.Add("Properties", new RebootStatus());
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1)
                {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}