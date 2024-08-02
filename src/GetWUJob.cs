using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using TaskScheduler;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Get invoked WUJob.</para>
    /// <para type="description">Use Get-WUJob cmdlet to get invoked WUJob in Task Scheduler.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Get invoked WUJob on remote machine.</para>
    ///
    /// Get-WUJob -ComputerName MG-PC
    ///
    /// </code>
    /// </example>
    [Cmdlet("Get", "WUJob", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(WUJob))]
    public class GetWUJob : PSCmdlet
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
        public PSCredential Credential { get; set; }

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
        public Hashtable PSWUSettings
        {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Specify custom name for Task Scheduler job. Default is 'PSWindowsUpdate'.</para>
        /// </summary>
        [Parameter]
        public string TaskName { get; set; }

        /// <summary>
        /// <para type="description">Clear expired WUJob.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter ClearExpired { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private string UserName { get; set; }

        private string Domain { get; set; }

        private string Password { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        /// <summary>Begin</summary>
        protected override void BeginProcessing()
        {
            CmdletStart = DateTime.Now;
            WriteDebug(DateTime.Now + " CmdletStart: " + MyInvocation.InvocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            OutputObj = [];
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new[]
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
            if (string.Equals(MyInvocation.InvocationName, "Clear-WUJob", StringComparison.OrdinalIgnoreCase))
            {
                ClearExpired = true;
            }

            var taskScheduler =
                (TaskScheduler.TaskScheduler)Activator.CreateInstance(
                    Marshal.GetTypeFromCLSID(new Guid("0F87369F-A4E5-4CFC-BD3E-73E6154572DD")));
            foreach (var target in ComputerName)
            {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try
                {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pswuModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                }
                catch
                {
                }

                for (var i = 1; i <= 3; ++i)
                {
                    try
                    {
                        if (Credential != null)
                        {
                            WriteVerbose("Try to connect " + target + ". Test " + i);
                            taskScheduler.Connect(target, UserName, Domain, Password);
                        }
                        else
                        {
                            WriteVerbose("Try to connect " + target + ". Test " + i);
                            taskScheduler.Connect(target, Type.Missing, Type.Missing, Type.Missing);
                        }
                    }
                    catch (COMException ex)
                    {
                        if (i >= 3)
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
                                ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                            }
                            else
                            {
                                WriteError(new ErrorRecord(new Exception(target + ": Unknown failure. Try debug."), "Unknown",
                                    ErrorCategory.CloseError, null));
                            }
                        }
                        else
                        {
                            Thread.Sleep(500);
                        }
                    }
                }

                try
                {
                    var taskCount = 0;
                    var folder = taskScheduler.GetFolder("\\");
                    if (!MyInvocation.BoundParameters.ContainsKey("TaskName"))
                    {
                        var tasks = folder.GetTasks(1);
                        foreach (IRegisteredTask registeredTask in tasks)
                        {
                            if (registeredTask.Definition.RegistrationInfo.Source == "PSWindowsUpdate")
                            {
                                var sendToPipeline = new PSObject(registeredTask);
                                sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                                switch (registeredTask.State)
                                {
                                    case _TASK_STATE.TASK_STATE_DISABLED:
                                        sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Disabled"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_QUEUED:
                                        sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Queued"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_READY:
                                        sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Ready"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_RUNNING:
                                        sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Running"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_UNKNOWN:
                                        sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Unknown"));
                                        break;
                                }

                                sendToPipeline.TypeNames.Clear();
                                sendToPipeline.TypeNames.Add("PSWindowsUpdate.WUJob");
                                WriteObject(sendToPipeline);
                                ++taskCount;
                            }
                        }
                    }
                    else
                    {
                        var task = folder.GetTask(TaskName);
                        var sendToPipeline = new PSObject(task);
                        sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                        switch (task.State)
                        {
                            case _TASK_STATE.TASK_STATE_DISABLED:
                                sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Disabled"));
                                break;
                            case _TASK_STATE.TASK_STATE_QUEUED:
                                sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Queued"));
                                break;
                            case _TASK_STATE.TASK_STATE_READY:
                                sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Ready"));
                                break;
                            case _TASK_STATE.TASK_STATE_RUNNING:
                                sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Running"));
                                break;
                            case _TASK_STATE.TASK_STATE_UNKNOWN:
                                sendToPipeline.Properties.Add(new PSNoteProperty("StateName", "Unknown"));
                                break;
                        }

                        sendToPipeline.TypeNames.Clear();
                        sendToPipeline.TypeNames.Add("PSWindowsUpdate.WUJob");
                        WriteObject(sendToPipeline);
                        ++taskCount;
                    }

                    WriteVerbose("Found " + taskCount + " tasks");
                    if (!MyInvocation.BoundParameters.ContainsKey("ClearExpired"))
                    {
                        var tasks = folder.GetTasks(1);
                        foreach (IRegisteredTask registeredTask in tasks)
                        {
                            if (registeredTask.Definition.RegistrationInfo.Source == "PSWindowsUpdate" &&
                                DateTime.Parse(registeredTask.Definition.Triggers[1].EndBoundary) < DateTime.Now &&
                                registeredTask.State != _TASK_STATE.TASK_STATE_RUNNING && ShouldProcess(target, "Clear expired task: " + registeredTask.Name))
                            {
                                folder.DeleteTask(registeredTask.Name, 0);
                            }
                        }
                    }
                }
                catch
                {
                    WriteError(new ErrorRecord(new Exception("Cant find task " + TaskName + " on " + target), "WUJob",
                        ErrorCategory.ResourceUnavailable, null));
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

            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}