using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using TaskScheduler;

namespace PSWindowsUpdate {
    [Cmdlet("Invoke", "WUJob", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "RunNow", SupportsShouldProcess = true)]
    public class InvokeWUJob : Cmdlet {
        private string _TaskName = "PSWindowsUpdate";
        private bool _Hidden = true;
        private string _Script = "ipmo PSWindowsUpdate; Get-WindowsUpdate -AcceptAll -Install | Out-File $Env:TEMP\\PSWindowsUpdate.log";

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        public PSCredential Credential { get; set; }

        [Parameter]
        public string TaskName {
            get => _TaskName;
            set => _TaskName = value;
        }

        [Parameter]
        public SwitchParameter Hidden {
            get => _Hidden;
            set => _Hidden = value;
        }

        [Parameter]
        [ValidateDateTime]
        public DateTime EndBoundary { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "TriggerDate")]
        [ValidateDateTime]
        public DateTime TriggerDate { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "TriggerAtStart")]
        public SwitchParameter TriggerAtStart { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "RunNow")]
        public SwitchParameter RunNow { get; set; }

        [Parameter]
        public string Script {
            get => _Script;
            set => _Script = value;
        }

        [Parameter]
        public SwitchParameter Force { get; set; }

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

        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            WriteDebug(DateTime.Now + " CmdletStart: Invoke-WUJob");
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (ComputerName != null) {
                return;
            }

            ComputerName = new string[1] {
                Environment.MachineName
            };
        }

        private void CoreProcessing() {
            var taskScheduler =
                (TaskScheduler.TaskScheduler)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("0F87369F-A4E5-4CFC-BD3E-73E6154572DD")));
            
            var taskDefinition = taskScheduler.NewTask(0u);
            
            var registrationInfo = taskDefinition.RegistrationInfo;
            registrationInfo.Description = TaskName;
            registrationInfo.Author = WindowsIdentity.GetCurrent().Name;
            registrationInfo.Source = "PSWindowsUpdate";
            
            var settings = taskDefinition.Settings;
            settings.Enabled = true;
            settings.StartWhenAvailable = true;
            settings.Hidden = Hidden;
            settings.DeleteExpiredTaskAfter = "PT30M";
            
            var execAction = (IExecAction)taskDefinition.Actions.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
            execAction.Path = "powershell.exe";
            execAction.Arguments = "-Command \"" + Script + "\"";
            
            if (TriggerDate != DateTime.MinValue) {
                var timeTrigger = (ITimeTrigger)taskDefinition.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);
                timeTrigger.StartBoundary = TriggerDate.ToString("yyyy-MM-ddTHH:mm:ss");
                timeTrigger.EndBoundary = EndBoundary != DateTime.MinValue
                    ? EndBoundary.ToString("yyyy-MM-ddTHH:mm:ss")
                    : TriggerDate.AddHours(1.0).ToString("yyyy-MM-ddTHH:mm:ss");
            } else if (TriggerAtStart) {
                var bootTrigger = (IBootTrigger)taskDefinition.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_BOOT);
                bootTrigger.Delay = "PT30S";
                bootTrigger.EndBoundary = EndBoundary != DateTime.MinValue
                    ? EndBoundary.ToString("yyyy-MM-ddTHH:mm:ss")
                    : DateTime.Now.AddDays(1.0).ToString("yyyy-MM-ddTHH:mm:ss");
            } else {
                var timeTrigger2 = (ITimeTrigger)taskDefinition.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);
                timeTrigger2.StartBoundary = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                timeTrigger2.EndBoundary = EndBoundary != DateTime.MinValue
                    ? EndBoundary.ToString("yyyy-MM-ddTHH:mm:ss")
                    : DateTime.Now.AddHours(1.0).ToString("yyyy-MM-ddTHH:mm:ss");
            }

            taskDefinition.Principal.RunLevel = _TASK_RUNLEVEL.TASK_RUNLEVEL_HIGHEST;
            
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                
                var pSObject = new PSObject();

                try {
                    var errorRecord = WUToolsObj.CheckPSWUModule(target);
                    if (errorRecord != null) {
                        WriteError(errorRecord);
                    }
                } catch {
                    WriteWarning(target + ": Can't check PSWindowsUpdate on destination machine - check WinRM. Can't guarantee Invoke success.");
                }

                if (!ShouldProcess(target, "(" + DateTime.Now + ") Invoke WU job: " + execAction.Path + " " + execAction.Arguments)) {
                    continue;
                }

                for (var j = 1; j <= 3; j++) {
                    try {
                        if (Credential != null) {
                            WriteVerbose("Try to connect " + target + ". Test " + j);
                            // todo: use PSCredMan.cs or switch to PSCredential?
                            taskScheduler.Connect(target, UserName, Domain, Password);
                        } else {
                            WriteVerbose("Try to connect " + target + ". Test " + j);
                            taskScheduler.Connect(target, Type.Missing, Type.Missing, Type.Missing);
                        }
                    } catch (COMException ex) {
                        if (j >= 3) {
                            var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                            if (wUApiCodeDetails != null) {
                                var codeType = wUApiCodeDetails.CodeType;
                                var num = codeType;
                                if (num == 2) {
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                                }
                            } else if (Debuger) {
                                var errorRecord2 = new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord2);
                            } else {
                                WriteError(new ErrorRecord(new Exception(target + ": Unknown failure. Try debug."), "Unknown", ErrorCategory.CloseError, null));
                            }
                        } else {
                            Thread.Sleep(500);
                        }
                    }
                }

                var folder = taskScheduler.GetFolder("\\");
                var num2 = 1;
                var runningTasks = taskScheduler.GetRunningTasks(0);
                IRunningTask runningTask = null;
                foreach (IRunningTask item in runningTasks) {
                    if (item.Name == TaskName) {
                        runningTask = item;
                        break;
                    }
                }

                if (runningTask != null) {
                    num2 = 0;
                    if (Force) {
                        WriteVerbose("Stopping task: " + TaskName);
                        folder.GetTask(TaskName).Stop(0);
                        num2 = 1;
                    } else {
                        var task = folder.GetTask(TaskName);
                        var path = ((IExecAction)task.Definition.Actions[0]).Path;
                        var arguments = ((IExecAction)task.Definition.Actions[0]).Arguments;
                        var errorRecord3 = new ErrorRecord(new Exception(target + ": Job " + TaskName + " (" + path + " " + arguments + ") is still running."),
                            "RegisterTaskDefinition", ErrorCategory.CloseError, null);
                        ThrowTerminatingError(errorRecord3);
                    }
                }

                if (num2 == 1) {
                    WriteVerbose("Registering task: " + TaskName);
                    folder.RegisterTaskDefinition(TaskName, taskDefinition, 6, "SYSTEM", null, _TASK_LOGON_TYPE.TASK_LOGON_PASSWORD, Type.Missing);
                    if (RunNow) {
                        Thread.Sleep(5000);
                        WriteVerbose("Starting task: " + TaskName);
                        folder.GetTask(TaskName).Run(0);
                    }
                }

                var tasks = folder.GetTasks(0);
                foreach (IRegisteredTask item2 in tasks) {
                    if (item2.Definition.RegistrationInfo.Source == "PSWindowsUpdate" && DateTime.Parse(item2.Definition.Triggers[1].EndBoundary) < DateTime.Now &&
                        item2.State != _TASK_STATE.TASK_STATE_RUNNING) {
                        folder.DeleteTask(item2.Name, 0);
                    }
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
    }
}