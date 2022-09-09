using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using TaskScheduler;

namespace PSWindowsUpdate {
    
    [Cmdlet("Get", "WUJob", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(WUJob))]
    public class GetWUJob : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        public PSCredential Credential { get; set; }

        [Parameter]
        private SwitchParameter SendReport { get; set; }

        [Parameter]
        public Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter]
        public string TaskName { get; set; }

        [Parameter]
        public SwitchParameter ClearExpired { get; set; }

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
                    PSWUSettings.Add("Properties", new string[1] {
                        "*"
                    });
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
            if (string.Equals(invocationName, "Clear-WUJob", StringComparison.OrdinalIgnoreCase)) {
                ClearExpired = true;
            }

            var taskScheduler = (TaskScheduler.TaskScheduler)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("0F87369F-A4E5-4CFC-BD3E-73E6154572DD")));
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                for (var j = 1; j <= 3; j++) {
                    try {
                        if (Credential != null) {
                            WriteVerbose("Try to connect " + target + ". Test " + j);
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
                                var errorRecord = new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord);
                            } else {
                                WriteError(new ErrorRecord(new Exception(target + ": Unknown failure. Try debug."), "Unknown", ErrorCategory.CloseError, null));
                            }
                        } else {
                            Thread.Sleep(500);
                        }
                    }
                }

                try {
                    var taskCount = 0;
                    var folder = taskScheduler.GetFolder("\\");
                    if (!MyInvocation.BoundParameters.ContainsKey("TaskName")) {
                        var tasks = folder.GetTasks(1);
                        foreach (IRegisteredTask item in tasks) {
                            if (item.Definition.RegistrationInfo.Source == "PSWindowsUpdate") {
                                var pSObject = new PSObject(item);
                                pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                                switch (item.State) {
                                    case _TASK_STATE.TASK_STATE_DISABLED:
                                        pSObject.Properties.Add(new PSNoteProperty("StateName", "Disabled"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_QUEUED:
                                        pSObject.Properties.Add(new PSNoteProperty("StateName", "Queued"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_READY:
                                        pSObject.Properties.Add(new PSNoteProperty("StateName", "Ready"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_RUNNING:
                                        pSObject.Properties.Add(new PSNoteProperty("StateName", "Running"));
                                        break;
                                    case _TASK_STATE.TASK_STATE_UNKNOWN:
                                        pSObject.Properties.Add(new PSNoteProperty("StateName", "Unknown"));
                                        break;
                                }

                                pSObject.TypeNames.Clear();
                                pSObject.TypeNames.Add("PSWindowsUpdate.WUJob");
                                WriteObject(pSObject);
                                taskCount++;
                            }
                        }
                    } else {
                        var task = folder.GetTask(TaskName);
                        var pSObject2 = new PSObject(task);
                        pSObject2.Properties.Add(new PSNoteProperty("ComputerName", target));
                        switch (task.State) {
                            case _TASK_STATE.TASK_STATE_DISABLED:
                                pSObject2.Properties.Add(new PSNoteProperty("StateName", "Disabled"));
                                break;
                            case _TASK_STATE.TASK_STATE_QUEUED:
                                pSObject2.Properties.Add(new PSNoteProperty("StateName", "Queued"));
                                break;
                            case _TASK_STATE.TASK_STATE_READY:
                                pSObject2.Properties.Add(new PSNoteProperty("StateName", "Ready"));
                                break;
                            case _TASK_STATE.TASK_STATE_RUNNING:
                                pSObject2.Properties.Add(new PSNoteProperty("StateName", "Running"));
                                break;
                            case _TASK_STATE.TASK_STATE_UNKNOWN:
                                pSObject2.Properties.Add(new PSNoteProperty("StateName", "Unknown"));
                                break;
                        }

                        pSObject2.TypeNames.Clear();
                        pSObject2.TypeNames.Add("PSWindowsUpdate.WUJob");
                        WriteObject(pSObject2);
                        taskCount++;
                    }

                    WriteVerbose("Found " + taskCount + " tasks");
                    if (MyInvocation.BoundParameters.ContainsKey("ClearExpired")) {
                        continue;
                    }

                    var tasks2 = folder.GetTasks(1);
                    foreach (IRegisteredTask item2 in tasks2) {
                        if (item2.Definition.RegistrationInfo.Source == "PSWindowsUpdate" && DateTime.Parse(item2.Definition.Triggers[1].EndBoundary) < DateTime.Now &&
                            item2.State != _TASK_STATE.TASK_STATE_RUNNING && ShouldProcess(target, "Clear expired task: " + item2.Name)) {
                            folder.DeleteTask(item2.Name, 0);
                        }
                    }
                } catch {
                    var errorRecord2 = new ErrorRecord(new Exception("Cant find task " + TaskName + " on " + target), "WUJob", ErrorCategory.ResourceUnavailable, null);
                    WriteError(errorRecord2);
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
                    PSWUSettings.Add("Properties", "*");
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