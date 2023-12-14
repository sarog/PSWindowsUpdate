using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    /// <summary>
    /// <para type="synopsis">Remove windows update service manager.</para>
    /// <para type="description">Use Remove-WUServiceManager cmdlet to unregister Windows Update Service Manager.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Try unregister Microsoft Update Service.</para>
    ///
    /// Remove-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d"
    ///
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Unregister Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".</para>
    /// <para></para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    /// </code>
    /// </example>
    [Cmdlet("Remove", "WUServiceManager", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class RemoveWUServiceManager : PSCmdlet {
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
        private Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Specify schedule time job.</para>
        /// </summary>
        [Parameter]
        public DateTime ScheduleJob { get; set; }

        /// <summary>
        /// <para type="description">Specify ServiceID of Service Manager that you want to remove.</para>
        /// </summary>
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string ServiceID { get; set; }

        /// <summary>
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateServiceManager ServiceManagerObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        /// <summary>Begin</summary>
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
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                try {
                    var pSWUModule = WUToolsObj.GetPSWUModule(target);
                    WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                    WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                } catch { }

                if (!ShouldProcess(target, "(" + DateTime.Now + ") Unregister Windows Update Service Manager: " + ServiceID)) {
                    continue;
                }

                if (WUToolsObj.IsLocalHost(target) && ScheduleJob == DateTime.MinValue) {
                    WriteVerbose("Run WUApi");
                    var wUApiServiceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                    WriteDebug(DateTime.Now + " ServiceManagerObj mode: " + wUApiServiceManagerObj.Mode);
                    if (wUApiServiceManagerObj.Status) {
                        ServiceManagerObj = (UpdateServiceManager)wUApiServiceManagerObj.Object;
                        try {
                            WUToolsObj.RestartService(target);
                            ServiceManagerObj.RemoveService(ServiceID);
                            WUToolsObj.RestartService(target);
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
                        }
                    } else if (Debuger) {
                        var errorRecord2 = new ErrorRecord(wUApiServiceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null);
                        WriteError(errorRecord2);
                    } else {
                        var error = wUApiServiceManagerObj.Error;
                        WriteError(error);
                    }

                    continue;
                }

                var errorRecord3 = WUToolsObj.CheckPSWUModule(target);
                if (errorRecord3 != null) {
                    WriteError(errorRecord3);
                    continue;
                }

                var text2 = "Remove-WUServiceManager -ServiceID " + ServiceID + " -Verbose -Confirm:$false *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                var invokeWUJob = new InvokeWUJob();
                invokeWUJob.ComputerName = new string[1] { target };
                if (Credential != null) {
                    invokeWUJob.Credential = Credential;
                }

                invokeWUJob.Script = text2;
                if (ScheduleJob != DateTime.MinValue) {
                    invokeWUJob.TriggerDate = ScheduleJob;
                    WriteVerbose("Invoke-WUJob: " + target + " (" + ScheduleJob + "): ");
                } else {
                    invokeWUJob.RunNow = true;
                    WriteVerbose("Invoke-WUJob: " + target + " (Now): ");
                }

                WriteVerbose("powershell.exe -Command \"" + text2 + "\"");
                var enumerable = invokeWUJob.Invoke();
                foreach (var item in enumerable) {
                    WriteObject(item);
                }
            }
        }

        /// <summary>Process</summary>
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

        /// <summary>End</summary>
        protected override void EndProcessing() {
            WriteDebug(DateTime.Now + " CmdletEnd");
        }
    }
}