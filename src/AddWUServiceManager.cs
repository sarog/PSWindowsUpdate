using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Add", "WUServiceManager", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "LocalServiceID", SupportsShouldProcess = true)]
    [OutputType(typeof(ServiceManager))]
    public class AddWUServiceManager : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();
        private int _AddServiceFlag = 2;
        private string _ServiceName = "Offline Sync Service";

        [Parameter(ParameterSetName = "RemoteServiceID", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "MicrosoftUpdate", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
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

        [Parameter(Mandatory = true, ParameterSetName = "LocalServiceID")]
        [Parameter(Mandatory = true, ParameterSetName = "RemoteServiceID")]
        [ValidateNotNullOrEmpty]
        public string ServiceID { get; set; }

        [Parameter(ParameterSetName = "LocalServiceID")]
        public int AddServiceFlag {
            get => _AddServiceFlag;
            set => _AddServiceFlag = value;
        }

        [Parameter(ParameterSetName = "LocalServiceID")]
        public string AuthorizationCabPath { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "MicrosoftUpdate")]
        public SwitchParameter MicrosoftUpdate { get; set; }

        [Parameter(ParameterSetName = "OfflineSync")]
        public string ServiceName {
            get => _ServiceName;
            set => _ServiceName = value;
        }

        [Parameter(Mandatory = true, ParameterSetName = "OfflineSync")]
        public string ScanFileLocation { get; set; }

        [Parameter]
        public SwitchParameter Silent { get; set; }

        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static UpdateServiceManager ServiceManagerObj { get; set; }

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
                    PSWUSettings.Add("Properties", new ServiceManager());
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = false;
                    WriteDebug(DateTime.Now + " Disabling -SendReport");
                }
            }

            if (ComputerName == null) {
                ComputerName = new string[1] {
                    Environment.MachineName
                };
            }

            if (MicrosoftUpdate) {
                return;
            }

            ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
        }

        private void CoreProcessing() {
            foreach (var target in ComputerName) {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                var pSWUModule = WUToolsObj.GetPSWUModule(target);
                WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                var wUApiServiceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                WriteDebug(DateTime.Now + " ServiceManagerObj mode: " + wUApiServiceManagerObj.Mode);
                if (wUApiServiceManagerObj.Status) {
                    ServiceManagerObj = (UpdateServiceManager)wUApiServiceManagerObj.Object;
                    IUpdateService obj = null;
                    var value = "";
                    WUToolsObj.RestartService(target);
                    if (ParameterSetName == "OfflineSync") {
                        if (ShouldProcess(target, "(" + DateTime.Now + ") Register Offline Windows Update Service Manager: " + ScanFileLocation)) {
                            try {
                                obj = ServiceManagerObj.AddScanPackageService(ServiceName, ScanFileLocation, 1);
                                value = "Registered";
                            } catch (COMException ex) {
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
                                }

                                continue;
                            }
                        }
                    } else if (ShouldProcess(target, "(" + CmdletStart + ") Register Windows Update Service Manager: " + ServiceID)) {
                        try {
                            var updateServiceRegistration = ServiceManagerObj.AddService2(ServiceID, AddServiceFlag, AuthorizationCabPath);
                            obj = updateServiceRegistration.Service;
                            switch (updateServiceRegistration.RegistrationState) {
                                case UpdateServiceRegistrationState.usrsNotRegistered:
                                    value = "Not Registered";
                                    break;
                                case UpdateServiceRegistrationState.usrsRegistrationPending:
                                    value = "Registration Pending";
                                    break;
                                case UpdateServiceRegistrationState.usrsRegistered:
                                    value = "Registered";
                                    break;
                            }
                        } catch (COMException ex2) {
                            var wUApiCodeDetails2 = WUToolsObj.GetWUApiCodeDetails(ex2.ErrorCode);
                            if (wUApiCodeDetails2 != null) {
                                var codeType2 = wUApiCodeDetails2.CodeType;
                                var num2 = codeType2;
                                if (num2 == 2) {
                                    WriteError(new ErrorRecord(new Exception(wUApiCodeDetails2.Description), wUApiCodeDetails2.HResult, ErrorCategory.CloseError, null));
                                }
                            } else if (Debuger) {
                                var errorRecord2 = new ErrorRecord(ex2, "Debug", ErrorCategory.CloseError, null);
                                ThrowTerminatingError(errorRecord2);
                            }

                            continue;
                        }
                    }

                    WUToolsObj.RestartService(target);
                    var pSObject = new PSObject(obj);
                    pSObject.Properties.Add(new PSNoteProperty("RegistrationStateName", value));
                    pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                    pSObject.TypeNames.Clear();
                    pSObject.TypeNames.Add("PSWindowsUpdate.ServiceManager");
                    OutputObj.Add(pSObject);
                    if (!Silent) {
                        WriteObject(pSObject, true);
                    }
                } else if (Debuger) {
                    var errorRecord3 = new ErrorRecord(wUApiServiceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null);
                    WriteError(errorRecord3);
                } else {
                    var error = wUApiServiceManagerObj.Error;
                    WriteError(error);
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
                    PSWUSettings.Add("Properties", new ServiceManager());
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