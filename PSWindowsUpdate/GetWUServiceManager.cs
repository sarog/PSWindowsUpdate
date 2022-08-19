using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUServiceManager", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(new Type[] { typeof(ServiceManager) })]
    public class GetWUServiceManager : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string[] ComputerName { get; set; }

        [Parameter]
        public PSCredential Credential { get; set; }

        [Parameter]
        public SwitchParameter SendReport { get; set; }

        [Parameter]
        public Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter]
        public string ServiceID { get; set; }

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
            WriteDebug(DateTime.Now.ToString() + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if ((bool)SendReport) {
                WriteDebug(DateTime.Now.ToString() + " Test smtp settings");
                if (!PSWUSettings.ContainsKey((object)"Properties")) {
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "Name",
                        "ServiceID",
                        "IsManaged",
                        "IsDefaultAUService"
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
            foreach (var Computer in ComputerName) {
                var now = DateTime.Now;
                WriteDebug(now.ToString() + " " + Computer + ": Connecting...");
                try {
                    var pswuModule = WUToolsObj.GetPSWUModule(Computer);
                    now = DateTime.Now;
                    WriteDebug(now.ToString() + " Module version: " + pswuModule.Properties["Version"].Value?.ToString());
                    now = DateTime.Now;
                    WriteDebug(now.ToString() + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value?.ToString());
                } catch { }

                var target = Computer;
                now = DateTime.Now;
                var action = "(" + now.ToString() + ") Get Windows Update ServiceManager";
                if (ShouldProcess(target, action)) {
                    var serviceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(Computer);
                    now = DateTime.Now;
                    WriteDebug(now.ToString() + " ServiceManagerObj mode: " + serviceManagerObj.Mode);
                    if (serviceManagerObj.Status) {
                        ServiceManagerObj = (UpdateServiceManager)serviceManagerObj.Object;
                        var source = new Collection<PSObject>();
                        foreach (IUpdateService service in ServiceManagerObj.Services) {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " ";
                            strArray[2] = Computer;
                            strArray[3] = ": Found service ";
                            strArray[4] = service.Name;
                            WriteDebug(string.Concat(strArray));
                            var psObject = new PSObject((object)service);
                            psObject.Properties.Add((PSPropertyInfo)new PSNoteProperty("ComputerName", (object)Computer));
                            psObject.TypeNames.Clear();
                            psObject.TypeNames.Add("PSWindowsUpdate.ServiceManager");
                            source.Add(psObject);
                        }

                        var collection1 = new Collection<PSObject>();
                        var collection2 = ServiceID == null
                            ? source
                            : new Collection<PSObject>((IList<PSObject>)source.Where<PSObject>((Func<PSObject, bool>)(x => x.Properties["ServiceID"].Value.ToString() == ServiceID))
                                .ToList<PSObject>());
                        WriteObject((object)collection2, true);
                        OutputObj = new Collection<PSObject>((IList<PSObject>)OutputObj.Concat<PSObject>((IEnumerable<PSObject>)collection2).ToList<PSObject>());
                    } else if ((bool)Debuger) {
                        WriteError(new ErrorRecord(serviceManagerObj.Exception, "Debug", ErrorCategory.CloseError, (object)null));
                    } else {
                        WriteError(serviceManagerObj.Error);
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
                    PSWUSettings.Add((object)"Properties", (object)new string[5] {
                        "ComputerName",
                        "Name",
                        "ServiceID",
                        "IsManaged",
                        "IsDefaultAUService"
                    });
                }

                var psObject = WUToolsObj.SendMail(PSWUSettings, OutputObj, CmdletInfo);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                }
            }

            WriteDebug(DateTime.Now.ToString() + " CmdletEnd");
        }

        protected override void StopProcessing() {
            base.StopProcessing();
        }
    }
}