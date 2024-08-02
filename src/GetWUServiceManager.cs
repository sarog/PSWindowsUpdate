using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Get Service Manager configuration.</para>
    /// <para type="description">Use Get-WUServiceManager cmdlet to get available configuration of update services.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Check currently available Windows Update Services on machine.</para>
    ///
    /// Get-WUServiceManager
    ///
    /// <para>ServiceID                            IsManaged IsDefault Name</para>
    /// <para>---------                            --------- --------- ----</para>
    /// <para>9482f4b4-e343-43b6-b170-9a65bc822c77 False     False     Windows Update</para>
    /// <para>7971f918-a847-4430-9279-4a52d1efe18d False     False     Microsoft Update</para>
    /// <para>3da21691-e39d-4da6-8a4b-b43877bcb1b7 True      True      Windows Server Update Service</para>
    /// <para>13df3d8f-78d7-4eb8-bb9c-2a101870d350 False     False     Offline Sync Service2</para>
    /// <para>a8f3b5e6-fb1f-4814-a047-2257d39c2460 False     False     Offline Sync Service</para>
    /// </code>
    /// </example>
    [Cmdlet("Get", "WUServiceManager", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(ServiceManager))]
    public class GetWUServiceManager : PSCmdlet
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
        /// <para type="description">Get specific Service Manager if it's available.</para>
        /// </summary>
        [Parameter]
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
        protected override void BeginProcessing()
        {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if (SendReport)
            {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties"))
                {
                    PSWUSettings.Add("Properties", new ServiceManager());
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

            ComputerName = new string[1]
            {
                Environment.MachineName
            };
        }

        private void CoreProcessing()
        {
            foreach (var target in ComputerName)
            {
                var now = DateTime.Now;
                WriteDebug(now + " " + target + ": Connecting...");
                try
                {
                    var pswuModule = WUToolsObj.GetPSWUModule(target);
                    now = DateTime.Now;
                    WriteDebug(now + " Module version: " + pswuModule.Properties["Version"].Value);
                    now = DateTime.Now;
                    WriteDebug(now + " Dll version: " + pswuModule.Properties["PSWUDllVersion"].Value);
                }
                catch
                {
                }

                now = DateTime.Now;
                var action = "(" + now + ") Get Windows Update ServiceManager";
                if (ShouldProcess(target, action))
                {
                    var serviceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                    now = DateTime.Now;
                    WriteDebug(now + " ServiceManagerObj mode: " + serviceManagerObj.Mode);
                    if (serviceManagerObj.Status)
                    {
                        ServiceManagerObj = (UpdateServiceManager)serviceManagerObj.Object;
                        var source = new Collection<PSObject>();
                        foreach (IUpdateService service in ServiceManagerObj.Services)
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " ";
                            strArray[2] = target;
                            strArray[3] = ": Found service ";
                            strArray[4] = service.Name;
                            WriteDebug(string.Concat(strArray));
                            var psObject = new PSObject(service);
                            psObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                            psObject.TypeNames.Clear();
                            psObject.TypeNames.Add("PSWindowsUpdate.ServiceManager");
                            source.Add(psObject);
                        }

                        var collection1 = new Collection<PSObject>();
                        var collection2 = ServiceID == null
                            ? source
                            : new Collection<PSObject>(source.Where(x => x.Properties["ServiceID"].Value.ToString() == ServiceID)
                                .ToList());
                        WriteObject(collection2, true);
                        OutputObj = new Collection<PSObject>(OutputObj.Concat(collection2).ToList());
                    }
                    else if (Debuger)
                    {
                        WriteError(new ErrorRecord(serviceManagerObj.Exception, "Debug", ErrorCategory.CloseError, null));
                    }
                    else
                    {
                        WriteError(serviceManagerObj.Error);
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
                    PSWUSettings.Add("Properties", new ServiceManager());
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