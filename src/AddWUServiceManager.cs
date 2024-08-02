using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using WUApiLib;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Register a new Windows Update API Service Manager.</para>
    /// <para type="description">Use Add-WUServiceManager cmdlet to register new Windows Update Service Manager.</para>
    /// <para type="description">It's combination old 'Add-WUServiceManager' and 'Add-WUOfflineSync' functions to register online  and offline ServiceManager </para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Try register Microsoft Update Service by custom ServiceID.</para>
    ///
    /// Add-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d"
    ///
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Register Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    ///
    /// <para>ServiceID                            IsManaged IsDefault Name</para>
    /// <para>---------                            --------- --------- ----</para>
    /// <para>7971f918-a847-4430-9279-4a52d1efe18d False     False     Microsoft Update</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Try register Microsoft Update service as Service Manager.</para>
    ///
    /// Add-WUServiceManager -MicrosoftUpdate
    ///
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Register Windows Update Service Manager: 7971f918-a847-4430-9279-4a52d1efe18d" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    ///
    /// <para>ServiceID                            IsManaged IsDefault Name</para>
    /// <para>---------                            --------- --------- ----</para>
    /// <para>7971f918-a847-4430-9279-4a52d1efe18d False     False     Microsoft Update</para>
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Try register Offline Sync Service from file C:\wsusscn2.cab.</para>
    ///
    /// Add-WUServiceManager -ScanFileLocation C:\wsusscn2.cab
    ///
    /// <para>Confirm</para>
    /// <para>Are you sure you want to perform this action?</para>
    /// <para>Performing the operation "Register Offline Windows Update Service Manager: C:\wsusscn2.cab" on target "MG-PC".</para>
    /// <para>[Y] Yes[A] Yes to All  [N] No[L] No to All  [S] Suspend[?] Help (default is "Y"): Y</para>
    ///
    /// <para>ServiceID                            IsManaged IsDefault Name</para>
    /// <para>---------                            --------- --------- ----</para>
    /// <para>7e1364ef-e30e-4f4e-9c66-84194eebcbbe False     False     Offline Sync Service</para>
    /// </code>
    /// </example>
    [Cmdlet("Add", "WUServiceManager", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "LocalServiceID",
        SupportsShouldProcess = true)]
    [OutputType(typeof(ServiceManager))]
    public class AddWUServiceManager : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();
        private int _AddServiceFlag = 2;
        private string _ServiceName = "Offline Sync Service";

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter(ParameterSetName = "RemoteServiceID", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "MicrosoftUpdate", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
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
        /// <para type="description">An identifier for the service to be registered. </para>
        /// <para type="description">Examples Of ServiceID:
        /// \r\n
        /// \r\n -- Windows Update 				    9482f4b4-e343-43b6-b170-9a65bc822c77
        /// \r\n -- Microsoft Update 				7971f918-a847-4430-9279-4a52d1efe18d
        /// \r\n -- Windows Store 					117cab2d-82b1-4b5a-a08c-4d62dbee7782
        /// \r\n -- Windows Server Update Service 	3da21691-e39d-4da6-8a4b-b43877bcb1b7</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "LocalServiceID")]
        [Parameter(Mandatory = true, ParameterSetName = "RemoteServiceID")]
        [ValidateNotNullOrEmpty]
        public string ServiceID { get; set; }

        /// <summary>
        /// <para type="description">A combination of AddServiceFlag values:
        /// \r\n0x1 - asfAllowPendingRegistration
        /// \r\n0x2 - asfAllowOnlineRegistration
        /// \r\n0x4 - asfRegisterServiceWithAU</para>
        /// </summary>
        [Parameter(ParameterSetName = "LocalServiceID")]
        public int AddServiceFlag
        {
            get => _AddServiceFlag;
            set => _AddServiceFlag = value;
        }

        /// <summary>
        /// <para type="description">The path of the Microsoft signed local cabinet file (.cab) that has the information that is required for a service registration. If empty, the update agent searches for the authorization cabinet file (.cab) during service registration when a network connection is available.</para>
        /// </summary>
        [Parameter(ParameterSetName = "LocalServiceID")]
        public string AuthorizationCabPath { get; set; }

        /// <summary>
        /// <para type="description">Register Microsoft Update Service Manager - '7971f918-a847-4430-9279-4a52d1efe18d'</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "MicrosoftUpdate")]
        public SwitchParameter MicrosoftUpdate { get; set; }

        /// <summary>
        /// <para type="description">Name under which it will be registered Windows Update offline service. Default name is 'Offline Sync Service'.</para>
        /// </summary>
        [Parameter(ParameterSetName = "OfflineSync")]
        public string ServiceName
        {
            get => _ServiceName;
            set => _ServiceName = value;
        }

        /// <summary>
        /// <para type="description">Path to Windows Update offline scan file (wsusscan.cab or wsusscn2.cab).</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "OfflineSync")]
        public string ScanFileLocation { get; set; }

        /// <summary>
        /// <para type="description">Don't return output.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Silent { get; set; }

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
                ThrowTerminatingError(new ErrorRecord(
                    new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
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

            if (ComputerName == null)
            {
                ComputerName = new string[1]
                {
                    Environment.MachineName
                };
            }

            if (!MicrosoftUpdate)
            {
                return;
            }

            ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d";
        }

        private void CoreProcessing()
        {
            foreach (var target in ComputerName)
            {
                WriteDebug(DateTime.Now + " " + target + ": Connecting...");
                var pSWUModule = WUToolsObj.GetPSWUModule(target);
                WriteDebug(DateTime.Now + " Module version: " + pSWUModule.Properties["Version"].Value);
                WriteDebug(DateTime.Now + " Dll version: " + pSWUModule.Properties["PSWUDllVersion"].Value);
                var serviceManagerObj = WUToolsObj.GetWUApiServiceManagerObj(target);
                WriteDebug(DateTime.Now + " ServiceManagerObj mode: " + serviceManagerObj.Mode);
                if (serviceManagerObj.Status)
                {
                    ServiceManagerObj = (UpdateServiceManager)serviceManagerObj.Object;
                    IUpdateService updateService = null;
                    var value = "";
                    WUToolsObj.RestartService(target);
                    if (ParameterSetName == "OfflineSync")
                    {
                        if (ShouldProcess(target,
                                "(" + DateTime.Now + ") Register Offline Windows Update Service Manager: " + ScanFileLocation))
                        {
                            try
                            {
                                updateService = ServiceManagerObj.AddScanPackageService(ServiceName, ScanFileLocation, 1);
                                value = "Registered";
                            }
                            catch (COMException ex)
                            {
                                var wUApiCodeDetails = WUToolsObj.GetWUApiCodeDetails(ex.ErrorCode);
                                if (wUApiCodeDetails != null)
                                {
                                    if (wUApiCodeDetails.CodeType == 2)
                                    {
                                        WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult,
                                            ErrorCategory.CloseError, null));
                                    }
                                }
                                else if (Debuger)
                                {
                                    ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                                }

                                continue;
                            }
                        }
                    }
                    else if (ShouldProcess(target, "(" + CmdletStart + ") Register Windows Update Service Manager: " + ServiceID))
                    {
                        try
                        {
                            var updateServiceRegistration = ServiceManagerObj.AddService2(ServiceID, AddServiceFlag, AuthorizationCabPath);
                            updateService = updateServiceRegistration.Service;
                            switch (updateServiceRegistration.RegistrationState)
                            {
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
                                ThrowTerminatingError(new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null));
                            }

                            continue;
                        }
                    }

                    WUToolsObj.RestartService(target);
                    var pSObject = new PSObject(updateService);
                    pSObject.Properties.Add(new PSNoteProperty("RegistrationStateName", value));
                    pSObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                    pSObject.TypeNames.Clear();
                    pSObject.TypeNames.Add("PSWindowsUpdate.ServiceManager");
                    OutputObj.Add(pSObject);
                    if (!Silent)
                    {
                        WriteObject(pSObject, true);
                    }
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