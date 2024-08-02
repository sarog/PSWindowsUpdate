using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;
using Microsoft.Win32;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Get Windows Update Client settings.</para>
    /// <para type="description">Use Get-WUSettings cmdlet to get current configuration of Windows Update Client.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Get current Windows Update Client configuration.</para>
    ///
    /// Get-WUSettings
    ///
    /// <para>ComputerName                  : MG-PC</para>
    /// <para>AcceptTrustedPublisherCerts   : 0</para>
    /// <para>WUServer                      : https://wsus.commandlinegeeks.com</para>
    /// <para>WUStatusServer                : https://wsus.commandlinegeeks.com</para>
    /// <para>DetectionFrequencyEnabled     : 1</para>
    /// <para>DetectionFrequency            : 2</para>
    /// <para>NoAutoRebootWithLoggedOnUsers : 1</para>
    /// <para>RebootRelaunchTimeoutEnabled  : 1</para>
    /// <para>RebootRelaunchTimeout         : 240</para>
    /// <para>IncludeRecommendedUpdates     : 0</para>
    /// <para>NoAutoUpdate                  : 0</para>
    /// <para>AUOptions                     : 2 - Notify before download</para>
    /// <para>ScheduledInstallDay           : 0 - Every Day</para>
    /// <para>ScheduledInstallTime          : 4</para>
    /// <para>UseWUServer                   : 1</para>
    /// </code>
    /// </example>
    [Cmdlet("Get", "WUSettings", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    [OutputType(typeof(WUSettings))]
    public class GetWUSettings : PSCmdlet
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
        /// <para type="description">Debuger return original exceptions.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

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

            ComputerName = new string[1]
            {
                Environment.MachineName
            };
        }

        private void CoreProcessing()
        {
            var invocationName = MyInvocation.InvocationName;
            var hashtable1 = new Hashtable();
            hashtable1.Add(0, "0 - Not configured");
            hashtable1.Add(1, "1 - Disabled");
            hashtable1.Add(2, "2 - Notify before download");
            hashtable1.Add(3, "3 - Notify before installation");
            hashtable1.Add(4, "4 - Scheduled installation");
            hashtable1.Add(5, "5 - Users configure");
            var hashtable2 = new Hashtable();
            hashtable2.Add(0, "0 - Every Day");
            hashtable2.Add(1, "1 - Every Sunday");
            hashtable2.Add(2, "2 - Every Monday");
            hashtable2.Add(3, "3 - Every Tuesday");
            hashtable2.Add(4, "4 - Every Wednesday");
            hashtable2.Add(5, "5 - Every Thursday");
            hashtable2.Add(6, "6 - Every Friday");
            hashtable2.Add(7, "7 - Every Saturday");

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

                if (ShouldProcess(target, "(" + DateTime.Now + ") Get Windows Update settings"))
                {
                    var registryKey1 = !WUToolsObj.IsLocalHost(target)
                        ? RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, target)
                        : RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\");
                    var registryKey3 = registryKey1.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU\\");
                    if (registryKey2 != null)
                    {
                        WriteVerbose("Some settings are managed by your system administrator.");
                        var sendToPipeline = new PSObject();
                        sendToPipeline.Properties.Add(new PSNoteProperty("ComputerName", target));
                        foreach (var valueName in registryKey2.GetValueNames())
                        {
                            var obj = registryKey2.GetValue(valueName);
                            sendToPipeline.Properties.Add(new PSNoteProperty(valueName, obj));
                        }

                        foreach (var valueName in registryKey3.GetValueNames())
                        {
                            var key = registryKey3.GetValue(valueName);
                            var str2 = valueName;
                            if (!(str2 == "AUOptions"))
                            {
                                if (str2 == "ScheduledInstallDay")
                                {
                                    key = hashtable2[key];
                                }
                            }
                            else
                            {
                                key = hashtable1[key];
                            }

                            sendToPipeline.Properties.Add(new PSNoteProperty(valueName, key));
                        }

                        WriteObject(sendToPipeline, true);
                        OutputObj.Add(sendToPipeline);
                    }
                    else
                    {
                        WriteVerbose("Settings not found.");
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