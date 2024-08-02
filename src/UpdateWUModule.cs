using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security.Principal;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Update PSWindowsUpdate module.</para>
    /// <para type="description">Use Use Update-WUModule cmdlet to remote update PSWindowsUpdate module.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Update PSWindowsUpdate module from PSGallery</para>
    ///
    /// Update-WUModule -ComputerName MG-PC -Online
    ///
    /// </code>
    /// </example>
    /// <example>
    /// <code>
    /// <para>Update PSWindowsUpdate module from current serwer. Require SMB connection to destination machine.</para>
    ///
    /// Update-WUModule -ComputerName MG-PC -Local
    ///
    /// </code>
    /// </example>
    [Cmdlet("Update", "WUModule", ConfirmImpact = ConfirmImpact.High, DefaultParameterSetName = "Online", SupportsShouldProcess = true)]
    public class UpdateWUModule : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "Local", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [Parameter(ParameterSetName = "Online", ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
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
        private Hashtable PSWUSettings
        {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Update from PSGallery.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "Online")]
        public SwitchParameter Online { get; set; }

        /// <summary>
        /// <para type="description">Update from current module.</para>
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "Local")]
        public SwitchParameter Local { get; set; }

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

            ComputerName =
            [
                Environment.MachineName
            ];
        }

        private void CoreProcessing()
        {
            if (Local)
            {
                var pswuModule = WUToolsObj.GetPSWUModule(Environment.MachineName);
                try
                {
                    var version1 = (Version)pswuModule.Properties["Version"].Value;
                    var version2 = (Version)pswuModule.Properties["PSWUDllVersion"].Value;
                    WriteVerbose(Environment.MachineName + ": Local module: " + version1 + " (" + version2 + ")");
                }
                catch
                {
                }
            }

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

                var pswuModule1 = WUToolsObj.GetPSWUModule(target);
                try
                {
                    var version3 = (Version)pswuModule1.Properties["Version"].Value;
                    var version4 = (Version)pswuModule1.Properties["PSWUDllVersion"].Value;
                    WriteVerbose(target + ": Remote module: " + version3 + " (" + version4 + ")");
                }
                catch
                {
                    WriteDebug(DateTime.Now + " Can't get PSWU module from " + target);
                }

                if (ShouldProcess(target, "(" + DateTime.Now + ") Update PSWindowsUpdate module"))
                {
                    string script;
                    if (Online)
                    {
                        WriteVerbose("Installing module...");
                        script =
                            "Invoke-Command -ComputerName $Computer -Command { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force; Install-Module -Name PSWindowsUpdate -Force -Verbose } -Authentication Negotiate"
                                .Replace("$Computer", target);
                        WriteDebug(DateTime.Now + " Online mode: " + script);
                    }
                    else
                    {
                        WriteVerbose("Cloning module...");
                        script =
                            "$ModuleBase = (Get-Module PSWindowsUpdate -ListAvailable | Sort Version -Descending | Select -First 1).ModuleBase\r\n$Destination = '\\\\$Computer\\' + $($ModuleBase -replace ':','$') \r\n\r\nCopy-Item $ModuleBase -Destination $Destination -Force\r\nCopy-Item ($ModuleBase + '\\*') -Destination $Destination -Force\r\n\r\nInvoke-Command -ComputerName $Computer -Command { ipmo PSWindowsUpdate; Enable-WURemoting -Verbose } -Authentication Negotiate"
                                .Replace("$Computer", target);
                        WriteDebug(DateTime.Now + " Local mode: " + script);
                    }

                    var errorRecord = (ErrorRecord)null;
                    using (var powerShell = PowerShell.Create())
                    {
                        powerShell.AddScript(script);
                        powerShell.Invoke();
                        WriteObject(powerShell.Streams.Verbose);
                        WriteObject(powerShell.Streams.Error);
                        if (powerShell.Streams.Error.Count > 0)
                        {
                            errorRecord = powerShell.Streams.Error[0];
                        }
                    }

                    var pswuModule2 = WUToolsObj.GetPSWUModule(target);
                    try
                    {
                        var version5 = (Version)pswuModule2.Properties["Version"].Value;
                        var version6 = (Version)pswuModule2.Properties["PSWUDllVersion"].Value;
                        WriteVerbose(target + ": Remote module: " + version5 + " (" + version6 + ")");
                    }
                    catch
                    {
                        WriteDebug(DateTime.Now + " Can't get PSWU module from " + target);
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