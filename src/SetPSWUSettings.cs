using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Security;
using System.Security.Principal;

namespace PSWindowsUpdate
{
    /// <summary>
    /// <para type="synopsis">Save PSWUSettings.</para>
    /// <para type="description">Use Set-PSWUSettings save PSWindowsUpdate module settings to XML file.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Set Office 365 as smtp server for PSWindowsUpdate module.</para>
    ///
    /// Set-PSWUSettings -SmtpServer smtp.office365.com -SmtpPort 587 -SmtpEnableSsl $true -SmtpSubject "PSWindowsUpdate Report" -SmtpTo mgajda@psmvp.pl -SmtpFrom mgajda@psmvp.pl -SmtpCredential (Get-Credential mgajda@psmvp.pl)
    ///
    /// </code>
    /// </example>
    [Cmdlet("Set", "PSWUSettings", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    public class SetPSWUSettings : PSCmdlet
    {
        private Hashtable _PSWUSettings = new Hashtable();
        private int _Port = 25;

        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        private string[] ComputerName { get; set; }

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
        private Hashtable PSWUSettings
        {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        /// <summary>
        /// <para type="description">Save smtp server to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public string SmtpServer { get; set; }

        /// <summary>
        /// <para type="description">Save smtp port to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public int SmtpPort
        {
            get => _Port;
            set => _Port = value;
        }

        /// <summary>
        /// <para type="description">Save enable ssl to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public bool SmtpEnableSsl { get; set; }

        /// <summary>
        /// <para type="description">Save smtp credential to Credential Manager.</para>
        /// </summary>
        [Parameter]
        public PSCredential SmtpCredential { get; set; }

        /// <summary>
        /// <para type="description">Save smtp recipient to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public string SmtpTo { get; set; }

        /// <summary>
        /// <para type="description">Save smtp sernder to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public string SmtpFrom { get; set; }

        /// <summary>
        /// <para type="description">Save alternative message subject to PSWUSettings file.</para>
        /// </summary>
        [Parameter]
        public string SmtpSubject { get; set; }

        /// <summary>
        /// <para type="description">Alternative report message propertie.s</para>
        /// </summary>
        [Parameter]
        public string Properties { get; set; }

        /// <summary>
        /// <para type="description">Alternative report message format style.</para>
        /// </summary>
        [ValidateSet("Table", "List", IgnoreCase = true)]
        [Parameter]
        public string Style { get; set; }

        /// <summary>
        /// <para type="description">Invoke-WUJob to save credential as system user</para>
        /// </summary>
        [Parameter]
        public string SaveAsSystem { get; set; }

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
                    PSWUSettings.Add("Properties", "*");
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
            var collection = new Collection<PSObject>();
            foreach (var target in ComputerName)
            {
                if (ShouldProcess(target, "Set PSWindowsUpdate module settings"))
                {
                    var num1 = 1;
                    var dateTime1 = new DateTime(1970, 1, 1);
                    var now = DateTime.Now;
                    var universalTime = now.ToUniversalTime();
                    TimeSpan timeSpan = universalTime - dateTime1;
                    var Seed1 = (int) Math.Floor((double) ((long) timeSpan.TotalMinutes / num1));
                    timeSpan = universalTime.AddMinutes(-1 * num1) - dateTime1;
                    var Seed2 = (int) Math.Floor((double) ((long) timeSpan.TotalMinutes / num1));
                    var randPass1 = new Random(Seed1).Next();
                    var randPass2 = new Random(Seed2).Next();
                    now = DateTime.Now;
                    WriteDebug(now + " " + randPass1);
                    now = DateTime.Now;
                    WriteDebug(now + " " + randPass2);
                    var psObject = new PSObject();
                    psObject.Properties.Add(new PSNoteProperty("ComputerName", target));
                    if (MyInvocation.BoundParameters.ContainsKey("SaveAsSystem"))
                    {
                        now = DateTime.Now;
                        WriteDebug(now + " Catched encrypted string: " + SaveAsSystem);
                        string str2;
                        try
                        {
                            str2 = WUTools.DecryptString(SaveAsSystem, randPass1.ToString());
                        }
                        catch (Exception ex1)
                        {
                            WriteDebug(DateTime.Now + " Cant decrypt with current secret.");
                            try
                            {
                                str2 = WUTools.DecryptString(SaveAsSystem, randPass2.ToString());
                            }
                            catch (Exception ex2)
                            {
                                WriteDebug(DateTime.Now + " Cant decrypt with previous secret. Job is too old.");
                                WriteError(new ErrorRecord(new Exception("Cant decrypt with previous secret. Job is too old."), "Decrypt",
                                    ErrorCategory.CloseError, null));
                                break;
                            }
                        }

                        psObject.Properties.Add(new PSNoteProperty("SaveAsSystem", SaveAsSystem));
                        var strArray = str2.Split('|');
                        var userName = strArray[0];
                        var str3 = strArray[1];
                        now = DateTime.Now;
                        WriteDebug(now + " Decrypted UserName: " + userName);
                        var password = new SecureString();
                        foreach (var c in str3.ToCharArray())
                        {
                            password.AppendChar(c);
                        }

                        SmtpCredential = new PSCredential(userName, password);
                    }

                    if (SmtpCredential != null)
                    {
                        now = DateTime.Now;
                        WriteDebug(now + " Save credential to Credential Manager");
                        WUToolsObj.SaveCredential(SmtpCredential.GetNetworkCredential().UserName,
                            SmtpCredential.GetNetworkCredential().Password);
                        if (!MyInvocation.BoundParameters.ContainsKey("SaveAsSystem"))
                        {
                            var str4 = WUTools.EncryptString(
                                SmtpCredential.GetNetworkCredential().UserName + "|" + SmtpCredential.GetNetworkCredential().Password,
                                randPass1.ToString());
                            now = DateTime.Now;
                            WriteDebug(now + " Encrypted string to SaveAsSystem: " + str4);
                            var str5 = "$DebugPreference = 'continue'; Set-PSWUSettings -SaveAsSystem '" + str4 + "'";
                            if (SendReport)
                            {
                                str5 += " -SendReport";
                            }

                            var str6 = str5 + " -Verbose *>&1 | Out-File $Env:TEMP\\PSWindowsUpdate.log";
                            var invokeWuJob1 = new InvokeWUJob();
                            invokeWuJob1.ComputerName =
                            [
                                target
                            ];
                            if (Credential != null)
                            {
                                invokeWuJob1.Credential = Credential;
                            }

                            invokeWuJob1.Script = str6;
                            invokeWuJob1.TaskName = "PSWindowsUpdate_Settings";
                            var invokeWuJob2 = invokeWuJob1;
                            now = DateTime.Now;
                            var dateTime2 = now.AddMinutes(2.0);
                            invokeWuJob2.EndBoundary = dateTime2;
                            invokeWuJob1.RunNow = true;
                            WriteVerbose("Invoke-WUJob: " + invokeWuJob1.TaskName + " " + target + " (Now)");
                            WriteVerbose("powershell.exe -Command \"" + str6 + "\"");
                            foreach (var sendToPipeline in invokeWuJob1.Invoke())
                            {
                                WriteObject(sendToPipeline);
                            }
                        }

                        psObject.Properties.Add(new PSNoteProperty("SmtpCredential", SmtpCredential.GetNetworkCredential().UserName));
                    }

                    var flag = false;
                    var pswuSettings = WUToolsObj.GetPSWUSettings();
                    if (MyInvocation.BoundParameters.ContainsKey("SmtpServer"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("SmtpServer"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change SmtpServer: ";
                            strArray[2] = pswuSettings["SmtpServer"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = SmtpServer;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["SmtpServer"] = SmtpServer;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add SmtpServer: " + SmtpServer);
                            pswuSettings.Add("SmtpServer", SmtpServer);
                        }

                        psObject.Properties.Add(new PSNoteProperty("SmtpServer", SmtpServer));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SmtpPort"))
                    {
                        flag = true;
                        int smtpPort;
                        if (pswuSettings.ContainsKey("Port"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change Port: ";
                            strArray[2] = pswuSettings["Port"]?.ToString();
                            strArray[3] = " -> ";
                            smtpPort = SmtpPort;
                            strArray[4] = smtpPort.ToString();
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["Port"] = SmtpPort;
                        }
                        else
                        {
                            now = DateTime.Now;
                            var str7 = now.ToString();
                            smtpPort = SmtpPort;
                            var str8 = smtpPort.ToString();
                            WriteDebug(str7 + " Add Port: " + str8);
                            pswuSettings.Add("Port", SmtpPort);
                        }

                        psObject.Properties.Add(new PSNoteProperty("Port", SmtpPort));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SmtpEnableSsl"))
                    {
                        flag = true;
                        bool smtpEnableSsl;
                        if (pswuSettings.ContainsKey("EnableSsl"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change Port: ";
                            strArray[2] = pswuSettings["EnableSsl"]?.ToString();
                            strArray[3] = " -> ";
                            smtpEnableSsl = SmtpEnableSsl;
                            strArray[4] = smtpEnableSsl.ToString();
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["EnableSsl"] = SmtpEnableSsl;
                        }
                        else
                        {
                            now = DateTime.Now;
                            var str9 = now.ToString();
                            smtpEnableSsl = SmtpEnableSsl;
                            var str10 = smtpEnableSsl.ToString();
                            WriteDebug(str9 + " Add Port: " + str10);
                            pswuSettings.Add("EnableSsl", SmtpEnableSsl);
                        }

                        psObject.Properties.Add(new PSNoteProperty("EnableSsl", SmtpEnableSsl));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SmtpTo"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("To"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change To: ";
                            strArray[2] = pswuSettings["To"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = SmtpTo;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["To"] = SmtpTo;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add To: " + SmtpTo);
                            pswuSettings.Add("To", SmtpTo);
                        }

                        psObject.Properties.Add(new PSNoteProperty("To", SmtpTo));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SmtpFrom"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("From"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change From: ";
                            strArray[2] = pswuSettings["From"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = SmtpFrom;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["From"] = SmtpFrom;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add From: " + SmtpFrom);
                            pswuSettings.Add("From", SmtpFrom);
                        }

                        psObject.Properties.Add(new PSNoteProperty("From", SmtpFrom));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("SmtpSubject"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("Subject"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change Subject: ";
                            strArray[2] = pswuSettings["Subject"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = SmtpSubject;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["Subject"] = SmtpSubject;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add Subject: " + SmtpSubject);
                            pswuSettings.Add("Subject", SmtpSubject);
                        }

                        psObject.Properties.Add(new PSNoteProperty("Subject", SmtpSubject));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("Properties"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("Properties"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change Properties: ";
                            strArray[2] = pswuSettings["Properties"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = Properties;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["Properties"] = Properties;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add Properties: " + Properties);
                            pswuSettings.Add("Properties", Properties);
                        }

                        psObject.Properties.Add(new PSNoteProperty("Properties", Properties));
                    }

                    if (MyInvocation.BoundParameters.ContainsKey("Style"))
                    {
                        flag = true;
                        if (pswuSettings.ContainsKey("Style"))
                        {
                            var strArray = new string[5];
                            now = DateTime.Now;
                            strArray[0] = now.ToString();
                            strArray[1] = " Change Style: ";
                            strArray[2] = pswuSettings["Style"]?.ToString();
                            strArray[3] = " -> ";
                            strArray[4] = Style;
                            WriteDebug(string.Concat(strArray));
                            pswuSettings["Style"] = Style;
                        }
                        else
                        {
                            now = DateTime.Now;
                            WriteDebug(now + " Add Style: " + Style);
                            pswuSettings.Add("Style", Style);
                        }

                        psObject.Properties.Add(new PSNoteProperty("Style", Style));
                    }

                    if (flag)
                    {
                        WUToolsObj.SetPSWUSettings(pswuSettings);
                        WriteVerbose("Save PSWUSetting to file: " + WUToolsObj.PSWUModulePath() + "\\PSWUSettings.xml");
                    }

                    collection.Add(psObject);
                }

                WriteObject(collection);
                OutputObj = new Collection<PSObject>(collection);
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