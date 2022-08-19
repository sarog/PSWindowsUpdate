using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.IO;
using System.Management.Automation;
using System.Security.Principal;

namespace PSWindowsUpdate {
    [Cmdlet("Reset", "WUComponents", ConfirmImpact = ConfirmImpact.Medium, SupportsShouldProcess = true)]
    public class ResetWUComponents : PSCmdlet {
        private Hashtable _PSWUSettings = new Hashtable();

        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        private string[] ComputerName { get; set; }

        [Parameter]
        private PSCredential Credential { get; set; }

        [Parameter]
        private SwitchParameter SendReport { get; set; }

        [Parameter]
        private Hashtable PSWUSettings {
            get => _PSWUSettings;
            set => _PSWUSettings = value;
        }

        [Parameter]
        private SwitchParameter Debuger { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now.ToString() + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, (object)null));
            }

            WUToolsObj = new WUTools();
            OutputObj = new Collection<PSObject>();
            if ((bool)SendReport) {
                WriteDebug(DateTime.Now.ToString() + " Test smtp settings");
                if (!PSWUSettings.ContainsKey((object)"Properties")) {
                    PSWUSettings.Add((object)"Properties", (object)new string[1] {
                        "*"
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
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 1: Stop Windows Update services");
                Console.ForegroundColor = ConsoleColor.White;
                WUToolsObj.StopService(Computer, "BITS");
                WriteVerbose("Background Intelligent Transfer Service (BITS)");
                WUToolsObj.StopService(Computer, "wuauserv");
                WriteVerbose("Windows Update (wuauserv)");
                WUToolsObj.StopService(Computer, "appidsvc");
                WriteVerbose("Application Identity (appidsvc)");
                WUToolsObj.StopService(Computer, "cryptsvc");
                WriteVerbose("Cryptographic Services (cryptsvc)");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 2: Delete the qmgr*.dat files");
                Console.ForegroundColor = ConsoleColor.White;
                var environmentVariable1 = Environment.GetEnvironmentVariable("ALLUSERSPROFILE");
                try {
                    foreach (var file in Directory.GetFiles(environmentVariable1 + "\\Microsoft\\Network\\Downloader", "qmgr*.dat")) {
                        WriteVerbose("Deleting the " + file + " files.");
                        File.Delete(file);
                    }
                } catch { }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 3: Backup softare distribution folders");
                Console.ForegroundColor = ConsoleColor.White;
                var environmentVariable2 = Environment.GetEnvironmentVariable("SYSTEMROOT");
                var str1 = environmentVariable2 + "\\SoftwareDistribution";
                var str2 = environmentVariable2 + "\\SoftwareDistribution.bak";
                if (Directory.Exists(str1)) {
                    var path = str2;
                    var num = 1;
                    while (Directory.Exists(path)) {
                        path = str2 + num.ToString();
                        ++num;
                    }

                    var str3 = path;
                    Directory.Move(str1, str3);
                    if (Directory.Exists(str3)) {
                        WriteVerbose("Renaming Software Distribution folder to " + str3);
                    }
                }

                var str4 = environmentVariable2 + "\\System32\\Catroot2";
                var str5 = environmentVariable2 + "\\System32\\Catroot2.bak";
                if (Directory.Exists(str4)) {
                    var path = str5;
                    var num = 1;
                    while (Directory.Exists(path)) {
                        path = str5 + num.ToString();
                        ++num;
                    }

                    var str6 = path;
                    Directory.Move(str4, str6);
                    if (Directory.Exists(str6)) {
                        WriteVerbose("Renaming CatRoot  folder to " + str6);
                    }
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 4: Remove old Windows Update logs");
                Console.ForegroundColor = ConsoleColor.White;
                try {
                    foreach (var file in Directory.GetFiles(environmentVariable2, "WindowsUpdate.log")) {
                        WriteVerbose("Deleting the " + file + " files.");
                        File.Delete(file);
                    }
                } catch { }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 5: Reset Windows Update services");
                Console.ForegroundColor = ConsoleColor.White;
                WriteVerbose("Reset BITS service");
                WUToolsObj.RunProcess("sc.exe", "sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)");
                WriteVerbose("Reset Windows Update service");
                WUToolsObj.RunProcess("sc.exe",
                    "sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 6: Reregister dll's");
                Console.ForegroundColor = ConsoleColor.White;
                WriteVerbose("regsvr32.exe / s atl.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\atl.dll");
                WriteVerbose("regsvr32.exe / s urlmon.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\urlmon.dll");
                WriteVerbose("regsvr32.exe / s mshtml.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\mshtml.dll");
                WriteVerbose("regsvr32.exe / s shdocvw.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\shdocvw.dll");
                WriteVerbose("regsvr32.exe / s browseui.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\browseui.dll");
                WriteVerbose("regsvr32.exe / s jscript.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\jscript.dll");
                WriteVerbose("regsvr32.exe / s vbscript.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\vbscript.dll");
                WriteVerbose("regsvr32.exe / s scrrun.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\scrrun.dll");
                WriteVerbose("regsvr32.exe / s msxml.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\msxml.dll");
                WriteVerbose("regsvr32.exe / s msxml3.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\msxml3.dll");
                WriteVerbose("regsvr32.exe / s msxml6.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\msxml6.dll");
                WriteVerbose("regsvr32.exe / s actxprxy.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\actxprxy.dll");
                WriteVerbose("regsvr32.exe / s softpub.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\softpub.dll");
                WriteVerbose("regsvr32.exe / s wintrust.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wintrust.dll");
                WriteVerbose("regsvr32.exe / s dssenh.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\dssenh.dll");
                WriteVerbose("regsvr32.exe / s rsaenh.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\rsaenh.dll");
                WriteVerbose("regsvr32.exe / s gpkcsp.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\gpkcsp.dll");
                WriteVerbose("regsvr32.exe / s sccbase.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\sccbase.dll");
                WriteVerbose("regsvr32.exe / s slbcsp.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\slbcsp.dll");
                WriteVerbose("regsvr32.exe / s cryptdlg.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\cryptdlg.dll");
                WriteVerbose("regsvr32.exe / s oleaut32.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\oleaut32.dll");
                WriteVerbose("regsvr32.exe / s ole32.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\ole32.dll");
                WriteVerbose("regsvr32.exe / s shell32.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\shell32.dll");
                WriteVerbose("regsvr32.exe / s initpki.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\initpki.dll");
                WriteVerbose("regsvr32.exe / s wuapi.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wuapi.dll");
                WriteVerbose("regsvr32.exe / s wuaueng.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wuaueng.dll");
                WriteVerbose("regsvr32.exe / s wuaueng1.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wuaueng1.dll");
                WriteVerbose("regsvr32.exe / s wucltui.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wucltui.dll");
                WriteVerbose("regsvr32.exe / s wups.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wups.dll");
                WriteVerbose("regsvr32.exe / s wups2.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wups2.dll");
                WriteVerbose("regsvr32.exe / s wuweb.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wuweb.dll");
                WriteVerbose("regsvr32.exe / s qmgr.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\qmgr.dll");
                WriteVerbose("regsvr32.exe / s qmgrprxy.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\qmgrprxy.dll");
                WriteVerbose("regsvr32.exe / s wucltux.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wucltux.dll");
                WriteVerbose("regsvr32.exe / s muweb.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\muweb.dll");
                WriteVerbose("regsvr32.exe / s wuwebv.dll");
                WUToolsObj.RunProcess("regsvr32.exe", "/s " + environmentVariable2 + "\\wuwebv.dll");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 7: Reset WinSock");
                Console.ForegroundColor = ConsoleColor.White;
                WriteVerbose("netsh winsock reset");
                WUToolsObj.RunProcess("netsh.exe", "winsock reset");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 8: Reset Proxy");
                Console.ForegroundColor = ConsoleColor.White;
                WriteVerbose("netsh winhttp reset proxy");
                WUToolsObj.RunProcess("netsh.exe", "winhttp reset proxy");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 9: Start Windows Update services");
                Console.ForegroundColor = ConsoleColor.White;
                WUToolsObj.StartService(Computer, "cryptsvc");
                WriteVerbose("Cryptographic Services (cryptsvc)");
                WUToolsObj.StartService(Computer, "appidsvc");
                WriteVerbose("Application Identity (appidsvc)");
                WUToolsObj.StartService(Computer, "wuauserv");
                WriteVerbose("Windows Update (wuauserv)");
                WUToolsObj.StartService(Computer, "BITS");
                WriteVerbose("Background Intelligent Transfer Service (BITS)");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Step 10: Start Windows Update services");
                Console.ForegroundColor = ConsoleColor.White;
                WriteVerbose("wuauclt /resetauthorization /detectnow");
                WUToolsObj.RunProcess("wuauclt.exe", "/resetauthorization /detectnow");
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
                    PSWUSettings.Add((object)"Properties", (object)"*");
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