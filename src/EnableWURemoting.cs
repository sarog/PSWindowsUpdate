using Microsoft.Win32;
using NetFwTypeLib;
using System;
using System.Collections.ObjectModel;
using System.Management;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace PSWindowsUpdate {
    /// <summary>
    /// <para type="synopsis">Enable firewall rules for PSWindowsUpdate remoting.</para>
    /// <para type="description">Use Enable-WURemoting cmdlet to enable nessesery firewall rules for PSWindowsUpdate remoting.</para>
    /// </summary>
    /// <para type="link" uri="https://commandlinegeeks.wordpress.com/">Author Blog</para>
    /// <example>
    /// <code>
    /// <para>Enable firewall rules for PSWindowsUpdate remoting.</para>
    ///
    /// Enable-WURemoting -Verbose
    /// </code>
    /// </example>
    [Cmdlet("Enable", "WURemoting", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    public class EnableWURemoting : PSCmdlet {
        /// <summary>
        /// <para type="description">Specify one or more computer names for remote connection.</para>
        /// </summary>
        [Parameter(ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        private string[] ComputerName { get; set; }

        /// <summary>
        /// <para type="description">Specify alternative credential.</para>
        /// </summary>
        [Parameter]
        private PSCredential Credential { get; set; }

        /// <summary>
        /// <para type="description">Set LocalAccountTokenFilterPolicy registry entry to builds an elevated token on the target remote computer.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter LocalAccountTokenFilterPolicy { get; set; }

        /// <summary>
        /// <para type="description">Enable WinRM Public access from all subnets. Default access is only enabled from local subnet. Required for workgroup computers.</para>
        /// </summary>
        [Parameter]
        public SwitchParameter WinRMPublic { get; set; }

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
        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                ThrowTerminatingError(new ErrorRecord(new Exception("To perform operations you must run an elevated Windows PowerShell console."), "AccessDenied",
                    ErrorCategory.PermissionDenied, null));
            }

            if (ComputerName != null) {
                return;
            }

            ComputerName = new string[1] {
                Environment.MachineName
            };
        }

        private void CoreProcessing() {
            var wUTools = new WUTools();
            var invocationName = MyInvocation.InvocationName;
            var computerName = ComputerName;
            foreach (var target in computerName) {
                var netFwPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                var rule1Enabled = false;
                foreach (INetFwRule rule in netFwPolicy.Rules) {
                    if (rule.Name == "PSWindowsUpdate (RPC Dynamics Ports)") {
                        rule1Enabled = true;
                    }
                }

                if (!rule1Enabled) {
                    WriteVerbose("Create firewall rule: PSWindowsUpdate (RPC Dynamics Ports)");
                    var netFwRule2 = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                    netFwRule2.Grouping = "PSWindowsUpdate";
                    netFwRule2.Enabled = true;
                    netFwRule2.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                    netFwRule2.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                    netFwRule2.Protocol = 6;
                    netFwRule2.LocalPorts = "RPC";
                    netFwRule2.ApplicationName = "%SystemRoot%\\System32\\dllhost.exe";
                    netFwRule2.Name = "PSWindowsUpdate (RPC Dynamics Ports)";
                    netFwPolicy.Rules.Add(netFwRule2);
                } else if (!netFwPolicy.IsRuleGroupCurrentlyEnabled["PSWindowsUpdate"]) {
                    WriteVerbose("Enable firewall rule: PSWindowsUpdate (RPC Dynamics Ports)");
                    netFwPolicy.EnableRuleGroup(int.MaxValue, "PSWindowsUpdate", true);
                } else {
                    WriteDebug(DateTime.Now + " PSWindowsUpdate (RPC Dynamics Ports) firewall rule is enabled");
                }

                if (WinRMPublic) {
                    var rule2Enabled = false;
                    foreach (INetFwRule rule2 in netFwPolicy.Rules) {
                        if (rule2.Name == "PSWindowsUpdate (WinRM Public)") {
                            rule2Enabled = true;
                        }
                    }

                    if (!rule2Enabled) {
                        WriteVerbose("Create firewall rule: PSWindowsUpdate (WinRM Public)");
                        var netFwRule4 = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                        netFwRule4.Grouping = "PSWindowsUpdate";
                        netFwRule4.Enabled = true;
                        netFwRule4.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                        netFwRule4.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                        netFwRule4.Protocol = 6;
                        netFwRule4.LocalPorts = "5985";
                        netFwRule4.ApplicationName = "System";
                        netFwRule4.Name = "PSWindowsUpdate (WinRM Public)";
                        netFwPolicy.Rules.Add(netFwRule4);
                    } else if (!netFwPolicy.IsRuleGroupCurrentlyEnabled["PSWindowsUpdate"]) {
                        WriteVerbose("Enable firewall rule: PSWindowsUpdate (WinRM Public)");
                        netFwPolicy.EnableRuleGroup(int.MaxValue, "PSWindowsUpdate", true);
                    } else {
                        WriteDebug(DateTime.Now + " PSWindowsUpdate (WinRM Public) firewall rule is enabled");
                    }
                }

                if (!netFwPolicy.IsRuleGroupCurrentlyEnabled["Remote Scheduled Tasks Management"]) {
                    WriteVerbose("Enable firewall rules: Remote Scheduled Tasks Management");
                    netFwPolicy.EnableRuleGroup(int.MaxValue, "Remote Scheduled Tasks Management", true);
                } else {
                    WriteDebug(DateTime.Now + " Remote Scheduled Tasks Management firewall rules are enabled");
                }

                if (!netFwPolicy.IsRuleGroupCurrentlyEnabled["Windows Management Instrumentation (WMI)"]) {
                    WriteVerbose("Enable firewall rules: Windows Management Instrumentation (WMI)");
                    netFwPolicy.EnableRuleGroup(int.MaxValue, "Windows Management Instrumentation (WMI)", true);
                } else {
                    WriteDebug(DateTime.Now + " Windows Management Instrumentation (WMI) firewall rules are enabled");
                }

                if (LocalAccountTokenFilterPolicy) {
                    WriteVerbose("Set LocalAccountTokenFilterPolicy=1 registry entry");
                    var registryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                    var registryKey2 = registryKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\system\\", true);
                    registryKey2.SetValue("LocalAccountTokenFilterPolicy", Convert.ToInt32(1), RegistryValueKind.DWord);
                }

                try {
                    var script = "(Get-Item WSMan:\\localhost\\Client\\TrustedHosts).Value -match '\\*';";
                    Collection<PSObject> collection;
                    using (var powerShell = PowerShell.Create()) {
                        powerShell.AddScript(script);
                        collection = powerShell.Invoke();
                    }

                    if (collection[0].ToString() == "False") {
                        WriteVerbose("Enable PSRemoting");
                        script =
                            "Enable -PSRemoting –force; \r\nStart-Service WinRM; \r\nSet-ItemProperty -Path HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System -Name LocalAccountTokenFilterPolicy -Value 1 -Type DWord;\r\nSet-ItemProperty -Path HKLM:\\System\\CurrentControlSet\\Control\\Lsa -Name ForceGuest -Value 0;\r\nSet-Item WSMan:\\localhost\\Client\\TrustedHosts -Value * -Force -Concatenate;";
                        using var powerShell2 = PowerShell.Create();
                        powerShell2.AddScript(script);
                        collection = powerShell2.Invoke();
                    } else {
                        WriteDebug(DateTime.Now + " PSRemoting TrustedHosts are set");
                    }
                } catch (Exception exception) {
                    WriteError(new ErrorRecord(exception, "PSRemoting", ErrorCategory.ResourceUnavailable, null));
                }

                try {
                    var num = 1u;
                    var queryString = "SELECT * FROM Win32_Service WHERE Name = 'RemoteRegistry'";
                    var managementObjectSearcher = new ManagementObjectSearcher(queryString);
                    var managementObjectCollection = managementObjectSearcher.Get();
                    foreach (ManagementObject item in managementObjectCollection) {
                        if (Convert.ToString(item.GetPropertyValue("StartMode")) == "Disabled") {
                            WriteVerbose("Change startmode service: Remote Registry");
                            var methodParameters = item.GetMethodParameters("ChangeStartMode");
                            methodParameters["startmode"] = "Automatic";
                            var managementBaseObject = item.InvokeMethod("ChangeStartMode", methodParameters, null);
                            num = Convert.ToUInt16(managementBaseObject.Properties["ReturnValue"].Value);
                        } else {
                            WriteDebug(DateTime.Now + " Remote Registry service start mode is ok");
                        }
                    }
                } catch (Exception exception2) {
                    WriteError(new ErrorRecord(exception2, "RemoteRegistry", ErrorCategory.ResourceUnavailable, null));
                }

                try {
                    var num2 = 1u;
                    var queryString2 = "SELECT * FROM Win32_Service WHERE Name = 'RemoteRegistry'";
                    var managementObjectSearcher2 = new ManagementObjectSearcher(queryString2);
                    var managementObjectCollection2 = managementObjectSearcher2.Get();
                    foreach (ManagementObject item2 in managementObjectCollection2) {
                        if (Convert.ToString(item2.GetPropertyValue("State")) == "Stopped") {
                            WriteVerbose("Start service: Remote Registry");
                            var managementBaseObject2 = item2.InvokeMethod("StartService", null, null);
                            num2 = Convert.ToUInt16(managementBaseObject2.Properties["ReturnValue"].Value);
                        } else {
                            WriteDebug(DateTime.Now + " Remote Registry service is Running");
                        }
                    }
                } catch (COMException ex) {
                    var wUApiCodeDetails = wUTools.GetWUApiCodeDetails(ex.ErrorCode);
                    if (wUApiCodeDetails != null) {
                        var codeType = wUApiCodeDetails.CodeType;
                        if (codeType == 2) {
                            WriteError(new ErrorRecord(new Exception(wUApiCodeDetails.Description), wUApiCodeDetails.HResult, ErrorCategory.CloseError, null));
                        }
                    } else if (Debuger) {
                        var errorRecord = new ErrorRecord(ex, "Debug", ErrorCategory.CloseError, null);
                        ThrowTerminatingError(errorRecord);
                    }
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