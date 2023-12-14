using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;
using WUApiLib;

namespace PSWindowsUpdate {
    /// <summary>
    /// Universal function used in rest of PSWindowsUpdate module.
    /// </summary>
    public class WUTools : PSCmdlet {
        // todo: 2022-08-22: change this
        private const string initVector = "SraNie_W-ban13!!";
        private const int keysize = 256;

        /// <summary>Encrypt string by key.</summary>
        /// <param name="plainText">Text to encrypt</param>
        /// <param name="passPhrase">Encryption key</param>
        public static string EncryptString(string plainText, string passPhrase) {
            var bytes1 = Encoding.UTF8.GetBytes(initVector);
            var bytes2 = Encoding.UTF8.GetBytes(plainText);

            // todo: replace with Rfc2898DeriveBytes
            var bytes3 = new PasswordDeriveBytes(passPhrase, null).GetBytes(32);
            var rijndaelManaged = new RijndaelManaged();
            rijndaelManaged.Mode = CipherMode.CBC;
            var encryptor = rijndaelManaged.CreateEncryptor(bytes3, bytes1);
            var memoryStream = new MemoryStream();
            var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
            cryptoStream.Write(bytes2, 0, bytes2.Length);
            cryptoStream.FlushFinalBlock();
            var array = memoryStream.ToArray();
            memoryStream.Close();
            cryptoStream.Close();
            return Convert.ToBase64String(array);
        }

        /// <summary>Decrypt string by key.</summary>
        /// <param name="cipherText">Encrypted text</param>
        /// <param name="passPhrase">Decryption key</param>
        public static string DecryptString(string cipherText, string passPhrase) {
            var bytes1 = Encoding.UTF8.GetBytes(initVector);
            var buffer = Convert.FromBase64String(cipherText);

            // todo: replace with Rfc2898DeriveBytes
            var bytes2 = new PasswordDeriveBytes(passPhrase, null).GetBytes(32);
            var rijndaelManaged = new RijndaelManaged();
            rijndaelManaged.Mode = CipherMode.CBC;
            var decryptor = rijndaelManaged.CreateDecryptor(bytes2, bytes1);
            var memoryStream = new MemoryStream(buffer);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            var numArray = new byte[buffer.Length];
            var count = cryptoStream.Read(numArray, 0, numArray.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(numArray, 0, count);
        }

        /// <summary>Run reboot.</summary>
        /// <param name="Computer">Computer name</param>
        public string RunReboot(string Computer) {
            var startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd";
            startInfo.Arguments = !IsLocalHost(Computer) ? "/C shutdown -m " + Computer + " -f -r -t 5" : "/C shutdown -f -r -t 5";
            var str = startInfo.FileName + " " + startInfo.Arguments;
            Process.Start(startInfo);
            return str;
        }

        /// <summary>Schedule reboot.</summary>
        /// <param name="Computer">Computer name</param>
        /// <param name="ScheduleReboot">Computer name</param>
        public string ScheduleReboot(string Computer, DateTime ScheduleReboot) {
            var now = DateTime.Now;
            var timeSpan = ScheduleReboot - now;
            string str;
            if ((int)timeSpan.TotalSeconds > 0) {
                var startInfo = new ProcessStartInfo();
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd";
                startInfo.Arguments = !IsLocalHost(Computer)
                    ? "/C shutdown -m " + Computer + " -f -r -t " + ((int)timeSpan.TotalSeconds)
                    : "/C shutdown -f -r -t " + ((int)timeSpan.TotalSeconds);
                str = startInfo.FileName + " " + startInfo.Arguments;
                Process.Start(startInfo);
            } else {
                str = "Restart time has expired: " + ScheduleReboot + " < " + now;
            }

            return str;
        }

        /// <summary>Cancel reboot.</summary>
        /// <param name="Computer">Computer name</param>
        public string CancelReboot(string Computer) {
            var startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd";
            startInfo.Arguments = !IsLocalHost(Computer) ? "/C shutdown -m " + Computer + " -f -a" : "/C shutdown -f -a";
            var str = startInfo.FileName + " " + startInfo.Arguments;
            Process.Start(startInfo);
            return str;
        }

        /// <summary>Check is computer is localhost or remotehost.</summary>
        /// <param name="Computer">Computer name</param>
        public bool IsLocalHost(string Computer) {
            if (Computer.ToLower() == "localhost" || Computer.ToLower() == "." || Computer.ToLower() == "127.0.0.1" || Computer.ToLower() == "::1" ||
                Computer.ToLower() == Environment.MachineName.ToLower()) {
                return true;
            }

            var lower = Dns.GetHostName().ToLower();
            if (Computer.ToLower() == lower) {
                return true;
            }

            var collection = new Collection<string>();
            var hostEntry = Dns.GetHostEntry("");
            if (Computer.ToLower() == hostEntry.HostName.ToLower()) {
                return true;
            }

            foreach (var address in hostEntry.AddressList) {
                collection.Add(address.ToString());
            }

            return collection.Contains(Computer.ToLower());
        }

        /// <summary>Restart Windows Update service on specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public void RestartService(string Computer) {
            if (IsLocalHost(Computer)) {
                try {
                    var serviceController = new ServiceController("Windows Update");
                    Thread.Sleep(500);
                    serviceController.Stop();
                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                    Thread.Sleep(500);
                    serviceController.Start();
                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                } catch { }
            } else {
                try {
                    var serviceController = new ServiceController("Windows Update", Computer);
                    Thread.Sleep(500);
                    serviceController.Stop();
                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                    Thread.Sleep(500);
                    serviceController.Start();
                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                } catch { }
            }
        }

        /// <summary>Stope service on specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        /// <param name="ServiceName">Service name</param>
        public void StopService(string Computer, string ServiceName) {
            if (IsLocalHost(Computer)) {
                try {
                    var serviceController = new ServiceController(ServiceName);
                    if (serviceController.Status != ServiceControllerStatus.Running) {
                        return;
                    }

                    Thread.Sleep(500);
                    serviceController.Stop();
                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                } catch { }
            } else {
                try {
                    var serviceController = new ServiceController(ServiceName, Computer);
                    if (serviceController.Status == ServiceControllerStatus.Running) {
                        Thread.Sleep(500);
                        serviceController.Stop();
                        serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                    }
                } catch { }
            }
        }

        /// <summary>Start service on specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        /// <param name="ServiceName">Service name</param>
        public void StartService(string Computer, string ServiceName) {
            if (IsLocalHost(Computer)) {
                try {
                    var serviceController = new ServiceController(ServiceName);
                    if (serviceController.Status != ServiceControllerStatus.Stopped) {
                        return;
                    }

                    Thread.Sleep(500);
                    serviceController.Start();
                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                } catch { }
            } else {
                try {
                    var serviceController = new ServiceController(ServiceName, Computer);
                    if (serviceController.Status == ServiceControllerStatus.Stopped) {
                        Thread.Sleep(500);
                        serviceController.Start();
                        serviceController.WaitForStatus(ServiceControllerStatus.Running);
                    }
                } catch { }
            }
        }

        /// <summary>Start service on specific machine.</summary>
        /// <param name="FileName">Name of starting file</param>
        /// <param name="Arguments">Arguments</param>
        public void RunProcess(string FileName, string Arguments) {
            try {
                new Process() {
                    StartInfo = new ProcessStartInfo() {
                        WindowStyle = ProcessWindowStyle.Hidden,
                        FileName = FileName,
                        Arguments = Arguments
                    }
                }.Start();
            } catch { }
        }

        /// <summary>Test ping connection for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public bool TestPing(string Computer) {
            bool flag;
            try {
                new Ping().Send(Computer);
                flag = true;
                WriteDebug(Computer + ": Ping connection Succeeded");
            } catch {
                flag = false;
                WriteDebug(Computer + ": Ping connection Failed");
            }

            return flag;
        }

        /// <summary>Create object ServiceManager for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        /// <param name="e">Current Exception</param>
        public ErrorRecord ParseException(string Computer, Exception e) {
            return !Regex.IsMatch(e.Message, "800706BA", RegexOptions.IgnoreCase)
                ? !Regex.IsMatch(e.Message, "80070005", RegexOptions.IgnoreCase)
                    ? new ErrorRecord(new Exception(Computer + ": Unknown failure. "), "Unknown", ErrorCategory.PermissionDenied, null)
                    : new ErrorRecord(new Exception(Computer + ": Access denied. You don't have permission to perform this task."), "80070005", ErrorCategory.PermissionDenied,
                        null)
                : new ErrorRecord(new Exception(Computer + ": Connection failure, check firewall on destination machine."), "800706BA", ErrorCategory.ResourceUnavailable,
                    null);
        }

        /// <summary>InvokeRestMethod</summary>
        /// <param name="requestUrl">Requested url</param>
        /// <param name="requestMethod">Request method</param>
        /// <param name="requestBody">Body</param>
        /// <param name="requestHeaders">Headers</param>
        public string InvokeRestMethod(
            string requestUrl,
            HttpWebRequestMethod requestMethod,
            Dictionary<string, string> requestBody,
            Dictionary<string, string> requestHeaders) {
            var str1 = string.Empty;
            if (requestMethod == HttpWebRequestMethod.GET && requestBody != null) {
                foreach (var keyValuePair in requestBody) {
                    if (str1 == string.Empty) {
                        str1 = str1 + keyValuePair.Key + "=" + keyValuePair.Value;
                    } else {
                        str1 = str1 + "&" + keyValuePair.Key + "=" + keyValuePair.Value;
                    }
                }

                requestUrl = requestUrl + "?" + str1;
            }

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUrl);
            httpWebRequest.ContentType = "application/x-www-form-urlencoded";
            httpWebRequest.Method = requestMethod.ToString();
            if (requestMethod == HttpWebRequestMethod.POST && requestBody != null) {
                foreach (var keyValuePair in requestBody) {
                    if (str1 == string.Empty) {
                        str1 = str1 + keyValuePair.Key + "=" + keyValuePair.Value;
                    } else {
                        str1 = str1 + ";" + keyValuePair.Key + "=" + keyValuePair.Value;
                    }
                }

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream())) {
                    streamWriter.Write(str1);
                }
            }

            var str2 = string.Empty;
            using (var response = (HttpWebResponse)httpWebRequest.GetResponse()) {
                using (var responseStream = response.GetResponseStream()) {
                    using (var streamReader = new StreamReader(responseStream)) {
                        str2 = streamReader.ReadToEnd();
                    }
                }
            }

            return str2;
        }

        /// <summary>Create object ServiceManager for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetServiceManagerObj(string Computer) {
            UpdateServiceManager updateServiceManager = null;
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    updateServiceManager = (UpdateServiceManager)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("F8D253D9-89A4-4DAA-87B6-1168369F0B21")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("ServiceManagerObj", updateServiceManager));
                } catch (Exception ex) {
                    var value = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.ServiceManager", Computer);
                    updateServiceManager = (UpdateServiceManager)Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("ServiceManagerObj", updateServiceManager));
                } catch (Exception ex2) {
                    var value2 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object ServiceManager for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiServiceManagerObj(string Computer) {
            UpdateServiceManager updateServiceManager = null;
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    updateServiceManager = (UpdateServiceManager)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("F8D253D9-89A4-4DAA-87B6-1168369F0B21")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateServiceManager;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.ServiceManager", Computer);
                    updateServiceManager = (UpdateServiceManager)Activator.CreateInstance(typeFromProgID);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateServiceManager;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>Create object UpdateSession for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        // todo: Change to UpdateSession instead of PSObject?
        public PSObject GetUpdateSessionObj(string Computer) {
            UpdateSession updateSession = null;
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    updateSession = (UpdateSession)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("4CB43D7F-7EEE-4906-8698-60DA1C38F2FE")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("UpdateSessionObj", updateSession));
                } catch (Exception ex) {
                    var value = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.Session", Computer);
                    updateSession = (UpdateSession)Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("UpdateSessionObj", updateSession));
                } catch (Exception ex2) {
                    var value2 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object UpdateSession for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiUpdateSessionObj(string Computer) {
            UpdateSession updateSession = null;
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    updateSession = (UpdateSession)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("4CB43D7F-7EEE-4906-8698-60DA1C38F2FE")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateSession;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.Session", Computer);
                    updateSession = (UpdateSession)Activator.CreateInstance(typeFromProgID);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateSession;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>Create object Installer for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetInstallerObj(string Computer) {
            UpdateInstaller updateInstaller = null;
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    updateInstaller = (UpdateInstaller)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("D2E0FE7F-D23E-48E1-93C0-6FA8CC346474")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("InstallerObj", updateInstaller));
                } catch (Exception ex) {
                    var value = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.Installer", Computer);
                    updateInstaller = (UpdateInstaller)Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("InstallerObj", updateInstaller));
                } catch (Exception ex2) {
                    var value2 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object Installer for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiInstallerObj(string Computer) {
            UpdateInstaller updateInstaller = null;
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    updateInstaller = (UpdateInstaller)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("D2E0FE7F-D23E-48E1-93C0-6FA8CC346474")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateInstaller;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.Installer", Computer);
                    updateInstaller = (UpdateInstaller)Activator.CreateInstance(typeFromProgID);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = updateInstaller;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>Create object SystemInfo for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetSystemInfoObj(string Computer) {
            SystemInformation systemInformation = null;
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    systemInformation = (SystemInformation)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("C01B9BA0-BEA7-41BA-B604-D0A36F469133")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("SystemInfoObj", systemInformation));
                } catch (Exception ex) {
                    var value = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.Installer", Computer);
                    systemInformation = (SystemInformation)Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("SystemInfoObj", systemInformation));
                } catch (Exception ex2) {
                    var value2 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object SystemInfo for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiSystemInfoObj(string Computer) {
            SystemInformation @object = null;
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    @object = (SystemInformation)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("C01B9BA0-BEA7-41BA-B604-D0A36F469133")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = @object;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.SystemInformation", Computer);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = @object;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>Create object AutoUpdate for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetAutoUpdateObj(string Computer) {
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    var value = (AutomaticUpdates)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("BFE18E9C-6D87-4450-B37C-E02F0B373803")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("AutoUpdateObj", value));
                } catch (Exception ex) {
                    var value2 = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.AutoUpdate", Computer);
                    var value3 = Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("AutoUpdateObj", value3));
                } catch (Exception ex2) {
                    var value4 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value4));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object AutoUpdate for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiAutoUpdateObj(string Computer) {
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    var @object = (AutomaticUpdates)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("BFE18E9C-6D87-4450-B37C-E02F0B373803")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = @object;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.AutoUpdate", Computer);
                    var object2 = Activator.CreateInstance(typeFromProgID);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = object2;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>Create object AgentInfo for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetAgentInfoObj(string Computer) {
            WindowsUpdateAgentInfo windowsUpdateAgentInfo = null;
            var pSObject = new PSObject();
            pSObject.Properties.Add(new PSNoteProperty("ComputerName", Computer));
            if (IsLocalHost(Computer)) {
                try {
                    windowsUpdateAgentInfo = (WindowsUpdateAgentInfo)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D")));
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("AgentInfoObj", windowsUpdateAgentInfo));
                } catch (Exception ex) {
                    var value = ParseException(Computer, ex);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Local"));
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.AgentInfo", Computer);
                    windowsUpdateAgentInfo = (WindowsUpdateAgentInfo)Activator.CreateInstance(typeFromProgID);
                    pSObject.Properties.Add(new PSNoteProperty("Status", true));
                    pSObject.Properties.Add(new PSNoteProperty("AgentInfoObj", windowsUpdateAgentInfo));
                } catch (Exception ex2) {
                    var value2 = ParseException(Computer, ex2);
                    pSObject.Properties.Add(new PSNoteProperty("Status", false));
                    pSObject.Properties.Add(new PSNoteProperty("Error", value2));
                    pSObject.Properties.Add(new PSNoteProperty("Exception", ex2));
                }

                pSObject.Properties.Add(new PSNoteProperty("Mode", "Activator"));
            }

            return pSObject;
        }

        /// <summary>Create object AgentInfo for specific machine.</summary>
        /// <param name="Computer">Computer name</param>
        public WUAPIObject GetWUApiAgentInfoObj(string Computer) {
            WindowsUpdateAgentInfo windowsUpdateAgentInfo = null;
            var wUAPIObject = new WUAPIObject();
            wUAPIObject.ComputerName = Computer;
            if (IsLocalHost(Computer)) {
                try {
                    windowsUpdateAgentInfo = (WindowsUpdateAgentInfo)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D")));
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = windowsUpdateAgentInfo;
                } catch (Exception ex) {
                    var error = ParseException(Computer, ex);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error;
                    wUAPIObject.Exception = ex;
                }

                wUAPIObject.Mode = "Local";
            } else {
                try {
                    var typeFromProgID = Type.GetTypeFromProgID("Microsoft.Update.AgentInfo", Computer);
                    windowsUpdateAgentInfo = (WindowsUpdateAgentInfo)Activator.CreateInstance(typeFromProgID);
                    wUAPIObject.Status = true;
                    wUAPIObject.Object = windowsUpdateAgentInfo;
                } catch (Exception ex2) {
                    var error2 = ParseException(Computer, ex2);
                    wUAPIObject.Status = false;
                    wUAPIObject.Error = error2;
                    wUAPIObject.Exception = ex2;
                }

                wUAPIObject.Mode = "Activator";
            }

            return wUAPIObject;
        }

        /// <summary>convert size to human friendly format.</summary>
        /// <param name="Size">Size in bytes</param>
        public string ConvertSize(decimal Size) {
            return !(Math.Round(Size, 0) < 1024M)
                ? !(Math.Round(Size, 0) < 1048576M)
                    ? !(Math.Round(Size, 0) < 1073741824M)
                        ? !(Math.Round(Size, 0) < 1099511627776M)
                            ? !(Math.Round(Size, 0) < 1125899906842624M)
                                ? Math.Round(Size / 1125899906842624M, 0) + "PB"
                                : Math.Round(Size / 1099511627776M, 0) + "TB"
                            : Math.Round(Size / 1073741824M, 0) + "GB"
                        : Math.Round(Size / 1048576M, 0) + "MB"
                    : Math.Round(Size / 1024M, 0) + "KB"
                : Math.Round(Size, 0) + "B";
        }

        /// <summary>
        /// Get PSWindowsUpdate module version on specific machine.
        /// </summary>
        /// <param name="Computer">Computer name</param>
        public PSObject GetPSWUModule(string Computer) {
            var str =
                "$Module = Get-Module PSWindowsUpdate -ListAvailable | Sort Version -Descending | Select -First 1; $Module | Add-Member -Type NoteProperty -Name PSWUDllVersion -Value ([Version]$(Try{[System.Diagnostics.FileVersionInfo]::GetVersionInfo($(Join-Path -Path $Module.ModuleBase -ChildPath $Module.RootModule)).FileVersion}Catch{'0.0.0.0'})); $Module";
            string script;
            if (IsLocalHost(Computer)) {
                script = str;
            } else {
                script = "Invoke-Command -ComputerName " + Computer + " -Command {" + str + "} -Authentication Negotiate";
            }

            var errorRecord = (ErrorRecord)null;
            Collection<PSObject> collection;
            using (var powerShell = PowerShell.Create()) {
                powerShell.AddScript(script);
                collection = powerShell.Invoke();
                if (powerShell.Streams.Error.Count > 0) {
                    errorRecord = powerShell.Streams.Error[0];
                }
            }

            var pswuModule = new PSObject();
            if (errorRecord == null) {
                try {
                    pswuModule = collection[0];
                } catch {
                    pswuModule.Properties.Add(new PSNoteProperty("PSWindowsUpdate", null));
                    pswuModule.Properties.Add(new PSNoteProperty("PSWUModuleDll", null));
                }
            } else {
                pswuModule.Properties.Add(new PSNoteProperty("Error", errorRecord));
            }

            return pswuModule;
        }

        /// <summary>
        /// Check PSWindowsUpdate module version on specific machine.
        /// </summary>
        /// <param name="Computer">Computer name</param>
        public ErrorRecord CheckPSWUModule(string Computer) {
            var pswuModule = GetPSWUModule(Computer);
            var errorRecord = (ErrorRecord)null;
            try {
                errorRecord = (ErrorRecord)pswuModule.Properties["Error"].Value;
            } catch { }

            var version1 = MinimumPSWUModule();
            if (errorRecord != null) {
                return errorRecord;
            }

            try {
                var version2 = (Version)pswuModule.Properties["Version"].Value;
                return version2 < version1
                    ? new ErrorRecord(
                        new Exception("To Invoke-WUJob you need PSWindowsUpdate " + version1 + " on destination machine! Current module version is " +
                                      version2), "LowModuleVersion", ErrorCategory.ResourceUnavailable, null)
                    : null;
            } catch {
                return new ErrorRecord(new Exception("PSWindowsUpdate module missing on destination machine"), "ModuleMissing", ErrorCategory.ResourceUnavailable, null);
            }
        }

        /// <summary>Define minimum PSWindowsUpdate module version.</summary>
        public Version MinimumPSWUModule() {
            return new Version("2.0.0.0");
        }

        /// <summary>Save Credentials to Credential Manager.</summary>
        /// <param name="Username">User name</param>
        /// <param name="Password">Password</param>
        public void SaveCredential(string Username, string Password) {
            using (var credential = new Credential()) {
                credential.Password = Password;
                credential.Username = Username;
                credential.Target = "PSWindowsUpdate";
                credential.Type = CredentialType.Generic;
                credential.PersistenceType = PersistenceType.LocalComputer;
                credential.Save();
            }
        }

        /// <summary>Get credentials from Credential Manager.</summary>
        public PSCredMan GetCredential() {
            using (var credential = new Credential()) {
                credential.Target = "PSWindowsUpdate";
                credential.Load();
                return new PSCredMan() {
                    Username = credential.Username,
                    Password = credential.Password
                };
            }
        }

        /// <summary>Get PSWindowsUpdate module path.</summary>
        public string PSWUModulePath() {
            return Path.GetDirectoryName(Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path));
        }

        /// <summary>Get GetPSWUSettings.</summary>
        public Hashtable GetPSWUSettings() {
            var path = PSWUModulePath() + "\\PSWUSettings.xml";
            if (File.Exists(path)) {
                try {
                    using (var streamReader = new StreamReader(path)) {
                        return (Hashtable)((PSObject)PSSerializer.Deserialize(streamReader.ReadToEnd())).BaseObject;
                    }
                } catch { }
            }

            return new Hashtable();
        }

        /// <summary>Save GetPSWUSettings.</summary>
        public void SetPSWUSettings(Hashtable PSWUSettings) {
            var path = PSWUModulePath() + "\\PSWUSettings.xml";
            var str = PSSerializer.Serialize(PSWUSettings);
            try {
                using (var streamWriter = new StreamWriter(path)) {
                    streamWriter.WriteLine(str);
                }
            } catch (Exception ex) {
                WriteDebug("Something goes wrong: " + ex.Message);
            }
        }

        /// <summary>Convert object to HTML formated table.</summary>
        /// <param name="PSObjects">Output object</param>
        /// <param name="PSObjectProperties">Choosed properties</param>
        /// <param name="CmdletInfo">Send ifno about used cmdlet</param>
        /// <param name="TableStyle">Format style: Table | List</param>
        public string ObjectToHtml(
            List<PSObject> PSObjects,
            List<string> PSObjectProperties,
            PSObject CmdletInfo = null,
            string TableStyle = null) {
            string str1 = "" +
                          "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\"  \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">\r\n<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n<head>\r\n<style>\r\n<!--\r\n    table {\r\n        font-family: Verdana;\r\n        border-style: dashed;\r\n        border-width: 1px;\r\n        border-color: #FF6600;\r\n        padding: 5px;\r\n        background-color: #FFFFCC;\r\n        table-layout: auto;\r\n        text-align: left;\r\n        font-size: 8pt;\r\nwidth: 100%;\r\n        }\r\n    table th {\r\n        border-bottom-style: solid;\r\n        border-bottom-width: 1px;\r\n        text-align: left;\r\n        }\r\n    table td {\r\n        border-top-style: solid;\r\n        border-top-width: 1px;\r\n        }\r\n    p {\r\n        margin-bottom: 2px\r\n        margin-top: 2px\r\n        }\r\n    body {\r\n        font-family: Verdana;\r\n        font-size: 8pt;\r\n        }\r\n-->\r\n</style>\r\n</head>\r\n<body>\r\n";
            System.Web.UI.WebControls.Table table = new System.Web.UI.WebControls.Table();
            if (TableStyle == null) {
                TableStyle = "Table";
                if (PSObjectProperties.Count > 5)
                    TableStyle = "List";
            }

            if (TableStyle == "Table") {
                TableHeaderRow row1 = new TableHeaderRow();
                foreach (string psObjectProperty in PSObjectProperties) {
                    TableHeaderCell cell = new TableHeaderCell();
                    cell.Text = psObjectProperty;
                    row1.Cells.Add((TableCell)cell);
                }

                table.Rows.AddAt(0, (TableRow)row1);
                foreach (PSObject psObject in PSObjects) {
                    TableRow row2 = new TableRow();
                    foreach (string psObjectProperty in PSObjectProperties) {
                        TableCell cell = new TableCell();
                        try {
                            cell.Text = psObject.Properties[psObjectProperty].Value.ToString();
                        } catch { }

                        row2.Cells.Add(cell);
                    }

                    table.Rows.Add(row2);
                }
            } else {
                foreach (PSObject psObject in PSObjects) {
                    foreach (string psObjectProperty in PSObjectProperties) {
                        TableRow row = new TableRow();
                        row.Cells.Add(new TableCell() {
                            Text = psObjectProperty
                        });
                        TableCell cell = new TableCell();
                        try {
                            cell.Text = psObject.Properties[psObjectProperty].Value.ToString();
                        } catch { }

                        row.Cells.Add(cell);
                        table.Rows.Add(row);
                    }
                }
            }

            using (StringWriter writer = new StringWriter()) {
                table.RenderControl(new HtmlTextWriter((TextWriter)writer));
                str1 += writer.ToString();
            }

            if (CmdletInfo != null) {
                string str2 = str1 + "<p>";
                foreach (PSPropertyInfo property in CmdletInfo.Properties)
                    str2 = str2 + property.Name.ToString() + ": " + CmdletInfo.Properties[property.Name].Value.ToString() + "<br />";
                str1 = str2 + "</p>";
            }

            return str1 + "</body>\r\n</html>";
        }

        /// <summary>Convert object to HTML formated table.</summary>
        /// <param name="PSObjects">Output object</param>
        /// <param name="PSObjectProperties">Choosed properties</param>
        /// <param name="CmdletInfo">Send ifno about used cmdlet</param>
        /// <param name="TableStyle">Format style: Table | List</param>
        public string ObjectToHtml(
            Collection<PSObject> PSObjects,
            Collection<string> PSObjectProperties,
            PSObject CmdletInfo = null,
            string TableStyle = null) {
            var str1 = "" +
                       "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\"  \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">\r\n<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n<head>\r\n<style>\r\n<!--\r\n    table {\r\n        font-family: Verdana;\r\n        border-style: dashed;\r\n        border-width: 1px;\r\n        border-color: #FF6600;\r\n        padding: 5px;\r\n        background-color: #FFFFCC;\r\n        table-layout: auto;\r\n        text-align: left;\r\n        font-size: 8pt;\r\nwidth: 100%;\r\n        }\r\n    table th {\r\n        border-bottom-style: solid;\r\n        border-bottom-width: 1px;\r\n        text-align: left;\r\n        }\r\n    table td {\r\n        border-top-style: solid;\r\n        border-top-width: 1px;\r\n        }\r\n    p {\r\n        margin-bottom: 2px\r\n        margin-top: 2px\r\n        }\r\n    body {\r\n        font-family: Verdana;\r\n        font-size: 8pt;\r\n        }\r\n-->\r\n</style>\r\n</head>\r\n<body>\r\n";
            var table = new Table();
            if (TableStyle == null) {
                TableStyle = "Table";
                if (PSObjectProperties.Count > 5) {
                    TableStyle = "List";
                }
            }

            if (TableStyle == "Table") {
                var row1 = new TableHeaderRow();
                foreach (var psObjectProperty in PSObjectProperties) {
                    var cell = new TableHeaderCell();
                    cell.Text = psObjectProperty;
                    row1.Cells.Add(cell);
                }

                table.Rows.AddAt(0, row1);
                foreach (var psObject in PSObjects) {
                    var row2 = new TableRow();
                    foreach (var psObjectProperty in PSObjectProperties) {
                        var cell = new TableCell();
                        try {
                            cell.Text = psObject.Properties[psObjectProperty].Value.ToString();
                        } catch { }

                        row2.Cells.Add(cell);
                    }

                    table.Rows.Add(row2);
                }
            } else {
                foreach (var psObject in PSObjects) {
                    foreach (var psObjectProperty in PSObjectProperties) {
                        var row = new TableRow();
                        row.Cells.Add(new TableCell {
                            Text = psObjectProperty
                        });
                        var cell = new TableCell();
                        try {
                            cell.Text = psObject.Properties[psObjectProperty].Value.ToString();
                        } catch { }

                        row.Cells.Add(cell);
                        table.Rows.Add(row);
                    }
                }
            }

            using (var writer = new StringWriter()) {
                table.RenderControl(new HtmlTextWriter(writer));
                str1 += writer.ToString();
            }

            if (CmdletInfo != null) {
                var str2 = str1 + "<p>";
                foreach (var property in CmdletInfo.Properties) {
                    str2 = str2 + property.Name + ": " + CmdletInfo.Properties[property.Name].Value + "<br />";
                }

                str1 = str2 + "</p>";
            }

            return str1 + "</body>\r\n</html>";
        }

        /// <summary>test email settings.</summary>
        /// <param name="LocalPSWUSettings">Send parameters</param>
        public PSObject TestMail(Hashtable LocalPSWUSettings) {
            var psObject = new PSObject();
            var mailOptionTypes = new string[8] {
                "SmtpServer",
                "Port",
                "EnableSsl",
                "To",
                "From",
                "Subject",
                "Properties",
                "Style"
            };
            var pswuSettings = GetPSWUSettings();
            var mailSettings = new Hashtable();
            foreach (var optionName in mailOptionTypes) {
                if (LocalPSWUSettings.ContainsKey(optionName)) {
                    mailSettings.Add(optionName, LocalPSWUSettings[optionName]);
                } else if (pswuSettings.ContainsKey(optionName)) {
                    if (optionName != "Properties") {
                        mailSettings.Add(optionName, pswuSettings[optionName]);
                    }
                } else {
                    if (!(optionName == "Port")) {
                        if (!(optionName == "EnableSsl")) {
                            if (!(optionName == "Properties")) {
                                if (!(optionName == "Style")) {
                                    if (optionName == "Subject") {
                                        mailSettings.Add("Subject", Environment.MachineName + ": Windows Update report " + DateTime.Now);
                                    } else {
                                        var errorRecord = new ErrorRecord(
                                            new Exception("Missing " + optionName + "; Use -PSWUSettings or declare PSWUSettings.xml in ModuleBase path."),
                                            optionName, ErrorCategory.CloseError, null);
                                        psObject.Properties.Add(new PSNoteProperty("ErrorRecord", errorRecord));
                                        return psObject;
                                    }
                                }
                            } else {
                                mailSettings.Add("Properties", "*");
                            }
                        } else {
                            mailSettings.Add("EnableSsl", false);
                        }
                    } else {
                        mailSettings.Add("Port", 25);
                    }
                }
            }

            return psObject;
        }

        /// <summary>Send email report.</summary>
        /// <param name="LocalPSWUSettings">Send parameters</param>
        /// <param name="PSObjects">Output object</param>
        /// <param name="CmdletInfo">Send ifno about used cmdlet</param>
        public PSObject SendMail(
            Hashtable LocalPSWUSettings,
            Collection<PSObject> PSObjects,
            PSObject CmdletInfo = null) {
            var psObject1 = new PSObject();
            var mailOptionTypes = new string[8] {
                "SmtpServer",
                "Port",
                "EnableSsl",
                "To",
                "From",
                "Subject",
                "Properties",
                "Style"
            };
            var pswuSettings = GetPSWUSettings();
            var mailSettings = new Hashtable();
            foreach (var optionName in mailOptionTypes) {
                if (LocalPSWUSettings.ContainsKey(optionName)) {
                    mailSettings.Add(optionName, LocalPSWUSettings[optionName]);
                } else if (pswuSettings.ContainsKey(optionName)) {
                    if (optionName != "Properties") {
                        mailSettings.Add(optionName, pswuSettings[optionName]);
                    }
                } else {
                    if (!(optionName == "Port")) {
                        if (!(optionName == "EnableSsl")) {
                            if (!(optionName == "Properties")) {
                                if (!(optionName == "Style")) {
                                    if (optionName == "Subject") {
                                        mailSettings.Add("Subject", Environment.MachineName + ": Windows Update report " + DateTime.Now);
                                    } else {
                                        var errorRecord = new ErrorRecord(
                                            new Exception("Missing " + optionName + "; Use -PSWUSettings or declare PSWUSettings.xml in ModuleBase path."),
                                            optionName, ErrorCategory.CloseError, null);
                                        psObject1.Properties.Add(new PSNoteProperty("ErrorRecord", errorRecord));
                                        return psObject1;
                                    }
                                }
                            } else {
                                mailSettings.Add("Properties", "*");
                            }
                        } else {
                            mailSettings.Add("EnableSsl", false);
                        }
                    } else {
                        mailSettings.Add("Port", 25);
                    }
                }
            }

            psObject1.Properties.Add(new PSNoteProperty("PSWUSettings", mailSettings));
            if (PSObjects != null) {
                var PSObjectProperties = new Collection<string>();
                if (mailSettings["Properties"].ToString() == "*") {
                    foreach (var property in PSObjects[0].Properties) {
                        if (!PSObjectProperties.Contains(property.Name)) {
                            PSObjectProperties.Add(property.Name);
                        }
                    }
                } else if (mailSettings["Properties"].ToString() == "**") {
                    foreach (var psObject2 in PSObjects) {
                        foreach (var property in psObject2.Properties) {
                            if (!PSObjectProperties.Contains(property.Name)) {
                                PSObjectProperties.Add(property.Name);
                            }
                        }
                    }
                } else {
                    foreach (var obj in (object[])mailSettings["Properties"]) {
                        PSObjectProperties.Add(obj.ToString());
                    }
                }

                var str = !mailSettings.ContainsKey("Style")
                    ? ObjectToHtml(PSObjects, PSObjectProperties, CmdletInfo)
                    : ObjectToHtml(PSObjects, PSObjectProperties, CmdletInfo, mailSettings["Style"].ToString());
                var message = new MailMessage(mailSettings["From"].ToString(), mailSettings["To"].ToString());
                var smtpClient = new SmtpClient();
                smtpClient.Host = mailSettings["SmtpServer"].ToString();
                smtpClient.Port = Convert.ToInt32(mailSettings["Port"].ToString());
                smtpClient.EnableSsl = (bool)mailSettings["EnableSsl"];
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                var credential = GetCredential();
                if (credential.Username != null) {
                    smtpClient.UseDefaultCredentials = false;
                    var networkCredential = new NetworkCredential(credential.Username, credential.Password);
                    smtpClient.Credentials = networkCredential;
                    message.Subject = mailSettings["Subject"] + ".";
                } else {
                    message.Subject = mailSettings["Subject"].ToString();
                }

                message.Body = str;
                message.BodyEncoding = new UTF8Encoding();
                message.IsBodyHtml = true;
                try {
                    smtpClient.Send(message);
                } catch (Exception ex) {
                    var errorRecord = new ErrorRecord(ex, "SmtpException", ErrorCategory.CloseError, null);
                    psObject1.Properties.Add(new PSNoteProperty("ErrorRecord", errorRecord));
                    return psObject1;
                }

                return psObject1;
            }

            var errorRecord1 = new ErrorRecord(new Exception("Missing Body"), "Body", ErrorCategory.CloseError, null);
            psObject1.Properties.Add(new PSNoteProperty("ErrorRecord", errorRecord1));
            return psObject1;
        }

        /// <summary>Send email report.</summary>
        /// <param name="LocalPSWUSettings">Send parameters</param>
        /// <param name="PSObjects">Output object</param>
        /// <param name="CmdletInfo">Send ifno about used cmdlet</param>
        public PSObject SendMail(
            Hashtable LocalPSWUSettings,
            List<PSObject> PSObjects,
            PSObject CmdletInfo = null) {
            PSObject psObject1 = new PSObject();
            string[] strArray = new string[8] {
                "SmtpServer",
                "Port",
                "EnableSsl",
                "To",
                "From",
                "Subject",
                "Properties",
                "Style"
            };
            Hashtable pswuSettings = this.GetPSWUSettings();
            Hashtable hashtable = new Hashtable();
            foreach (string str in strArray) {
                if (LocalPSWUSettings.ContainsKey((object)str))
                    hashtable.Add((object)str, LocalPSWUSettings[(object)str]);
                else if (pswuSettings.ContainsKey((object)str)) {
                    if (str != "Properties")
                        hashtable.Add((object)str, pswuSettings[(object)str]);
                } else {
                    switch (str) {
                        case "Port":
                            hashtable.Add((object)"Port", (object)25);
                            goto case "Style";
                        case "EnableSsl":
                            hashtable.Add((object)"EnableSsl", (object)false);
                            goto case "Style";
                        case "Properties":
                            hashtable.Add((object)"Properties", (object)"*");
                            goto case "Style";
                        case "Style":
                            continue;
                        case "Subject":
                            hashtable.Add((object)"Subject", (object)(Environment.MachineName + ": Windows Update report " + DateTime.Now.ToString()));
                            goto case "Style";
                        default:
                            ErrorRecord errorRecord = new ErrorRecord(new Exception("Missing " + str + "; Use -PSWUSettings or declare PSWUSettings.xml in ModuleBase path."), str,
                                ErrorCategory.CloseError, (object)null);
                            psObject1.Properties.Add((PSPropertyInfo)new PSNoteProperty("ErrorRecord", (object)errorRecord));
                            return psObject1;
                    }
                }
            }

            psObject1.Properties.Add((PSPropertyInfo)new PSNoteProperty("PSWUSettings", (object)hashtable));
            if (PSObjects != null) {
                List<string> PSObjectProperties = new List<string>();
                if (hashtable[(object)"Properties"].ToString() == "*") {
                    foreach (PSPropertyInfo property in PSObjects[0].Properties) {
                        if (!PSObjectProperties.Contains(property.Name.ToString()))
                            PSObjectProperties.Add(property.Name.ToString());
                    }
                } else if (hashtable[(object)"Properties"].ToString() == "**") {
                    foreach (PSObject psObject2 in PSObjects) {
                        foreach (PSPropertyInfo property in psObject2.Properties) {
                            if (!PSObjectProperties.Contains(property.Name.ToString()))
                                PSObjectProperties.Add(property.Name.ToString());
                        }
                    }
                } else {
                    foreach (object obj in (object[])hashtable[(object)"Properties"])
                        PSObjectProperties.Add(obj.ToString());
                }

                string str = !hashtable.ContainsKey((object)"Style")
                    ? this.ObjectToHtml(PSObjects, PSObjectProperties, CmdletInfo)
                    : this.ObjectToHtml(PSObjects, PSObjectProperties, CmdletInfo, hashtable[(object)"Style"].ToString());
                MailMessage message = new MailMessage(hashtable[(object)"From"].ToString(), hashtable[(object)"To"].ToString());
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = hashtable[(object)"SmtpServer"].ToString();
                smtpClient.Port = Convert.ToInt32(hashtable[(object)"Port"].ToString());
                smtpClient.EnableSsl = (bool)hashtable[(object)"EnableSsl"];
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                PSCredMan credential = this.GetCredential();
                if (credential.Username != null) {
                    smtpClient.UseDefaultCredentials = false;
                    NetworkCredential networkCredential = new NetworkCredential(credential.Username, credential.Password);
                    smtpClient.Credentials = (ICredentialsByHost)networkCredential;
                    message.Subject = hashtable[(object)"Subject"].ToString() + ".";
                } else
                    message.Subject = hashtable[(object)"Subject"].ToString();

                message.Body = str;
                message.BodyEncoding = (Encoding)new UTF8Encoding();
                message.IsBodyHtml = true;
                try {
                    smtpClient.Send(message);
                } catch (Exception ex) {
                    ErrorRecord errorRecord = new ErrorRecord(ex, "SmtpException", ErrorCategory.CloseError, (object)null);
                    psObject1.Properties.Add((PSPropertyInfo)new PSNoteProperty("ErrorRecord", (object)errorRecord));
                    return psObject1;
                }

                return psObject1;
            }

            ErrorRecord errorRecord1 = new ErrorRecord(new Exception("Missing Body"), "Body", ErrorCategory.CloseError, (object)null);
            psObject1.Properties.Add((PSPropertyInfo)new PSNoteProperty("ErrorRecord", (object)errorRecord1));
            return psObject1;
        }

        /// <summary>Get description from error code</summary>
        /// <param name="ErrorCode">Error code</param>
        public WUApiCode GetWUApiCodeDetails(int ErrorCode) {
            WUApiCode wuApiCode;
            return new Dictionary<string, WUApiCode> {
                {
                    "0x00000000",
                    new WUApiCode {
                        CodeType = 1,
                        HResult = "0x00000000",
                        Message = "",
                        Description = "Success"
                    }
                }, {
                    "0x00240001",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240001",
                        Message = "WU_S_SERVICE_STOP",
                        Description = "WUA was stopped successfully."
                    }
                }, {
                    "0x00240002",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240002",
                        Message = "WU_S_SELFUPDATE",
                        Description = "WUA updated itself."
                    }
                }, {
                    "0x00240003",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240003",
                        Message = "WU_S_UPDATE_ERROR",
                        Description = "The operation completed successfully but errors occurred applying the updates."
                    }
                }, {
                    "0x00240004",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240004",
                        Message = "WU_S_MARKED_FOR_DISCONNECT",
                        Description = " callback was marked to be disconnected later because the request to disconnect the operation came while a callback was executing."
                    }
                }, {
                    "0x00240005",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240005",
                        Message = "WU_S_REBOOT_REQUIRED",
                        Description = "The system must be restarted to complete installation of the update."
                    }
                }, {
                    "0x00240006",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240006",
                        Message = "WU_S_ALREADY_INSTALLED",
                        Description = "The update to be installed is already installed on the system."
                    }
                }, {
                    "0x00240007",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240007",
                        Message = "WU_S_ALREADY_UNINSTALLED",
                        Description = "The update to be removed is not installed on the system."
                    }
                }, {
                    "0x00240008",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00240008",
                        Message = "WU_S_ALREADY_DOWNLOADED",
                        Description = "The update to be downloaded has already been downloaded."
                    }
                }, {
                    "0x00242015",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x00242015",
                        Message = "WU_S_UH_INSTALLSTILLPENDING",
                        Description = "The installation operation for the update is still in progress."
                    }
                }, {
                    "0x80040154",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80040154",
                        Message = "REGDB_E_CLASSNOTREG",
                        Description = "CreateInstance failed for UpdateServiceManager Class not registered. Probably you use wrong Service Manager."
                    }
                }, {
                    "0x80070005",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80070005",
                        Message = "ERROR_ACCESS_DENIED",
                        Description = "Probably you don't have permission to perform this task or the authentication method is not supported."
                    }
                }, {
                    "0x800706BE",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x800706BE",
                        Message = "RPC_S_SERVER_UNAVAILABLE",
                        Description = "RPC connection failure, check firewall on destination machine."
                    }
                }, {
                    "0x80072EE2",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80072EE2",
                        Message = "ERROR_INTERNET_TIMEOUT",
                        Description = "The request has timed out. Probably you don't have connection to Windows Update server."
                    }
                }, {
                    "0x80240001",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240001",
                        Message = "WU_E_NO_SERVICE",
                        Description = "WUA was unable to provide the service."
                    }
                }, {
                    "0x80240002",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240002",
                        Message = "WU_E_MAX_CAPACITY_REACHED",
                        Description = "The maximum capacity of the service was exceeded."
                    }
                }, {
                    "0x80240003",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240003",
                        Message = "WU_E_UNKNOWN_ID",
                        Description = "WUA cannot find an ID."
                    }
                }, {
                    "0x80240004",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240004",
                        Message = "WU_E_NOT_INITIALIZED",
                        Description = "The object could not be initialized."
                    }
                }, {
                    "0x80240005",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240005",
                        Message = "WU_E_RANGEOVERLAP",
                        Description = "The update handler requested a byte range overlapping a previously requested range."
                    }
                }, {
                    "0x80240006",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240006",
                        Message = "WU_E_TOOMANYRANGES",
                        Description = "The requested number of byte ranges exceeds the maximum number (2\u00B3\u00B9 - 1)."
                    }
                }, {
                    "0x80240007",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240007",
                        Message = "WU_E_INVALIDINDEX",
                        Description = "The index to a collection was invalid."
                    }
                }, {
                    "0x80240008",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240008",
                        Message = "WU_E_ITEMNOTFOUND",
                        Description = "The key for the item queried could not be found."
                    }
                }, {
                    "0x80240009",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240009",
                        Message = "WU_E_OPERATIONINPROGRESS",
                        Description = "Another conflicting operation was in progress. Some operations such as installation cannot be performed twice simultaneously."
                    }
                }, {
                    "0x8024000A",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000A",
                        Message = "WU_E_COULDNOTCANCE",
                        Description = "Cancellation of the operation was not allowed."
                    }
                }, {
                    "0x8024000B",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000B",
                        Message = "WU_E_CALL_CANCELLED",
                        Description = "Operation was cancelled."
                    }
                }, {
                    "0x8024000C",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000C",
                        Message = "WU_E_NOOP",
                        Description = "No operation was required."
                    }
                }, {
                    "0x8024000D",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000D",
                        Message = "WU_E_XML_MISSINGDATA",
                        Description = "WUA could not find required information in the update's XML data."
                    }
                }, {
                    "0x8024000E",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000E",
                        Message = "WU_E_XML_INVALID",
                        Description = "WUA found invalid information in the update's XML data."
                    }
                }, {
                    "0x8024000F",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024000F",
                        Message = "WU_E_CYCLE_DETECTED",
                        Description = "Circular update relationships were detected in the metadata."
                    }
                }, {
                    "0x80240010",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240010",
                        Message = "WU_E_TOO_DEEP_RELATION",
                        Description = "Update relationships too deep to evaluate were evaluated."
                    }
                }, {
                    "0x80240011",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240011",
                        Message = "WU_E_INVALID_RELATIONSHIP",
                        Description = "An invalid update relationship was detected."
                    }
                }, {
                    "0x80240012",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240012",
                        Message = "WU_E_REG_VALUE_INVALID",
                        Description = "An invalid registry value was read."
                    }
                }, {
                    "0x80240013",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240013",
                        Message = "WU_E_DUPLICATE_ITEM",
                        Description = "Operation tried to add a duplicate item to a list."
                    }
                }, {
                    "0x80240014",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240014",
                        Message = "WU_E_INVALID_INSTALL_REQUESTED",
                        Description = "Updates that are requested for install are not installable by the caller."
                    }
                }, {
                    "0x80240016",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240016",
                        Message = "WU_E_INSTALL_NOT_ALLOWED",
                        Description = "Operation tried to install while another installation was in progress or the system was pending a mandatory restart."
                    }
                }, {
                    "0x80240017",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240017",
                        Message = "WU_E_NOT_APPLICABLE",
                        Description = "Operation was not performed because there are no applicable updates."
                    }
                }, {
                    "0x80240018",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240018",
                        Message = "WU_E_NO_USERTOKEN",
                        Description = "Operation failed because a required user token is missing."
                    }
                }, {
                    "0x80240019",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240019",
                        Message = "WU_E_EXCLUSIVE_INSTALL_CONFLICT",
                        Description = "An exclusive update can't be installed with other updates at the same time."
                    }
                }, {
                    "0x8024001A",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024001A",
                        Message = "WU_E_POLICY_NOT_SET",
                        Description = "A policy value was not set."
                    }
                }, {
                    "0x8024001B",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024001B",
                        Message = "WU_E_SELFUPDATE_IN_PROGRESS",
                        Description = "The operation could not be performed because the Windows Update Agent is self-updating."
                    }
                }, {
                    "0x8024001D",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024001D",
                        Message = "WU_E_INVALID_UPDATE",
                        Description = "An update contains invalid metadata."
                    }
                }, {
                    "0x8024001E",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024001E",
                        Message = "WU_E_SERVICE_STOP",
                        Description = "Operation did not complete because the service or system was being shut down."
                    }
                }, {
                    "0x8024001F",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024001F",
                        Message = "WU_E_NO_CONNECTION",
                        Description = "Operation did not complete because the network connection was unavailable."
                    }
                }, {
                    "0x80240020",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240020",
                        Message = "WU_E_NO_INTERACTIVE_USER",
                        Description = "Operation did not complete because there is no logged-on interactive user."
                    }
                }, {
                    "0x80240021",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240021",
                        Message = "WU_E_TIME_OUT",
                        Description = "Operation did not complete because it timed out."
                    }
                }, {
                    "0x80240022",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240022",
                        Message = "WU_E_ALL_UPDATES_FAILED",
                        Description = "Operation failed for all the updates."
                    }
                }, {
                    "0x80240023",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240023",
                        Message = "WU_E_EULAS_DECLINED",
                        Description = "The license terms for all updates were declined."
                    }
                }, {
                    "0x80240024",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240024",
                        Message = "WU_E_NO_UPDATE",
                        Description = "There are no updates."
                    }
                }, {
                    "0x80240025",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240025",
                        Message = "WU_E_USER_ACCESS_DISABLED",
                        Description = "Group Policy settings prevented access to Windows Update."
                    }
                }, {
                    "0x80240026",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240026",
                        Message = "WU_E_INVALID_UPDATE_TYPE",
                        Description = "The type of update is invalid."
                    }
                }, {
                    "0x80240027",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240027",
                        Message = "WU_E_URL_TOO_LONG",
                        Description = "The URL exceeded the maximum length."
                    }
                }, {
                    "0x80240028",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240028",
                        Message = "WU_E_UNINSTALL_NOT_ALLOWED",
                        Description = "The update could not be uninstalled because the request did not originate from a Windows Server Update Services (WSUS) server."
                    }
                }, {
                    "0x80240029",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240029",
                        Message = "WU_E_INVALID_PRODUCT_LICENSE",
                        Description = "Search may have missed some updates before there is an unlicensed application on the system."
                    }
                }, {
                    "0x8024002A",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002A",
                        Message = "WU_E_MISSING_HANDLER",
                        Description = "A component required to detect applicable updates was missing."
                    }
                }, {
                    "0x8024002B",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002B",
                        Message = "WU_E_LEGACYSERVER",
                        Description = "An operation did not complete because it requires a newer version of server."
                    }
                }, {
                    "0x8024002C",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002C",
                        Message = "WU_E_BIN_SOURCE_ABSENT",
                        Description = "A delta-compressed update could not be installed because it required the source."
                    }
                }, {
                    "0x8024002D",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002D",
                        Message = "WU_E_SOURCE_ABSENT",
                        Description = "A full-file update could not be installed because it required the source."
                    }
                }, {
                    "0x8024002E",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002E",
                        Message = "WU_E_WU_DISABLED",
                        Description = "Access to an unmanaged server is not allowed."
                    }
                }, {
                    "0x8024002F",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024002F",
                        Message = "WU_E_CALL_CANCELLED_BY_POLICY",
                        Description = "Operation did not complete because the DisableWindowsUpdateAccess policy was set in the registry."
                    }
                }, {
                    "0x80240030",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240030",
                        Message = "WU_E_INVALID_PROXY_SERVER",
                        Description = "The format of the proxy list was invalid."
                    }
                }, {
                    "0x80240031",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240031",
                        Message = "WU_E_INVALID_FILE",
                        Description = "The file is in the wrong format."
                    }
                }, {
                    "0x80240032",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240032",
                        Message = "WU_E_INVALID_CRITERIA",
                        Description = "The search criteria string was invalid. Probably you have wrong pre search query."
                    }
                }, {
                    "0x80240033",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240033",
                        Message = "WU_E_EULA_UNAVAILABLE",
                        Description = "License terms could not be downloaded."
                    }
                }, {
                    "0x80240034",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240034",
                        Message = "WU_E_DOWNLOAD_FAILED",
                        Description = "Update failed to download."
                    }
                }, {
                    "0x80240035",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240035",
                        Message = "WU_E_UPDATE_NOT_PROCESSED",
                        Description = "The update was not processed."
                    }
                }, {
                    "0x80240036",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240036",
                        Message = "WU_E_INVALID_OPERATION",
                        Description = "The object's current state did not allow the operation."
                    }
                }, {
                    "0x80240037",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240037",
                        Message = "WU_E_NOT_SUPPORTED",
                        Description = "The functionality for the operation is not supported."
                    }
                }, {
                    "0x80240039",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240039",
                        Message = "WU_E_TOO_MANY_RESYNC",
                        Description = "Agent is asked by server to resync too many times."
                    }
                }, {
                    "0x80240040",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240040",
                        Message = "WU_E_NO_SERVER_CORE_SUPPORT",
                        Description = "The WUA API method does not run on the server core installation."
                    }
                }, {
                    "0x80240041",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240041",
                        Message = "WU_E_SYSPREP_IN_PROGRESS",
                        Description = "Service is not available while sysprep is running."
                    }
                }, {
                    "0x80240042",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240042",
                        Message = "WU_E_UNKNOWN_SERVICE",
                        Description = "The update service is no longer registered with automatic updates."
                    }
                }, {
                    "0x80240043",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240043",
                        Message = "WU_E_NO_UI_SUPPORT",
                        Description = "No support for the WUA user interface."
                    }
                }, {
                    "0x80240044",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240044",
                        Message = "WU_E_PER_MACHINE_UPDATE_ACCESS_DENIED",
                        Description = "Only administrators can perform this operation on per-computer updates. Use run as administrator."
                    }
                }, {
                    "0x80240045",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240045",
                        Message = "WU_E_UNSUPPORTED_SEARCHSCOPE",
                        Description = "A search was attempted with a scope that is not currently supported for this type of search."
                    }
                }, {
                    "0x80240046",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240046",
                        Message = "WU_E_BAD_FILE_URL",
                        Description = "The URL does not point to a file."
                    }
                }, {
                    "0x80240047",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240047",
                        Message = "WU_E_NOTSUPPORTED",
                        Description = "The operation requested is not supported."
                    }
                }, {
                    "0x80240048",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240048",
                        Message = "WU_E_INVALID_NOTIFICATION_INFO",
                        Description = "The featured update notification info returned by the server is invalid."
                    }
                }, {
                    "0x80240049",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240049",
                        Message = "WU_E_OUTOFRANGE",
                        Description = "The data is out of range."
                    }
                }, {
                    "0x8024004A",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024004A",
                        Message = "WU_E_SETUP_IN_PROGRESS",
                        Description = "WUA operations are not available while operating system setup is running."
                    }
                }, {
                    "0x80240FFF",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80240FFF",
                        Message = "WU_E_UNEXPECTED",
                        Description = "An operation failed due to reasons not covered by another error code."
                    }
                }, {
                    "0x80244016",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80244016",
                        Message = "WU_E_PT_HTTP_STATUS_BAD_REQUEST",
                        Description = "Same as HTTP status 400 - the server could not process the request due to invalid syntax."
                    }
                }, {
                    "0x80244017",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80244017",
                        Message = "WU_E_PT_HTTP_STATUS_DENIED",
                        Description = "Same as HTTP status 401 - the requested resource requires user authentication."
                    }
                }, {
                    "0x80244018",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80244018",
                        Message = "WU_E_PT_HTTP_STATUS_FORBIDDEN",
                        Description = "Same as HTTP status 403 - server understood the request, but declined to fulfill it."
                    }
                }, {
                    "0x80244019",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80244019",
                        Message = "WU_E_PT_HTTP_STATUS_NOT_FOUND",
                        Description = "Same as HTTP status 404 - the server cannot find the requested URI."
                    }
                }, {
                    "0x8024401A",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024401A",
                        Message = "WU_E_PT_HTTP_STATUS_BAD_METHOD",
                        Description = "Same as HTTP status 405 - the HTTP method is not allowed."
                    }
                }, {
                    "0x8024401B",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024401B",
                        Message = "WU_E_PT_HTTP_STATUS_PROXY_AUTH_REQ",
                        Description = "Same as HTTP status 407 - proxy authentication is required."
                    }
                }, {
                    "0x8024401C",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x8024401C",
                        Message = "WU_E_PT_HTTP_STATUS_REQUEST_TIMEOUT",
                        Description = "Same as HTTP status 408 - the server timed out waiting for the request."
                    }
                }, {
                    "0x80244022",
                    new WUApiCode {
                        CodeType = 2,
                        HResult = "0x80244022",
                        Message = "WU_E_PT_HTTP_STATUS_SERVICE_UNAVAIL",
                        Description = "Same as HTTP status 503 - the service is temporarily overloaded."
                    }
                }, {
                    "0x80248014",
                    new WUApiCode {
                        CodeType = 2, // CodeType.Error,
                        HResult = "0x80248014",
                        Message = "WU_E_DS_UNKNOWNSERVICE",
                        Description =
                            "An operation did not complete because the service is not in the data store. Probably you don't have permission for remote connection to Windows Update Agent or used wrong ServiceID."
                    }
                }
            }.TryGetValue("0x" + Convert.ToString(ErrorCode, 16), out wuApiCode)
                ? wuApiCode
                : null;
        }

        /// <summary>Possible http web request methods</summary>
        public enum HttpWebRequestMethod {
            /// <summary>Post</summary>
            POST,

            /// <summary>Get</summary>
            GET
        }

        /// <summary>Specifies the BuiltinUser type.</summary>
        public enum CodeType {
            /// <summary>Success stream.</summary>
            Success = 1,

            /// <summary>Error stream.</summary>
            Error = 2,

            /// <summary>Warning stream.</summary>
            Warning = 3,

            /// <summary>Verbose stream.</summary>
            Verbose = 4,

            /// <summary>Debug stream.</summary>
            Debug = 5,

            /// <summary>Information stream.</summary>
            Information = 6,

            /// <summary>Throw stream.</summary>
            Throw = 7
        }
    }
}