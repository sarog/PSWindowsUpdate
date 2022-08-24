using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using MSHTML;

namespace PSWindowsUpdate {
    [Cmdlet("Get", "WUOfflineMSU", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class GetWUOfflineMSU : PSCmdlet {
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

        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        public string KBArticleID { get; set; }

        [Parameter(Mandatory = true)]
        public string Destination { get; set; }

        [Parameter]
        public SwitchParameter AcceptAll { get; set; }

        private static WUImpersonator ImpersonatorObj { get; set; }

        private static WUTools WUToolsObj { get; set; }

        private static WUBits WUBitsObj { get; set; }

        private static Collection<PSObject> OutputObj { get; set; }

        private static DateTime CmdletStart { get; set; }

        private static DateTime CmdletEnd { get; set; }

        protected override void BeginProcessing() {
            CmdletStart = DateTime.Now;
            var invocationName = MyInvocation.InvocationName;
            WriteDebug(DateTime.Now + " CmdletStart: " + invocationName);
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)) {
                WriteWarning("To perform some operations you must run an elevated Windows PowerShell console.");
            }

            WUToolsObj = new WUTools();
            WUBitsObj = new WUBits();
            OutputObj = new Collection<PSObject>();
            if (SendReport) {
                WriteDebug(DateTime.Now + " Test smtp settings");
                if (!PSWUSettings.ContainsKey("Properties")) {
                    PSWUSettings.Add("Properties", new string[1] {
                        "*"
                    });
                }

                var psObject = WUToolsObj.TestMail(PSWUSettings);
                if (psObject.Properties.Match("ErrorRecord").Count == 1) {
                    WriteError((ErrorRecord)psObject.Properties["ErrorRecord"].Value);
                    SendReport = false;
                    WriteDebug(DateTime.Now + " Disabling -SendReport");
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
            var computerNames = ComputerName;
            foreach (var target in computerNames) {
                if (!Regex.IsMatch(KBArticleID, "KB", RegexOptions.IgnoreCase)) {
                    KBArticleID = "KB" + KBArticleID;
                }

                var reqUrl = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + KBArticleID;
                WriteDebug(DateTime.Now + " " + reqUrl);
                var text3 = WUToolsObj.InvokeRestMethod(reqUrl, WUTools.HttpWebRequestMethod.GET, null, null);
                var psarray = new object[1] { text3 };
                var hTMLDocument = (HTMLDocument)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("25336920-03F9-11CF-8FD0-00AA00686F13")));
                var iHTMLDocument = (IHTMLDocument2)hTMLDocument;
                iHTMLDocument.write(psarray);
                var regex = new Regex("(id=\\\"(?<id>.*?)\\\"(.*?)value='Download')", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                var matchCollection = regex.Matches(text3);
                var count = matchCollection.Count;
                WriteVerbose("Found [" + count + "] Updates in Microsoft Update Catalog");
                var num = 0;
                var activityId = 1;
                var activity = "Choose updates for " + target;
                var statusDescription = "[" + num + "/" + count + "]";
                var progressRecord = new ProgressRecord(activityId, activity, statusDescription);
                var collection = new Collection<PSObject>();
                var num2 = 0;
                foreach (Match item in matchCollection) {
                    var pSObject = new PSObject();
                    pSObject.Properties.Add(new PSNoteProperty("X", 1));
                    var groups = item.Groups;
                    var updateId = groups["id"].Value;
                    pSObject.Properties.Add(new PSNoteProperty("Identity", updateId));
                    var elementById = hTMLDocument.getElementById(groups["id"].Value + "_link");
                    var updateTitle = elementById.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("Title", updateTitle));
                    var elementById2 = hTMLDocument.getElementById(groups["id"].Value + "_C2_R" + num2);
                    var product = elementById2.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("Product", product));
                    var elementById3 = hTMLDocument.getElementById(groups["id"].Value + "_C3_R" + num2);
                    var classification = elementById3.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("Classification", classification));
                    var elementById4 = hTMLDocument.getElementById(groups["id"].Value + "_C4_R" + num2);
                    var lastUpdated = elementById4.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("LastUpdated", lastUpdated));
                    var elementById5 = hTMLDocument.getElementById(groups["id"].Value + "_C5_R" + num2);
                    var updateVersion = elementById5.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("Version", updateVersion));
                    var elementById6 = hTMLDocument.getElementById(groups["id"].Value + "_size");
                    var updateSize = elementById6.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("Size", updateSize));
                    var elementById7 = hTMLDocument.getElementById(groups["id"].Value + "_originalSize");
                    var originalSize = elementById7.innerText;
                    pSObject.Properties.Add(new PSNoteProperty("OriginalSize", originalSize));
                    progressRecord.StatusDescription = "[" + num + "/" + count + "] " + updateTitle + " " + updateSize;
                    progressRecord.PercentComplete = num * 100 / count;
                    WriteProgress(progressRecord);
                    num++;
                    WriteDebug(DateTime.Now + " Show update to accept: " + updateTitle);
                    var flag = false;
                    flag = AcceptAll || (ShouldProcess(target, "(" + DateTime.Now + ") " + updateTitle + "[" + updateSize + "]") ? true : false);
                    var pStatus = "";
                    var pResult = "";
                    if (flag) {
                        pStatus += "A";
                        pResult = "Accepted";
                        WriteDebug(DateTime.Now + " " + pResult);
                        pSObject.Properties.Add(new PSNoteProperty("ChooseResult", pResult));
                        pSObject.Properties.Add(new PSNoteProperty("Status", pStatus));
                        pSObject.Properties.Add(new PSNoteProperty("Result", pResult));
                        var requestUrl = "http://www.catalog.update.microsoft.com/DownloadDialog.aspx";
                        var dictionary = new Dictionary<string, string>();
                        dictionary.Add("updateIDs", "[{\"uidInfo\":\"" + updateId + "\",\"updateID\":\"" + updateId + "\",\"size\":0}]");
                        WriteDebug(DateTime.Now + " " + dictionary["updateIDs"]);
                        var input = WUToolsObj.InvokeRestMethod(requestUrl, WUTools.HttpWebRequestMethod.POST, dictionary, null);
                        var regex2 = new Regex("(?<url>http[s]?\\://download\\.windowsupdate\\.com/[^\\'\\\"]*)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                        var matchCollection2 = regex2.Matches(input);
                        var downloadUrl = matchCollection2[0].Groups["url"].Value;
                        var fileName = Path.GetFileName(downloadUrl);
                        pSObject.Properties.Add(new PSNoteProperty("DownloadUrl", downloadUrl));
                        pSObject.Properties.Add(new PSNoteProperty("FileName", fileName));
                        WriteDebug(DateTime.Now + " " + downloadUrl);
                    } else {
                        pStatus += "R";
                        pResult = "Rejected";
                        WriteDebug(DateTime.Now + " " + pResult);
                        pSObject.Properties.Add(new PSNoteProperty("ChooseResult", pResult));
                        pSObject.Properties.Add(new PSNoteProperty("Status", pStatus));
                        pSObject.Properties.Add(new PSNoteProperty("Result", pResult));
                        pSObject.Properties.Add(new PSNoteProperty("DownloadUrl", null));
                        pSObject.Properties.Add(new PSNoteProperty("FileName", null));
                    }

                    pSObject.TypeNames.Clear();
                    pSObject.TypeNames.Add("PSWindowsUpdate.OfflineMSU");
                    collection.Add(pSObject);
                    num2++;
                }

                var num3 = collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted").Count();
                WriteObject(collection, true);
                WriteVerbose("Accepted [" + num3 + "] Updates ready to Download");
                var num4 = 0;
                var activityId2 = 1;
                var activity2 = "Download updates for " + target;
                var statusDescription2 = "[" + num4 + "/" + num3 + "]";
                var progressRecord2 = new ProgressRecord(activityId2, activity2, statusDescription2);
                foreach (var item2 in collection.Where(x => x.Properties["Result"].Value.ToString() == "Accepted")) {
                    item2.Properties.Add(new PSNoteProperty("X", 2));
                    var source = item2.Properties["DownloadUrl"].Value.ToString();
                    var path = item2.Properties["FileName"].Value.ToString();
                    var text7 = Path.Combine(Destination, path);
                    WriteDebug(DateTime.Now + " " + text7);
                    var text8 = item2.Properties["Title"].Value.ToString();
                    var text9 = item2.Properties["Size"].Value.ToString();
                    progressRecord2.StatusDescription = "[" + num4 + "/" + num3 + "] " + text8 + " " + text9;
                    progressRecord2.PercentComplete = num4 * 100 / num3;
                    WriteProgress(progressRecord2);
                    num4++;
                    if (!Directory.Exists(Destination)) {
                        Directory.CreateDirectory(Destination);
                        WriteDebug(DateTime.Now + " Creatig destination: " + Destination);
                    }

                    if (Directory.Exists(Destination)) {
                        if (File.Exists(text7)) {
                            WriteDebug(DateTime.Now + " File exist and will be overwited: " + text7);
                        }

                        var text10 = WUBitsObj.StartBitsTransfer(source, text7);
                        WriteDebug(DateTime.Now + " " + text10);
                        var status = item2.Properties["Status"].Value.ToString();
                        switch (text10) {
                            case "ERROR":
                                status += "E";
                                item2.Properties.Add(new PSNoteProperty("DownloadResult", "Error"));
                                item2.Properties.Add(new PSNoteProperty("Status", status));
                                item2.Properties.Add(new PSNoteProperty("Result", "Error"));
                                break;
                            case "CANCELLED":
                                status += "C";
                                item2.Properties.Add(new PSNoteProperty("DownloadResult", "Cancelled"));
                                item2.Properties.Add(new PSNoteProperty("Status", status));
                                item2.Properties.Add(new PSNoteProperty("Result", "Cancelled"));
                                break;
                            case "ACKNOWLEDGED":
                                status += "D";
                                item2.Properties.Add(new PSNoteProperty("DownloadResult", "Downloaded"));
                                item2.Properties.Add(new PSNoteProperty("Status", status));
                                item2.Properties.Add(new PSNoteProperty("Result", "Downloaded"));
                                break;
                        }
                    }
                }

                var dlCount = collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded").Count();
                WriteObject(collection.Where(x => x.Properties["Result"].Value.ToString() == "Downloaded"), true);
                WriteVerbose("Downloaded [" + dlCount + "] Updates");
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
                    PSWUSettings.Add("Properties", "*");
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