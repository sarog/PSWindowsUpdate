using Microsoft.Management.Infrastructure;
using System;
using System.ComponentModel;
using System.Globalization;

namespace PSWindowsUpdate {
    internal class WUCimSession {
        private string ComputerName;

        public WUCimSession(string computerName) {
            ComputerName = computerName;
        }

        public CimMethodResult InvokeCimMethod(
            string cimClassName,
            string cimNamespace,
            string cimMethod,
            CimMethodParametersCollection cimParameters) {
            var instance = new CimInstance(cimClassName, cimNamespace);
            var cimMethodResult = CimSession.Create(ComputerName).InvokeMethod(instance, cimMethod, cimParameters);
            var int32 = Convert.ToInt32(cimMethodResult.ReturnValue.Value, (IFormatProvider)CultureInfo.CurrentCulture);
            if (int32 != 0) {
                Console.WriteLine(new Win32Exception(int32).Message);
            }

            return cimMethodResult;
        }

        public CimInstance[] ScanForUpdates(string searchCriteria = "IsInstalled = 0") {
            var className = "MSFT_WUOperations";
            var namespaceName = "root/Microsoft/Windows/WindowsUpdate";
            var methodName = nameof(ScanForUpdates);
            var methodParameters = new CimMethodParametersCollection();
            methodParameters.Add(CimMethodParameter.Create("SearchCriteria", (object)searchCriteria, CimType.String, CimFlags.None));
            var instance = new CimInstance(className, namespaceName);
            return (CimInstance[])CimSession.Create(ComputerName).InvokeMethod(instance, methodName, methodParameters).OutParameters["Updates"].Value;
        }

        public CimMethodResult InstallUpdates(
            object updates,
            bool downloadOnly = false,
            bool rebootRequired = false) {
            var className = "MSFT_WUOperations";
            var namespaceName = "root/Microsoft/Windows/WindowsUpdate";
            var methodName = nameof(InstallUpdates);
            var methodParameters = new CimMethodParametersCollection();
            methodParameters.Add(CimMethodParameter.Create("Updates", updates, CimFlags.None));
            if (downloadOnly) {
                methodParameters.Add(CimMethodParameter.Create("DownloadOnly", (object)true, CimFlags.None));
            }

            if (rebootRequired) {
                methodParameters.Add(CimMethodParameter.Create("RebootRequired", (object)true, CimFlags.None));
            }

            var instance = new CimInstance(className, namespaceName);
            var cimSession = CimSession.Create(ComputerName);
            var cimMethodResult = (CimMethodResult)null;
            try {
                cimMethodResult = cimSession.InvokeMethod(instance, methodName, methodParameters);
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }

            return cimMethodResult;
        }
    }
}