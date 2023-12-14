using System;
using System.ServiceProcess;
using System.Threading;

namespace PSWindowsUpdate {
    internal class WUServiceController {
        private string ComputerName;
        private string ServiceName;

        public WUServiceController(string computerName, string serviceName) {
            ComputerName = computerName;
            ServiceName = serviceName;
        }

        /// <summary>Restart service on specific machine.</summary>
        public void RestartService() {
            try {
                var serviceController = new ServiceController(ServiceName, ComputerName);
                Thread.Sleep(500);
                serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                Thread.Sleep(500);
                serviceController.WaitForStatus(ServiceControllerStatus.Running);
                serviceController.Dispose();
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
        }

        /// <summary>Stope service on specific machine.</summary>
        public void StopService() {
            try {
                var serviceController = new ServiceController(ServiceName, ComputerName);
                if (serviceController.Status == ServiceControllerStatus.Running) {
                    Thread.Sleep(500);
                    serviceController.Stop();
                    serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                }

                serviceController.Dispose();
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
        }

        /// <summary>Start service on specific machine.</summary>
        public void StartService() {
            try {
                var serviceController = new ServiceController(ServiceName, ComputerName);
                if (serviceController.Status == ServiceControllerStatus.Stopped) {
                    Thread.Sleep(500);
                    serviceController.Start();
                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                }

                serviceController.Dispose();
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
        }
    }
}