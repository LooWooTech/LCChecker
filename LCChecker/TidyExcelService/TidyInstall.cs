using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace TidyExcelService
{
    [RunInstaller(true)]
    public class TidyInstall:Installer
    {
        ServiceProcessInstaller processInstall;
        ServiceInstaller serviceInstall;

        public TidyInstall() {
            this.processInstall = new ServiceProcessInstaller();
            this.serviceInstall = new ServiceInstaller();

            processInstall.Account = ServiceAccount.LocalSystem;
            this.serviceInstall.ServiceName = "LCChecker TIDY Service";

            this.Installers.Add(this.serviceInstall);
            this.Installers.Add(this.processInstall);
        }
    }
}
