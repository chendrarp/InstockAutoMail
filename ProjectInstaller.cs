using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace InstockAutoMailService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            HostInstaller = new ServiceInstaller();
            HostInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            HostInstaller.ServiceName = System.Configuration.ConfigurationManager.AppSettings["SERVICE_NAME"].ToString();
            HostInstaller.DisplayName = System.Configuration.ConfigurationManager.AppSettings["DISPLAY_NAME"].ToString();
            HostInstaller.Description = System.Configuration.ConfigurationManager.AppSettings["DISPLAY_NAME"].ToString();

            Installers.Add(HostInstaller);

            HostProcessInstaller = new ServiceProcessInstaller();
            HostProcessInstaller.Account = ServiceAccount.LocalSystem;
            Installers.Add(HostProcessInstaller);
        }
    }
}
