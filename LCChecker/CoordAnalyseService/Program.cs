using ESRI.ArcGIS.esriSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace CoordAnalyseService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            var init = new AoInitializeClass();
            if (init.IsProductCodeAvailable(esriLicenseProductCode.esriLicenseProductCodeEngine) == esriLicenseStatus.esriLicenseAvailable)
            {
                init.Initialize(esriLicenseProductCode.esriLicenseProductCodeEngine);
            }
            else if (init.IsProductCodeAvailable(esriLicenseProductCode.esriLicenseProductCodeArcEditor) == esriLicenseStatus.esriLicenseAvailable)
            {
                init.Initialize(esriLicenseProductCode.esriLicenseProductCodeArcEditor);
            }

            var manager = new ServiceManager();
            Console.WriteLine("服务正在运行中，任意键停止");
            manager.Start();
            Console.ReadLine();
            manager.Stop();
            //Analyser.ProcessNext();
            init.Shutdown();
            

            /*ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new Service1() 
            };
            ServiceBase.Run(ServicesToRun);*/
        }
    }
}
