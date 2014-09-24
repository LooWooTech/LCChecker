using ESRI.ArcGIS.esriSystem;
using ICSharpCode.SharpZipLib.Zip;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace CoordAnalyseService
{
    public class ServiceManager
    {
        private Thread thread;

        private bool signal = false;

        public void Start()
        {
            signal = true;
            thread = new Thread(MainLoop);
            thread.Start();
        }

        public void Stop()
        {
            signal = false;
            thread.Join();
        }

        public void MainLoop()
        {
            var loopInterval = int.Parse(ConfigurationManager.AppSettings["LoopInterval"]);
            var init = new AoInitializeClass();
            if (init.IsProductCodeAvailable(esriLicenseProductCode.esriLicenseProductCodeEngine) == esriLicenseStatus.esriLicenseAvailable)
            {
                init.Initialize(esriLicenseProductCode.esriLicenseProductCodeEngine);
            }
            else if(init.IsProductCodeAvailable(esriLicenseProductCode.esriLicenseProductCodeArcEditor) == esriLicenseStatus.esriLicenseAvailable )
            {
                init.Initialize(esriLicenseProductCode.esriLicenseProductCodeArcEditor);
            }
            else
            {
                var logerror = log4net.LogManager.GetLogger("logerror");
                logerror.Error("没有合适的ArcGIS Desktop或ArcGIS Engine授权.");
                return;
            }
                
            while (signal)
            {
                Analyser.ProcessNext();
                Thread.Sleep(loopInterval);
            }
            init.Shutdown();
        }

       
    }
}
