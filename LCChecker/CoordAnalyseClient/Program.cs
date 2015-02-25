using ESRI.ArcGIS.esriSystem;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CoordAnalyseClient
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            init.Shutdown();
        }
    }
}
