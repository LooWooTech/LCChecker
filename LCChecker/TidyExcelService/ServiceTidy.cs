using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace TidyExcelService
{
    public partial class ServiceTidy : ServiceBase
    {
        private TidyService TidyManager = new TidyService();
        public ServiceTidy()
        {
            InitializeComponent();
            this.ServiceName = "LCChecker Tidy Service";
        }

        protected override void OnStart(string[] args)
        {
            TidyManager.Start();
        }

        protected override void OnStop()
        {
            TidyManager.Stop();
        }
    }
}
