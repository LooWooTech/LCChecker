using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;

namespace TidyExcelService
{
    public class TidyService
    {
        private Thread thread;
        public bool Signal = false;
        public void Start() {
            Signal = true;
            thread = new Thread(MainLoop);
            thread.Start();
        }
        public void Stop() {
            Signal = false;
            thread.Join();
        }

        public void MainLoop() {
            var loopInterval = int.Parse(ConfigurationManager.AppSettings["LoopInterval"]);
            while (Signal) {
                TidyExcel.Process();
                Thread.Sleep(loopInterval);
            }
        }
    }
}
