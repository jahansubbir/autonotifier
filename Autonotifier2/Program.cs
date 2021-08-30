using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autofac;
using Autonotifier2.BusinessLogic;
using Autonotifier2.DataAccess;
using Autonotifier2.Models;
using Autonotifier2.Utilities;

namespace Autonotifier2
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var container=DependancyInjector.GetInjector();
            var excelGateway=container.Resolve<IExcelGateway>();
            var jsonColumnHeaderReader = container.Resolve<JsonColumnHeaderReader>();
            var outstandingDataManager = container.Resolve<OutstandingDataManager>();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(outstandingDataManager));
        }
    }
}
