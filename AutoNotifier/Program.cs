using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autofac;
using AutoNotifier.BusinessLogic;
using AutoNotifier.Models;
using AutoNotifier.Utilities;

namespace AutoNotifier
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            var container = DependancyInjector.GetInjector();
            var excelGateway = container.Resolve<IExcelGateway>();
            var jsonColumnHeaderReader = container.Resolve<JsonColumnHeaderReader>();
            var outstandingDataManager = container.Resolve<OutstandingDataManager>();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(outstandingDataManager));
        }
    }
}
