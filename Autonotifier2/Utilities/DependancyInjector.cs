using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autofac;
using Autonotifier2.BusinessLogic;
using Autonotifier2.DataAccess;
using Autonotifier2.Models;

namespace Autonotifier2.Utilities
{
    class DependancyInjector
    {
        public static IContainer GetInjector()
        {


            var builder = new ContainerBuilder();
            builder.RegisterType<ExcelGateway>().As<IExcelGateway>();
            builder.RegisterType<JsonColumnHeaderReader>().AsSelf();
            builder.RegisterType<OutstandingDataManager>().AsSelf();

            var container = builder.Build();
        //    var x=container.Resolve<ExcelGateway>();
            return container;
        }

       

     
    }
}
