using Autofac;
using AutoNotifier.BusinessLogic;
using AutoNotifier.DataAccess;
using AutoNotifier.Models;

namespace AutoNotifier.Utilities
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
