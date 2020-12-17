using System.Configuration.Install;
using System.Reflection;
using System.ServiceProcess;

namespace appmon
{
    static class Program
    {
        static void Main(string[] args)
        {
            if (System.Environment.UserInteractive)
            {
                if (args.Length > 0)
                {
                    switch (args[0])
                    {
                        case "install":
                            ManagedInstallerClass.InstallHelper(new string[] { Assembly.GetExecutingAssembly().Location });
                            break;
                        case "uninstall":
                            ManagedInstallerClass.InstallHelper(new string[] { "/u", Assembly.GetExecutingAssembly().Location });
                            break;
                        case "test":
                            var svc = new AppMonService();
                            svc.Init();
                            for (int i = 0; i < 2; i++)
                            {
                                AppMonService.Monitor(svc);
                            }
                            System.Console.ReadLine();
                            svc.Stop();
                            break;
                        default:
                            break;
                    }
                }
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] { new AppMonService() };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
