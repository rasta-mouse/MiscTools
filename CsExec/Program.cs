using System;
using System.Threading;

using static CsExec.NativeMethods;

namespace CsExec
{
    class Program
    {
        static void Main(string[] args)
        {

            if (args.Length < 4)
            {
                Console.WriteLine(" [x] Invalid number of arguments");
                Console.WriteLine("     Usage: CsExec.exe <targetMachine> <serviceName> <serviceDisplayName> <binPath>");
                return;
            }

            string target = $@"\\{args[0]}";
            string serviceName = args[1];
            string serviceDisplayName = args[2];
            string binPath = args[3];

            try
            {
                // Connect to Service Manager
                IntPtr hSCManager = OpenSCManager(target, null, SCM_ACCESS.SC_MANAGER_ALL_ACCESS);

                if (hSCManager == IntPtr.Zero)
                {
                    Console.WriteLine($" [x] Could not open Service Manager on {target}");
                    return;
                }

                // Create Service
                IntPtr hService = CreateService(hSCManager, serviceName, serviceDisplayName, SERVICE_ACCESS.SERVICE_ALL_ACCESS, SERVICE_TYPE.SERVICE_WIN32_OWN_PROCESS, SERVICE_START.SERVICE_DEMAND_START, SERVICE_ERROR.SERVICE_ERROR_NORMAL, binPath, null, null, null, null, null);

                if (hService == IntPtr.Zero)
                {
                    Console.WriteLine($" [x] Could not create service {serviceName}");
                    return;
                }

                // Start Service
                bool running = StartService(hService, 0, null);

                Thread.Sleep(1000);

                if (running)
                {
                    // Stop Service
                    SERVICE_STATUS serviceStatus = new SERVICE_STATUS();
                    bool stopped = ControlService(hService, SERVICE_CONTROL.STOP, ref serviceStatus);

                    if (!stopped)
                        Console.WriteLine($" [x] Could not stop service {serviceName}");
                }
                else
                {
                    Console.WriteLine($" [x] Could not start service {serviceName}");
                }

                Thread.Sleep(1000);

                //Delete Service
                bool deleted = DeleteService(hService);

                if (!deleted)
                    Console.WriteLine($" [x] Could not delete service {serviceName}");

                // Close Handles
                CloseServiceHandle(hService);
                CloseServiceHandle(hSCManager);

            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
            
        }
    }
}