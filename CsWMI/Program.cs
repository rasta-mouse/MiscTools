using System;
using System.Management;

namespace CsWMI
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine(" [x] Invalid number of arguments");
                Console.WriteLine("     Usage: WMI.exe <targetMachine> <command> <method>");
                return;
            }

            string target = args[0];
            string command = args[1];
            string method = args[2];

            ManagementScope scope = null;
            ManagementBaseObject result = null;

            try
            {
                scope = new ManagementScope($@"\\{target}\root\cimv2");
                scope.Connect();
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
                return;
            }

            if (method.ToLower() == "processcallcreate")
            {
                result = ProcessCallCreate(scope, command);
            }

            Console.WriteLine(" [*] Return Value: {0}", result["returnValue"]);
            Console.WriteLine(" [*] ProcessId: {0}", result["ProcessId"]);
        }

        public static ManagementBaseObject ProcessCallCreate(ManagementScope scope, string command)
        {
            ManagementBaseObject result = null;

            try
            {
                ManagementClass mClass = new ManagementClass(scope, new ManagementPath("Win32_Process"), new ObjectGetOptions());
                ManagementBaseObject parameters = mClass.GetMethodParameters("Create");
                PropertyDataCollection properties = parameters.Properties;
                parameters["CommandLine"] = command;

                result = mClass.InvokeMethod("Create", parameters, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }

            return result;
        }
    }
}