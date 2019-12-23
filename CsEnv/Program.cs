using System;

namespace CsEnv
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine(" [x] Invalid number of arguments");
                Console.WriteLine("     Usage: CsEnv.exe <variable> <value> <target>");
                return;
            }

            string variable = args[0];
            string value = args[1];
            string target = args[2];

            SetEnvVar(variable, value, target);
        }

        private static void SetEnvVar(string variable, string value, string target)
        {
            EnvironmentVariableTarget envTarget;

            switch (target) {
                case "user":
                    envTarget = EnvironmentVariableTarget.User;
                    break;

                case "machine":
                    envTarget = EnvironmentVariableTarget.Machine;
                    break;

                case "process":
                    envTarget = EnvironmentVariableTarget.Process;
                    break;

                default:
                    return;
                
            }

            try
            {
                Environment.SetEnvironmentVariable(variable, value, envTarget);
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }

        }
    }
}