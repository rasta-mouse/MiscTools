using System;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

using NDesk.Options;

namespace CsPosh
{
    class Program
    {
        public static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage:");
            p.WriteOptionDescriptions(Console.Out);
        }

        static void Main(string[] args)
        {
            var help = false;
            var outstring = false;
            var target = string.Empty;
            var code = string.Empty;

            var options = new OptionSet(){
                {"t|target=","Target Machine", o => target = o},
                {"c|code=","Code: Get-Process", o => code = o},
                {"o|outstring", "Append Out-String to code", o => outstring = true },
                {"h|?|help","Show Help", o => help = true},
            };

            try
            {
                options.Parse(args);

                if (help)
                {
                    ShowHelp(options);
                    return;
                }

                if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(code))
                {
                    ShowHelp(options);
                    return;
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                ShowHelp(options);
                return;
            }

            try
            {
                var uri = new Uri($"http://{target}:5985/WSMAN");
                var conn = new WSManConnectionInfo(uri);

                using (var runspace = RunspaceFactory.CreateRunspace(conn))
                {
                    runspace.Open();

                    using (var posh = PowerShell.Create())
                    {
                        posh.Runspace = runspace;
                        posh.AddScript(code);
                        if (outstring) { posh.AddCommand("Out-String"); }
                        var results = posh.Invoke();
                        var output = string.Join(Environment.NewLine, results.Select(R => R.ToString()).ToArray());
                        Console.WriteLine(output);
                    }

                    runspace.Close();
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }
        }
    }
}