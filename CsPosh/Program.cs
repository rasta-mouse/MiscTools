using System;
using System.Linq;
using System.Security;
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
            var encoded = string.Empty;
            var domain = string.Empty;
            var username = string.Empty;
            var password = string.Empty;

            var options = new OptionSet(){
                {"t|target=", "Target machine", o => target = o},
                {"c|code=", "Code to execute", o => code = o},
                {"e|encoded=", "Encoded Code to execute", o => encoded = o},
                {"o|outstring", "Append Out-String to code", o => outstring = true },
                {"d|domain=", "Domain for alternate credentials", o => domain = o },
                {"u|username=", "Username for alternate credentials", o => username = o },
                {"p|password=", "Password for alternate credentials", o => password = o },
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
                
                if (!string.IsNullOrEmpty(encoded))
                {
                    code = System.Text.ASCIIEncoding.ASCII.GetString(System.Convert.FromBase64String(encoded));
                    Console.WriteLine("Encoded command to execute: " + code);
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

                if (!string.IsNullOrEmpty(domain) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    var pass = new SecureString();
                    foreach (char c in password.ToCharArray())
                        pass.AppendChar(c);

                    var cred = new PSCredential($"{domain}\\{username}", pass);

                    conn.Credential = cred;
                }

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
