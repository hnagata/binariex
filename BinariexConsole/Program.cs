using System;
using System.Collections.Generic;
using System.Linq;

namespace binariex
{
    class Program
    {
        static void Main(string[] args)
        {
            var inputPaths = new List<string>();
            var settingsPath = null as string;
            var quitOnEnd = false;

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-q":
                        quitOnEnd = true;
                        break;
                    case "-s":
                        if (i + 1 >= args.Length)
                        {
                            Console.WriteLine("Invalid arguments: -s");
                        }
                        else
                        {
                            settingsPath = args[++i];
                        }
                        break;
                    default:
                        inputPaths.Add(args[i]);
                        break;
                }
            }

            new BinariexApp(inputPaths, settingsPath).Run();

            if (!quitOnEnd)
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }
    }
}
