using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace binariex
{
    class Program
    {
        static void Main(string[] args)
        {
            var quitOnEnd = args.Any(e => e == "-q");
            var inputPaths = args.Where(e => e[0] != '-').ToArray();

            new BinariexApp(inputPaths).Run();

            if (!quitOnEnd)
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }
    }
}
