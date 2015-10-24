using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMAutoTagPhotos
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (WindowsSearchSession session = new WindowsSearchSession(@"\\Ganymede\Archive\Photos"))
                {
                    string[] keywords = session.GetAllKeywords();
                    foreach(string keyword in keywords)
                    {
                        Console.WriteLine(keyword);
                    }
                }

            }
            catch(Exception err)
            {
#if DEBUG
                Console.Error.WriteLine(err.ToString());
#else
                Console.Error.WriteLine(err.Message);
#endif
            }

            if (Win32Interop.ConsoleHelper.IsSoleConsoleOwner)
            {
                Console.Write("Press any key to exit.");
                Console.ReadKey(true);
            }
        }
    }
}
