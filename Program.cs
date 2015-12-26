using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FMAutoTagPhotos
{
    class Program
    {
        const string c_syntax =
@"Syntax:
    FMAutoTagPhotos -l <LibraryPath> [options]
Options:
    -h                  Display this help test.
    -lib <LibraryRoot>  Path to the root folder of the photos library.
                        The folder must be included in Windows Search.
    -alltags            List all tags presently in use.
    -match <Folder>     Folder of photos to match in the library.
    -tag <keyword>      Tag to apply to all of the matched photos.

Locates patches for all of the photos in the ""match"" folder and applies
the specified tag to them. Success or failure to find matches is listed
as output.
";

        static bool s_verbose = true;
                        
        static void Main(string[] args)
        {
            bool writeSyntax = false;
            bool writeAllTags = false;
            string libraryPath = null;
            string matchPath = null;

            try
            {
                for (int nArg = 0; nArg < args.Length; ++nArg)
                {
                    switch (args[nArg].ToLower())
                    {
                        case "-h":
                            writeSyntax = true;
                            break;

                        case "-lib":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-lib'");
                            libraryPath = Path.GetFullPath(args[nArg]);
                            if (!Directory.Exists(libraryPath))
                            {
                                throw new ArgumentException(String.Format("Folder '{0}' does not exist.", libraryPath));
                            }
                            break;

                        case "-alltags":
                            writeAllTags = true;
                            break;

                        case "-match":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-match'");
                            matchPath = Path.GetFullPath(args[nArg]);
                            if (!Directory.Exists(matchPath))
                            {
                                throw new ArgumentException(String.Format("Folder '{0}' does not exist.", matchPath));
                            }
                            break;

                    }
                }

                if (writeSyntax)
                {
                    // Do nothing here
                }
                else
                {
                    if (libraryPath == null) throw new ArgumentException("Library path not specified. Use '-l'");

                    using (WindowsSearchSession session = new WindowsSearchSession(libraryPath))
                    {

                        if (writeAllTags)
                        {
                            Console.Error.WriteLine("All Tags:");
                            Console.Error.WriteLine();
                            string[] keywords = session.GetAllKeywords();
                            foreach (string keyword in keywords)
                            {
                                Console.WriteLine(keyword);
                            }
                        }
                        else if (matchPath != null)
                        {
                            TagAllMatches(matchPath, session);
                        }
                        else
                        {
                            throw new ArgumentException("No operation option specified.");
                        }
                    }
                }
            }
            catch(Exception err)
            {
                if (err is ArgumentException) writeSyntax = true;
#if DEBUG
                Console.Error.WriteLine(err.ToString());
#else
                Console.Error.WriteLine(err.Message);
#endif
                Console.Error.WriteLine();
            }

            if (writeSyntax) Console.Error.Write(c_syntax);

            if (Win32Interop.ConsoleHelper.IsSoleConsoleOwner)
            {
                Console.Error.WriteLine();
                Console.Error.Write("Press any key to exit.");
                Console.ReadKey(true);
            }
        }

        static void TagAllMatches(string matchFolder, WindowsSearchSession session)
        {
            int nJpeg = 0;
            int nTagged = 0;
            foreach (string pattern in new string[] { "*.jpg", "*.jpeg"})
            {
                foreach (string filename in Directory.GetFiles(matchFolder, pattern))
                {
                    ++nJpeg;
                    if (s_verbose) Console.WriteLine(filename);
                    DumpProperties(filename);
                    Console.WriteLine();
                }
            }
        }

        static void DumpProperties(string filename)
        {
            using (WinShell.PropertySystem propsys = new WinShell.PropertySystem())
            {
                using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(filename))
                {
                    int count = store.Count;
                    for (int i = 0; i < count; ++i)
                    {
                        WinShell.PROPERTYKEY key = store.GetAt(i);

                        string name;
                        try
                        {
                            using (WinShell.PropertyDescription desc = propsys.GetPropertyDescription(key))
                            {
                                name = string.Concat(desc.CanonicalName, " ", desc.DisplayName);
                            }
                        }
                        catch
                        {
                            name = string.Format("({0}:{1})", key.fmtid, key.pid);
                        }

                        object value = store.GetValue(key);
                        string strValue;

                        if (value is string[])
                        {
                            strValue = string.Join(";", (string[])value);
                        }
                        else
                        {
                            strValue = value.ToString();
                        }
                        Console.WriteLine("{0}: {1}", name, strValue);
                    }
                }
            }

        }


    }
}
