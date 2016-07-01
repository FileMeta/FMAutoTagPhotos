using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using WindowsSearch;

namespace FMAutoTagPhotos
{
    class Program
    {
        const string c_syntax =
@"Syntax:
    FMAutoTagPhotos -lib <LibraryPath> [options]
Options:
    -h                  Display this help test.
    -lib <LibraryRoot>  Path to the root folder of the photos library.
                        The folder must be included in Windows Search.
    -alltags            List all tags presently in use.
    -match <Path>       Path to a folder or filename of photos to match in
                        the library. May include wildcards in the filename.
    -dump <Path>        Path to a folder or filename to photos for which all
                        metadata will be dumped. May include wildcards.
    -tag <keyword>      Tag to apply to all of the matched photos.

Locates patches for all of the photos in the ""match"" folder and applies
the specified tag to them. Success or failure to find matches is listed
as output.
";
// 78 Columns                                                                |

        static readonly char[] s_invalidTagChars = new char[]
        {
            '#',
            '/',
            '\\'
        };

        static bool s_verbose = true;
                        
        static void Main(string[] args)
        {
            bool writeSyntax = false;
            bool writeAllTags = false;
            string libraryPath = null;
            string matchPath = null;
            string dumpPath = null;
            string tag = null;

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
                                throw new ArgumentException(String.Format("Folder '{0}' not found.", libraryPath));
                            }
                            break;

                        case "-alltags":
                            writeAllTags = true;
                            break;

                        case "-match":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-match'");
                            matchPath = PhotoEnumerable.GetFullPath(args[nArg]);
                            break;

                        case "-dump":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-dump'");
                            dumpPath = PhotoEnumerable.GetFullPath(args[nArg]);
                            break;

                        case "-tag":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-tag'");
                            tag = args[nArg];
                            if (tag.IndexOfAny(s_invalidTagChars) >= 0) throw new ArgumentException("Prohibited character in tag.");
                            break;

                        default:
                            throw new ArgumentException(string.Format("Unexpected command-line parameter '{0}'", args[nArg]));
                    }
                }

                if (writeSyntax)
                {
                    // Do nothing here
                }
                else if (writeAllTags)
                {
                    using (WindowsSearchSession session = new WindowsSearchSession(libraryPath))
                    {
                        Console.Error.WriteLine("All Tags:");
                        Console.Error.WriteLine();
                        string[] keywords = session.GetAllKeywords();
                        foreach (string keyword in keywords)
                        {
                            Console.WriteLine(keyword);
                        }
                    }
                }
                else if (matchPath != null)
                {
                    PhotoTagger tagger = new PhotoTagger(libraryPath);
                    tagger.Verbose = s_verbose;
                    tagger.TagAllMatches(matchPath, tag);
                }
                else if (dumpPath != null)
                {
                    PhotoTagger tagger = new PhotoTagger(libraryPath);
                    tagger.Verbose = s_verbose;
                    tagger.Dump(dumpPath, tag);
                }
                else
                {
                    throw new ArgumentException("No operation option specified.");
                }
            }
            catch (Exception err)
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


    }
}
