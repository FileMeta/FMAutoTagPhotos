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
// 78 Columns                                                                |
@"Syntax:
To add a keyword tag to all photos in a library that match a photo or
collection of photos.
    FMAutoTagPhotos -lib <LibraryRoot> -match <Path> -tag <Keyword> [-backup <Folder>]

To list all keyword tags used in the library.
    FMAutoTagPhotos -lib <LibraryRoot> -alltags

To dump the raw Windows Property System metadata for a set of files/photos
    FMAutoTagPhotos -dump <Path>

Arguments:
    -h                  Display this help text.
    -lib <LibraryRoot>  Path to the root folder of the photos library.
                        The folder must be included in Windows Search.
    -match <Path>       Path to filenames of photos to match in
                        the library. Usually includes wildcards in the
                        filename.
    -tag <Keyword>      Tag to apply to all of the matched photos.
    -delsrc             Delete the source file once a match is found and
                        tagged.
    -backup <Folder>    Path to a folder that will be loaded with originals
                        of each photo that is tagged and with a ""revert.bat""
                        file that will revert all photos back to their
                        pre-autotag state. Folder should be empty. It will be
                        created if it doesn't already exist.
    -simulate           Simulate the tagging operation but don't actually
                        perform it.
    -alltags            List all tags presently in use.
    -dump <Path>        Path to a folder or filename to photos for which all
                        metadata will be dumped. May include wildcards.

Photos are located by their metadata (date taken, camera model, etc.) and by
filename then matches are verified by comparing pixels from the images.
The filenames need not match unless metadata is not present.
";
// 78 Columns                                                                |

        static readonly char[] s_invalidTagChars = new char[]
        {
            '#',
            '/',
            '\\'
        };

        static bool s_verbose = false;
                        
        static void Main(string[] args)
        {
            bool writeSyntax = false;
            bool writeAllTags = false;
            string libraryPath = null;
            string matchPath = null;
            string dumpPath = null;
            string tag = null;
            bool delsrc = false;
            string backupPath = null;
            bool simulate = false;

            try
            {
                for (int nArg = 0; nArg < args.Length; ++nArg)
                {
                    switch (args[nArg].ToLower())
                    {
                        case "-?":
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

                        case "-tag":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-tag'");
                            tag = args[nArg];
                            if (tag.IndexOfAny(s_invalidTagChars) >= 0) throw new ArgumentException("Prohibited character in tag.");
                            break;

                        case "-delsrc":
                            delsrc = true;
                            break;

                        case "-backup":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-backup'");
                            backupPath = PhotoEnumerable.GetFullPath(args[nArg]);
                            break;

                        case "-dump":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-dump'");
                            dumpPath = PhotoEnumerable.GetFullPath(args[nArg]);
                            break;

                        case "-simulate":
                            simulate = true;
                            break;

                        case "-verbose":
                            s_verbose = true;
                            break;

                        default:
                            throw new ArgumentException(string.Format("Unexpected command-line parameter '{0}'", args[nArg]));
                    }
                }

                // Make sure only one command was specified
                {
                    int nCommands = 0;
                    if (writeAllTags) ++nCommands;
                    if (matchPath != null) ++nCommands;
                    if (dumpPath != null) ++nCommands;
                    if (nCommands != 1) throw new ArgumentException(string.Format("Must specify one and only one command: alltags, match, or dump."));
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
                    // If a backup folder was specified. Create it if necessary. Make sure it is empty.
                    if (backupPath != null)
                    {
                        if (!Directory.Exists(backupPath))
                        {
                            Directory.CreateDirectory(backupPath);
                        }
                        else if (Directory.GetFileSystemEntries(backupPath).Length != 0)
                        {
                            throw new ArgumentException(string.Format("Backup folder '{0}' must be empty.", backupPath));
                        }
                    }

                    PhotoTagger tagger = new PhotoTagger(libraryPath);
                    tagger.Verbose = s_verbose;
                    tagger.Simulate = simulate;
                    tagger.TagAllMatches(matchPath, tag, delsrc, backupPath);
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
#if DEBUG
                Console.Error.WriteLine(err.ToString());
#else
                Console.Error.WriteLine(err.Message);
#endif
                Console.Error.WriteLine();
                if (err is ArgumentException) Console.Error.Write("For syntax: FMAutoTagPhotos -h");
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
