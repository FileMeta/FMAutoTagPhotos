using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using WindowsSearch;
using System.Diagnostics;
using System.Data.OleDb;
using System.Globalization;

namespace FMAutoTagPhotos
{
    class PhotoTagger
    {
        const string c_BackupRevertBatName = "Revert.bat";
        const string c_BackupRevertHeader =
@"@ECHO OFF
ECHO This batch will undo FMAutoTagPhotos by replacing previously tagged
photos with backups.
ECHO.
ECHO lib:   {0}
ECHO match: {1}
ECHO tag:   {2}
ECHO.
CHOICE /C YN /M ""Proceed to revert changes?""
IF %ERRORLEVEL% == 1 GOTO PROCEED
:CANCEL
ECHO Canceled
GOTO FINISH
:PROCEED
ECHO.
ECHO Restoring backups...
@ECHO ON
";

        const string c_BackupRevertFooter =
@"@ECHO Revert complete.
:FINISH
@PAUSE
";

        // Property Keys
        // From: https://msdn.microsoft.com/en-us/library/windows/desktop/dd561977(v=vs.85).aspx
        static WinShell.PROPERTYKEY s_pkFilename = new WinShell.PROPERTYKEY("41CF5AE0-F75A-4806-BD87-59C7D9248EB9", 100); // System.FileName
        static WinShell.PROPERTYKEY s_pkHorizontalSize = new WinShell.PROPERTYKEY("6444048F-4C8B-11D1-8B70-080036B11A03", 3); // System.Image.HorizontalSize
        static WinShell.PROPERTYKEY s_pkVerticalSize = new WinShell.PROPERTYKEY("6444048F-4C8B-11D1-8B70-080036B11A03", 4); // System.Image.VerticalSize
        static WinShell.PROPERTYKEY s_pkCameraModel = new WinShell.PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 272); // System.Photo.CameraModel
        static WinShell.PROPERTYKEY s_pkDateTaken = new WinShell.PROPERTYKEY("14B81DA1-0135-4D31-96D9-6CBFC9671A99", 36867); // System.Photo.DateTaken
        static WinShell.PROPERTYKEY s_pkKeywords = new WinShell.PROPERTYKEY("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 5); // System.Keywords

        string m_libraryPath;

        public PhotoTagger(string libraryPath)
        {
            m_libraryPath = libraryPath;
        }

        public bool Verbose { get; set; }
        public bool Simulate { get; set; }

        public void TagAllMatches(string path, string tag, bool delsrc, string backupPath)
        {
            int nJpeg = 0;
            int nJpegsMatched = 0;

            TextWriter backupBatch = null;
            try
            {
                if (Simulate)
                {
                    Console.WriteLine("Simulating Operations.");
                }

                if (!string.IsNullOrEmpty(backupPath))
                {
                    backupBatch = new StreamWriter(Path.Combine(backupPath, c_BackupRevertBatName), false);
                    backupBatch.WriteLine(c_BackupRevertHeader, m_libraryPath, path, tag);
                }

                // Open the Windows Search system
                using (WindowsSearchSession winSrchSession = new WindowsSearchSession(m_libraryPath))
                {
                    // Process each matching file
                    foreach (string filename in new PhotoEnumerable(path))
                    {
                        ++nJpeg;
                        Console.WriteLine();
                        Console.WriteLine(filename);

                        // Open the property store for this file and retrieve the key properties
                        string propFilename = null;
                        UInt32? propHorizontalSize = null;
                        UInt32? propVerticalSize = null;
                        string propCameraModel = null;
                        DateTime? propDateTaken = null;
                        using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(filename))
                        {
                            propFilename = store.GetValue(s_pkFilename) as string;
                            propVerticalSize = store.GetValue(s_pkVerticalSize) as UInt32?;
                            propHorizontalSize = store.GetValue(s_pkHorizontalSize) as UInt32?;
                            propCameraModel = store.GetValue(s_pkCameraModel) as string;
                            propDateTaken = store.GetValue(s_pkDateTaken) as DateTime?;
                        }

                        // Get the pixel sample
                        byte[] pixelSample = GetPixelSample(filename);

                        if (Verbose)
                        {
                            Console.WriteLine("Filename:       {0}", propFilename ?? "(null)");
                            Console.WriteLine("HorizontalSize: {0}", propHorizontalSize ?? 0);
                            Console.WriteLine("VerticalSize:   {0}", propVerticalSize ?? 0);
                            Console.WriteLine("CameraModel:    {0}", propCameraModel ?? "(null)");
                            Console.WriteLine("DateTaken:      {0}", (propDateTaken == null) ? "(null)" : propDateTaken.ToString());
                        }

                        if (propFilename == null || propHorizontalSize == null || propVerticalSize == null)
                        {
                            if (Verbose) Console.WriteLine("Invalid image metadata. Unable to match.");
                            continue;
                        }

                        var matches = new List<string>();

                        // Manage lifetime of dataReader
                        OleDbDataReader dataReader = null;
                        try
                        {
                            // Attempt to find using camera model and date taken
                            if (propCameraModel != null && propDateTaken != null)
                            {
                                /* Windows Search expects all times in Universal Time. This applies to Date Taken
                                 * even tough it is stored in local time and does not include timezone info. Presumably,
                                 * Windows Search converts back to local time using whatever timezone is current on the
                                 * local machine. This might create a problem if the local computer and the computer being
                                 * queried are not set to the same timezone.
                                 * Ways to compensate would be to figure out the timezone of the other computer or to specify
                                 * a time range. So far, however, this is working.
                                 */
                                string sql = string.Format(CultureInfo.InvariantCulture,
                                    "SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.ContentType = 'image/jpeg' AND System.Photo.CameraModel = '{0}' AND System.Photo.DateTaken = '{1}' AND System.Image.HorizontalSize = {2} AND System.Image.VerticalSize = {3}",
                                    SqlEncode(propCameraModel), propDateTaken.Value.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                                    propHorizontalSize, propVerticalSize);

                                dataReader = winSrchSession.Query(sql);
                                if (!dataReader.HasRows)
                                {
                                    dataReader.Close();
                                    dataReader = null;
                                }
                            }

                            // If no matches, attempt to find based on filename and dimensions
                            if (dataReader != null && dataReader.HasRows)
                            {
                                Console.WriteLine("Matching by metadata.");
                            }
                            else
                            {
                                string sql = string.Format(CultureInfo.InvariantCulture,
                                    "SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.ContentType = 'image/jpeg' AND System.FileName LIKE '{0}' AND System.Image.HorizontalSize = {1} AND System.Image.VerticalSize = {2}",
                                    SqlEncode(propFilename), propHorizontalSize, propVerticalSize);

                                dataReader = winSrchSession.Query(sql);
                                if (dataReader != null && dataReader.HasRows)
                                {
                                    Console.WriteLine("Matching by filename.");
                                }
                            }

                            while (dataReader.Read())
                            {
                                string matchFilename = dataReader.GetString(0);
                                byte[] matchSample = GetPixelSample(matchFilename);
                                if (EqualsPixelSample(pixelSample, matchSample))
                                {
                                    matches.Add(matchFilename);
                                }
                                else
                                {
                                    Console.WriteLine("   '{0}' fails pixel match.", matchFilename);
                                }
                            }
                        }
                        finally
                        {
                            if (dataReader != null)
                            {
                                dataReader.Close();
                                dataReader = null;
                            }
                        }

                        // Process the matches
                        int nTagged = 0;
                        foreach (string matchFilename in matches)
                        {
                            Console.WriteLine("Match: " + matchFilename);

                            try
                            {
                                // See if the match is already tagged with this keyword
                                string[] keywords;
                                using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(matchFilename))
                                {
                                    keywords = store.GetValue(s_pkKeywords) as string[];
                                }
                                if (keywords == null)
                                {
                                    keywords = new string[0];
                                }
                                else
                                {
                                    bool alreadyTagged = false;
                                    foreach (string keyword in keywords)
                                    {
                                        if (keyword.Equals(tag, StringComparison.OrdinalIgnoreCase))
                                        {
                                            alreadyTagged = true;
                                            break;
                                        }
                                    }
                                    if (alreadyTagged)
                                    {
                                        Console.WriteLine("   Already tagged with keyword '{0}'", tag);
                                        continue;
                                    }
                                }

                                // Simulate setting the tag
                                if (Simulate)
                                {
                                    if (backupBatch != null)
                                    {
                                        string backupFilename = Path.Combine(backupPath, string.Format(CultureInfo.InvariantCulture, "({0:D3}) {1}", nTagged, Path.GetFileName(matchFilename)));
                                        Console.WriteLine("   (Simulate on) Would be backed up to '{0}'.", backupFilename);
                                        backupBatch.WriteLine("REM MOVE /Y \"{0}\" \"{1}\"", backupFilename, matchFilename);
                                    }
                                    Console.WriteLine("   (Simulate on) Would be tagged with keyword '{0}'", tag);
                                    ++nTagged;
                                }

                                // Actually set the tag
                                else
                                {
                                    if (backupBatch != null)
                                    {
                                        string backupFilename = Path.Combine(backupPath, string.Format(CultureInfo.InvariantCulture, "({0:D3}) {1}", nTagged, Path.GetFileName(matchFilename)));
                                        Console.WriteLine("   Backing up to '{0}'.", backupFilename);
                                        File.Copy(matchFilename, backupFilename);
                                        backupBatch.WriteLine("MOVE /Y \"{0}\" \"{1}\"", backupFilename, matchFilename);
                                    }

                                    Console.WriteLine("   Tagging with keyword '{0}'.", tag);
                                    using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(matchFilename, true))
                                    {
                                        var keywordList = new List<string>(keywords);
                                        keywordList.Add(tag);
                                        keywords = keywordList.ToArray();
                                        store.SetValue(s_pkKeywords, keywords);
                                        store.Commit();
                                    }

                                    ++nTagged;
                                }
                            }
                            catch (Exception err)
                            {
                                Console.WriteLine("Failed to tag match:");
                                Console.WriteLine("   " + err.Message);
                                Console.WriteLine();
                            }


                        } // foreach match

                        // If specified, delete the original
                        if (delsrc && nTagged > 0)
                        {
                            Console.WriteLine("Deleting '{0}'.", filename);
                            try
                            {
                                File.Delete(filename);
                            }
                            catch (Exception err)
                            {
                                Console.WriteLine("Failed to delete:");
                                Console.WriteLine("   " + err.Message);
                                Console.WriteLine();
                            }

                        }

                        Console.WriteLine("{0} matches.", matches.Count);
                        Console.WriteLine("{0} tagged.", nTagged);

                        if (nTagged != 0) ++nJpegsMatched;
                    } // foreach Jpeg

                } // using WindowsSearchSession
            }
            finally
            {
                if (backupBatch != null)
                {
                    backupBatch.Write(c_BackupRevertFooter);
                    backupBatch.Dispose();
                    backupBatch = null;
                }
            }

            Console.WriteLine();
            Console.WriteLine("{0} JPEG Photos Found.", nJpeg);
            Console.WriteLine("{0} Matched and tagged.", nJpegsMatched);
        }

        /*
        Potential Fields on which to match.

        For digital camera photos:
        System.Photo.CameraManufacturer
        System.Photo.CameraModel
        System.Photo.EXIFVersion (probably not because it could be changed by photo software)
        System.Image.HorizontalSize
        System.Image.VerticalSize

        For others:
        System.FileName
        */

        public void Dump(string path, string tag)
        {
            int nJpeg = 0;

            foreach (string filename in new PhotoEnumerable(path))
            {
                ++nJpeg;
                if (Verbose) Console.WriteLine(filename);
                DumpProperties(filename, 3);
                Console.WriteLine();

                //if (nJpeg > 5) break;
            }
        }

        static void DumpProperties(string filename, int indent=0)
        {
            List<KeyValuePair<string, string>> values = new List<KeyValuePair<string, string>>();

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
                                name = string.Concat(desc.CanonicalName, " (", desc.DisplayName, ")");
                            }
                        }
                        catch
                        {
                            name = string.Concat("{", key.fmtid.ToString(), "}");
                        }

                        object value = store.GetValue(key);
                        string strValue;

                        if (value == null)
                        {
                            strValue = "(Null)";
                        }
                        else if (value is string[])
                        {
                            strValue = string.Join(";", (string[])value);
                        }
                        else if (value is double[])
                        {
                            strValue = string.Join(";", (IEnumerable<double>)value);
                        }
                        else
                        {
                            strValue = value.ToString();
                        }
                        values.Add(new KeyValuePair<string, string>(name, strValue));
                    }
                }
            } // using PropSys

            values.Sort(delegate (KeyValuePair<string, string> a, KeyValuePair<string, string> b)
            {
                return string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            });
            string indentStr = new string(' ', indent);
            foreach(var pair in values)
            {
                Console.WriteLine("{0}{1}: {2}", indentStr, pair.Key, pair.Value);
            }
        }

        // Take 128 samples in the form of 16 rows by 8 samples, each sample being 4 pixels wide
        // pixels are three bytes long in rgb form so each sample is 12 bytes and the total pixel
        // sample set is 16 x 8 x 4 x 3 or 1,536 bytes. We prefix the sample with the original
        // bitmap width and height each as a 32-bit integer. This makes the whole sample 1,542 bytes
        // long.
        const int c_pixelSampleColumns = 8;
        const int c_pixelSampleRows = 16;
        const int c_pixelSamplePixels = 4;  // Keeps things on a 32-bit boundary which is handy
        const int c_bytesPerPixel = 3;
        const int c_pixelPrefixSize = sizeof(int) * 2;  // Height and width prefix
        const int c_sampleSize = c_bytesPerPixel * c_pixelSamplePixels;
        const int c_totalSampleSize = c_pixelPrefixSize + (c_sampleSize * c_pixelSampleRows * c_pixelSampleColumns);

        /// <summary>
        /// Return a sample of pixels from a JPEG that can be compared with other samples
        /// with a very low probability that matching samples don't represent matching images.
        /// </summary>
        /// <param name="jpegFilename"></param>
        /// <returns>A byte string that may be compared with other bitmap samples.</returns>
        static byte[] GetPixelSample(string jpegFilename)
        {
            // Open the JPEG
            using (var jpegStream = new FileStream(jpegFilename, FileMode.Open, FileAccess.Read))
            {
                // Decode the jpeg and get the bitmap
                var jpegDecoder = new System.Windows.Media.Imaging.JpegBitmapDecoder(jpegStream, System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat, System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                System.Windows.Media.Imaging.BitmapFrame bitmap = jpegDecoder.Frames[0];

                if (bitmap.PixelWidth < 4) throw new InvalidDataException("Image is too narrow to match.");

                byte[] pixels = new byte[c_totalSampleSize];

                // Copy in the width and height
                {
                    int[] wh = new int[2];
                    wh[0] = bitmap.PixelWidth;
                    wh[1] = bitmap.PixelHeight;
                    Buffer.BlockCopy(wh, 0, pixels, 0, 8);
                }

                // This will take us close to the right and bottom edges within the limits of integer rounding
                int xInterval = ((bitmap.PixelWidth / c_pixelSamplePixels - 1) / (c_pixelSampleColumns - 1)) * c_pixelSamplePixels;
                int yInterval = (bitmap.PixelHeight - 1) / (c_pixelSampleRows - 1);
                for (int row = 0; row<c_pixelSampleRows; ++row)
                {
                    for (int col = 0; col<c_pixelSampleColumns; ++col)
                    {
                        var rect = new System.Windows.Int32Rect(col * xInterval, row * yInterval, c_pixelSamplePixels, 1);
                        int offset = c_pixelPrefixSize + ((row * c_pixelSampleColumns) + col) * c_sampleSize;
                        bitmap.CopyPixels(rect, pixels, c_sampleSize, offset);
                    }
                }

                return pixels;
            }
        }

        static void DumpPixelSample(byte[] pixels)
        {
            for (int i = 0; i < c_totalSampleSize; ++i) Console.Write(pixels[i].ToString("x2"));
            Console.WriteLine();
        }

        static bool EqualsPixelSample(byte[] pixels1, byte[] pixels2)
        {
            int len = pixels1.Length;
            if (len != pixels2.Length) return false;
            for (int i=0; i<len; ++i)
            {
                if (pixels1[i] != pixels2[i]) return false;
            }
            return true;
        }

        static string SqlEncode(string value)
        {
            return value.Replace("'", "''");
        }

    } // Class PhotoTagger


    class PhotoEnumerable : IEnumerable<string>
    {
        static public readonly char[] s_wildcards = new char[] { '*', '?' };

        string m_path;

        public PhotoEnumerable(string path)
        {
            m_path = path;
        }

        public IEnumerator<string> GetEnumerator()
        {
            return new PhotoEnumerator(m_path);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<string>)this).GetEnumerator();
        }

        // Much like Path.GetFullPath except that it also handles wildcards
        public static string GetFullPath(string path)
        {
            if (path.IndexOfAny(s_wildcards) >= 0)
            {
                return Path.Combine(Path.GetFullPath(Path.GetDirectoryName(path)), Path.GetFileName(path));
            }
            else
            {
                return Path.GetFullPath(path);
            }
        }
    }

    class PhotoEnumerator : IEnumerator<string>
    {
        string[] m_paths;
        int m_pathIndex = -1;
        string[] m_matches;
        int m_matchIndex = -1;

        public PhotoEnumerator(string path)
        {
            // If the path has wildcards then use it as-is
            if (path.IndexOfAny(PhotoEnumerable.s_wildcards) >= 0)
            {
                m_paths = new string[] { path };
            }
            else if (Directory.Exists(path))
            {
                m_paths = new string[2];
                m_paths[0] = Path.Combine(path, "*.jpg");
                m_paths[1] = Path.Combine(path, "*.jpeg");
            }
            else if (File.Exists(path))
            {
                m_paths = new string[] { path };
            }
            else
            {
                throw new ArgumentException(string.Format("Path \"{0}\" does not exist.", path));
            }
        }

        public string Current
        {
            get
            {
                if (m_pathIndex < 0) throw new InvalidOperationException("Enumerator before first item.");
                if (m_pathIndex >= m_paths.Length) throw new InvalidOperationException("Enumerator after last item.");
                return m_matches[m_matchIndex];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return ((IEnumerator<string>)this).Current;
            }
        }

        public bool MoveNext()
        {
            if (m_pathIndex >= m_paths.Length) return false;
            ++m_matchIndex;
            while (m_matches == null || m_matchIndex >= m_matches.Length)
            {
                m_matches = null;
                ++m_pathIndex;
                if (m_pathIndex >= m_paths.Length) return false;
                m_matches = Directory.GetFiles(Path.GetDirectoryName(m_paths[m_pathIndex]),
                    Path.GetFileName(m_paths[m_pathIndex]));
                m_matchIndex = 0;
            }
            return true;
        }

        public void Reset()
        {
            m_pathIndex = -1;
            m_matchIndex = -1;
            m_matches = null;
        }

        public void Dispose()
        {
            Reset();
        }

    }

    static class CollectionHelp
    {
        public static void Increment(this Dictionary<string, int> dict, string key)
        {
            int count;
            if (!dict.TryGetValue(key, out count)) count = 0;
            dict[key] = count + 1;
        }

        public static int Count(this Dictionary<string, int> dict, string key)
        {
            int count;
            if (!dict.TryGetValue(key, out count)) count = 0;
            return count;
        }

        public static void Dump(this Dictionary<string, int> dict, TextWriter writer)
        {
            List<KeyValuePair<string, int>> list = new List<KeyValuePair<string, int>>(dict);
            list.Sort(delegate (KeyValuePair<string, int> a, KeyValuePair<string, int> b)
            {
                int diff = b.Value - a.Value;
                return (diff != 0) ? diff : string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            });
            foreach (var pair in list)
            {
                writer.WriteLine("{0,6}: {1}", pair.Value, pair.Key);
            }
        }
    }

}

