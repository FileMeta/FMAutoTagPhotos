using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using WindowsSearch;

namespace FMAutoTagPhotos
{
    class PhotoTagger
    {
        string m_libraryPath;

        public PhotoTagger(string libraryPath)
        {
            m_libraryPath = libraryPath;
        }

        public bool Verbose { get; set; }

        public void TagAllMatches(string path, string tag)
        {
            int nJpeg = 0;
            int nTagged = 0;

            // Open the Windows Search system
            using (WindowsSearchSession winSrchSession = new WindowsSearchSession(m_libraryPath))
            {
                // Open the property system
                using (WinShell.PropertySystem propsys = new WinShell.PropertySystem())
                {
                    // Get the property keys
                    WinShell.PROPERTYKEY pkFilename = propsys.GetPropertyKeyByName("System.FileName");
                    WinShell.PROPERTYKEY pkHorizontalSize = propsys.GetPropertyKeyByName("System.Image.HorizontalSize");
                    WinShell.PROPERTYKEY pkVerticalSize = propsys.GetPropertyKeyByName("System.Image.VerticalSize");
                    WinShell.PROPERTYKEY pkCameraModel = propsys.GetPropertyKeyByName("System.Photo.CameraModel");
                    WinShell.PROPERTYKEY pkDateTaken = propsys.GetPropertyKeyByName("System.Photo.DateTaken");

                    // Process each matching file
                    foreach (string filename in new PhotoEnumerable(path))
                    {
                        ++nJpeg;
                        if (Verbose)
                        {
                            Console.WriteLine();
                            Console.WriteLine(filename);
                        }

                        // Open the property store for this file and retrieve the key properties
                        string propFilename = null;
                        UInt32? propHorizontalSize = null;
                        UInt32? propVerticalSize = null;
                        string propCameraModel = null;
                        DateTime? propDateTaken = null;
                        using (WinShell.PropertyStore store = WinShell.PropertyStore.Open(filename))
                        {
                            propFilename = store.GetValue(pkFilename) as string;
                            propVerticalSize = store.GetValue(pkVerticalSize) as UInt32?;
                            propHorizontalSize = store.GetValue(pkHorizontalSize) as UInt32?;
                            propCameraModel = store.GetValue(pkCameraModel) as string;
                            propDateTaken = store.GetValue(pkDateTaken) as DateTime?;
                        }

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

                        int matches = 0;

                        // Attempt to find using camera model and date taken
                        if (propCameraModel != null && propDateTaken != null)
                        {
                            /* Windows Search expects all times in Universal Time. This applies to Date Taken
                             * even tough it is stored in local time and does not include timezone info. Presumably,
                             * Windows Search converts back to local time using whatever timezone is current on the
                             * local machine. This might create a problem if the local computer and the computer being
                             * queried are not set to the same timezone.
                             */
                            string sql = string.Format(System.Globalization.CultureInfo.InvariantCulture,
                                "SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.ContentType = 'image/jpeg' AND System.Photo.CameraModel='{0}' AND System.Photo.DateTaken = '{1}' AND System.Image.HorizontalSize = {2} AND System.Image.VerticalSize = {3}",
                                propCameraModel, propDateTaken.Value.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss",System.Globalization.CultureInfo.InvariantCulture),
                                propHorizontalSize, propVerticalSize);
                            Console.WriteLine(sql);

                            using (var reader = winSrchSession.Query(sql))
                            {
                                matches += reader.WriteRowsToCsv(Console.Out);
                            }
                        }

                        if (Verbose) Console.WriteLine("{0} matches.", matches);
                    } // foreach Jpeg

                } // using PropSys
            } // using WindowsSearchSession
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
            int nTagged = 0;

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

/* Potential Fields on which to match.
Field presence count out of 19 sample .jpegs from a variety of cameras and scanners

    19: System.ContentType (Content type)
    19: System.FileName (Filename)
    19: System.Image.BitDepth (Bit depth)
    19: System.Image.Dimensions (Dimensions)
    19: System.Image.HorizontalResolution (Horizontal resolution)
    19: System.Image.HorizontalSize (Width)
    19: System.Image.VerticalResolution (Vertical resolution)
    19: System.Image.VerticalSize (Height)
    19: System.Kind (Kind)
    17: System.Image.ColorSpace (Color representation)
    17: System.Photo.CameraManufacturer (Camera maker)
    17: System.Photo.CameraModel (Camera model)
    17: System.Photo.EXIFVersion (EXIF version)
    16: System.Image.ResolutionUnit (Resolution unit)
    15: System.Photo.DateTaken (Date taken)
    15: System.Photo.Orientation (Orientation)
    14: System.Photo.ExposureTime (Exposure time)
    14: System.Photo.Flash (Flash mode)
    14: System.Photo.FNumber (F-stop)
    14: System.Photo.MeteringMode (Metering mode)
    12: System.Photo.ExposureBias (Exposure bias)
    11: System.Photo.Aperture (Aperture)
    11: System.Photo.ShutterSpeed (Shutter speed)
    11: System.Photo.WhiteBalance (White balance)
    10: System.ApplicationName (Program name)
    10: System.Photo.FocalLength (Focal length)
    10: System.Photo.ISOSpeed (ISO speed)
     9: System.Photo.DigitalZoom (Digital zoom)
     7: System.Photo.LightSource (Light source)
     7: System.Photo.MaxAperture (Max aperture)
     6: System.Image.CompressedBitsPerPixel (Compressed bits/pixel)
     3: System.GPS.Altitude (Altitude)
     3: System.GPS.Latitude (Latitude)
     3: System.GPS.Longitude (Longitude)
     3: System.Photo.ExposureProgram (Exposure program)
     3: System.Photo.FocalLengthInFilm (35mm focal length)
     3: System.Photo.ProgramMode (Program mode)
     3: System.Photo.Sharpness (Sharpness)
     2: System.Photo.Contrast (Contrast)
     2: System.Photo.Saturation (Saturation)
     1: System.Copyright (Copyright)
     1: System.Image.ImageID (Image ID)
     1: System.Photo.Brightness (Brightness)
     1: System.Photo.SubjectDistance (Subject distance)
*/