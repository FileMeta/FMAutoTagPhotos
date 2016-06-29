using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

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

}