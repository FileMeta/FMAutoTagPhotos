using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;

namespace FMAutoTagPhotos
{
    class WindowsSearchSession : IDisposable
    {
        OleDbConnection m_dbConnection = null;
        string m_pathInUrlForm;
        string m_hostPrefix;

        public WindowsSearchSession(string path)
        {
            path = Path.GetFullPath(path);
            m_pathInUrlForm = path.Replace('\\', '/');
            
            // Get host prefix (empty string if localhost)
            if (m_pathInUrlForm.StartsWith("//", StringComparison.Ordinal))
            {
                int slash = m_pathInUrlForm.IndexOf('/', 2);
                if (slash > 1)
                {
                    m_hostPrefix = string.Concat(m_pathInUrlForm.Substring(2, slash - 2), ".");
                }
                else
                {
                    throw new ArgumentException(string.Format("WindowsSearchSession - Invalid Path: '{0}'", path), "path");
                }
            }
            else
            {
                m_hostPrefix = string.Empty;
            }

            m_dbConnection = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';");
            m_dbConnection.Open();
        }

        public string[] GetAllKeywords()
        {
            string query = string.Format("SELECT System.Keywords FROM {0}SystemIndex WHERE SCOPE='file:{1}'", m_hostPrefix, m_pathInUrlForm);
            Debug.WriteLine(query);

            HashSet<string> keywords = new HashSet<string>();

            int nReads = 0;
            int nValues = 0;
            int maxValuesPerRead = 0;

            using (OleDbCommand cmd = new OleDbCommand(query, m_dbConnection))
            {
                using (OleDbDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        ++nReads;

                        string[] values = rdr[0] as string[];
                        if (values != null)
                        {
                            foreach (string value in values)
                            {
                                ++nValues;
                                keywords.Add(value);
                            }
                            if (maxValuesPerRead < values.Length) maxValuesPerRead = values.Length;
                        }
                    }
                    rdr.Close();
                }
            }

            Debug.WriteLine("{0} reads, {1} values, {2} maxValuesPerRead, {3} distinct values", nReads, nValues, maxValuesPerRead, keywords.Count);

            List<string> kwList = new List<string>(keywords);
            kwList.Sort();

            return kwList.ToArray();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~WindowsSearchSession()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_dbConnection != null)
            {
                m_dbConnection.Dispose();
                m_dbConnection = null;
                GC.SuppressFinalize(this);
#if DEBUG
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose WindowsSearchSession.");
                }
#endif
            }
        }

    }
}
