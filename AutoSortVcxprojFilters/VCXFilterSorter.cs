using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AutoSortVcxprojFilters
{
    class VCXFilterSorter : IDisposable
    {
        private string Path { get; set; }
        private string ProjectName { get; set; }
        public  string FullProjectName { get; private set; }
        private string VcxprojFiltersName { get; set; }
        private string FullVcxprojFiltersName { get; set; }

        private FileSystemWatcher m_FileSystemWatcher;

        public VCXFilterSorter(Project project)
        {
            ProjectName = System.IO.Path.GetFileName(project.FullName);
            Path = System.IO.Path.GetDirectoryName(project.FullName);

            FullProjectName = Path + System.IO.Path.DirectorySeparatorChar + ProjectName;

            VcxprojFiltersName = ProjectName + @".filters";
            FullVcxprojFiltersName = Path + System.IO.Path.DirectorySeparatorChar + VcxprojFiltersName;

            m_FileSystemWatcher = new FileSystemWatcher(Path, VcxprojFiltersName);
            m_FileSystemWatcher.Created += m_FileSystemWatcher_Created;
            m_FileSystemWatcher.Changed += m_FileSystemWatcher_Changed;
            m_FileSystemWatcher.EnableRaisingEvents = true;
        }

        private void m_FileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (e.Name.EndsWith(".vcxproj.filters"))
            {
                Sort();
            }
        }

        private void m_FileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            if (e.Name.EndsWith(".vcxproj.filters"))
            {
                Sort();
            }
        }

        private void StopWatching()
        {
            m_FileSystemWatcher.EnableRaisingEvents = false;
        }

        private void StartWatching()
        {
            m_FileSystemWatcher.EnableRaisingEvents = true;
        }

        public void Sort()
        {
            StopWatching();

            if (System.IO.File.Exists(FullVcxprojFiltersName))
            {
                System.Threading.Thread.Sleep(200); // Avoid load fail

                XDocument xmlDocument = XDocument.Load(FullVcxprojFiltersName);

                foreach (var itemgroup in xmlDocument.Root.Elements("{http://schemas.microsoft.com/developer/msbuild/2003}ItemGroup"))
                {
                    itemgroup.ReplaceNodes(from elem in itemgroup.Elements()
                                           orderby elem.Attribute("Include")?.Value
                                           select elem);
                }

                xmlDocument.Save(FullVcxprojFiltersName);
            }

            StartWatching();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    if (m_FileSystemWatcher != null)
                    {
                        m_FileSystemWatcher.EnableRaisingEvents = false;
                        m_FileSystemWatcher.Changed -= m_FileSystemWatcher_Changed;
                        m_FileSystemWatcher.Created -= m_FileSystemWatcher_Created;
                        m_FileSystemWatcher.Dispose();
                        m_FileSystemWatcher = null;
                    }
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~VCXFilterSorter() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
