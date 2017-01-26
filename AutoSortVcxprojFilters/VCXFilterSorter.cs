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
        public string Path { get; }
        public string ProjectName { get; }

        private FileSystemWatcher _watcher;

        public VCXFilterSorter(Project project)
        {
            ProjectName = System.IO.Path.GetFileName(project.FullName);
            Path = System.IO.Path.GetDirectoryName(project.FullName);

            _watcher = new FileSystemWatcher(Path, "*.vcxproj*");
            _watcher.Created += _watcher_Created;
            _watcher.Changed += _watcher_Changed;
            _watcher.EnableRaisingEvents = true;
        }

        private void _watcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (e.Name.EndsWith(".vcxproj.filters"))
            {
                // file changed try resort
                _watcher.EnableRaisingEvents = false;
                Sort(e.FullPath);
                _watcher.EnableRaisingEvents = true;
            }
        }

        private void _watcher_Created(object sender, FileSystemEventArgs e)
        {
            if (e.Name.EndsWith(".vcxproj.filters"))
            {
                // file created - try resort
                _watcher.EnableRaisingEvents = false;
                Sort(e.FullPath);
                _watcher.EnableRaisingEvents = true;
            }
        }

        public void StopWatching()
        {
            _watcher.EnableRaisingEvents = false;
        }

        public void StartWatching()
        {
            _watcher.EnableRaisingEvents = true;
        }

        public static void Sort(string filename)
        {
            try
            {
                if (System.IO.File.Exists(filename))
                {
                    var xmldoc = XDocument.Load(filename);
                    foreach(var itemgroup in xmldoc.Root.Elements("{http://schemas.microsoft.com/developer/msbuild/2003}ItemGroup"))
                    {
                        itemgroup.ReplaceNodes(from elem in itemgroup.Elements()
                                               orderby elem.Attribute("Include")?.Value
                                               select elem);
                    }
                    xmldoc.Save(filename);
                }
            } catch (Exception ex)
            {
            }
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
                    if (_watcher != null)
                    {
                        _watcher.EnableRaisingEvents = false;
                        _watcher.Changed -= _watcher_Changed;
                        _watcher.Created -= _watcher_Created;
                        _watcher.Dispose();
                        _watcher = null;
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
