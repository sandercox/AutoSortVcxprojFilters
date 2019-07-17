//------------------------------------------------------------------------------
// <copyright file="AutoSortPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;

namespace AutoSortVcxprojFilters
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(AutoSortPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string)]
    public sealed class AutoSortPackage : Package, IVsSolutionEvents3
    {
        /// <summary>
        /// AutoSortPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "01a71080-33cd-4769-883b-07242c7c6c3e";

        private EnvDTE.DTE _dte;
        private IVsSolution _solution;
        private uint _hSolutionEvents = uint.MaxValue;

        private List<VCXFilterSorter> sorters = new List<VCXFilterSorter>();
        internal List<VCXFilterSorter> Sorters { get { return sorters; } }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutoSortPackage"/> class.
        /// </summary>
        public AutoSortPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }
        #region IVSSolutionEvents3
        public int OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterClosingChildren(IVsHierarchy pHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterMergeSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpeningChildren(IVsHierarchy pHierarchy)
        {
            UpdateTrackedFilters();
            return VSConstants.S_OK;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            UpdateTrackedFilters();
            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            UpdateTrackedFilters();
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            UpdateTrackedFilters();
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeClosingChildren(IVsHierarchy pHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeOpeningChildren(IVsHierarchy pHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        void AdviseSolutionEvents()
        {
            UnAdviseSolutionEvents();

            _solution = this.GetService(typeof(SVsSolution)) as IVsSolution;

            if (_solution != null)
            {
                _solution.AdviseSolutionEvents(this, out _hSolutionEvents);
                if (_hSolutionEvents == uint.MaxValue)
                {
                    _solution = null;
                }
            }

        }

        void UnAdviseSolutionEvents()
        {
            if ( _solution != null)
            {
                _solution.UnadviseSolutionEvents(_hSolutionEvents);
                _solution = null;
            }
        }

        public EnvDTE.Project[] GetProjects()
        {
            return _dte.Solution.Projects
                .Cast<EnvDTE.Project>()
                .Where(x => { return x?.Object != null; })
                .ToArray();
        }

        private void UpdateTrackedFilters()
        {
            var projects = GetProjects();
            var newSorters = new List<VCXFilterSorter>();
            if (projects != null)
            {
                foreach(var proj in projects)
                {
                    var obj = sorters.Find(x => { return x.FullProjectName == proj.FullName; });
                    if (obj != null)
                    {
                        newSorters.Add(obj);
                        sorters.Remove(obj);
                    }
                    else
                    {
                        newSorters.Add(new VCXFilterSorter(proj));
                    }
                }
            }
            foreach(var sort in sorters)
            {
                sort.Dispose();
            }
            sorters = newSorters;
        }
        #endregion

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
            SortAllCommand.Initialize(this);

            this._dte = GetService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            AdviseSolutionEvents();
        }

        protected override void Dispose(bool disposing)
        {
            UnAdviseSolutionEvents();

            base.Dispose(disposing);
        }

        #endregion
    }
}
