using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace AutoSortVcxprojFilters
{
    internal class SortFilterOnAfterSave : IVsRunningDocTableEvents
    {
        IVsRunningDocumentTable m_IVsRunningDocumentTable;

        public SortFilterOnAfterSave(IVsRunningDocumentTable i_IVsRunningDocumentTable)
        {
            m_IVsRunningDocumentTable = i_IVsRunningDocumentTable;
        }

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            uint pgrfRDTFlags, pdwReadLocks, pdwEditLocks, pitemid;
            IVsHierarchy ppHier;
            IntPtr ppunkDocData;

            string fullDocumentName;

            // Get document name
            m_IVsRunningDocumentTable.GetDocumentInfo(
                docCookie,
                out pgrfRDTFlags,
                out pdwReadLocks,
                out pdwEditLocks,
                out fullDocumentName,
                out ppHier,
                out pitemid,
                out ppunkDocData);

            if (fullDocumentName.EndsWith(@"vcxproj.filters"))
            {
                VCXFilterSorter.Sort(fullDocumentName);
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            return VSConstants.S_OK;
        }
    }
}
