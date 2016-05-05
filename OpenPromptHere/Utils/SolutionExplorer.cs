using System;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace OpenPromptHere.Utils
{
    public sealed class SolutionExplorer
    {
        public static Project GetSelectedProject()
        {
            return GetSelectedItem() as Project;
        }

        private static object GetSelectedItem()
        {
            var monitorSelection = (IVsMonitorSelection)Package.GetGlobalService(typeof(SVsShellMonitorSelection));

            IntPtr hierarchyPointer;
            IntPtr selectionContainerPointer;
            IVsMultiItemSelect multiItemSelect;
            uint projectItemId;

            monitorSelection.GetCurrentSelection(out hierarchyPointer,
                out projectItemId,
                out multiItemSelect,
                out selectionContainerPointer);

            using (var container = TypedObjectForIUnknown<ISelectionContainer>.Attach(selectionContainerPointer))
            using (var hierarchy = TypedObjectForIUnknown<IVsHierarchy>.Attach(hierarchyPointer))
            {
                object selection = null;
                ErrorHandler.ThrowOnFailure(
                    hierarchy.Instance.GetProperty(
                        projectItemId,
                        (int)__VSHPROPID.VSHPROPID_ExtObject,
                        out selection)
                    );

                return selection;
            }
        }
    }
}