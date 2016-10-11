#region Namespaces
using System;
using System.Collections.Generic;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace ManageSheetSpecs
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            CreateRibbonTab cTab = new CreateRibbonTab();
            cTab.tabAndButtons(a);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
