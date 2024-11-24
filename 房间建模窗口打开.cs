using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Collections;
using System;
using Autodesk.Revit.UI;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;


namespace CodeAgent

{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class openWindow : IExternalCommand
    {


        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIApplication uIApp = commandData.Application;
            Application application = uIApp.Application;
            UIDocument uIDoc = uIApp.ActiveUIDocument;
            Document document = uIDoc.Document;
            Selection selection = uIDoc.Selection;
            View view = uIDoc.ActiveView;

            UserControl userControl1 = new UserControl(document, view);
            userControl1.Show();

            return Result.Succeeded;
        }
    }
}