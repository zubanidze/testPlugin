using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using WPF = System.Windows;
using Autodesk.Revit.UI;

namespace testPlugin
{
    [Transaction(TransactionMode.Manual)]
    public sealed partial class ApartmentColoring :  IExternalCommand
    {

        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Result result = Result.Failed;

            try
            {

                UIApplication ui_app = commandData?.Application;
                UIDocument ui_doc = ui_app?.ActiveUIDocument;
                Application app = ui_app?.Application;
                Document doc = ui_doc?.Document;

                using (var tr_gr = new TransactionGroup(doc, "Окрашивание квартир"))
                {

                    if (TransactionStatus.Started == tr_gr.Start())
                    {
                        if (DoWork(commandData, ref message, elements))
                        {
                            if (TransactionStatus.Committed == tr_gr.Assimilate())
                            {
                                result = Result.Succeeded;
                            }
                        }
                        else
                        {
                            tr_gr.RollBack();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                TaskDialog.Show("Возникла ошибка", ex.Message+"\n"+ex.StackTrace);

                result = Result.Failed;
            }
            return result;
        }
    }
}
