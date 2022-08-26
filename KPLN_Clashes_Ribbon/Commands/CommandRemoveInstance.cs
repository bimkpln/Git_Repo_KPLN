using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandRemoveInstance : IExecutableCommand
    {
        private Document Document { get; }
        private FamilyInstance Instance { get; }
        public CommandRemoveInstance(Document document, FamilyInstance instance)
        {
            Document = document;
            Instance = instance;
        }
        public Result Execute(UIApplication app)
        {
            try
            {
                if (Document != null)
                {
                    Transaction t = new Transaction(Document, "Указатель");
                    t.Start();
                    try
                    {
                        if (Instance.Symbol.FamilyName == "ClashPoint")
                        {
                            Document.Delete(Instance.Id);
                            t.Commit();
                            return Result.Succeeded;
                        }  
                    }
                    catch (Exception)
                    {
                        t.RollBack();
                        return Result.Failed;
                    }
                    
                }
                return Result.Failed;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
    }
}
