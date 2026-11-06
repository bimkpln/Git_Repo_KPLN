using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Forms;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using System;
using System.Linq;

namespace KPLN_OpeningHoleManager.ExecutableCommand
{
    internal sealed class LoadOpeningFamiliesExcCmd : IExecutableCommand
    {
        private readonly MainViewModel _viewModel;
        private readonly Document _document;

        public LoadOpeningFamiliesExcCmd(MainViewModel viewModel, Document document)
        {
            _viewModel = viewModel;
            _document = document;
        }

        public Result Execute(UIApplication app)
        {
            if (!_document.IsValidObject || app.ActiveUIDocument == null || !_document.Equals(app.ActiveUIDocument.Document))
                return Result.Cancelled;

            try
            {
                using (Transaction transaction = new Transaction(_document, "KPLN: Загрузить семейства отверстий"))
                {
                    transaction.Start();
                    foreach (bool round in new[] { false, true })
                    {
                        string name = OpeningTypeMapping.GetFamilyName(_document, round);
                        bool exists = new FilteredElementCollector(_document).OfClass(typeof(Family)).Cast<Family>().Any(f => f.Name == name);
                        if (!exists && !_document.LoadFamily(OpeningTypeMapping.GetFamilyPath(_document, round)))
                            throw new InvalidOperationException("Не удалось загрузить семейство: " + name);
                    }
                    transaction.Commit();
                }

                _viewModel.RefreshTypeMapping();
                new ParamsOpenHoleByIOSElemsForm(_viewModel).ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Семейства отверстий", KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex));
                return Result.Failed;
            }
        }
    }
}
