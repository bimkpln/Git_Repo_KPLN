using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools.Common.AR_PyatnGraph;
using KPLN_Tools.Forms.Models;
using KPLN_Tools.Forms.Models.Core;
using System.Windows.Forms;

namespace KPLN_Tools.ExecutableCommand
{
    internal class ExcCmdARPG_PreSetData : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly AR_PyatnGraph_VM _pgVM;
        private readonly ARPG_Room[] _aRPGRooms;
        private readonly ARPG_Flat[] _aRPGFlats;
        private readonly ARPG_TZ_FlatData[] _arpgTZFlatDatas;

        public ExcCmdARPG_PreSetData(Document doc, AR_PyatnGraph_VM pgVM, ARPG_Room[] aRPGRooms, ARPG_Flat[] aRPGFlats, ARPG_TZ_FlatData[] arpgTZFlatDatas)
        {
            _doc = doc;
            _pgVM = pgVM;
            _aRPGRooms = aRPGRooms;
            _aRPGFlats = aRPGFlats;
            _arpgTZFlatDatas = arpgTZFlatDatas;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction trans = new Transaction(_doc, "KPLN: Пятнография_Предустановка"))
            {
                trans.Start();

                ARPG_Flat.SetFlatCodeData(_pgVM.ARPG_TZ_MainData, _aRPGFlats, _arpgTZFlatDatas);
                if (ARPG_Flat.ErrorDict_Flat.Keys.Count != 0)
                {
                    HtmlOutput.PrintMsgDict("ОШИБКА", MessageType.Critical, ARPG_Flat.ErrorDict_Flat);

                    MessageBox.Show(
                        $"Запуск невозможен. Список критических ошибок - выведен отдельным окном",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    trans.RollBack();

                    return Result.Cancelled;
                }

                _doc.Regenerate();

                if (ARPG_Room.CheckRoomCodesByTZMain(_pgVM, _aRPGRooms))
                {
                    MessageBox.Show(
                        $"Плагин успешно завершил предустановку кодов квартир",
                        "Результат",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"Плагин завершил предустановку кодов квартир с ПРЕДУПРЕЖДЕНИЯМИ. Они появились отдельным окном",
                        "Результат",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
