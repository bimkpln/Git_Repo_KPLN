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

                string checkRoomCodesByTZ = ARPG_Room.CheckRoomCodesByTZMain(_pgVM, _aRPGRooms);
                bool checkWarningDict_Room = ARPG_Room.WarnDict_Room.Keys.Count != 0;
                bool checkWarningDict_Flat = ARPG_Flat.WarnDict_Flat.Keys.Count != 0;

                if (string.IsNullOrEmpty(checkRoomCodesByTZ) && !checkWarningDict_Room && !checkWarningDict_Flat)
                {
                    MessageBox.Show(
                            $"Плагин завершил рассчёт БЕЗ ошибок",
                            "Результат",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"Плагин завершил рассчёт с ПРЕДУПРЕЖДЕНИЯМИ. Они будут выведены отдельным окном",
                        "Результат",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    if (!string.IsNullOrEmpty(checkRoomCodesByTZ))
                        HtmlOutput.Print(checkRoomCodesByTZ, MessageType.Warning);

                    HtmlOutput.PrintMsgDict("ПРЕДУПРЕЖДЕНИЕ", MessageType.Warning, ARPG_Flat.WarnDict_Flat);
                    HtmlOutput.PrintMsgDict("ПРЕДУПРЕЖДЕНИЕ", MessageType.Warning, ARPG_Room.WarnDict_Room);
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
