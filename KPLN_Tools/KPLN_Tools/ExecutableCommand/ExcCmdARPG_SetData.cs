﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools.Common;
using KPLN_Tools.Forms.Models.Core;
using System.Windows.Forms;

namespace KPLN_Tools.ExecutableCommand
{
    internal class ExcCmdARPG_SetData : IExecutableCommand
    {
        private readonly Document _doc;
        private readonly ARPG_Room[] _aRPGRooms;
        private readonly ARPG_Flat[] _aRPGFlats;
        private readonly ARPG_TZ_FlatData[] _arpgTZFlatDatas;

        public ExcCmdARPG_SetData(Document doc, ARPG_Room[] aRPGRooms, ARPG_Flat[] aRPGFlats, ARPG_TZ_FlatData[] arpgTZFlatDatas)
        {
            _doc = doc;
            _aRPGRooms = aRPGRooms;
            _aRPGFlats = aRPGFlats;
            _arpgTZFlatDatas = arpgTZFlatDatas;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction trans = new Transaction(_doc, "KPLN: Пятнография"))
            {
                trans.Start();

                ARPG_Flat.SetMainFlatData(_aRPGFlats, _arpgTZFlatDatas);
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

                ARPG_Room.SetCountedRoomData(_aRPGRooms);

                MessageBox.Show(
                        $"Плагин успешно завершил работу",
                        "Результат",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
