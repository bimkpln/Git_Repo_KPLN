using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public class ColorPresset
    {
        public int RoomId { get; }
        public static KPLNRandom RandomR { get; set; }
        public static KPLNRandom RandomG { get; set; }
        public static KPLNRandom RandomB { get; set; }
        public OverrideGraphicSettings Settings = new OverrideGraphicSettings();
        public ColorPresset(int id, FillPatternElement fill)
        {
#if Debug2020 || Debug2023 || Revit2020 || Revit2023
            RoomId = id;
            Settings.SetSurfaceForegroundPatternColor(new Autodesk.Revit.DB.Color((byte)RandomR.GetRandom(), (byte)RandomG.GetRandom(), (byte)RandomB.GetRandom()));
            Settings.SetProjectionLineWeight(1);
            Settings.SetHalftone(false);
            Settings.SetSurfaceForegroundPatternId(fill.Id);
            Settings.SetSurfaceTransparency(50);
#endif
#if Debug2018 || Revit2018
            RoomId = id;
            Settings.SetProjectionFillColor(new Autodesk.Revit.DB.Color((byte)RandomR.GetRandom(), (byte)RandomG.GetRandom(), (byte)RandomB.GetRandom()));
            Settings.SetProjectionLineWeight(1);
            Settings.SetHalftone(false);
            Settings.SetProjectionFillPatternId(fill.Id);
            Settings.SetSurfaceTransparency(50);
#endif
        }
        public static ColorPresset GetRoomColor(List<ColorPresset> list, int id, FillPatternElement fill)
        {
            foreach (ColorPresset preset in list)
            {
                if (preset.RoomId == id)
                {
                    return preset;
                }
            }
            ColorPresset newPreset = new ColorPresset(id, fill);
            list.Add(newPreset);
            return newPreset;
        }
    }
}
