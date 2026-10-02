using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CalculateTEP.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_CalculateTEP.Common.Geometry.TepClipper.Clipper;
using CP = KPLN_CalculateTEP.Common.Geometry.TepClipper.IntPoint;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;
using PT = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyType;
using PF = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyFillType;

using TepClipper = KPLN_CalculateTEP.Common.Geometry.TepClipper;
using KPLN_CalculateTEP.Common.Methodologies;
namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class Settings : System.ComponentModel.INotifyPropertyChanged
        {
            private static readonly Dictionary<string,string> fixedParameters = new Dictionary<string,string> {
                {"building","ПОМ_Корпус"}, {"apartment","КВ_Номер"}, {"section","ПОМ_Секция"},
                {"parking","@ParkingMark"}, {"function","@Department"}, {"part",""}, {"profile",""}, {"coefficient",""}
            };
            public static bool IsFixedParameter(string key) { return fixedParameters.ContainsKey(key); }
            public static string FixedParameterName(string key) { return fixedParameters[key]; }
            public ObservableCollection<ContourSketch> Contours {get;set;} = new ObservableCollection<ContourSketch>();
            public int Version {get;set;} = 2;
            public int RoomWorkflowVersion {get;set;}
            public List<DepartmentAssignment> Departments {get;set;} = new List<DepartmentAssignment>();
            public string PreviousWorkflowSettings {get;set;}
            public string LoggiaCoefficient {get;set;} = "0.5";
            public string BalconyCoefficient {get;set;} = "0.3";
            public string Method {get;set;} = "new";
            public string Profile {get;set;} = "residential";
            public string Grouping {get;set;} = "links";
            public string ManualBuilding {get;set;} = "Корпус 1";
            public string Geometry {get;set;} = "auto";
            public string Phase {get;set;} public string AreaScheme {get;set;}
            public string ZeroMode {get;set;} = "level"; public string ZeroValue {get;set;} = "0";
            public string ZeroLevel {get;set;} public string ZeroParameter {get;set;}
            public string GroundMode {get;set;} = "manual"; public string GroundValue {get;set;} = "0";
            public string GroundLevel {get;set;} public string GroundParameter {get;set;}
            public string ApartmentMode {get;set;} = "id"; public string ParkingMode {get;set;} = "id";
            private bool createViews;
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            public bool CreateViews {get{return createViews;}set{if(createViews==value)return;createViews=value;PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs(nameof(CreateViews)));}}
            public string Graphics {get;set;} = "replace"; public bool Create3D {get;set;} = true;
            public bool CreateSchedule {get;set;} = false;
            public string Prefix {get;set;} = "ТЭП_";
            public int Decimals {get;set;} = 2;
            public string IssueDetail {get;set;} = "full";
            public List<Metric> Metrics {get;set;} = Catalog();
            public List<ParameterMap> Parameters {get;set;} = DefaultParameters();
            public ObservableCollection<Rule> Rules {get;set;} = new ObservableCollection<Rule>();
            public ObservableCollection<BuildingMap> Buildings {get;set;} = new ObservableCollection<BuildingMap>();
            public ObservableCollection<LevelSetting> Levels {get;set;} = new ObservableCollection<LevelSetting>();
            public ObservableCollection<Correction> Corrections {get;set;} = new ObservableCollection<Correction>();
            public List<SourceOption> Sources {get;set;} = new List<SourceOption>();
            public string IncludedColor {get;set;} = "#347EC4";
            public string ExcludedColor {get;set;} = "#CC5555";
            public string ManualColor {get;set;} = "#A56ED3";
            public string PublicColor {get;set;} = "#C59043";
            public string SummerColor {get;set;} = "#36A38E";
            public string TechnicalColor {get;set;} = "#8993A0";
            public string ErrorColor {get;set;} = "#E24040";
            public string PlanTemplate {get;set;} public string View3DTemplate {get;set;}
            public static List<ParameterMap> DefaultParameters()
            {
                string[] keys={"building","profile","function","part","apartment","section","parking","embedded","standalone","vertical","coefficient","height","slope","width","roof-ratio","partial-floor","transition","include","mezzanine-ratio"};
                string[] labels={"Корпус","Тип здания / профиль","Назначение помещения","Жилая / нежилая часть","Номер / ID квартиры","Секция","ID машино-места","Признак встроенно-пристроенной части","Признак отдельно стоящего объекта","ID вертикального пространства","Коэффициент летнего помещения","Высота в свету","Угол наклона потолка, градусы","Ширина проёма / просвета","Доля площади надстройки от кровли","Техническое пространство занимает часть этажа","Корпуса перехода через ;","Признак включения в ТЭП","Доля площади антресоли от этажа"};
                var result=keys.Select((k,i)=>new ParameterMap {Key=k,Title=labels[i],Name="",Description="Параметр экземпляра или типа: "+labels[i]+". Длины Revit переводятся в метры; текстовые длины задаются в метрах, доли в диапазоне 0..1. Для признаков: 1/0, да/нет."}).ToList();
                result.Add(new ParameterMap{Key="service-access",Title="Нужен проход для обслуживания коммуникаций",Name="",Description="Для общественного технического пространства ниже 1,8 м: да/1 - включать, нет/0 - исключать. Отсутствие признака вызывает ошибку."});
                result.Add(new ParameterMap{Key="stair-width",Title="Ширина лестничного марша",Name="",Description="Длина Revit или текст в метрах. Для лестничного просвета проверяются ширина марша и порог 1,5 м."});
                result.Add(new ParameterMap{Key="level",Title="Расчётный уровень объекта",Name="",Description="Параметр-ссылка на уровень либо его точное имя. При заполнении уточняет встроенный LevelId; полезен для расчётных семейств и DirectShape."});
                result.Add(new ParameterMap{Key="floor-offset",Title="Смещение чистого пола от уровня помещения",Name="",Description="Необязательный параметр Room: длина Revit или текст в метрах относительно уровня помещения в исходной модели. При заполненном имени требуется значение у каждого помещения. Без назначения берётся низ Room (уровень + смещение снизу), с проверкой совпадения с верхней поверхностью пола. Не задаёт произвольную высоту сечения.\nПример: ТЭП_СмещениеЧистогоПола; значение 0,150 м."});
                result.First(x=>x.Key=="apartment").Description="Сквозной номер / ID квартиры в пределах корпуса и секции. Один ID объединяет помещения на разных этажах. Пусто, 0 или - означает помещение вне квартиры.\nПример: КВ_Номер; значение 125.";
                result.First(x=>x.Key=="function").Description="Значение назначения или имя помещения сопоставляется с классификацией. Проектные названия задайте через правила.\nПример: @Name; значение Лоджия - правило назначения loggia.";
                result.First(x=>x.Key=="building").Description="Разделяет помещения по корпусам. Пустое имя - отдельный корпус для каждого источника модели.\nПример: ПОМ_Корпус; значение 2.";
                result.First(x=>x.Key=="section").Description="Разделяет одинаковые номера квартир в разных секциях одного корпуса. Если параметр назначен, секция квартиры должна быть заполнена.\nПример: ПОМ_Секция; значение 1.";
                result.First(x=>x.Key=="part").Description="Определяет жилую / нежилую часть для функциональных показателей. Для обычных квартир определяется автоматически; для неоднозначных назначений нужна классификация.\nПример: ТЭП_Часть; значение residential или nonresidential.";
                result.First(x=>x.Key=="profile").Description="Уточняет тип здания для отдельных помещений и меняет применяемые нормативные правила. Пустое имя - тип из первого шага.\nПример: ТЭП_ТипЗдания; значение public для общественной части.";
                result.First(x=>x.Key=="parking").Description="Общий ID машино-места позволяет учесть его один раз. Нужен при выборе показателя количества машино-мест.\nПример: ТЭП_МашиноМесто; значение ММ-015.";
                result.First(x=>x.Key=="embedded").Description="Показывает, относится ли нежилое помещение к встроенно-пристроенной части. Нужен для показателей ННП.\nПример: ТЭП_ВстроеннаяЧасть; значение да или 1.";
                result.First(x=>x.Key=="standalone").Description="Разделяет отдельно стоящие и остальные объекты при расчёте ННП. Заполняется вместе с признаком встроенно-пристроенной части.\nПример: ТЭП_ОтдельноСтоящий; значение нет или 0.";
                result.First(x=>x.Key=="vertical").Description="Связывает по этажам одну шахту, проём или многосветное пространство. Позволяет применить правило учёта нижнего этажа и исключения верхних.\nПример: ТЭП_ВертикальноеПространство; значение ШАХТА-03 одинаково на всех её этажах.";
                result.First(x=>x.Key=="height").Description="Высота в свету нужна для проверки ниш и низких технических пространств. Принимается параметр длины Revit либо текст в метрах.\nПример: ТЭП_ВысотаВСвету; текстовое значение 1,75.";
                result.First(x=>x.Key=="slope").Description="Угол наклона потолка включает высотные проверки мансардных помещений. Число в градусах от 0 до 90 либо параметр угла Revit; для обычного помещения оставьте значение пустым.\nПример: ТЭП_УголПотолка; значение 45.";
                result.First(x=>x.Key=="width").Description="Ширина нужна для проверки арочного проёма или лестничного просвета. Параметр длины либо текст в метрах.\nПример: ТЭП_ШиринаПросвета; текстовое значение 1,6.";
                result.First(x=>x.Key=="roof-ratio").Description="Доля площади надстройки от кровли участвует в правилах общественных зданий. Число от 0 до 1.\nПример: ТЭП_ДоляНадстройки; значение 0,2 означает 20%.";
                result.First(x=>x.Key=="partial-floor").Description="Указывает, занимает ли техническое пространство только часть этажа; используется для высотных зданий.\nПример: ТЭП_ЧастьЭтажа; значение да или 1.";
                result.First(x=>x.Key=="transition").Description="Перечисляет корпуса, соединённые переходом. Его учитываемая площадь распределяется между указанными корпусами.\nПример: ТЭП_КорпусаПерехода; значение Корпус 1;Корпус 2.";
                result.First(x=>x.Key=="include").Description="Позволяет явно исключить помещение из расчёта. Значение нет/0 исключает его; да/1 оставляет нормативные проверки включения.\nПример: ТЭП_Учитывать; значение нет.";
                result.First(x=>x.Key=="mezzanine-ratio").Description="Доля площади антресоли от этажа участвует в проверке общей площади общественного здания. Число от 0 до 1.\nПример: ТЭП_ДоляАнтресоли; значение 0,45 означает 45%.";
                result.First(x=>x.Key=="service-access").Description="Для общественного технического пространства ниже 1,8 м определяет необходимость прохода обслуживания коммуникаций: да/1 - учитывать, нет/0 - исключать.\nПример: ТЭП_ПроходОбслуживания; значение да.";
                result.First(x=>x.Key=="stair-width").Description="Ширина марша сравнивается с шириной лестничного просвета при его исключении. Параметр длины либо текст в метрах.\nПример: ТЭП_ШиринаМарша; текстовое значение 1,2.";
                result.First(x=>x.Key=="level").Description="Необязательное уточнение расчётного уровня. Обычно используется уровень самого помещения; заполняйте при необходимости другого расчётного отнесения.\nПример: ТЭП_РасчётныйУровень; точное имя уровня «02 Этаж» либо ссылка на него.";
                result.First(x=>x.Key=="function").Description="Выберите параметр, в котором записано назначение помещения, либо @Name для имени. Непонятные плагину обозначения сопоставьте во вкладке «Классификация».\nПример: ТЭП_Назначение; значение ЛДЖ - правило с назначением «Лоджия».";
                foreach(var map in result.Where(m=>IsFixedParameter(m.Key)))map.Name=FixedParameterName(map.Key);
                return result;
            }
            public string Parameter(string key) { var p=Parameters.FirstOrDefault(x=>x.Key==key); return p==null?"":p.Name; }
        }
    }
}
