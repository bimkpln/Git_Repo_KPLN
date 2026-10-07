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
        public partial class Engine
        {
            public string SettingsStatus {get;private set;}
            private void Progress(string message){progressMessage=message;reportProgress?.Invoke(message);}
            public void Initialize(Action<string> progress)
            {
                if(IsInitialized)return;
                reportProgress=progress;
                try
                {
                Progress("Загрузка модели: поиск источников...");
                Sources.Add(new Source {Key="host",Name=doc.Title,Document=doc,Transform=Transform.Identity});
                Discover(doc,"host",Transform.Identity,null,new HashSet<Document>{doc});
                ReadSourceElements();
                try
                {
                    var saved=Read<Settings>("settings");Config=saved??new Settings();
                    var changes=UpgradeSettings(Config);Normalize(Config);
                    SettingsStatus=saved==null?"Настройки RVT ещё не сохранены.":"Настройки прочитаны из RVT.";
                    if(changes.Count>0)
                    {
                        SettingsStatus+=" "+string.Join(" ",changes);
                        startupIssues.Add(new Issue{Code="SETTINGS_MIGRATION",Severity="Предупреждение",Message=string.Join(" ",changes),Action="Проверьте восстановленные настройки и нажмите «Сохранить настройки в RVT»."});
                    }
                }
                catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Config=new Settings();settingsLoadFailed=true;SettingsStatus="Сохранённые настройки не прочитаны: "+ex.Message+" Загружены начальные значения; исходная запись RVT сохранена.";startupIssues.Add(new Issue{Code="SETTINGS_RECOVERY",Severity="Предупреждение",Message=SettingsStatus,
                    Action="Проверьте начальные настройки и нажмите «Сохранить настройки в RVT», чтобы явно заменить неподдерживаемую или повреждённую запись."});}
                RestoreSourceOptions();RefreshCatalogs();PrepareRoomWorkflow();
                Progress("Загрузка сохранённого отчёта...");
                try{Last=Read<Run>("run");RestoreReviewInputs();}catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){startupIssues.Add(new Issue{Code="RUN_RECOVERY",Severity="Предупреждение",Message=ex.Message,Action="Выполните новый расчёт."});}
                IsInitialized=true;
                }
                finally {reportProgress=null;}
            }
            private void ReadSourceElements()
            {
                var documents=new Dictionary<Document,List<Element>>();
                foreach(var source in Sources.Where(x=>x.Loaded))
                {
                    Progress("Загрузка модели: "+source.Name+" - чтение элементов...");
                    try
                    {
                        List<Element> elements;
                        if(!documents.TryGetValue(source.Document,out elements))
                        {
                            var ownIds=new HashSet<ElementId>(new FilteredElementCollector(source.Document).WherePasses(new ExtensibleStorageFilter(StorageId)).ToElementIds());
                            elements=new List<Element>();int scanned=0;
                            foreach(var e in new FilteredElementCollector(source.Document).WhereElementIsNotElementType())
                            {
                                if(++scanned%200==0)Progress("Загрузка модели: "+source.Name+" - прочитано объектов: "+scanned);
                                if(e.Category!=null&&!(e is View)&&!(e is DataStorage)&&!ownIds.Contains(e.Id))elements.Add(e);
                            }
                            documents.Add(source.Document,elements);
                        }
                        source.Elements=elements;
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){source.LoadError="Не удалось прочитать элементы источника: "+ex.Message;}
                }
            }
            public static string LinkLoadWarning(IEnumerable<string> loadable,IEnumerable<string> nested)
            {
                var direct=loadable.Distinct().ToList();var indirect=nested.Distinct().ToList();
                string message=direct.Count==0?"Незагруженные связи недоступны для автоматической загрузки из текущего проекта.":
                    "Для расчёта нужно загрузить связи по сохранённым путям:\n\n"+string.Join("\n",direct.Select(n=>"• "+n));
                if(indirect.Count>0)message+="\n\nВложенные связи (их загрузкой управляет родительская модель):\n"+string.Join("\n",indirect.Select(n=>"• "+n));
                if(direct.Count>0)message+="\n\nЗагрузка может занять время и увеличить расход памяти. Revit очищает историю отмены (Ctrl+Z). Загруженные связи останутся загруженными в текущем сеансе. Плагин не сохраняет файлы и не меняет пути связей.\n\nЗагрузить связи и продолжить?";
                else message+="\n\nПродолжить проверку доступных данных? Недоступные источники будут указаны в отчёте.";
                return message;
            }
            // Must be called BEFORE Calculate opens its TransactionGroup: RevitLinkType.Load owns its transactions.
            public bool LoadMissingLinks(Action<string> progress,Func<string,bool> confirm)
            {
                var missing=Sources.Where(s=>s.Mode!="exclude"&&!s.Loaded).ToList();
                if(missing.Count==0)return false;
                var direct=missing.Where(s=>s.ParentDocument==doc&&s.LinkTypeId!=null)
                    .GroupBy(s=>IDHelper.ElIdValue(s.LinkTypeId)).Select(g=>g.First()).ToList();
                var nested=missing.Where(s=>s.ParentDocument!=doc||s.LinkTypeId==null).ToList();
                if(confirm==null||!confirm(LinkLoadWarning(direct.Select(s=>s.Name),nested.Select(s=>s.Name))))
                    throw new System.OperationCanceledException();
                foreach(var source in nested)source.LoadError="Связь «"+source.Name+"» не загружена. Вложенную связь нужно загрузить в её родительской модели; документ связи доступен только для чтения.";
                if(direct.Count==0)return false;
                if(doc.IsModifiable||doc.IsReadOnly)throw new InvalidOperationException("Загрузка связей недоступна: документ находится в транзакции или доступен только для чтения.");
                var failures=new Dictionary<long,string>();
                bool cancelled=false;var oldProgress=reportProgress;
                // Once Load has run, complete the rescan even if cancellation arrives: cached Revit objects can be invalidated.
                reportProgress=message=>{try{progress?.Invoke(message);}catch(System.OperationCanceledException){cancelled=true;}};
                SnapshotSources();
                startupIssues.RemoveAll(i=>i.Code=="LINK_AUTOLOAD");
                DisposeSpatialCalculators();
                try
                {
                    int index=0;
                    foreach(var source in direct)
                    {
                        Progress("Загрузка связи "+(++index)+" / "+direct.Count+": "+source.Name);
                        if(cancelled)break;
                        try
                        {
                            var type=doc.GetElement(source.LinkTypeId) as RevitLinkType;
                            if(type==null)throw new InvalidOperationException("Тип связи не найден в текущем проекте.");
                            if(RevitLinkType.IsLoaded(doc,type.Id))continue;
                            using(var result=type.Load())
                            {
                                if(result==null||!LinkLoadResult.IsCodeSuccess(result.LoadResult))
                                    failures[IDHelper.ElIdValue(source.LinkTypeId)]="Результат Revit: "+(result==null?"результат загрузки отсутствует":result.LoadResult.ToString());
                            }
                        }
                        catch(Autodesk.Revit.Exceptions.OperationCanceledException){cancelled=true;break;}
                        catch(System.OperationCanceledException){cancelled=true;break;}
                        catch(Exception ex){failures[IDHelper.ElIdValue(source.LinkTypeId)]=ex.Message;}
                    }
                    parameterCache.Clear();phaseCache.Clear();catalogsReady=false;shapes.Clear();ClearVolumeCaches();
                    Sources.Clear();Sources.Add(new Source{Key="host",Name=doc.Title,Document=doc,Transform=Transform.Identity});
                    Discover(doc,"host",Transform.Identity,null,new HashSet<Document>{doc});
                    RestoreSourceOptions();ReadSourceElements();
                    foreach(var source in Sources.Where(s=>!s.Loaded))
                    {
                        string reason;
                        if(source.ParentDocument==doc&&source.LinkTypeId!=null&&failures.TryGetValue(IDHelper.ElIdValue(source.LinkTypeId),out reason))
                            source.LoadError="Не удалось загрузить связь «"+source.Name+"»: "+reason;
                        else source.LoadError="Связь «"+source.Name+"» осталась незагруженной. "+(source.ParentDocument!=doc?"Проверьте загрузку вложенной связи в родительской модели.":"Проверьте доступность файла и состояние рабочего набора в управлении связями.");
                    }
                    RefreshCatalogs();SnapshotSources();
                    var nowLoaded=direct.Where(old=>Sources.Any(s=>s.Key==old.Key&&s.Loaded)).Select(s=>s.Name).ToList();
                    startupIssues.Add(new Issue{Code="LINK_AUTOLOAD",Severity="Информация",Message="После подготовки доступны связи: "+(nowLoaded.Count==0?"ни одна из запрошенных":string.Join(", ",nowLoaded))+". Источники и параметры перечитаны.",Action="Загрузка связей не сохраняет RVT на диск."});
                }
                finally{reportProgress=oldProgress;}
                if(cancelled)throw new System.OperationCanceledException();
                return true;
            }
            private void Discover(Document parent,string path,Transform transform,ElementId root,HashSet<Document> chain)
            {
                foreach(RevitLinkInstance link in new FilteredElementCollector(parent).OfClass(typeof(RevitLinkInstance)))
                {
                    Progress("Загрузка модели: поиск связи "+link.Name);
                    try
                    {
                    // Nested overlay links are not displayed through their parent attachment.
                    var type=parent.GetElement(link.GetTypeId()) as RevitLinkType;
                    if(root!=null&&type!=null&&type.AttachmentType==AttachmentType.Overlay) continue;
                    var linked=link.GetLinkDocument();var key=path+"/"+link.UniqueId;
                    var source=new Source {Key=key,Name=path=="host"?link.Name:path+" / "+link.Name,
                        Document=linked,Transform=transform.Multiply(link.GetTotalTransform()),RootLink=root??link.Id,ParentDocument=parent,LinkTypeId=link.GetTypeId()};
                    Sources.Add(source);
                    if(linked!=null&&!chain.Contains(linked)) {var next=new HashSet<Document>(chain){linked};Discover(linked,key,source.Transform,source.RootLink,next);}
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Sources.Add(new Source{Key=path+"/"+link.UniqueId,Name=link.Name,RootLink=root??link.Id,LoadError=ex.Message,Transform=transform,ParentDocument=parent,LinkTypeId=link.GetTypeId()});}
                }
            }
            private static void Normalize(Settings s)
            {
                if(s.Version!=Settings.CurrentVersion)UpgradeSettings(s);
                s.BuildingAssignments=s.BuildingAssignments??new List<BuildingAssignment>();s.IssueDetail="full";
                s.Departments=s.Departments??new List<DepartmentAssignment>();
                if(s.Departments.Any(d=>d==null||!Roles().Any(r=>r.Key==d.Role)) || s.Departments.GroupBy(d=>(d.Value??"" ).Trim(),StringComparer.OrdinalIgnoreCase).Any(g=>g.Count()>1))
                    throw new InvalidOperationException("В классификации назначений есть неизвестные категории или повторяющиеся значения.");
                s.LoggiaCoefficient=s.LoggiaCoefficient??"0.5";s.BalconyCoefficient=s.BalconyCoefficient??"0.3";
                s.Metrics=s.Metrics??Catalog();var enabled=s.Metrics.ToDictionary(x=>x.Key,x=>x);
                foreach(var official in Catalog())
                {
                    Metric m;if(enabled.TryGetValue(official.Key,out m)){m.Name=official.Name;m.Unit=official.Unit;m.Description=official.Description;}
                    else s.Metrics.Add(official);
                }
                s.Metrics.RemoveAll(m=>!Enum.IsDefined(typeof(Indicator),m.Key));
                s.Contours=s.Contours??new ObservableCollection<ContourSketch>();
                foreach(var m in s.Metrics)
                {
                    m.ContourMode=m.ContourMode??"current";m.WallSelectionMode=m.WallSelectionMode??"exterior";
                    m.WallBoundary=m.WallBoundary??"auto";m.WallCutHeight=m.WallCutHeight??"0";
                    m.ContourExclusions=m.ContourExclusions??"normative";m.VolumeMaskHeight=m.VolumeMaskHeight??"actual";
                    m.SelectedWalls=m.SelectedWalls??new List<string>();
                    foreach(var pair in new[]{Tuple.Create("contour",m.ContourMode),Tuple.Create("wall-selection",m.WallSelectionMode),Tuple.Create("wall-boundary",m.WallBoundary),Tuple.Create("contour-exclusions",m.ContourExclusions),Tuple.Create("mask-height",m.VolumeMaskHeight)})
                        if(!Choices(pair.Item1).Any(x=>x.Key==pair.Item2))throw new InvalidOperationException("Неизвестная настройка контура: "+pair.Item1);
                    if(m.ContourMode!="current"&&!SupportsContours(m.Key))throw new InvalidOperationException("Для этого показателя нужен расчёт по классифицированным помещениям: "+m.Name);
                }
                foreach(var c in s.Contours)
                    if(!s.Metrics.Any(m=>m.Key==c.Metric)||!Choices("sketch-kind").Any(x=>x.Key==c.Kind))throw new InvalidOperationException("Некорректная привязка ручного контура.");
                s.RoomClassificationParameter="@Department";
                s.Parameters=s.Parameters??Settings.DefaultParameters();
                s.StairFamilies=s.StairFamilies??new List<StairFamilySetting>();
                s.CategorySources=s.CategorySources??new List<CategorySource>();s.CategoryFamilies=s.CategoryFamilies??new List<CategoryFamily>();
                if(s.CategorySources.Any(c=>c==null||!Roles().Any(r=>r.Key==c.Role)||c.Role=="unknown"&&c.Families)||s.CategorySources.GroupBy(c=>c.Role).Any(g=>g.Count()>1))
                    throw new InvalidOperationException("Некорректный источник категории в настройках.");
                if(s.CategoryFamilies.Any(c=>c==null||!Roles().Any(r=>r.Key==c.Role)||c.Role=="unknown"||string.IsNullOrWhiteSpace(c.Family)||string.IsNullOrWhiteSpace(c.Type)))
                    throw new InvalidOperationException("Некорректное назначение типа семейства в настройках.");
                if(s.CategoryFamilies.GroupBy(c=>c.Caption,StringComparer.OrdinalIgnoreCase).Any(g=>g.Select(c=>c.Role).Distinct().Count()>1))
                    throw new InvalidOperationException("Тип семейства назначен разным категориям.");
                foreach(var p in Settings.DefaultParameters()) {var existing=s.Parameters.FirstOrDefault(x=>x.Key==p.Key);if(existing==null)s.Parameters.Add(p);else {existing.Title=p.Title;existing.Description=p.Description;if(Settings.IsFixedParameter(p.Key))existing.Name=Settings.FixedParameterName(p.Key);}}
                s.Rules=s.Rules??new ObservableCollection<Rule>();s.Buildings=s.Buildings??new ObservableCollection<BuildingMap>();
                s.Levels=s.Levels??new ObservableCollection<LevelSetting>();s.Corrections=s.Corrections??new ObservableCollection<Correction>();
                // Level settings are shared by all buildings and sections. Discard legacy overrides
                // so settings loaded from an earlier version cannot silently alter the calculation.
                foreach(var refinement in s.Levels.Where(l=>!string.IsNullOrWhiteSpace(l.Building)||!string.IsNullOrWhiteSpace(l.Section)).ToList())
                    s.Levels.Remove(refinement);
                foreach(var level in s.Levels)
                {
                    if(string.IsNullOrEmpty(level.TopSlabMode))level.TopSlabMode=string.IsNullOrWhiteSpace(level.TopSlab)?"auto":"manual";
                    if(string.IsNullOrEmpty(level.HeightMode))level.HeightMode=string.IsNullOrWhiteSpace(level.Height)?"auto":"manual";
                    if(string.IsNullOrEmpty(level.RoofAreaMode))level.RoofAreaMode=string.IsNullOrWhiteSpace(level.RoofArea)?"auto":"manual";
                    if(string.IsNullOrEmpty(level.RoofRatioMode))level.RoofRatioMode=string.IsNullOrWhiteSpace(level.RoofRatio)?"auto":"manual";
                    level.RoofSources=level.RoofSources??new List<string>();level.NormativeProfile=s.Profile;
                }
                s.Sources=s.Sources??new List<SourceOption>();s.Decimals=Math.Max(0,Math.Min(6,s.Decimals));
                var defaults=new Settings();
                foreach(var name in new[]{"IncludedColor","ExcludedColor","ManualColor","PublicColor","SummerColor","TechnicalColor","ErrorColor"})
                {var property=typeof(Settings).GetProperty(name);if(string.IsNullOrWhiteSpace((string)property.GetValue(s)))property.SetValue(s,property.GetValue(defaults));}
                foreach(var name in new[]{"method","profile","group","geometry","graphics","count"})
                {
                    var value=name=="method"?s.Method:name=="profile"?s.Profile:name=="group"?s.Grouping:name=="geometry"?s.Geometry:name=="graphics"?s.Graphics:s.ApartmentMode;
                    if(!Choices(name).Any(x=>x.Key==value))throw new InvalidOperationException("Неизвестный вариант настройки: "+name+" = "+value);
                }
                // Earlier settings used a combo-box entry to disable graphics.
                if(s.Graphics=="none"){s.CreateViews=false;s.Graphics="replace";}
                foreach(var rule in s.Rules.Where(x=>x.Enabled))
                {
                    if(rule.Metric!="all"&&!s.Metrics.Any(m=>m.Key==rule.Metric))throw new InvalidOperationException("Неизвестный показатель правила: "+rule.Metric);
                    if(!Choices("operation").Any(x=>x.Key==rule.Operation)||!Choices("action").Any(x=>x.Key==rule.Action)||!Roles().Any(x=>x.Key==rule.Role)||!Choices("part").Any(x=>x.Key==rule.Part))throw new InvalidOperationException("Неизвестное условие / действие / назначение в правиле.");
                    if(rule.Coefficient.HasValue&&(double.IsNaN(rule.Coefficient.Value)||double.IsInfinity(rule.Coefficient.Value)||rule.Coefficient<0||rule.Coefficient>1))throw new InvalidOperationException("Коэффициент правила должен быть конечным числом от 0 до 1.");
                }
                foreach(var source in s.Sources)if(!Choices("source").Any(x=>x.Key==source.Mode))throw new InvalidOperationException("Неизвестный режим участия источника: "+source.Mode);
            }
            private void RestoreSourceOptions()
            {
                foreach(var s in Sources) {var saved=Config.Sources.FirstOrDefault(x=>x.Key==s.Key);if(saved!=null){s.Mode=saved.Mode;s.Profile=saved.Profile;}}
            }
            private void SnapshotSources(){Config.Sources=Sources.Select(s=>new SourceOption{Key=s.Key,Mode=s.Mode,Profile=s.Profile}).ToList();}
            public void ImportSettings(string path)
            {var value=Deserialize<Settings>(File.ReadAllText(path,Encoding.UTF8));UpgradeSettings(value);Normalize(value);Config=value;settingsLoadFailed=false;RestoreSourceOptions();RefreshCatalogs();PrepareRoomWorkflow();}
            public void ExportSettings(string path){PrepareRoomWorkflow();SnapshotSources();File.WriteAllText(path,Serialize(Config),new UTF8Encoding(true));}
            public void SaveSettings(){Normalize(Config);PrepareRoomWorkflow();SnapshotSources();Normalize(Config);Write("settings",Config);settingsLoadFailed=false;startupIssues.RemoveAll(i=>i.Code=="SETTINGS_RECOVERY"||i.Code=="SETTINGS_MIGRATION"||i.Code=="SETTINGS_REFERENCES");SettingsStatus="Настройки записаны в текущую модель RVT.";}
            private void RefreshCatalogs()
            {
                var actualLevels=new HashSet<string>(StringComparer.Ordinal);
                var scannedDocuments=new HashSet<Document>();
                var names=new HashSet<string>(Parameters,StringComparer.OrdinalIgnoreCase){"@Name","@Category","@Type","@Family","@Workset"};
                foreach(var s in Sources.Where(x=>x.Loaded&&x.LoadError==null))
                {
                    if(!catalogsReady&&scannedDocuments.Add(s.Document))
                    {
                    int scanned=0;
                    foreach(var e in s.Elements.OfType<Room>().Where(IsPlacedRoom))
                    {if(++scanned%100==0)Progress("Загрузка параметров: "+s.Name+" - "+scanned+" / "+s.Elements.Count);
                    try{foreach(Parameter p in e.Parameters) if(p.Definition!=null)names.Add(p.Definition.Name);}
                        catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){startupIssues.Add(new Issue{Code="PARAM_SCAN",Severity="Ошибка",Source=s.Name,Element=IDHelper.ElIdValue(e.Id).ToString(),Message=ex.Message,Action="Проверьте повреждённый объект модели."});}}
                    foreach(var e in s.Elements.OfType<Room>().Where(IsPlacedRoom).Select(room=>room.GetTypeId()).Where(id=>id!=ElementId.InvalidElementId).Distinct().Select(id=>s.Document.GetElement(id)).Where(type=>type!=null))
                    {if(++scanned%100==0)Progress("Загрузка параметров типов: "+s.Name+" - "+scanned);
                        foreach(Parameter p in e.Parameters)if(p.Definition!=null)names.Add(p.Definition.Name);}
                    }
                    var occupied=new HashSet<ElementId>(s.Elements.OfType<Room>().Where(IsPlacedRoom).Select(e=>e.LevelId));
                    foreach(Level level in new FilteredElementCollector(s.Document).OfClass(typeof(Level)))
                    {
                        string key=s.Key+"/"+level.UniqueId;
                        if(Owned(level)&&Kind(level)=="review-level")
                        {foreach(var obsolete in Config.Levels.Where(l=>l.Key==key).ToList())Config.Levels.Remove(obsolete);continue;}
                        actualLevels.Add(key);
                        var existing=Config.Levels.FirstOrDefault(x=>x.Key==key&&string.IsNullOrWhiteSpace(x.Building)&&string.IsNullOrWhiteSpace(x.Section));
                        if(existing==null){existing=new LevelSetting{Key=key};Config.Levels.Add(existing);
                            var story=level.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY);
                            existing.Include=(story!=null&&story.AsInteger()==1)||occupied.Contains(level.Id);}
                        foreach(var setting in Config.Levels.Where(x=>x.Key==key))
                        {
                            var point=s.Transform.OfPoint(new XYZ(0,0,level.ProjectElevation));
                            setting.Source=s.Name;setting.Name=level.Name;setting.Elevation=point.Z;
                            setting.AbsoluteElevationMeters=doc.ActiveProjectLocation.GetProjectPosition(point).Elevation*.3048;
                        }
                    }
                }
                int removed=RemoveMissingSettingsReferences(Config,new HashSet<string>(Sources.Select(s=>s.Key)),new HashSet<string>(Sources.Where(s=>s.Loaded&&s.LoadError==null).Select(s=>s.Key)),actualLevels);
                if(removed>0)
                {
                    string message="Очищены ссылки настроек на удалённые уровни и источники: "+removed+". Настройки незагруженных связей сохранены.";
                    SettingsStatus=(SettingsStatus??"")+" "+message;
                    startupIssues.Add(new Issue{Code="SETTINGS_REFERENCES",Severity="Предупреждение",Message=message,Action="Проверьте первый этаж и уровень земли; сохраните настройки в RVT."});
                }
                catalogsReady=true;
                Parameters=names.OrderBy(x=>x).ToList();Categories=Sources.Where(s=>s.Loaded).GroupBy(s=>s.Document).Select(g=>g.First()).SelectMany(s=>s.Elements).Select(e=>e.Category.Name).Distinct().OrderBy(x=>x).ToList();
                Phases=Sources.Where(s=>s.Loaded).GroupBy(s=>s.Document).Select(g=>g.First()).SelectMany(s=>s.Document.Phases.Cast<Phase>()).Select(p=>p.Name).Distinct().OrderBy(x=>x).ToList();
                AreaSchemes=Sources.Where(s=>s.Loaded).GroupBy(s=>s.Document).Select(g=>g.First()).SelectMany(s=>new FilteredElementCollector(s.Document).OfClass(typeof(AreaScheme)).Cast<AreaScheme>()).Select(a=>a.Name).Distinct().OrderBy(x=>x).ToList();
                Templates=new List<Choice>{new Choice("","Без шаблона","Сохраняет цвета и оформление, созданные расчётом.")};
                Templates.AddRange(new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Where(v=>v.IsTemplate).Select(v=>new Choice(v.UniqueId,v.Name,"Применяет шаблон к служебному виду. Управляемые шаблоном категории и фильтры могут изменить видимость расчётной графики.")));
            }
        }
    }
}
