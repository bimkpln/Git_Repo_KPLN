using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP : Window
    {
        public TEP.Engine Engine {get;private set;}
        public TEP.Settings Config {get{return Engine.Config;}}
        public List<TEP.Choice> Methods {get{return TEP.Choices("method");}}
        public List<TEP.Choice> Profiles {get{return TEP.Choices("profile").Where(x=>!x.Key.StartsWith("by-")).ToList();}}
        public List<TEP.Choice> GroupModes {get{return TEP.Choices("group");}}
        public List<TEP.Choice> DatumModes {get{return TEP.Choices("datum");}}
        public List<TEP.Choice> ZeroModes {get{return DatumModes.Where(x=>x.Key!="terrain").ToList();}}
        public List<TEP.Choice> GeometryModes {get{return TEP.Choices("geometry");}}
        public List<TEP.Choice> CountModes {get{return TEP.Choices("count");}}
        public List<TEP.Choice> GraphicsModes {get{return TEP.Choices("graphics").Where(x=>x.Key!="none").ToList();}}
        public List<TEP.Choice> DiagnosticModes {get{return TEP.Choices("diagnostics");}}
        public List<TEP.Choice> MetricChoices {get{return new List<TEP.Choice>{new TEP.Choice("all","Все показатели","Общая классификация применяется до правил конкретного показателя.")}.Concat(TEP.Catalog().Select(m=>new TEP.Choice(m.Key,m.Name,m.Description))).ToList();}}
        public List<TEP.Metric> ContourMetrics {get{return Config.Metrics.Where(m=>TEP.SupportsContours(m.Key)).ToList();}}
        public List<TEP.Choice> ContourModes {get{return TEP.Choices("contour");}}
        public List<TEP.Choice> WallSelectionModes {get{return TEP.Choices("wall-selection");}}
        public List<TEP.Choice> WallBoundaries {get{return TEP.Choices("wall-boundary");}}
        public List<TEP.Choice> ContourExclusionModes {get{return TEP.Choices("contour-exclusions");}}
        public List<TEP.Choice> MaskHeightModes {get{return TEP.Choices("mask-height");}}
        private static readonly Brush ErrorTextBrush=new SolidColorBrush(Color.FromRgb(180,35,24));
        private bool changing,busy,cancelRequested;
        private string checkedConfiguration;
        public IEnumerable<TEP.ParameterMap> RoomParameters {get{return Config.Parameters.Where(p=>p.Key!="vertical"&&!TEP.Settings.IsFixedParameter(p.Key)).Concat(new[]{new TEP.ParameterMap{Key="$stairs",Title="Ширина лестничного марша",Description="Стандартные лестницы Revit читаются автоматически.\nДля загружаемых семейств выберите типы и способ определения ширины."}});}}
        public string PhaseDescription {get{return "Стадия определяет, какие помещения и конструкции учитывать. «Последняя стадия» - последняя стадия каждого источника. «Учитывать все» - без фильтра по стадии.";}}
        public List<TEP.Choice> PhaseChoices {get{return new[]{new TEP.Choice("","Последняя стадия","Последняя стадия каждого источника."),new TEP.Choice(TEP.Settings.AllPhases,"Учитывать все","Все стадии выбранных источников, без фильтра.")}.Concat((Engine.Phases??new List<string>()).Select(p=>new TEP.Choice(p,p,"Помещения и конструкции выбранной стадии."))).ToList();}}
        public string PhaseSelection {get{return Config.Phase??"";}set{Config.Phase=value;}}
        private void RefreshRequiredParameters()
        {
            var roles=(departmentGroups??new List<TEP.DepartmentGroup>()).Select(g=>Engine.DepartmentRole(g.Value))
                .Concat((Config.CategoryFamilies??new List<TEP.CategoryFamily>()).Where(f=>Engine.CategoryUsesFamilies(f.Role)).Select(f=>f.Role));
            var required=TEP.Engine.RequiredParameterKeys(Config,roles);
            foreach(var parameter in Config.Parameters)parameter.IsRequired=required.Contains(parameter.Key);
        }
        public IEnumerable<TEP.LevelSetting> ZeroLevels {get{return Config.Levels.GroupBy(l=>l.Key).Select(g=>g.First());}}
        private List<TEP.DepartmentGroup> departmentGroups;
        private string scannedPhase;
        private string scannedRoomParameter;
        private string selectedDepartmentRole="unknown";
        private bool unassignedExpanded=true;
        public List<TEP.Choice> DepartmentRoles {get{return TEP.Roles().Select(r=>r.Key=="unknown"?new TEP.Choice("unknown","Нераспределённые","Назначьте категорию вручную. Нераспределённые помещения блокируют зависимые показатели; остальные рассчитываются."):r).ToList();}}
        public List<TEP.Choice> CategoryTargetChoices {get{return DepartmentRoles.Select(r=>new TEP.Choice(r.Key,r.Label+(r.Key=="unknown"?"":" | "+(Engine.CategoryUsesFamilies(r.Key)?"Семейства":"Помещение")),r.Description)).ToList();}}
        private void ScanRoomDepartments()
        {
            // Publish only a complete scan; a cancelled or failed scan must not look up to date.
            List<TEP.DepartmentGroup> scanned;
            try{scanned=Engine.ScanDepartments(ReportProgress);}
            catch{departmentGroups=null;scannedPhase=null;CoefficientInputs.IsEnabled=false;RefreshDepartmentCards();throw;}
            departmentGroups=scanned;scannedPhase=Config.Phase;scannedRoomParameter=Config.RoomClassificationParameter;CoefficientInputs.IsEnabled=true;
            RefreshDepartmentCards();
            Engine.RefreshTopSlabs(ReportProgress);
            Engine.PreviewLevelGeometry(ReportProgress);
        }
        private void OpenReview_Click(object s,RoutedEventArgs e)
        {Try(()=>{var item=ReviewGrid.SelectedItem as TEP.ReviewInput;if(item==null)throw new InvalidOperationException("Выберите строку контура.");Engine.RequestReviewNavigation(item);Close();});}
        private void BindReview_Click(object s,RoutedEventArgs e)
        {Try(()=>{var item=ReviewGrid.SelectedItem as TEP.ReviewInput;if(item==null)throw new InvalidOperationException("Выберите строку контура.");Engine.PendingReviewKey=item.Key;Engine.PendingContourAction="review-bind";Close();});}
        private void ScanRooms_Click(object s,RoutedEventArgs e)
        {Try(()=>{SetBusy(true);try{PrepareLinkedSources();ScanRoomDepartments();Status.Text="Назначения считаны. Распределите значения по категориям.";}catch(System.OperationCanceledException){Status.Text="Сканирование отменено. Уже загруженные связи остаются загруженными.";}finally{SetBusy(false);}});}
        private void PrepareLinkedSources()
        {
            if(!Engine.Sources.Any(source=>source.Mode!="exclude"&&!source.Loaded))return;
            try
            {
                Engine.LoadMissingLinks(ReportProgress,message=>MessageBox.Show(this,message,"KPLN | ТЭП: незагруженные связи",
                    MessageBoxButton.YesNo,MessageBoxImage.Warning,MessageBoxResult.No)==MessageBoxResult.Yes);
            }
            finally
            {
                checkedConfiguration=null;departmentGroups=null;scannedPhase=null;
                Rebind();RefreshDepartmentCards();
            }
        }
        private void RefreshDepartmentCards()
        {
            RefreshRequiredParameters();
            if(CategoryCards==null)return;
            CategoryCards.Children.Clear();
            var groups=departmentGroups??new List<TEP.DepartmentGroup>();
            var destination=DepartmentTarget.SelectedValue;DepartmentTarget.ItemsSource=CategoryTargetChoices;DepartmentTarget.SelectedValue=destination??"unknown";
            int unassigned=groups.Where(g=>Engine.DepartmentRole(g.Value)=="unknown").Sum(g=>g.Count);
            int assigned=groups.Sum(g=>g.Count)-unassigned;
            var unknown=new Button{Content="Нераспределённые ("+unassigned+")",Tag="unknown",MinWidth=220,Height=48,Margin=new Thickness(0,0,8,0),ToolTip="Показывает ниже назначения, которым ещё не назначена категория расчёта."};
            if(unassigned>0){unknown.Background=new SolidColorBrush(Color.FromRgb(255,230,230));unknown.Foreground=ErrorTextBrush;}
            unknown.IsEnabled=unassigned>0;
            unknown.Click+=(s,e)=>{unassignedExpanded=!unassignedExpanded;RefreshDepartmentCards();};CategoryCards.Children.Add(unknown);
            var distributed=new Button{Content="Распределённые ("+assigned+")",Tag="assigned",MinWidth=220,Height=48,Margin=new Thickness(0,0,8,0),ToolTip="Открывает все категории в отдельном окне. Назначения можно переносить между категориями или вернуть в нераспределённые."};
            distributed.Click+=AssignedDepartments_Click;CategoryCards.Children.Add(distributed);
            var settings=new Button{Content="Настройка категорий",MinWidth=220,Height=48,Margin=new Thickness(0),ToolTip="Правила категорий, источники семейств, общий словарь и запись назначения в параметр."};
            settings.Click+=CategorySettings_Click;CategoryCards.Children.Add(settings);
            ClassificationScanText.Text=departmentGroups==null?"Помещения ещё не просканированы.":"Помещений: "+groups.Sum(g=>g.Count)+". Разных назначений: "+groups.Count+".";
            selectedDepartmentRole="unknown";
            CategoryTitle.Text="Нераспределённые";CategoryDescription.Text=unassigned>0?"Выберите назначения и категорию, затем нажмите «Распределить».":"Все найденные назначения распределены.";
            RoomCategoryPanel.Visibility=unassigned>0&&unassignedExpanded?Visibility.Visible:Visibility.Collapsed;
            DepartmentValues.ItemsSource=groups.Where(g=>Engine.DepartmentRole(g.Value)=="unknown").ToList();
            RefreshDepartmentRooms();
        }
        private void DepartmentValues_Changed(object s,SelectionChangedEventArgs e){if(DepartmentRoomsGrid!=null)RefreshDepartmentRooms();}
        private void RefreshDepartmentRooms()
        {
            var selected=DepartmentValues.SelectedItems.Cast<TEP.DepartmentGroup>().ToList();
            var visible=selected.Count>0?selected:DepartmentValues.Items.Cast<TEP.DepartmentGroup>().ToList();
            var rooms=Engine.CategoryUsesFamilies(selectedDepartmentRole)?Engine.CategoryFamilyInstances(selectedDepartmentRole):visible.SelectMany(g=>g.Rooms).ToList();DepartmentRoomsGrid.ItemsSource=rooms;
            DepartmentRoomsTitle.Text=(Engine.CategoryUsesFamilies(selectedDepartmentRole)?"Экземпляров семейств: ":"Помещений: ")+rooms.Count;
            AssignDepartmentButton.IsEnabled=selected.Count>0;
        }
        private void AssignDepartment_Click(object s,RoutedEventArgs e)
        {Try(()=>{string role=DepartmentTarget.SelectedValue as string;if(role==null)return;
            foreach(var group in DepartmentValues.SelectedItems.Cast<TEP.DepartmentGroup>().ToList())Engine.AssignDepartment(group.Value,role);
            Engine.SetCategorySource(role,false);checkedConfiguration=null;RefreshDepartmentCards();Status.Text="Распределение обновлено. Перед расчётом помещения будут проверены.";AutoSaveCategoryAssignments();});}
        private string ScanConfiguration()
        {
            var copy=TEP.Engine.Deserialize<TEP.Settings>(TEP.Engine.Serialize(Config));
            copy.LoggiaCoefficient="";copy.BalconyCoefficient="";
            return TEP.Engine.Serialize(copy);
        }
        private bool CheckBeforeCalculation(bool openCoefficients)
        {
            ScanRoomDepartments();
            var issues=Engine.CheckParameters(ReportProgress);
            CoefficientInputs.IsEnabled=true;
            checkedConfiguration=ScanConfiguration();ShowIssues(issues);
            string message=TEP.Engine.PreflightMessage(Config.Metrics,issues);
            if(issues.Any(i=>i.Severity=="Ошибка"))
            {
                SummaryGrid.ItemsSource=null;RefreshObjectReview(null);FloorReport.Document=new FlowDocument();
                if(ResultsTab.IsEnabled){Steps.SelectedItem=ResultsTab;ReportTabs.SelectedItem=IssuesTab;}ReportStatus.Text=message;ReportStatus.Foreground=ErrorTextBrush;
                Status.Text=message+" Нажмите «Рассчитать ТЭП».";
            }
            else
            {
                Steps.SelectedIndex=1;SettingsScroll.ScrollToTop();Status.Text=message;
            }
            return true;
        }
        private readonly System.Diagnostics.Stopwatch progressClock=System.Diagnostics.Stopwatch.StartNew();
        private List<TEP.Issue> displayedIssues;
        public AR_CalculateTEP(TEP.Engine engine)
        {
            Engine=engine;InitializeComponent();DataContext=this;ConfigureTables();
            RoomNavigationPanel.Children.Add(CreateRoomNavigationButton(DepartmentRoomsGrid,this));
            Loaded+=(s,e)=>Dispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>
            {
                SetBusy(true);
                try {
                    if(!Engine.IsInitialized)Engine.Initialize(ReportProgress);
                    Config.CreateViews=false;Rebind();ConfigureTables();ShowReport(Engine.Last);
                    try{ScanRoomDepartments();Status.Text=(Engine.SettingsStatus??"Модель загружена.")+" Распределите назначения помещений по категориям.";}
                    catch(System.OperationCanceledException){Status.Text="Сканирование отменено. Повторите его во вкладке «Классификация».";}
                    catch(Exception scanError){Status.Text="Не удалось считать назначения: "+scanError.Message;MessageBox.Show(this,Status.Text,"KPLN | Сканирование помещений",MessageBoxButton.OK,MessageBoxImage.Warning);}
                }
                catch(System.OperationCanceledException){SetBusy(false);Close();return;}
                catch(Exception ex){SetBusy(false);MessageBox.Show(this,ex.Message,"KPLN | Не удалось загрузить модель ТЭП",MessageBoxButton.OK,MessageBoxImage.Error);Close();return;}
                finally {SetBusy(false);}
            }));
            Closing+=(s,e)=>{if(busy)e.Cancel=true;};
            Closed+=(s,e)=>StopObservingSlabRows();
        }
        private void SetBusy(bool value)
        {
            busy=value;cancelRequested=false;Steps.IsEnabled=!value;
            CalculateButton.IsEnabled=!value;
            SaveSettingsButton.IsEnabled=!value;
            CancelButton.Visibility=value?Visibility.Visible:Visibility.Collapsed;CancelButton.IsEnabled=value;
            WorkProgress.Visibility=value?Visibility.Visible:Visibility.Collapsed;
            if(value)progressClock.Restart();
        }
        private void Cancel_Click(object sender,RoutedEventArgs e)
        {cancelRequested=true;CancelButton.IsEnabled=false;Status.Text="Отмена после текущей операции Revit...";}
        private void ReportProgress(string message)
        {
            if(cancelRequested)throw new System.OperationCanceledException();
            if(progressClock.ElapsedMilliseconds<80&&!message.StartsWith("Загрузка связи")&&!message.StartsWith("Доступен частичный расчёт")&&!message.StartsWith("Недостаточно данных"))return;
            Status.Text=message;progressClock.Restart();
            // Revit API remains on the command thread. Pump only while all editing actions are disabled.
            var frame=new DispatcherFrame();
            Dispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>frame.Continue=false));
            Dispatcher.PushFrame(frame);
            if(cancelRequested)throw new System.OperationCanceledException();
        }
        private static Binding Bind(string path,string format=null)
        {
            var binding=new Binding(path){Mode=BindingMode.TwoWay,UpdateSourceTrigger=UpdateSourceTrigger.PropertyChanged,StringFormat=format,ValidatesOnExceptions=true,NotifyOnValidationError=true};
            int digits;if(format!=null&&format.StartsWith("{0:N")&&int.TryParse(format.Substring(4).TrimEnd('}'),out digits))binding.Converter=new RoundedNumber{Digits=digits};
            return binding;
        }
        private sealed class RoundedNumber : IValueConverter
        {
            internal int Digits;
            public object Convert(object value,Type targetType,object parameter,System.Globalization.CultureInfo culture)
            {return value is double?Math.Round((double)value,Digits,MidpointRounding.AwayFromZero):value;}
            public object ConvertBack(object value,Type targetType,object parameter,System.Globalization.CultureInfo culture){return Binding.DoNothing;}
        }
        private void RefreshPrecision(DataGrid grid)
        {foreach(var column in grid.Columns.OfType<DataGridTextColumn>())if(new[]{"Value","DisplayValue"}.Contains((column.Binding as Binding)?.Path.Path)){var binding=Bind(((Binding)column.Binding).Path.Path,"{0:N"+Config.Decimals+"}");binding.Mode=BindingMode.OneWay;column.Binding=binding;}}
        private static string ColumnTip(DataGrid grid,string title,string path,bool editable,string extra=null)
        {
            if(grid?.Name=="ContoursGrid")
            {
                switch(path)
                {
                    case "ElementLabel":return "ID цветовой области активной модели. Расчёт читает её актуальные границы при каждом запуске.";
                    case "Kind":return "Основной контур задаёт исходную площадь в ручном режиме; вырез вычитается также в режиме наружных стен.";
                    case "LevelName":return "Уровень плана, на котором находится область. Изменение уровня требует повторного закрепления области.";
                    case "Building":return "Имя расчётного корпуса. Вырезы помещений применяются при совпадении корпуса и секции с основным контуром.";
                    case "Section":return "Имя секции. Пустое значение означает секцию без имени; оно не является подстановкой для всех секций.";
                    case "Profile":return "Профиль здания для проверки правил включения площади этажа и нормативных исключений.";
                    case "BuildingClass":return "Жилое или нежилое здание. Определяет принадлежность поэтажных площадей к соответствующему показателю.";
                    case "Height":return "Высота вертикального участка в метрах. Пусто - до верха участка, заданного таблицей этажей или следующим включённым уровнем, с учётом смещения низа. Для последнего участка высота обязательна; для плоских площадей не применяется.";
                    case "BottomOffset":return "Смещение низа ручного выреза относительно его уровня, в метрах. Положительное поднимает, отрицательное опускает вырез. У основного контура должно быть нулём; для площадей не применяется.";
                }
                return extra??title;
            }
            string meaning=null,example=null;
            switch(path)
            {
                case "Source":meaning=grid==null?"":grid.Name=="BuildingsGrid"?"Укажите имя источника или *. Ограничивает строку соответствия выбранной моделью; * действует для всех источников.":"Укажите имя модели или экземпляра связи из списка. Определяет, к какому источнику относится корректировка.";example="Корпус 1.rvt";break;
                case "MatchValue":meaning="Укажите исходное значение корпуса, рабочего набора или имени связи либо *. Совпадение назначает объектам корпус и профиль этой строки.";example="Секция А";break;
                case "Building":meaning=grid.Name=="LevelsGrid"?"Укажите корпус уточнения уровня. Пусто - общая настройка; имя ограничивает уточнение одним корпусом.":"Введите итоговое имя корпуса. Под этим именем объединяются площади и другие показатели здания.";example="Корпус 1";break;
                case "Section":meaning="Укажите секцию уточнения уровня. Позволяет отдельно задать наземность и параметры этажа этой секции; пусто - все секции.";example="А";break;
                case "TopSlab":meaning="Введите абсолютную отметку верха перекрытия в метрах основной модели. Используется при определении наземности цоколя.";example="158,20";break;
                case "Height":meaning="Введите высоту этажа в метрах. Используется в правилах включения технического этажа / надстройки; пустое значение требует данных модели.";example="2,40";break;
                case "RoofRatio":meaning="Введите долю площади надстройки от кровли от 0 до 1. Влияет на включение технической надстройки общественного здания.";example="0,15";break;
                case "RoofArea":meaning="Введите площадь надстройки последнего верхнего этажа в м². Вместе с высотой применяется для правил высотного здания.";example="7,5";break;
                case "Name":meaning="Введите точное имя параметра экземпляра или типа для назначения, указанного в этой строке. Правило чтения и влияние описаны в соседнем столбце; пусто - параметр не назначен.";example="ТЭП_Назначение";break;
                case "Category":meaning="Выберите или введите точное имя категории. Правило проверяет только эту категорию; пусто - все категории.";example="Помещения";break;
                case "Parameter":meaning="Выберите или введите имя параметра экземпляра / типа. По его значению правило определяет назначение и включение объекта; @Name означает имя объекта.";example="ТЭП_Назначение";break;
                case "Value":meaning=grid.Name=="CorrectionsGrid"?"Введите значение для выбранного действия: назначение, имя корпуса или числовую дельту в единицах показателя. Корректирует только расчёт.":"Введите сравниваемое значение параметра. Используется условием «Равно» или «Содержит»; для «Заполнен» / «Пусто» не требуется.";example=grid.Name=="CorrectionsGrid"?"12,5 (дельта площади в м²)":"Лоджия";break;
                case "Coefficient":meaning="Введите коэффициент от 0 до 1 либо оставьте пустым для нормативного значения назначения. Меняет площадь квартиры с летними помещениями; параметр коэффициента объекта имеет приоритет.";example="0,5";break;
                case "Element":meaning="Введите ElementId или UniqueId объекта. Для числовой дельты укажите корпус, для дельты этажности - Корпус|Секция. Определяет объект корректировки.";example="123456";break;
                case "Reason":meaning="Введите причину ручного решения. Обоснование сохраняется в отчёте вместе с автором, датой и прежним значением.";example="Исключён дублирующий контур лоджии";break;
                case "AreaScheme":meaning="Выберите или введите схему зон этого показателя. Пусто - общая схема; отдельное значение позволяет считать разные показатели по разным контурам.";example="ГНС - наружный контур";break;
                case "Mode":meaning="Выберите участие источника. Включение даёт вклад в итог; исключение убирает источник; сверка оставляет только проверочную графику.";break;
                case "Profile":meaning="Выберите нормативный профиль источника или корпуса. Определяет включение помещений и этажей при соответствующем способе выбора типа здания.";break;
                case "Class":meaning="Выберите класс здания. От него зависит отнесение ГНС и НП к жилым или нежилым зданиям.";break;
                case "Kind":meaning="Выберите вид этажа. Меняет нормативные правила включения его площадей и этажности.";break;
                case "Above":meaning="Выберите наземность этажа. Автоматический режим использует землю и параметры этажа; ручной вариант переопределяет его для этого уровня.";break;
                case "Geometry":meaning="Выберите источник площади показателя. Пустой выбор наследует общую настройку; Rooms, Spaces и Areas не должны дублировать один этаж.";break;
                case "Metric":meaning="Выберите показатель, на который действует строка. «Все показатели» задаёт общую классификацию / корректировку до правил отдельного показателя.";break;
                case "Operation":meaning="Выберите условие сравнения параметра. При совпадении применяется действие и назначение этой строки.";break;
                case "Action":meaning=grid.Name=="CorrectionsGrid"?"Выберите вид ручной корректировки. Действие определяет смысл полей объекта и нового значения.":"Выберите действие правила. Классификация задаёт назначение; включение или исключение переопределяет участие объекта.";break;
                case "Role":meaning="Выберите назначение совпавших объектов. От него зависят нормативное включение и коэффициент площади.";break;
                case "Part":meaning="Выберите часть здания. Определяет распределение площади жилого здания на жилую и нежилую части.";break;
                case "Loaded":meaning="Показывает, доступна ли модель связи. Незагруженный источник нельзя рассчитать; загрузите связь в Revit либо исключите её.";break;
                case "Enabled":meaning="Включите флажок, чтобы правило участвовало в классификации объектов. Снятый флажок отключает строку без удаления.";break;
                case "Include":meaning=grid.Name=="LevelsGrid"?"Включите флажок, чтобы уровень участвовал в расчёте. Снятие исключает его из состава расчётных этажей.":"Включите флажок, чтобы учитывать объекты этого соответствия корпуса. Снятие исключает совпавшие объекты.";break;
            }
            if(meaning==null)meaning=extra??("«"+title+"»: значение в строке результата.");
            if(!string.IsNullOrWhiteSpace(extra)&&!meaning.Contains(extra))meaning+="\n"+extra;
            return meaning+(editable&&example!=null?"\nПример: "+example:"");
        }
        private static void TextColumn(DataGrid grid,string title,string path,double width=150,bool readOnly=false,string format=null,string tip=null)
        {
            tip=readOnly?(tip??("«"+title+"»: значение в строке результата.")):ColumnTip(grid,title,path,true,tip);
            var header=new TextBlock{Text=title,ToolTip=tip,VerticalAlignment=VerticalAlignment.Center};
            var binding=Bind(path,format);if(readOnly)binding.Mode=BindingMode.OneWay;
            var elementStyle=new Style(typeof(TextBlock));elementStyle.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center));
            elementStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty,new Thickness(7,0,7,0)));elementStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty,TextTrimming.CharacterEllipsis));
            elementStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,readOnly?(object)new Binding(path):tip));
            string errorPath=grid.Name=="IssuesGrid"?"Severity":grid.Name=="SummaryGrid"?"Status":grid.Name=="DetailsGrid"?"HasError":null;
            if(errorPath!=null)
            {
                var error=new DataTrigger{Binding=new Binding(errorPath),Value=errorPath=="HasError"?(object)true:errorPath=="Severity"?"Ошибка":"Неполный результат"};
                error.Setters.Add(new Setter(TextBlock.ForegroundProperty,ErrorTextBrush));elementStyle.Triggers.Add(error);
                if(grid.Name=="SummaryGrid")
                {var skipped=new DataTrigger{Binding=new Binding("Status"),Value="Не рассчитано"};skipped.Setters.Add(new Setter(TextBlock.ForegroundProperty,ErrorTextBrush));elementStyle.Triggers.Add(skipped);}
            }
            var editStyle=new Style(typeof(TextBox));editStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,tip));editStyle.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center));editStyle.Setters.Add(new Setter(FrameworkElement.MinHeightProperty,28.0));editStyle.Setters.Add(new Setter(Control.PaddingProperty,new Thickness(4,2,4,2)));
            grid.Columns.Add(new DataGridTextColumn{Header=header,Binding=binding,Width=width,IsReadOnly=readOnly,ElementStyle=elementStyle,EditingElementStyle=editStyle});
        }
        private static void CheckColumn(DataGrid grid,string title,string path,double width=70,bool readOnly=false)
        {
            var binding=Bind(path);if(readOnly)binding.Mode=BindingMode.OneWay;
            var style=new Style(typeof(CheckBox));style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center));
            style.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty,HorizontalAlignment.Center));style.Setters.Add(new Setter(FrameworkElement.MarginProperty,new Thickness(0)));
            style.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,ColumnTip(grid,title,path,false)));
            var display=new Style(typeof(CheckBox),style);display.Setters.Add(new Setter(UIElement.IsHitTestVisibleProperty,false));display.Setters.Add(new Setter(UIElement.FocusableProperty,false));
            grid.Columns.Add(new DataGridCheckBoxColumn{Header=new TextBlock{Text=title,ToolTip=ColumnTip(grid,title,path,false)},Binding=binding,Width=width,IsReadOnly=readOnly,ElementStyle=display,EditingElementStyle=style});
        }
        private void ChoiceColumn(DataGrid grid,string title,string path,IEnumerable<TEP.Choice> choices,double width=180)
        {
            var combo=new FrameworkElementFactory(typeof(ComboBox));combo.SetValue(ItemsControl.ItemsSourceProperty,choices.ToList());
            combo.SetValue(ItemsControl.DisplayMemberPathProperty,"Label");combo.SetValue(System.Windows.Controls.Primitives.Selector.SelectedValuePathProperty,"Key");
            combo.SetBinding(System.Windows.Controls.Primitives.Selector.SelectedValueProperty,Bind(path));
            if(grid==LevelsGrid&&(path=="Kind"||path=="Above"))combo.AddHandler(ComboBox.SelectionChangedEvent,new SelectionChangedEventHandler(LevelKind_Changed));
            combo.SetValue(Control.PaddingProperty,new Thickness(4,2,4,2));combo.SetValue(FrameworkElement.MinHeightProperty,28.0);combo.SetValue(FrameworkElement.HeightProperty,28.0);combo.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);combo.SetValue(FrameworkElement.MarginProperty,new Thickness(4,0,4,0));
            combo.SetValue(FrameworkElement.TagProperty,ColumnTip(grid,title,path,false));
            var template=new DataTemplate{VisualTree=combo};grid.Columns.Add(new DataGridTemplateColumn{Header=new TextBlock{Text=title,ToolTip=ColumnTip(grid,title,path,false)},CellTemplate=template,Width=width});
        }
        private static void EditableColumn(DataGrid grid,string title,string path,IEnumerable<string> values,double width=220,string tooltip=null)
        {
            var combo=new FrameworkElementFactory(typeof(ComboBox));combo.SetValue(ItemsControl.ItemsSourceProperty,values.ToList());combo.SetValue(ComboBox.IsEditableProperty,true);
            combo.SetBinding(ComboBox.TextProperty,Bind(path));combo.SetValue(Control.PaddingProperty,new Thickness(4,2,4,2));combo.SetValue(FrameworkElement.MinHeightProperty,28.0);combo.SetValue(FrameworkElement.HeightProperty,28.0);combo.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);combo.SetValue(FrameworkElement.MarginProperty,new Thickness(4,0,4,0));
            tooltip=ColumnTip(grid,title,path,true,tooltip);
            if(grid.Name=="ParametersGrid")
            {var tip=new MultiBinding{StringFormat="{0}\n{1}"};tip.Bindings.Add(new Binding{Source=tooltip});tip.Bindings.Add(new Binding("Description"));combo.SetBinding(FrameworkElement.ToolTipProperty,tip);}
            else combo.SetValue(FrameworkElement.ToolTipProperty,tooltip);
            grid.Columns.Add(new DataGridTemplateColumn{Header=new TextBlock{Text=title,ToolTip=tooltip},CellTemplate=new DataTemplate{VisualTree=combo},Width=width});
        }
        private void ParameterSourceColumn(DataGrid grid)
        {
            var panel=new FrameworkElementFactory(typeof(Grid));
            var selector=new FrameworkElementFactory(typeof(ComboBox));selector.SetValue(ComboBox.IsEditableProperty,true);selector.SetValue(ItemsControl.ItemsSourceProperty,Engine.Parameters);selector.SetBinding(ComboBox.TextProperty,Bind("Name"));
            selector.SetValue(FrameworkElement.MarginProperty,new Thickness(4));selector.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);
            selector.SetValue(Control.PaddingProperty,new Thickness(4,2,4,2));selector.SetValue(Control.FontSizeProperty,12.0);selector.SetValue(FrameworkElement.HeightProperty,28.0);
            var selectorStyle=new Style(typeof(ComboBox));var hide=new DataTrigger{Binding=new Binding("Key"),Value="$stairs"};hide.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));selectorStyle.Triggers.Add(hide);selector.SetValue(FrameworkElement.StyleProperty,selectorStyle);panel.AppendChild(selector);
            var button=new FrameworkElementFactory(typeof(Button));button.SetValue(ContentControl.ContentProperty,"Лестницы и Марши");button.AddHandler(Button.ClickEvent,new RoutedEventHandler(Stairs_Click));button.SetValue(FrameworkElement.MarginProperty,new Thickness(4));button.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);
            button.SetBinding(FrameworkElement.ToolTipProperty,new Binding("Description"));
            var style=new Style(typeof(Button));style.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));var show=new DataTrigger{Binding=new Binding("Key"),Value="$stairs"};show.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Visible));style.Triggers.Add(show);button.SetValue(FrameworkElement.StyleProperty,style);panel.AppendChild(button);
            grid.Columns.Add(new DataGridTemplateColumn{Header="Параметр / настройка",Width=280,CellTemplate=new DataTemplate{VisualTree=panel}});
        }
        private void ConfigureTables()
        {
            foreach(var grid in new[]{LevelsGrid,ParametersGrid,DepartmentRoomsGrid,SummaryGrid,IssuesGrid})grid.Columns.Clear();
            CheckColumn(LevelsGrid,"Учесть","Include");TextColumn(LevelsGrid,"Источник","Source",433.755,true);TextColumn(LevelsGrid,"Уровень","Name",85,true);TextColumn(LevelsGrid,"Отметка, м","ElevationMeters",110,true,"{0:0.000}");
            ChoiceColumn(LevelsGrid,"Вид этажа","Kind",TEP.Choices("level"),230);ChoiceColumn(LevelsGrid,"Наземность","Above",TEP.Choices("above"),145);
            TopSlabColumn();
            LevelGeometryColumns();
            foreach(var grid in new[]{ParametersGrid})
            {
                var titleLayout=new FrameworkElementFactory(typeof(DockPanel));titleLayout.SetValue(FrameworkElement.MarginProperty,new Thickness(7,6,7,6));titleLayout.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);
                var marker=new FrameworkElementFactory(typeof(TextBlock));marker.SetValue(TextBlock.TextProperty,"[!]");marker.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);marker.SetValue(TextBlock.ForegroundProperty,ErrorTextBrush);marker.SetValue(TextBlock.FontWeightProperty,FontWeights.Bold);marker.SetValue(FrameworkElement.MarginProperty,new Thickness(0,0,6,0));marker.SetValue(DockPanel.DockProperty,Dock.Left);
                var markerStyle=new Style(typeof(TextBlock));markerStyle.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));
                foreach(string key in new[]{"embedded","standalone"}){var trigger=new DataTrigger{Binding=new Binding("Key"),Value=key};trigger.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Visible));markerStyle.Triggers.Add(trigger);}
                marker.SetValue(FrameworkElement.StyleProperty,markerStyle);titleLayout.AppendChild(marker);
                var title=new FrameworkElementFactory(typeof(TextBlock));title.SetBinding(TextBlock.TextProperty,new Binding("Title"));title.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);title.SetBinding(FrameworkElement.ToolTipProperty,new Binding("Description"));title.SetValue(TextBlock.TextWrappingProperty,TextWrapping.Wrap);
                var titleStyle=new Style(typeof(TextBlock));
                var requiredTrigger=new DataTrigger{Binding=new Binding("IsRequired"),Value=true};requiredTrigger.Setters.Add(new Setter(TextBlock.FontWeightProperty,FontWeights.Bold));titleStyle.Triggers.Add(requiredTrigger);
                foreach(string key in new[]{"embedded","standalone"}){var trigger=new DataTrigger{Binding=new Binding("Key"),Value=key};trigger.Setters.Add(new Setter(TextBlock.FontWeightProperty,FontWeights.Normal));titleStyle.Triggers.Add(trigger);}
                title.SetValue(FrameworkElement.StyleProperty,titleStyle);titleLayout.AppendChild(title);
                grid.Columns.Add(new DataGridTemplateColumn{Header="Назначение",Width=343.85,MinWidth=343.85,IsReadOnly=true,CellTemplate=new DataTemplate{VisualTree=titleLayout}});
                ParameterSourceColumn(grid);TextColumn(grid,"Правило чтения","Description",500,true);
                grid.RowHeight=double.NaN;grid.MinRowHeight=52;grid.Columns[1].MinWidth=280;
                var description=(DataGridTextColumn)grid.Columns.Last();description.Width=new DataGridLength(1,DataGridLengthUnitType.Star);description.MinWidth=260;
                var wrap=new Style(typeof(TextBlock),description.ElementStyle);wrap.Setters.Add(new Setter(TextBlock.TextWrappingProperty,TextWrapping.Wrap));wrap.Setters.Add(new Setter(TextBlock.TextTrimmingProperty,TextTrimming.None));wrap.Setters.Add(new Setter(FrameworkElement.MarginProperty,new Thickness(7,6,7,6)));description.ElementStyle=wrap;
            }
            TextColumn(DepartmentRoomsGrid,"Номер","Number",85,true);TextColumn(DepartmentRoomsGrid,"Имя","Name",160,true);
            TextColumn(DepartmentRoomsGrid,"Назначение","Department",140,true);TextColumn(DepartmentRoomsGrid,"Этаж","Level",130,true);
            TextColumn(DepartmentRoomsGrid,"Источник","Source",180,true);TextColumn(DepartmentRoomsGrid,"ID","Element",90,true);
            TextColumn(DepartmentRoomsGrid,"Площадь Revit, м²","Area",145,true,"{0:N3}");
            TextColumn(SummaryGrid,"Показатель","Name",540,true);TextColumn(SummaryGrid,"Значение","DisplayValue",135,true,"{0:N"+Config.Decimals+"}");TextColumn(SummaryGrid,"Ед.","Unit",65,true);TextColumn(SummaryGrid,"Статус","Status",180,true);TextColumn(SummaryGrid,"Методика","Method",170,true);TextColumn(SummaryGrid,"Пояснение","Comment",400,true);
            TextColumn(IssuesGrid,"Важность","Severity",135,true);TextColumn(IssuesGrid,"Код","Code",185,true);TextColumn(IssuesGrid,"Сообщение","Message",540,true);TextColumn(IssuesGrid,"Показатель","MetricName",300,true);TextColumn(IssuesGrid,"Источник","Source",220,true);TextColumn(IssuesGrid,"Корпус","Building",130,true);TextColumn(IssuesGrid,"ElementId","Element",110,true);TextColumn(IssuesGrid,"Что сделать","Action",480,true);
        }
        private IEnumerable<T> Children<T>(DependencyObject parent) where T:DependencyObject
        {for(int i=0;i<VisualTreeHelper.GetChildrenCount(parent);i++){var child=VisualTreeHelper.GetChild(parent,i);if(child is T)yield return (T)child;foreach(var nested in Children<T>(child))yield return nested;}}
        private void CommitInputs()
        {
            foreach(var grid in new[]{LevelsGrid,ParametersGrid})
                if(!grid.CommitEdit(DataGridEditingUnit.Cell,true)||!grid.CommitEdit(DataGridEditingUnit.Row,true))throw new InvalidOperationException("Исправьте значение в таблице.");
            foreach(var box in Children<TextBox>(this))box.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            if(Children<FrameworkElement>(this).Any(Validation.GetHasError))throw new InvalidOperationException("Исправьте поля с ошибками ввода.");
            if(Config.Decimals<0||Config.Decimals>6)throw new InvalidOperationException("Точность отображения: от 0 до 6 знаков.");

        }
        private void Try(Action action)
        {try{CommitInputs();action();}catch(Exception ex){MessageBox.Show(this,ex.Message,"KPLN | ТЭП",MessageBoxButton.OK,MessageBoxImage.Warning);Status.Text=ex.Message;}}
        private void SaveSettings_Click(object sender,RoutedEventArgs e)
        {Try(()=>{SetBusy(true);try{Engine.SaveSettings();Rebind();Status.Text="Настройки ТЭП записаны в модель. Сохраните RVT обычной командой Revit, чтобы они остались в файле.";}finally{SetBusy(false);}});}
        private void Rebind(){RefreshRequiredParameters();DataContext=null;DataContext=this;if(CoefficientInputs!=null)CoefficientInputs.IsEnabled=departmentGroups!=null;}
        private void SelectAll_Click(object sender,RoutedEventArgs e){changing=true;foreach(var m in Config.Metrics)m.Enabled=true;Rebind();changing=false;}
        private void SelectNone_Click(object sender,RoutedEventArgs e){changing=true;foreach(var m in Config.Metrics)m.Enabled=false;Rebind();changing=false;}
        private void Metric_Checked(object sender,RoutedEventArgs e)
        {
            if(changing||!IsLoaded)return;var check=sender as CheckBox;var metric=check?.DataContext as TEP.Metric;if(metric==null)return;
            var dependencies=new Dictionary<string,string[]>{{"Gns",new[]{"GnsResidential","GnsNonresidential"}},{"GnsResidential",new[]{"GnsLivingPart","GnsNonlivingPart"}},{"Volume",new[]{"VolumeAbove","VolumeBelow"}},{"Gross",new[]{"GrossAbove","GrossBelow"}},{"Np",new[]{"NpResidential","NpNonresidential"}}};
            string[] keys;if(dependencies.TryGetValue(metric.Key,out keys))
            {changing=true;foreach(var key in keys)Config.Metrics.First(x=>x.Key==key).Enabled=true;Rebind();changing=false;}
            RefreshRequiredParameters();
        }
        private void Calculate_Click(object s,RoutedEventArgs e)
        {
            Try(()=>
            {
                if(!Config.Metrics.Any(x=>x.Enabled))throw new InvalidOperationException("Выберите хотя бы один показатель.");
                // Verification output is disabled for the current workflow, including imported settings.
                Config.CreateViews=false;
                SetBusy(true);
                try{PrepareLinkedSources();ScanRoomDepartments();SetBusy(false);ResolveAmbiguousTopSlabs();ResolveLevelGeometryChoices();SetBusy(true);var audit=Engine.AuditInputs(ReportProgress);SetBusy(false);if(!ShowInputAudit(audit))return;Engine.InputAuditForRun=audit;SetBusy(true);var run=Engine.Calculate(ReportProgress,false);ShowReport(run);Steps.SelectedItem=ResultsTab;Status.Text="Расчёт завершён. Выполнено показателей: "+run.Summary.Count(x=>!x.NotCalculated)+", пропущено: "+run.Summary.Count(x=>x.NotCalculated)+". Ошибок: "+run.Issues.Count(x=>x.Severity=="Ошибка")+", предупреждений: "+run.Issues.Count(x=>x.Severity=="Предупреждение")+".";}
                catch(System.OperationCanceledException){Status.Text="Расчёт отменён; предыдущий отчёт сохранён. Уже загруженные связи остаются загруженными.";}
                finally{Engine.ClearSingleBuildingAssumption();SetBusy(false);}
            });
        }
        private void ShowReport(TEP.Run run)
        {
            ResultsTab.IsEnabled=run!=null;if(run==null)return;RefreshPrecision(SummaryGrid);SummaryGrid.ItemsSource=run.Summary;RefreshObjectReview(run);ShowIssues(run.Issues);FloorReport.Document=ColoredReport(run,Config.Decimals);
            ResultAuditContent.Children.Clear();ResultAuditContent.Children.Add(run.InputAudit==null?(UIElement)new TextBlock{Text="В сохранённом результате нет проверки параметров. Выполните новый расчёт.",Margin=new Thickness(12)}:
                CreateInputAuditTable(run.InputAudit,row=>Try(()=>{Engine.RequestAuditNavigation(row);Close();})));
            string state=run.Summary.Count>0&&run.Summary.All(x=>x.NotCalculated)?"Нет доступных показателей":run.Summary.Any(x=>x.NotCalculated||x.Status=="Неполный результат")?"Частичный расчёт":run.Issues.Any(x=>x.Severity!="Информация")?"Расчёт с замечаниями":"Расчёт завершён";
            ReportStatus.Text=state+" | "+run.Date+" | "+run.Method;ReportStatus.Foreground=run.Issues.Any(x=>x.Severity=="Ошибка")?ErrorTextBrush:(Brush)FindResource("Ink");ReportTabs.SelectedIndex=0;
        }
        private static void AppendReportText(Paragraph paragraph,string text,IEnumerable<string> errors)
        {
            // Use complete structured error messages, not code/keyword guesses: a message can contain | and newlines.
            var ranges=new List<Tuple<int,int>>();
            foreach(var error in errors.Where(e=>!string.IsNullOrEmpty(e)).Distinct())
                for(int position=0;(position=text.IndexOf(error,position,StringComparison.Ordinal))>=0;position+=error.Length)
                    ranges.Add(Tuple.Create(position,position+error.Length));
            int cursor=0;
            foreach(var range in ranges.OrderBy(r=>r.Item1).ThenByDescending(r=>r.Item2))
            {
                if(range.Item2<=cursor)continue;
                if(range.Item1>cursor){paragraph.Inlines.Add(new Run(text.Substring(cursor,range.Item1-cursor)));cursor=range.Item1;}
                paragraph.Inlines.Add(new Run(text.Substring(cursor,range.Item2-cursor)){Foreground=ErrorTextBrush});cursor=range.Item2;
            }
            if(cursor<text.Length)paragraph.Inlines.Add(new Run(text.Substring(cursor)));
        }
        private static FlowDocument ColoredReport(TEP.Run report,int decimals)
        {
            string text=report.TextReport(decimals);
            var paragraph=new Paragraph{Margin=new Thickness(0)};
            var document=new FlowDocument(paragraph){FontFamily=new FontFamily("Consolas"),FontSize=13,PagePadding=new Thickness(10),ColumnWidth=double.PositiveInfinity};
            // Preserve the original text, including rounding and the per-floor breakdown. Scope errors to each metric.
            var sections=new List<Tuple<int,TEP.Summary>>();int search=0;
            foreach(var summary in report.Summary)
            {
                string header=summary.Name+": "+summary.ValueText(decimals)+" | "+summary.Status;
                int position=text.IndexOf(header,search,StringComparison.Ordinal);
                if(position<0)continue;sections.Add(Tuple.Create(position,summary));search=position+header.Length;
            }
            if(sections.Count==0){paragraph.Inlines.Add(new Run(text));return document;}
            paragraph.Inlines.Add(new Run(text.Substring(0,sections[0].Item1)));
            for(int i=0;i<sections.Count;i++)
            {
                var summary=sections[i].Item2;int start=sections[i].Item1,end=i+1<sections.Count?sections[i+1].Item1:text.Length;
                var errors=(report.Issues??new List<TEP.Issue>()).Where(e=>e.Severity=="Ошибка"&&TEP.Engine.IssueAffectsMetric(e,summary.Key))
                    .Select(e=>e.Code+": "+e.Message).ToList();
                if(summary.Status=="Неполный результат")errors.Add("Неполный результат");
                if(summary.NotCalculated){errors.Add("не рассчитано");errors.Add("Не рассчитано");}
                AppendReportText(paragraph,text.Substring(start,end-start),errors);
            }
            return document;
        }
        private void ShowIssues(IEnumerable<TEP.Issue> issues){displayedIssues=issues.OrderByDescending(i=>i.Code=="SPATIAL_UNBOUNDED"||i.Code=="ZERO_AREA").ToList();RefreshIssues();}
        private void DiagnosticMode_Changed(object s,SelectionChangedEventArgs e){RefreshIssues();}
        private void RefreshIssues()
        {
            if(IssuesGrid==null||displayedIssues==null)return;
            Config.IssueDetail="full";IssuesGrid.ItemsSource=displayedIssues.OrderByDescending(i=>i.Code=="SPATIAL_UNBOUNDED"||i.Code=="ZERO_AREA").ToList();
        }
        private void ExportRun_Click(object s,RoutedEventArgs e)
        {Try(()=>{if(Engine.Last==null)throw new InvalidOperationException("Сначала выполните расчёт.");var dialog=new SaveFileDialog{Title="KPLN | Экспорт ТЭП",Filter="Книга Excel (*.xlsx)|*.xlsx|CSV, все разделы (*.csv)|*.csv",FileName="ТЭП_"+DateTime.Now.ToString("yyyyMMdd_HHmm")+".xlsx"};if(dialog.ShowDialog(this)!=true)return;
            TEP.Engine.ExportRun(Engine.Last,dialog.FileName,Config.Decimals);Status.Text="Экспортированы 11 разделов: "+dialog.FileName;});}
        private void Cleanup_Click(object s,RoutedEventArgs e)
        {Try(()=>{if(MessageBox.Show(this,"Удалить созданные этой версией плагина планы, 3D-виды, цветовые области, расчётные тела и сводные спецификации? Настройки и последний отчёт останутся в RVT.","KPLN | Удаление результатов ТЭП",MessageBoxButton.YesNo,MessageBoxImage.Question)!=MessageBoxResult.Yes)return;
            Status.Text="Удалено служебных элементов: "+Engine.Cleanup();if(Engine.Last!=null)IssuesGrid.ItemsSource=Engine.Last.Issues.ToList();});}
        private void Navigate_Click(object s,RoutedEventArgs e)
        {
            TEP.Detail detail=null;
            if(ReportTabs.SelectedItem==IssuesTab)
            {
                var issue=IssuesGrid.SelectedItem as TEP.Issue;
                if(issue!=null){var source=Engine.Sources.FirstOrDefault(x=>x.Name==issue.Source);detail=new TEP.Detail{SourceKey=source?.Key,Source=issue.Source,Element=issue.Element};}
            }
            if(detail==null){Status.Text="Выберите строку объекта или ошибки с ElementId.";return;}
            Engine.RequestedDetail=detail;Close();
        }
        private void Issues_DoubleClick(object s,MouseButtonEventArgs e){Navigate_Click(s,e);}
    }
}
