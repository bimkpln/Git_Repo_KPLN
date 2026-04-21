using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    internal static class IDHelper
    {
#if !Revit2024 && !Debug2024
        internal static int InvalidIdInt => ElementId.InvalidElementId.IntegerValue;
        internal static long InvalidIdValue => ElementId.InvalidElementId.IntegerValue;

        internal static int ElIdInt(ElementId id) => id == null ? InvalidIdInt : id.IntegerValue;
        internal static long ElIdValue(ElementId id) => id == null ? InvalidIdValue : id.IntegerValue;

        internal static ElementId ToElementId(int id) => new ElementId(id);
        internal static ElementId ToElementId(long id) => new ElementId((int)id);
#else
        internal static int InvalidIdInt => (int)ElementId.InvalidElementId.Value;
        internal static long InvalidIdValue => ElementId.InvalidElementId.Value;

        internal static int ElIdInt(ElementId id) => id == null ? InvalidIdInt : (int)id.Value;
        internal static long ElIdValue(ElementId id) => id == null ? InvalidIdValue : id.Value;

        internal static ElementId ToElementId(int id) => new ElementId((long)id);
        internal static ElementId ToElementId(long id) => new ElementId(id);
#endif
    }

    internal enum SheetStampMode
    {
        StandardFull,
        Form12SingleRevision,
        RevisionPermissionSingleRevision,
        GipCertificateSingleRevision,
        MultiStampCustom
    }

    internal class FrmIZMFillingViewModel : INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly Dictionary<int, SheetRowItem> _sheetRowCache = new Dictionary<int, SheetRowItem>();
        private readonly List<SignatureComboItem> _signatureDefinitions = new List<SignatureComboItem>();

        public ObservableCollection<SheetTreeNode> RootNodes { get; private set; }
        public ObservableCollection<SheetRowItem> SheetRows { get; private set; }
        public ObservableCollection<RevisionComboItem> RevisionItems { get; private set; }
        public ObservableCollection<SheetStatusItem> StatusItems { get; private set; }
        public ObservableCollection<SheetStatusItem> ManualStatusItems { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler SelectedSheetsChanged;
        public event EventHandler DuplicateRevisionDetected;

        public FrmIZMFillingViewModel(Document doc)
        {
            _doc = doc;
            RootNodes = new ObservableCollection<SheetTreeNode>();
            SheetRows = new ObservableCollection<SheetRowItem>();
            RevisionItems = new ObservableCollection<RevisionComboItem>();
            StatusItems = new ObservableCollection<SheetStatusItem>();
            ManualStatusItems = new ObservableCollection<SheetStatusItem>();

            FillRevisionItems();
            FillStatusItems();
            FillSignatureItems();
            BuildTree();
            RefreshSheetRows();
        }

        private void FillRevisionItems()
        {
            RevisionItems.Clear();
            RevisionItems.Add(new RevisionComboItem(ElementId.InvalidElementId, "*Пусто*", int.MinValue));

            List<Revision> revisions = new FilteredElementCollector(_doc)
                .OfClass(typeof(Revision))
                .Cast<Revision>()
                .OrderByDescending(x => x.SequenceNumber)
                .ToList();

            foreach (Revision revision in revisions)
            {
                string approvedFor = GetRevisionApprovedForDisplay(revision);
                string displayName = "Строка " + revision.SequenceNumber.ToString(CultureInfo.InvariantCulture) + ": ИЗМ № " + approvedFor;
                RevisionItems.Add(new RevisionComboItem(revision.Id, displayName, revision.SequenceNumber));
            }
        }

        private void FillSignatureItems()
        {
            _signatureDefinitions.Clear();
            _signatureDefinitions.Add(new SignatureComboItem(ElementId.InvalidElementId, "*Пусто*"));

            List<FamilySymbol> symbols = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                .Cast<FamilySymbol>()
                .OrderBy(x => GetSignatureDisplayName(x), StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            foreach (FamilySymbol symbol in symbols)
                _signatureDefinitions.Add(new SignatureComboItem(symbol.Id, GetSignatureDisplayName(symbol)));
        }

        private string GetSignatureDisplayName(FamilySymbol symbol)
        {
            if (symbol == null)
                return string.Empty;

            string familyName = string.Empty;
            try
            {
                Parameter familyNameParam = symbol.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                if (familyNameParam != null)
                    familyName = familyNameParam.AsString();
            }
            catch
            {
            }

            string typeName = string.Empty;
            try
            {
                typeName = symbol.Name;
            }
            catch
            {
            }

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
                return familyName + " : " + typeName;

            if (!string.IsNullOrWhiteSpace(typeName))
                return typeName;

            return IDHelper.ElIdInt(symbol.Id).ToString(CultureInfo.InvariantCulture);
        }

        private string GetRevisionApprovedForDisplay(Revision revision)
        {
            if (revision == null)
                return "(без значения)";

            string value = GetParameterString(
                revision,
                "УТВЕРЖДЕНО ДЛЯ",
                "Утверждено для",
                "Approved For",
                "ApprovedFor");

            if (!string.IsNullOrWhiteSpace(value))
                return value;

            if (!string.IsNullOrWhiteSpace(revision.Description))
                return revision.Description;

            return "(без значения)";
        }

        private void FillStatusItems()
        {
            StatusItems.Clear();
            StatusItems.Add(new SheetStatusItem(0, "*Пусто*"));
            StatusItems.Add(new SheetStatusItem(1, "-"));
            StatusItems.Add(new SheetStatusItem(2, "Зам."));
            StatusItems.Add(new SheetStatusItem(3, "Нов."));

            ManualStatusItems.Clear();
            ManualStatusItems.Add(new SheetStatusItem(0, "*Пусто*"));
            ManualStatusItems.Add(new SheetStatusItem(1, "-"));
            ManualStatusItems.Add(new SheetStatusItem(2, "Зам."));
            ManualStatusItems.Add(new SheetStatusItem(3, "Нов."));
            ManualStatusItems.Add(new SheetStatusItem(-1, "Ошибка"));
        }

        private void BuildTree()
        {
            RootNodes.Clear();

            BrowserOrganization browserOrg = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(_doc);

            List<ViewSheet> sheets = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(x => !x.IsPlaceholder)
                .ToList();

            if (browserOrg != null)
            {
                sheets = sheets
                    .Where(x => browserOrg.AreFiltersSatisfied(x.Id))
                    .ToList();
            }

            Dictionary<string, SheetTreeNode> rootMap = new Dictionary<string, SheetTreeNode>();

            foreach (ViewSheet sheet in sheets)
            {
                IList<FolderItemInfo> folders = null;

                if (browserOrg != null)
                    folders = browserOrg.GetFolderItems(sheet.Id);

                SheetTreeNode currentParent = null;
                string pathKey = string.Empty;

                if (folders != null && folders.Count > 0)
                {
                    foreach (FolderItemInfo folder in folders)
                    {
                        pathKey += "|" + folder.Name;

                        SheetTreeNode folderNode;

                        if (currentParent == null)
                        {
                            if (!rootMap.TryGetValue(pathKey, out folderNode))
                            {
                                folderNode = new SheetTreeNode(folder.Name, null, false);
                                folderNode.PropertyChanged += OnTreeNodePropertyChanged;
                                rootMap[pathKey] = folderNode;
                                RootNodes.Add(folderNode);
                            }
                        }
                        else
                        {
                            folderNode = currentParent.Children
                                .FirstOrDefault(x => !x.IsLeaf &&
                                                     string.Equals(x.DisplayName, folder.Name, StringComparison.CurrentCultureIgnoreCase));

                            if (folderNode == null)
                            {
                                folderNode = new SheetTreeNode(folder.Name, null, false);
                                folderNode.Parent = currentParent;
                                folderNode.PropertyChanged += OnTreeNodePropertyChanged;
                                currentParent.Children.Add(folderNode);
                            }
                        }

                        currentParent = folderNode;
                    }
                }

                SheetTreeNode leaf = new SheetTreeNode(
                    sheet.SheetNumber + " - " + sheet.Name,
                    sheet,
                    true);

                leaf.PropertyChanged += OnTreeNodePropertyChanged;

                if (currentParent == null)
                {
                    RootNodes.Add(leaf);
                }
                else
                {
                    leaf.Parent = currentParent;
                    currentParent.Children.Add(leaf);
                }
            }

            SortTreeNodes();
        }

        private void SortTreeNodes()
        {
            List<SheetTreeNode> sortedRoots = RootNodes
                .OrderBy(x => x.DisplayName, new NodeDisplayNameComparer())
                .ToList();

            RootNodes.Clear();

            foreach (SheetTreeNode node in sortedRoots)
            {
                SortNodeChildrenRecursive(node);
                RootNodes.Add(node);
            }
        }

        private void SortNodeChildrenRecursive(SheetTreeNode node)
        {
            if (node == null || node.Children == null || node.Children.Count == 0)
                return;

            List<SheetTreeNode> sortedChildren = node.Children
                .OrderBy(x => x.DisplayName, new NodeDisplayNameComparer())
                .ToList();

            node.Children.Clear();

            foreach (SheetTreeNode child in sortedChildren)
            {
                SortNodeChildrenRecursive(child);
                node.Children.Add(child);
            }
        }

        private void OnTreeNodePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsChecked")
            {
                RefreshSheetRows();

                EventHandler handler = SelectedSheetsChanged;
                if (handler != null)
                    handler(this, EventArgs.Empty);
            }
        }

        public List<ViewSheet> GetSelectedSheets()
        {
            List<ViewSheet> result = new List<ViewSheet>();

            foreach (SheetTreeNode root in RootNodes)
                CollectSelectedSheets(root, result);

            return result;
        }

        private void CollectSelectedSheets(SheetTreeNode node, List<ViewSheet> result)
        {
            if (node.IsLeaf && node.IsChecked == true && node.Sheet != null)
                result.Add(node.Sheet);

            foreach (SheetTreeNode child in node.Children)
                CollectSelectedSheets(child, result);
        }

        public void RefreshSheetRows()
        {
            List<ViewSheet> selectedSheets = GetSelectedSheets();

            SheetRows.Clear();

            foreach (ViewSheet sheet in selectedSheets)
            {
                int key = IDHelper.ElIdInt(sheet.Id);
                SheetRowItem row;

                if (_sheetRowCache.ContainsKey(key))
                {
                    row = _sheetRowCache[key];
                }
                else
                {
                    row = new SheetRowItem(_doc, sheet, RevisionItems.ToList(), _signatureDefinitions);
                    row.DuplicateRevisionDetected += Row_DuplicateRevisionDetected;
                    _sheetRowCache[key] = row;
                }

                SheetRows.Add(row);
            }

            OnPropertyChanged("SheetRows");
        }

        private void Row_DuplicateRevisionDetected(object sender, EventArgs e)
        {
            EventHandler handler = DuplicateRevisionDetected;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        private string GetParameterString(Element element, params string[] parameterNames)
        {
            if (element == null || parameterNames == null)
                return string.Empty;

            foreach (string name in parameterNames)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                Parameter p = element.LookupParameter(name);
                if (p == null)
                    continue;

                string value = GetParameterValueAsString(p);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return string.Empty;
        }

        private string GetParameterValueAsString(Parameter p)
        {
            if (p == null)
                return string.Empty;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                string valueString = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                    return valueString;

                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.Double)
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.ElementId)
                    return IDHelper.ElIdInt(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
            }

            return string.Empty;
        }

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }
    }

    internal class NodeDisplayNameComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            string left = Normalize(x);
            string right = Normalize(y);

            int leftGroup = GetGroup(left);
            int rightGroup = GetGroup(right);

            int groupCompare = leftGroup.CompareTo(rightGroup);
            if (groupCompare != 0)
                return groupCompare;

            return StringComparer.CurrentCultureIgnoreCase.Compare(left, right);
        }

        private string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private int GetGroup(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return 2;

            char first = value[0];

            if (char.IsDigit(first))
                return 0;

            if (char.IsLetter(first))
                return 1;

            return 2;
        }
    }

    public class SheetTreeNode : INotifyPropertyChanged
    {
        private bool? _isChecked;
        private bool _isInternalUpdate;

        public string DisplayName { get; set; }
        public ViewSheet Sheet { get; set; }
        public bool IsLeaf { get; set; }
        public bool IsFolder { get { return !IsLeaf; } }
        public SheetTreeNode Parent { get; set; }
        public ObservableCollection<SheetTreeNode> Children { get; set; }

        public bool? IsChecked
        {
            get { return _isChecked; }
            set
            {
                bool? normalizedValue = value;

                if (IsLeaf && normalizedValue == null)
                    normalizedValue = false;

                if (!IsLeaf && !_isInternalUpdate && normalizedValue == null)
                    normalizedValue = false;

                if (_isChecked == normalizedValue)
                    return;

                _isChecked = normalizedValue;
                OnPropertyChanged("IsChecked");

                if (_isInternalUpdate)
                    return;

                if (!IsLeaf)
                {
                    bool childValue = normalizedValue == true;

                    foreach (SheetTreeNode child in Children)
                        child.SetCheckedFromParent(childValue);
                }

                UpdateParentState();
            }
        }

        public SheetTreeNode(string displayName, ViewSheet sheet, bool isLeaf)
        {
            DisplayName = displayName;
            Sheet = sheet;
            IsLeaf = isLeaf;
            Parent = null;
            Children = new ObservableCollection<SheetTreeNode>();
            _isChecked = false;
            _isInternalUpdate = false;
        }

        private void SetCheckedFromParent(bool isChecked)
        {
            _isInternalUpdate = true;
            _isChecked = isChecked;
            OnPropertyChanged("IsChecked");
            _isInternalUpdate = false;

            foreach (SheetTreeNode child in Children)
                child.SetCheckedFromParent(isChecked);
        }

        private void UpdateParentState()
        {
            if (Parent == null)
                return;

            if (Parent.Children == null || Parent.Children.Count == 0)
            {
                Parent._isInternalUpdate = true;
                Parent._isChecked = false;
                Parent.OnPropertyChanged("IsChecked");
                Parent._isInternalUpdate = false;
                Parent.UpdateParentState();
                return;
            }

            bool allChecked = Parent.Children.All(x => x.IsChecked == true);
            bool allUnchecked = Parent.Children.All(x => x.IsChecked == false);

            Parent._isInternalUpdate = true;

            if (allChecked)
                Parent._isChecked = true;
            else if (allUnchecked)
                Parent._isChecked = false;
            else
                Parent._isChecked = null;

            Parent.OnPropertyChanged("IsChecked");
            Parent._isInternalUpdate = false;

            Parent.UpdateParentState();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }
    }

    internal class RevisionComboItem : INotifyPropertyChanged
    {
        private bool _isEnabled;

        public ElementId Id { get; set; }
        public int IdValue { get; set; }
        public string Name { get; set; }
        public int SequenceNumber { get; set; }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (_isEnabled == value)
                    return;

                _isEnabled = value;
                OnPropertyChanged("IsEnabled");
            }
        }

        public RevisionComboItem(ElementId id, string name, int sequenceNumber)
        {
            Id = id ?? ElementId.InvalidElementId;
            IdValue = IDHelper.ElIdInt(Id);
            Name = name;
            SequenceNumber = sequenceNumber;
            _isEnabled = true;
        }

        public RevisionComboItem Clone()
        {
            return new RevisionComboItem(Id, Name, SequenceNumber)
            {
                IsEnabled = IsEnabled
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }
    }

    internal class SignatureComboItem
    {
        public ElementId Id { get; set; }
        public int IdValue { get; set; }
        public string Name { get; set; }

        public SignatureComboItem(ElementId id, string name)
        {
            Id = id ?? ElementId.InvalidElementId;
            IdValue = IDHelper.ElIdInt(Id);
            Name = name ?? string.Empty;
        }

        public SignatureComboItem Clone()
        {
            return new SignatureComboItem(Id, Name);
        }
    }

    internal class SheetStatusItem
    {
        public int Code { get; set; }
        public string Name { get; set; }

        public SheetStatusItem(int code, string name)
        {
            Code = code;
            Name = name;
        }
    }

    internal class SheetRowItem : INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly List<RevisionComboItem> _revisionDefinitions;
        private readonly List<SignatureComboItem> _signatureDefinitions;

        private const string _form12TitleBlockName = "020_Основная надпись_Форма_12";
        private const string _revisionPermissionTitleBlockName = "020_Разрешение на внесение изменений";
        private const string _gipCertificateTitleBlockName = "020_Справка_ГИП";

        public ViewSheet Sheet { get; private set; }
        public string SheetNumber { get; private set; }
        public string SheetName { get; private set; }

        public ObservableCollection<SheetStampBlockItem> Blocks { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler DuplicateRevisionDetected;

        public SheetRowItem(
            Document doc,
            ViewSheet sheet,
            List<RevisionComboItem> revisionDefinitions,
            List<SignatureComboItem> signatureDefinitions)
        {
            _doc = doc;
            _revisionDefinitions = revisionDefinitions ?? new List<RevisionComboItem>();
            _signatureDefinitions = signatureDefinitions ?? new List<SignatureComboItem>();

            Sheet = sheet;
            SheetNumber = sheet.SheetNumber;
            SheetName = sheet.Name;

            Blocks = new ObservableCollection<SheetStampBlockItem>();

            BuildBlocks();
        }

        private void BuildBlocks()
        {
            Blocks.Clear();

            List<FamilyInstance> titleBlocks = GetTitleBlocksOnSheet();

            if (titleBlocks.Count <= 1)
            {
                FamilyInstance titleBlock = titleBlocks.FirstOrDefault();

                SheetStampMode mode = GetSingleStampMode(titleBlock);
                int rowCount = mode == SheetStampMode.StandardFull
                    ? DetectRowCount(titleBlock)
                    : 1;

                SheetStampBlockItem block = new SheetStampBlockItem(
                    _doc,
                    Sheet,
                    titleBlock,
                    titleBlock != null ? GetTitleBlockHeader(titleBlock) : string.Empty,
                    mode,
                    rowCount,
                    _revisionDefinitions,
                    _signatureDefinitions);

                block.ShowHeader = false;
                block.DuplicateRevisionDetected += Block_DuplicateRevisionDetected;

                Blocks.Add(block);

                OnPropertyChanged("Blocks");
                return;
            }

            foreach (FamilyInstance titleBlock in titleBlocks)
            {
                int rowCount = DetectRowCount(titleBlock);

                SheetStampBlockItem block = new SheetStampBlockItem(
                    _doc,
                    Sheet,
                    titleBlock,
                    GetTitleBlockHeader(titleBlock),
                    SheetStampMode.MultiStampCustom,
                    rowCount,
                    _revisionDefinitions,
                    _signatureDefinitions);

                block.ShowHeader = true;
                block.DuplicateRevisionDetected += Block_DuplicateRevisionDetected;

                Blocks.Add(block);
            }

            OnPropertyChanged("Blocks");
        }

        private void Block_DuplicateRevisionDetected(object sender, EventArgs e)
        {
            EventHandler handler = DuplicateRevisionDetected;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        public void ApplyToSheet()
        {
            foreach (SheetStampBlockItem block in Blocks)
                block.Apply();
        }

        public bool HasDuplicateRevisions()
        {
            return Blocks.Any(x => x.HasDuplicateRevisions());
        }

        private List<FamilyInstance> GetTitleBlocksOnSheet()
        {
            return new FilteredElementCollector(_doc, Sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .OrderBy(x => GetTitleBlockHeader(x), StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private SheetStampMode GetSingleStampMode(FamilyInstance titleBlock)
        {
            if (MatchesTitleBlock(titleBlock, _form12TitleBlockName))
                return SheetStampMode.Form12SingleRevision;

            if (MatchesTitleBlock(titleBlock, _revisionPermissionTitleBlockName))
                return SheetStampMode.RevisionPermissionSingleRevision;

            if (MatchesTitleBlock(titleBlock, _gipCertificateTitleBlockName))
                return SheetStampMode.GipCertificateSingleRevision;

            return SheetStampMode.StandardFull;
        }

        private bool MatchesTitleBlock(FamilyInstance titleBlock, string targetName)
        {
            if (titleBlock == null || string.IsNullOrWhiteSpace(targetName))
                return false;

            string familyName = GetTitleBlockFamilyName(titleBlock);
            string typeName = GetTitleBlockTypeName(titleBlock);

            if (string.Equals(familyName, targetName, StringComparison.CurrentCultureIgnoreCase))
                return true;

            if (string.Equals(typeName, targetName, StringComparison.CurrentCultureIgnoreCase))
                return true;

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
            {
                string combined = familyName + " : " + typeName;
                if (string.Equals(combined, targetName, StringComparison.CurrentCultureIgnoreCase))
                    return true;
            }

            return false;
        }

        private int DetectRowCount(FamilyInstance titleBlock)
        {
            if (HasTitleBlockParameterForRow4(titleBlock))
                return 4;

            return 2;
        }

        private bool HasTitleBlockParameterForRow4(FamilyInstance titleBlock)
        {
            if (titleBlock == null)
                return false;

            Parameter p = titleBlock.LookupParameter("Изм4 подпись<Типовые аннотации>");
            if (p != null)
                return true;

            p = titleBlock.LookupParameter("Изм4 подпись");
            return p != null;
        }

        private string GetTitleBlockHeader(FamilyInstance titleBlock)
        {
            if (titleBlock == null)
                return string.Empty;

            string manualNumber = GetParameterString(titleBlock, "Номер листа вручную");
            if (!string.IsNullOrWhiteSpace(manualNumber))
                return manualNumber;

            string familyName = GetTitleBlockFamilyName(titleBlock);
            if (!string.IsNullOrWhiteSpace(familyName))
                return familyName;

            return "Основная надпись " + IDHelper.ElIdInt(titleBlock.Id).ToString(CultureInfo.InvariantCulture);
        }

        private string GetTitleBlockFamilyName(FamilyInstance titleBlock)
        {
            if (titleBlock == null)
                return string.Empty;

            Element typeElement = _doc.GetElement(titleBlock.GetTypeId());
            FamilySymbol symbol = typeElement as FamilySymbol;
            if (symbol == null)
                return string.Empty;

            try
            {
                Parameter p = symbol.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);
                if (p != null)
                {
                    string familyName = p.AsString();
                    if (!string.IsNullOrWhiteSpace(familyName))
                        return familyName;
                }
            }
            catch
            {
            }

            try
            {
                if (symbol.Family != null && !string.IsNullOrWhiteSpace(symbol.Family.Name))
                    return symbol.Family.Name;
            }
            catch
            {
            }

            return string.Empty;
        }

        private string GetTitleBlockTypeName(FamilyInstance titleBlock)
        {
            if (titleBlock == null)
                return string.Empty;

            Element typeElement = _doc.GetElement(titleBlock.GetTypeId());
            FamilySymbol symbol = typeElement as FamilySymbol;
            if (symbol == null)
                return string.Empty;

            try
            {
                return symbol.Name ?? string.Empty;
            }
            catch
            {
            }

            return string.Empty;
        }

        private string GetParameterString(Element element, params string[] parameterNames)
        {
            if (element == null || parameterNames == null)
                return string.Empty;

            foreach (string name in parameterNames)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                Parameter p = element.LookupParameter(name);
                if (p == null)
                    continue;

                string value = GetParameterValueAsString(p);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return string.Empty;
        }

        private string GetParameterValueAsString(Parameter p)
        {
            if (p == null)
                return string.Empty;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                string valueString = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                    return valueString;

                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.Double)
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.ElementId)
                    return IDHelper.ElIdInt(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
            }

            return string.Empty;
        }

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }
    }

    internal class SheetStampBlockItem : INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly ViewSheet _sheet;
        private readonly FamilyInstance _titleBlock;
        private readonly List<RevisionComboItem> _revisionDefinitions;
        private readonly List<SignatureComboItem> _signatureDefinitions;

        private bool _isSorting;
        private bool _isResettingDuplicateRevision;
        private bool _isRebuildingRevisionUi;
        private bool _manualEditEnabled;
        private bool _manualDocDataEnabled;

        public string Header { get; private set; }
        public SheetStampMode Mode { get; private set; }
        public bool ShowHeader { get; set; }

        public bool ManualEditEnabled
        {
            get { return _manualEditEnabled; }
            set
            {
                if (_manualEditEnabled == value)
                    return;

                _manualEditEnabled = value;
                OnPropertyChanged("ManualEditEnabled");
            }
        }

        public bool ManualDocDataEnabled
        {
            get { return _manualDocDataEnabled; }
            set
            {
                if (_manualDocDataEnabled == value)
                    return;

                _manualDocDataEnabled = value;
                OnPropertyChanged("ManualDocDataEnabled");
            }
        }

        public bool IsStandardFull { get { return Mode == SheetStampMode.StandardFull; } }

        public bool IsForm12SingleRevision
        {
            get { return Mode == SheetStampMode.Form12SingleRevision; }
        }

        public bool IsRevisionPermissionSingleRevision
        {
            get { return Mode == SheetStampMode.RevisionPermissionSingleRevision; }
        }

        public bool IsGipCertificateSingleRevision
        {
            get { return Mode == SheetStampMode.GipCertificateSingleRevision; }
        }

        public bool IsSingleRevisionOnly
        {
            get
            {
                return Mode == SheetStampMode.Form12SingleRevision
                    || Mode == SheetStampMode.RevisionPermissionSingleRevision
                    || Mode == SheetStampMode.GipCertificateSingleRevision;
            }
        }

        public bool IsMultiStampCustom
        {
            get { return Mode == SheetStampMode.MultiStampCustom; }
        }

        public ObservableCollection<SheetRevisionLine> Lines { get; private set; }

        public SheetRevisionLine SingleLine
        {
            get { return Lines.FirstOrDefault(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler DuplicateRevisionDetected;

        public SheetStampBlockItem(
            Document doc,
            ViewSheet sheet,
            FamilyInstance titleBlock,
            string header,
            SheetStampMode mode,
            int rowCount,
            List<RevisionComboItem> revisionDefinitions,
            List<SignatureComboItem> signatureDefinitions)
        {
            _doc = doc;
            _sheet = sheet;
            _titleBlock = titleBlock;
            _revisionDefinitions = revisionDefinitions ?? new List<RevisionComboItem>();
            _signatureDefinitions = signatureDefinitions ?? new List<SignatureComboItem>();

            Header = header ?? string.Empty;
            Mode = mode;
            ShowHeader = false;
            ManualEditEnabled = false;
            ManualDocDataEnabled = false;

            if (rowCount <= 0)
                rowCount = 1;

            Lines = new ObservableCollection<SheetRevisionLine>();

            for (int displayNumber = rowCount; displayNumber >= 1; displayNumber--)
                Lines.Add(new SheetRevisionLine(_doc, displayNumber, _revisionDefinitions, _signatureDefinitions));

            foreach (SheetRevisionLine line in Lines)
                line.PropertyChanged += Line_PropertyChanged;

            RefreshDisplayNumbers();

            if (Mode == SheetStampMode.StandardFull)
            {
                RebuildRevisionUi();
                LoadStandardFull();
            }
            else if (IsSingleRevisionOnly)
            {
                RebuildRevisionUi();
                LoadSingleRevisionOnly();
            }
            else if (Mode == SheetStampMode.MultiStampCustom)
            {
                LoadMultiStampCustom();
            }
        }

        private void Line_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "RevisionIdValue")
                return;

            if (!IsStandardFull)
                return;

            if (_isSorting || _isResettingDuplicateRevision || _isRebuildingRevisionUi)
                return;

            SheetRevisionLine changedLine = sender as SheetRevisionLine;
            if (changedLine == null)
                return;

            if (HasDuplicateRevisions())
            {
                _isResettingDuplicateRevision = true;
                changedLine.RevertRevisionSelection();
                _isResettingDuplicateRevision = false;

                RebuildRevisionUi();

                EventHandler handler = DuplicateRevisionDetected;
                if (handler != null)
                    handler(this, EventArgs.Empty);

                return;
            }

            SortLinesByRevisionOrder();
        }

        public bool HasDuplicateRevisions()
        {
            if (!IsStandardFull)
                return false;

            List<int> ids = Lines
                .Where(x => IsValidRevisionId(x.RevisionId))
                .Select(x => IDHelper.ElIdInt(x.RevisionId))
                .ToList();

            return ids.Count != ids.Distinct().Count();
        }

        public void Apply()
        {
            if (IsStandardFull)
                ApplyStandardFull();
            else if (IsSingleRevisionOnly)
                ApplySingleRevisionOnly();
            else if (IsMultiStampCustom)
                ApplyMultiStampCustom();
        }

        private void LoadStandardFull()
        {
            List<Revision> revisions = _sheet.GetAdditionalRevisionIds()
                .Select(x => _doc.GetElement(x) as Revision)
                .Where(x => x != null)
                .OrderByDescending(x => x.SequenceNumber)
                .ToList();

            _isSorting = true;
            try
            {
                for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
                    Lines[rowIndex].ResetAllData();

                for (int i = 0; i < revisions.Count && i < Lines.Count; i++)
                    Lines[i].RevisionId = revisions[i].Id;
            }
            finally
            {
                _isSorting = false;
            }

            SortLinesByRevisionOrder();
            LoadStaticColumnsFromSheet();
            CommitAllRevisionSelections();
        }

        private void LoadSingleRevisionOnly()
        {
            foreach (SheetRevisionLine line in Lines)
                line.ResetAllData();

            Revision revision = _sheet.GetAdditionalRevisionIds()
                .Select(x => _doc.GetElement(x) as Revision)
                .Where(x => x != null)
                .OrderByDescending(x => x.SequenceNumber)
                .FirstOrDefault();

            if (revision != null && SingleLine != null)
                SingleLine.RevisionId = revision.Id;

            CommitAllRevisionSelections();
        }

        private void LoadMultiStampCustom()
        {
            ManualEditEnabled = GetBoolFromTitleBlock(_titleBlock, "Изм_Вручную_Вкл");
            ManualDocDataEnabled = GetBoolFromTitleBlock(_titleBlock, "ИзмДокДата_Вкл_Вручную");

            foreach (SheetRevisionLine line in Lines)
            {
                line.ResetAllData();
                line.SignatureTypeId = GetSignatureTypeIdFromTitleBlock(_titleBlock, line.DisplayNumber);
                line.QuantityText = GetParameterString(_titleBlock, GetManualQtyParamName(line.DisplayNumber));
                line.StatusCode = GetStatusCodeByManualValue(GetParameterString(_titleBlock, GetManualSheetParamName(line.DisplayNumber)));
                line.ManualRevisionText = GetParameterString(_titleBlock, GetManualRevisionParamName(line.DisplayNumber));
                line.ManualDocNumberText = GetParameterString(_titleBlock, GetManualDocParamName(line.DisplayNumber));
                line.ManualDateText = GetParameterString(_titleBlock, GetManualDateParamName(line.DisplayNumber));
            }
        }

        private void ApplyStandardFull()
        {
            SortLinesByRevisionOrder();

            List<ElementId> selectedRevisionIds = GetSelectedRevisionIds();
            _sheet.SetAdditionalRevisionIds(selectedRevisionIds);

            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int mirroredSlotNumber = GetMirroredSlotNumber(rowIndex);

                SetParameterString(_sheet, GetQtyParamName(mirroredSlotNumber), line.QuantityText);
            }

            SetStatusStringToSheet();

            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int displayNumber = line.DisplayNumber;

                SetSignatureTypeIdToTitleBlock(_titleBlock, displayNumber, line.SignatureTypeId);
            }
        }

        private void ApplySingleRevisionOnly()
        {
            List<ElementId> ids = new List<ElementId>();

            if (SingleLine != null && IsValidRevisionId(SingleLine.RevisionId))
                ids.Add(SingleLine.RevisionId);

            _sheet.SetAdditionalRevisionIds(ids);
        }

        private void ApplyMultiStampCustom()
        {
            SetBoolToTitleBlock(_titleBlock, "КолУчЛист_Вручную_Вкл", ManualEditEnabled);
            SetBoolToTitleBlock(_titleBlock, "ИзмДокДата_Вкл_Вручную", ManualDocDataEnabled);

            foreach (SheetRevisionLine line in Lines)
            {
                SetSignatureTypeIdToTitleBlock(_titleBlock, line.DisplayNumber, line.SignatureTypeId);

                if (ManualEditEnabled)
                {
                    SetParameterString(_titleBlock, GetManualQtyParamName(line.DisplayNumber), line.QuantityText);

                    if (line.StatusCode != -1)
                        SetParameterString(_titleBlock, GetManualSheetParamName(line.DisplayNumber), GetManualSheetValueByStatusCode(line.StatusCode));
                }

                if (ManualDocDataEnabled)
                {
                    SetParameterString(_titleBlock, GetManualRevisionParamName(line.DisplayNumber), line.ManualRevisionText);
                    SetParameterString(_titleBlock, GetManualDocParamName(line.DisplayNumber), line.ManualDocNumberText);
                    SetParameterString(_titleBlock, GetManualDateParamName(line.DisplayNumber), line.ManualDateText);
                }
            }
        }

        public void SortLinesByRevisionOrder()
        {
            if (!IsStandardFull)
                return;

            _isSorting = true;

            try
            {
                List<RevisionSortableState> sortedStates = Lines
                    .Select((x, index) => new RevisionSortableState
                    {
                        RevisionId = x.RevisionId,
                        SortWeight = GetVisualSortWeight(x),
                        SequenceNumber = GetRevisionSequence(x.RevisionId),
                        OriginalIndex = index
                    })
                    .OrderBy(x => x.SortWeight)
                    .ThenByDescending(x => x.SequenceNumber)
                    .ThenBy(x => x.OriginalIndex)
                    .ToList();

                for (int i = 0; i < Lines.Count; i++)
                {
                    RevisionSortableState state = sortedStates[i];
                    Lines[i].ApplyRevisionState(state.RevisionId);
                }
            }
            finally
            {
                _isSorting = false;
            }

            RefreshDisplayNumbers();
            RebuildRevisionUi();
            CommitAllRevisionSelections();
            OnPropertyChanged("Lines");
        }

        public List<ElementId> GetSelectedRevisionIds()
        {
            return Lines
                .Where(x => IsValidRevisionId(x.RevisionId))
                .Select(x => x.RevisionId)
                .Distinct(new ElementIdEqualityComparer())
                .OrderByDescending(x => GetRevisionSequence(x))
                .ToList();
        }

        private void CommitAllRevisionSelections()
        {
            foreach (SheetRevisionLine line in Lines)
                line.CommitRevisionSelection();
        }

        private void RebuildRevisionUi()
        {
            if (!IsStandardFull && !IsSingleRevisionOnly)
                return;

            _isRebuildingRevisionUi = true;

            try
            {
                foreach (SheetRevisionLine line in Lines)
                    line.RebuildAvailableRevisionItems(_revisionDefinitions);

                foreach (SheetRevisionLine currentLine in Lines)
                {
                    HashSet<int> selectedByOtherLines = new HashSet<int>(
                        Lines.Where(x => !ReferenceEquals(x, currentLine) && IsValidRevisionId(x.RevisionId))
                             .Select(x => IDHelper.ElIdInt(x.RevisionId)));

                    currentLine.ApplyRevisionAvailability(selectedByOtherLines);
                }

                foreach (SheetRevisionLine line in Lines)
                    line.SyncSelectedRevisionItem();
            }
            finally
            {
                _isRebuildingRevisionUi = false;
            }
        }

        private int GetVisualSortWeight(SheetRevisionLine line)
        {
            if (line == null || !line.HasRevisionSelected)
                return 0;

            return 1;
        }

        private int GetRevisionSequence(ElementId revisionId)
        {
            if (revisionId == null || revisionId == ElementId.InvalidElementId)
                return int.MinValue;

            Revision revision = _doc.GetElement(revisionId) as Revision;
            if (revision == null)
                return int.MinValue;

            return revision.SequenceNumber;
        }

        private void RefreshDisplayNumbers()
        {
            for (int i = 0; i < Lines.Count; i++)
                Lines[i].SetDisplayNumber(Lines.Count - i);
        }

        private int GetMirroredSlotNumber(int rowIndex)
        {
            return Lines.Count - rowIndex;
        }

        private SheetRevisionLine GetLineByMirroredSlotNumber(int slotNumber)
        {
            int rowIndex = Lines.Count - slotNumber;

            if (rowIndex < 0 || rowIndex >= Lines.Count)
                return null;

            return Lines[rowIndex];
        }

        private string GetQtyParamName(int slotNumber)
        {
            switch (slotNumber)
            {
                case 1: return "Ш.КолвоУч1Текст";
                case 2: return "Ш.КолвоУч2Текст";
                case 3: return "Ш.КолвоУч3Текст";
                case 4: return "Ш.КолвоУч4Текст";
                default: return string.Empty;
            }
        }

        private string GetManualQtyParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 1: return "КолУч1_Вручную";
                case 2: return "КолУч2_Вручную";
                case 3: return "КолУч3_Вручную";
                case 4: return "КолУч4_Вручную";
                default: return string.Empty;
            }
        }

        private string GetManualSheetParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 1: return "Лист1_Вручную";
                case 2: return "Лист2_Вручную";
                case 3: return "Лист3_Вручную";
                case 4: return "Лист4_Вручную";
                default: return string.Empty;
            }
        }

        private string GetManualRevisionParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 1: return "№Изм1_Вручную";
                case 2: return "№Изм2_Вручную";
                case 3: return "№Изм3_Вручную";
                case 4: return "№Изм4_Вручную";
                default: return string.Empty;
            }
        }

        private string GetManualDocParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 1: return "Док1_Вручную";
                case 2: return "Док2_Вручную";
                case 3: return "Док3_Вручную";
                case 4: return "Док4_Вручную";
                default: return string.Empty;
            }
        }

        private string GetManualDateParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 1: return "Дата1_Вручную";
                case 2: return "Дата2_Вручную";
                case 3: return "Дата3_Вручную";
                case 4: return "Дата4_Вручную";
                default: return string.Empty;
            }
        }

        private string GetManualSheetValueByStatusCode(int statusCode)
        {
            switch (statusCode)
            {
                case 1: return "-";
                case 2: return "Зам.";
                case 3: return "Нов.";
                case 0: return string.Empty;
                default: return null;
            }
        }

        private int GetStatusCodeByManualValue(string value)
        {
            string normalized = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

            if (string.IsNullOrWhiteSpace(normalized))
                return 0;

            if (string.Equals(normalized, "-", StringComparison.CurrentCultureIgnoreCase))
                return 1;

            if (string.Equals(normalized, "Зам.", StringComparison.CurrentCultureIgnoreCase))
                return 2;

            if (string.Equals(normalized, "Нов.", StringComparison.CurrentCultureIgnoreCase))
                return 3;

            return -1;
        }

        private bool GetBoolFromTitleBlock(FamilyInstance titleBlock, string parameterName)
        {
            if (titleBlock == null || string.IsNullOrWhiteSpace(parameterName))
                return false;

            Parameter p = titleBlock.LookupParameter(parameterName);
            if (p == null)
                return false;

            try
            {
                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger() != 0;

                if (p.StorageType == StorageType.String)
                {
                    string value = p.AsString();
                    if (string.IsNullOrWhiteSpace(value))
                        return false;

                    value = value.Trim();

                    if (string.Equals(value, "1", StringComparison.CurrentCultureIgnoreCase))
                        return true;
                    if (string.Equals(value, "true", StringComparison.CurrentCultureIgnoreCase))
                        return true;
                    if (string.Equals(value, "да", StringComparison.CurrentCultureIgnoreCase))
                        return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private void SetBoolToTitleBlock(FamilyInstance titleBlock, string parameterName, bool value)
        {
            if (titleBlock == null || string.IsNullOrWhiteSpace(parameterName))
                return;

            Parameter p = titleBlock.LookupParameter(parameterName);
            if (p == null || p.IsReadOnly)
                return;

            try
            {
                if (p.StorageType == StorageType.Integer)
                {
                    p.Set(value ? 1 : 0);
                    return;
                }

                if (p.StorageType == StorageType.String)
                {
                    p.Set(value ? "1" : "0");
                }
            }
            catch
            {
            }
        }

        private string GetTitleBlockSignatureParamName(int displayNumber)
        {
            switch (displayNumber)
            {
                case 4: return "Изм4 подпись";
                case 3: return "Изм3 подпись";
                case 2: return "Изм2 подпись";
                case 1: return "Изм1 подпись";
                default: return string.Empty;
            }
        }

        private Parameter LookupTitleBlockSignatureParameter(FamilyInstance titleBlock, int displayNumber)
        {
            if (titleBlock == null)
                return null;

            string baseName = GetTitleBlockSignatureParamName(displayNumber);
            if (string.IsNullOrWhiteSpace(baseName))
                return null;

            Parameter p = titleBlock.LookupParameter(baseName + "<Типовые аннотации>");
            if (p != null)
                return p;

            p = titleBlock.LookupParameter(baseName);
            if (p != null)
                return p;

            return null;
        }

        private ElementId GetSignatureTypeIdFromTitleBlock(FamilyInstance titleBlock, int displayNumber)
        {
            Parameter p = LookupTitleBlockSignatureParameter(titleBlock, displayNumber);
            if (p == null)
                return ElementId.InvalidElementId;

            try
            {
                if (p.StorageType == StorageType.ElementId)
                {
                    ElementId id = p.AsElementId();
                    return id ?? ElementId.InvalidElementId;
                }

                if (p.StorageType == StorageType.String)
                {
                    string textValue = p.AsString();
                    SignatureComboItem item = _signatureDefinitions
                        .FirstOrDefault(x => string.Equals(x.Name, textValue, StringComparison.CurrentCultureIgnoreCase));

                    return item != null ? item.Id : ElementId.InvalidElementId;
                }
            }
            catch
            {
            }

            return ElementId.InvalidElementId;
        }

        private void SetSignatureTypeIdToTitleBlock(FamilyInstance titleBlock, int displayNumber, ElementId signatureTypeId)
        {
            Parameter p = LookupTitleBlockSignatureParameter(titleBlock, displayNumber);
            if (p == null || p.IsReadOnly)
                return;

            ElementId valueToSet = signatureTypeId ?? ElementId.InvalidElementId;

            try
            {
                if (p.StorageType == StorageType.ElementId)
                {
                    p.Set(valueToSet);
                    return;
                }

                if (p.StorageType == StorageType.String)
                {
                    SignatureComboItem item = _signatureDefinitions
                        .FirstOrDefault(x => x.IdValue == IDHelper.ElIdInt(valueToSet));

                    p.Set(item != null ? item.Name : string.Empty);
                }
            }
            catch
            {
            }
        }

        private void LoadStaticColumnsFromSheet()
        {
            int[] statusDigits = GetStatusDigitsFromSheet();

            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int mirroredSlotNumber = GetMirroredSlotNumber(rowIndex);

                line.QuantityText = GetParameterString(_sheet, GetQtyParamName(mirroredSlotNumber));
                line.StatusCode = GetStatusCodeBySlotNumber(statusDigits, mirroredSlotNumber);

                int displayNumber = line.DisplayNumber;
                line.SignatureTypeId = GetSignatureTypeIdFromTitleBlock(_titleBlock, displayNumber);
            }
        }

        private int[] GetStatusDigitsFromSheet()
        {
            string value = GetParameterString(_sheet, "Ш.ШифрСтатусЛиста");

            if (Lines.Count == 4)
                return ParseStatusDigits(value, 4);

            return ParseStatusDigits(value, 2);
        }

        private int[] ParseStatusDigits(string input, int expectedCount)
        {
            int[] result = new int[expectedCount];

            for (int i = 0; i < result.Length; i++)
                result[i] = 0;

            if (string.IsNullOrWhiteSpace(input))
                return result;

            string digitsOnly = new string(input.Where(char.IsDigit).ToArray());

            for (int i = 0; i < result.Length; i++)
            {
                if (i < digitsOnly.Length)
                {
                    int digit;
                    if (int.TryParse(digitsOnly[i].ToString(), out digit) && digit >= 0 && digit <= 3)
                        result[i] = digit;
                }
            }

            return result;
        }

        private int GetStatusCodeBySlotNumber(int[] digits, int slotNumber)
        {
            if (digits == null || digits.Length == 0)
                return 0;

            int index = slotNumber - 1;

            if (index < 0 || index >= digits.Length)
                return 0;

            return digits[index];
        }

        private void SetStatusStringToSheet()
        {
            StringBuilder statusBuilder = new StringBuilder();

            for (int slotNumber = 1; slotNumber <= Lines.Count; slotNumber++)
            {
                SheetRevisionLine line = GetLineByMirroredSlotNumber(slotNumber);
                int status = line != null ? line.StatusCode : 0;

                if (status < 0 || status > 3)
                    status = 0;

                statusBuilder.Append(status.ToString(CultureInfo.InvariantCulture));
            }

            SetParameterString(_sheet, "Ш.ШифрСтатусЛиста", statusBuilder.ToString());
        }

        private bool IsValidRevisionId(ElementId revisionId)
        {
            if (revisionId == null)
                return false;

            if (revisionId == ElementId.InvalidElementId)
                return false;

            if (IDHelper.ElIdInt(revisionId) <= 0)
                return false;

            return _doc.GetElement(revisionId) is Revision;
        }

        private string GetParameterString(Element element, string parameterName)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
                return string.Empty;

            Parameter p = element.LookupParameter(parameterName);
            if (p == null)
                return string.Empty;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                string valueString = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                    return valueString;

                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.Double)
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.ElementId)
                    return IDHelper.ElIdInt(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
            }

            return string.Empty;
        }

        private void SetParameterString(Element element, string parameterName, string value)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
                return;

            Parameter p = element.LookupParameter(parameterName);
            if (p == null)
                return;

            if (p.IsReadOnly)
                return;

            try
            {
                if (p.StorageType == StorageType.String)
                {
                    p.Set(value ?? string.Empty);
                }
                else if (p.StorageType == StorageType.Integer)
                {
                    int intValue = 0;
                    int.TryParse(value, out intValue);
                    p.Set(intValue);
                }
            }
            catch
            {
            }
        }

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }

        private class RevisionSortableState
        {
            public ElementId RevisionId { get; set; }
            public int SortWeight { get; set; }
            public int SequenceNumber { get; set; }
            public int OriginalIndex { get; set; }
        }
    }

    internal class SheetRevisionLine : INotifyPropertyChanged
    {
        private readonly Document _doc;

        private int _displayNumber;
        private ElementId _revisionId;
        private int _revisionIdValue;
        private int _lastCommittedRevisionIdValue;
        private string _quantityText;
        private int _statusCode;
        private string _revisionDocNumber;
        private string _revisionDate;

        private ElementId _signatureTypeId;
        private int _signatureTypeIdValue;

        private string _manualRevisionText;
        private string _manualDocNumberText;
        private string _manualDateText;

        private RevisionComboItem _selectedRevisionItem;
        private SignatureComboItem _selectedSignatureItem;

        private bool _isSyncingSelectedItem;
        private bool _ignoreSelectedRevisionItemChanges;

        private bool _isSyncingSelectedSignatureItem;
        private bool _ignoreSelectedSignatureItemChanges;

        public ObservableCollection<RevisionComboItem> AvailableRevisionItems { get; private set; }
        public ObservableCollection<SignatureComboItem> AvailableSignatureItems { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public int DisplayNumber
        {
            get { return _displayNumber; }
            private set
            {
                if (_displayNumber == value)
                    return;

                _displayNumber = value;
                OnPropertyChanged("DisplayNumber");
            }
        }

        public bool HasRevisionSelected
        {
            get
            {
                return IsValidRevisionId(_revisionId);
            }
        }

        public RevisionComboItem SelectedRevisionItem
        {
            get { return _selectedRevisionItem; }
            set
            {
                if (_ignoreSelectedRevisionItemChanges || _isSyncingSelectedItem)
                    return;

                if (ReferenceEquals(_selectedRevisionItem, value))
                    return;

                if (value != null && !value.IsEnabled)
                {
                    SyncSelectedRevisionItem();
                    return;
                }

                _selectedRevisionItem = value;
                OnPropertyChanged("SelectedRevisionItem");

                int newIdValue = value != null
                    ? value.IdValue
                    : IDHelper.InvalidIdInt;

                ApplyRevisionIdInternal(newIdValue, true);
            }
        }

        public SignatureComboItem SelectedSignatureItem
        {
            get { return _selectedSignatureItem; }
            set
            {
                if (_ignoreSelectedSignatureItemChanges || _isSyncingSelectedSignatureItem)
                    return;

                if (ReferenceEquals(_selectedSignatureItem, value))
                    return;

                _selectedSignatureItem = value;
                OnPropertyChanged("SelectedSignatureItem");

                int newIdValue = value != null
                    ? value.IdValue
                    : IDHelper.InvalidIdInt;

                ApplySignatureTypeIdInternal(newIdValue, true);
            }
        }

        public ElementId RevisionId
        {
            get { return _revisionId; }
            set
            {
                int newIdValue = value != null
                    ? IDHelper.ElIdInt(value)
                    : IDHelper.InvalidIdInt;

                ApplyRevisionIdInternal(newIdValue, true);
            }
        }

        public int RevisionIdValue
        {
            get { return _revisionIdValue; }
            set
            {
                ApplyRevisionIdInternal(value, true);
            }
        }

        public ElementId SignatureTypeId
        {
            get { return _signatureTypeId; }
            set
            {
                int newIdValue = value != null
                    ? IDHelper.ElIdInt(value)
                    : IDHelper.InvalidIdInt;

                ApplySignatureTypeIdInternal(newIdValue, true);
            }
        }

        public int SignatureTypeIdValue
        {
            get { return _signatureTypeIdValue; }
            set
            {
                ApplySignatureTypeIdInternal(value, true);
            }
        }

        public string QuantityText
        {
            get { return _quantityText; }
            set
            {
                string newValue = value ?? string.Empty;

                if (_quantityText == newValue)
                    return;

                _quantityText = newValue;
                OnPropertyChanged("QuantityText");
            }
        }

        public int StatusCode
        {
            get { return _statusCode; }
            set
            {
                if (_statusCode == value)
                    return;

                _statusCode = value;
                OnPropertyChanged("StatusCode");
            }
        }

        public string RevisionDocNumber
        {
            get { return _revisionDocNumber; }
            private set
            {
                if (_revisionDocNumber == value)
                    return;

                _revisionDocNumber = value;
                OnPropertyChanged("RevisionDocNumber");
            }
        }

        public string RevisionDate
        {
            get { return _revisionDate; }
            private set
            {
                if (_revisionDate == value)
                    return;

                _revisionDate = value;
                OnPropertyChanged("RevisionDate");
            }
        }

        public string ManualRevisionText
        {
            get { return _manualRevisionText; }
            set
            {
                string newValue = value ?? string.Empty;
                if (_manualRevisionText == newValue)
                    return;

                _manualRevisionText = newValue;
                OnPropertyChanged("ManualRevisionText");
            }
        }

        public string ManualDocNumberText
        {
            get { return _manualDocNumberText; }
            set
            {
                string newValue = value ?? string.Empty;
                if (_manualDocNumberText == newValue)
                    return;

                _manualDocNumberText = newValue;
                OnPropertyChanged("ManualDocNumberText");
            }
        }

        public string ManualDateText
        {
            get { return _manualDateText; }
            set
            {
                string newValue = value ?? string.Empty;
                if (_manualDateText == newValue)
                    return;

                _manualDateText = newValue;
                OnPropertyChanged("ManualDateText");
            }
        }

        public SheetRevisionLine(
            Document doc,
            int displayNumber,
            List<RevisionComboItem> revisionDefinitions,
            List<SignatureComboItem> signatureDefinitions)
        {
            _doc = doc;
            _displayNumber = displayNumber;

            _revisionId = ElementId.InvalidElementId;
            _revisionIdValue = IDHelper.InvalidIdInt;
            _lastCommittedRevisionIdValue = IDHelper.InvalidIdInt;

            _signatureTypeId = ElementId.InvalidElementId;
            _signatureTypeIdValue = IDHelper.InvalidIdInt;

            _quantityText = string.Empty;
            _statusCode = 0;
            _revisionDocNumber = string.Empty;
            _revisionDate = string.Empty;

            _manualRevisionText = string.Empty;
            _manualDocNumberText = string.Empty;
            _manualDateText = string.Empty;

            AvailableRevisionItems = new ObservableCollection<RevisionComboItem>();
            AvailableSignatureItems = new ObservableCollection<SignatureComboItem>();

            if (revisionDefinitions != null)
            {
                foreach (RevisionComboItem item in revisionDefinitions)
                    AvailableRevisionItems.Add(item.Clone());
            }

            if (signatureDefinitions != null)
            {
                foreach (SignatureComboItem item in signatureDefinitions)
                    AvailableSignatureItems.Add(item.Clone());
            }

            SyncSelectedRevisionItem();
            SyncSelectedSignatureItem();
        }

        public void CommitRevisionSelection()
        {
            _lastCommittedRevisionIdValue = _revisionIdValue;
        }

        public void RevertRevisionSelection()
        {
            ApplyRevisionIdInternal(_lastCommittedRevisionIdValue, true);
        }

        public void RebuildAvailableRevisionItems(IEnumerable<RevisionComboItem> revisionDefinitions)
        {
            _ignoreSelectedRevisionItemChanges = true;

            try
            {
                AvailableRevisionItems.Clear();

                if (revisionDefinitions != null)
                {
                    foreach (RevisionComboItem item in revisionDefinitions)
                        AvailableRevisionItems.Add(item.Clone());
                }
            }
            finally
            {
                _ignoreSelectedRevisionItemChanges = false;
            }

            SyncSelectedRevisionItem();
        }

        public void ApplyRevisionAvailability(HashSet<int> selectedByOtherLines)
        {
            if (selectedByOtherLines == null)
                selectedByOtherLines = new HashSet<int>();

            foreach (RevisionComboItem item in AvailableRevisionItems)
            {
                if (item.IdValue == IDHelper.InvalidIdInt || item.IdValue <= 0)
                {
                    item.IsEnabled = true;
                    continue;
                }

                bool isCurrentSelection = _revisionId != null &&
                                          _revisionId != ElementId.InvalidElementId &&
                                          IDHelper.ElIdInt(_revisionId) == item.IdValue;

                item.IsEnabled = isCurrentSelection || !selectedByOtherLines.Contains(item.IdValue);
            }

            SyncSelectedRevisionItem();
        }

        public void SyncSelectedRevisionItem()
        {
            _isSyncingSelectedItem = true;
            try
            {
                RevisionComboItem item = FindRevisionItem(_revisionIdValue);

                if (!ReferenceEquals(_selectedRevisionItem, item))
                {
                    _selectedRevisionItem = item;
                    OnPropertyChanged("SelectedRevisionItem");
                }
            }
            finally
            {
                _isSyncingSelectedItem = false;
            }
        }

        public void SyncSelectedSignatureItem()
        {
            _isSyncingSelectedSignatureItem = true;
            try
            {
                SignatureComboItem item = FindSignatureItem(_signatureTypeIdValue);

                if (!ReferenceEquals(_selectedSignatureItem, item))
                {
                    _selectedSignatureItem = item;
                    OnPropertyChanged("SelectedSignatureItem");
                }
            }
            finally
            {
                _isSyncingSelectedSignatureItem = false;
            }
        }

        private RevisionComboItem FindRevisionItem(int idValue)
        {
            foreach (RevisionComboItem item in AvailableRevisionItems)
            {
                if (item.IdValue == idValue)
                    return item;
            }

            return null;
        }

        private SignatureComboItem FindSignatureItem(int idValue)
        {
            foreach (SignatureComboItem item in AvailableSignatureItems)
            {
                if (item.IdValue == idValue)
                    return item;
            }

            return AvailableSignatureItems.FirstOrDefault(x => x.IdValue == IDHelper.InvalidIdInt);
        }

        private void ApplyRevisionIdInternal(int revisionIdValue, bool raiseEvents)
        {
            int normalizedIdValue = NormalizeRevisionIdValue(revisionIdValue);

            if (_revisionIdValue == normalizedIdValue)
            {
                SyncSelectedRevisionItem();
                return;
            }

            _revisionIdValue = normalizedIdValue;
            _revisionId = normalizedIdValue == IDHelper.InvalidIdInt
                ? ElementId.InvalidElementId
                : IDHelper.ToElementId(normalizedIdValue);

            if (!HasRevisionSelected)
                ClearRevisionDependentData();
            else
                UpdateRevisionData();

            SyncSelectedRevisionItem();

            if (raiseEvents)
            {
                OnPropertyChanged("RevisionIdValue");
                OnPropertyChanged("RevisionId");
                OnPropertyChanged("HasRevisionSelected");
            }
        }

        private void ApplySignatureTypeIdInternal(int signatureTypeIdValue, bool raiseEvents)
        {
            int normalizedIdValue = NormalizeSignatureTypeIdValue(signatureTypeIdValue);

            if (_signatureTypeIdValue == normalizedIdValue)
            {
                SyncSelectedSignatureItem();
                return;
            }

            _signatureTypeIdValue = normalizedIdValue;
            _signatureTypeId = normalizedIdValue == IDHelper.InvalidIdInt
                ? ElementId.InvalidElementId
                : IDHelper.ToElementId(normalizedIdValue);

            SyncSelectedSignatureItem();

            if (raiseEvents)
            {
                OnPropertyChanged("SignatureTypeIdValue");
                OnPropertyChanged("SignatureTypeId");
            }
        }

        private int NormalizeRevisionIdValue(int revisionIdValue)
        {
            if (revisionIdValue == IDHelper.InvalidIdInt)
                return IDHelper.InvalidIdInt;

            if (revisionIdValue <= 0)
                return IDHelper.InvalidIdInt;

            Revision revision = _doc.GetElement(IDHelper.ToElementId(revisionIdValue)) as Revision;
            if (revision == null)
                return IDHelper.InvalidIdInt;

            return revisionIdValue;
        }

        private int NormalizeSignatureTypeIdValue(int signatureTypeIdValue)
        {
            if (signatureTypeIdValue == IDHelper.InvalidIdInt)
                return IDHelper.InvalidIdInt;

            if (signatureTypeIdValue <= 0)
                return IDHelper.InvalidIdInt;

            Element e = _doc.GetElement(IDHelper.ToElementId(signatureTypeIdValue));
            if (e == null)
                return IDHelper.InvalidIdInt;

            return signatureTypeIdValue;
        }

        private bool IsValidRevisionId(ElementId revisionId)
        {
            if (revisionId == null)
                return false;

            if (revisionId == ElementId.InvalidElementId)
                return false;

            if (IDHelper.ElIdInt(revisionId) <= 0)
                return false;

            return _doc.GetElement(revisionId) is Revision;
        }

        public void SetDisplayNumber(int number)
        {
            DisplayNumber = number;
        }

        public void ResetAllData()
        {
            _revisionId = ElementId.InvalidElementId;
            _revisionIdValue = IDHelper.InvalidIdInt;
            _lastCommittedRevisionIdValue = IDHelper.InvalidIdInt;

            _signatureTypeId = ElementId.InvalidElementId;
            _signatureTypeIdValue = IDHelper.InvalidIdInt;

            _quantityText = string.Empty;
            _statusCode = 0;
            _revisionDocNumber = string.Empty;
            _revisionDate = string.Empty;

            _manualRevisionText = string.Empty;
            _manualDocNumberText = string.Empty;
            _manualDateText = string.Empty;

            SyncSelectedRevisionItem();
            SyncSelectedSignatureItem();

            OnPropertyChanged("RevisionIdValue");
            OnPropertyChanged("RevisionId");
            OnPropertyChanged("HasRevisionSelected");

            OnPropertyChanged("SignatureTypeIdValue");
            OnPropertyChanged("SignatureTypeId");

            OnPropertyChanged("QuantityText");
            OnPropertyChanged("StatusCode");
            OnPropertyChanged("RevisionDocNumber");
            OnPropertyChanged("RevisionDate");

            OnPropertyChanged("ManualRevisionText");
            OnPropertyChanged("ManualDocNumberText");
            OnPropertyChanged("ManualDateText");
        }

        public void ApplyRevisionState(ElementId revisionId)
        {
            int newIdValue = revisionId != null
                ? IDHelper.ElIdInt(revisionId)
                : IDHelper.InvalidIdInt;

            newIdValue = NormalizeRevisionIdValue(newIdValue);

            bool revisionChanged = _revisionIdValue != newIdValue;

            _revisionIdValue = newIdValue;
            _revisionId = newIdValue == IDHelper.InvalidIdInt
                ? ElementId.InvalidElementId
                : IDHelper.ToElementId(newIdValue);

            if (!HasRevisionSelected)
            {
                RevisionDocNumber = string.Empty;
                RevisionDate = string.Empty;
            }
            else
            {
                UpdateRevisionData();
            }

            _lastCommittedRevisionIdValue = _revisionIdValue;

            SyncSelectedRevisionItem();

            if (revisionChanged)
            {
                OnPropertyChanged("RevisionIdValue");
                OnPropertyChanged("RevisionId");
                OnPropertyChanged("HasRevisionSelected");
            }
        }

        private void ClearRevisionDependentData()
        {
            RevisionDocNumber = string.Empty;
            RevisionDate = string.Empty;
        }

        private void UpdateRevisionData()
        {
            Revision revision = null;

            if (IsValidRevisionId(_revisionId))
                revision = _doc.GetElement(_revisionId) as Revision;

            if (revision == null)
            {
                RevisionDocNumber = string.Empty;
                RevisionDate = string.Empty;
                return;
            }

            RevisionDocNumber = GetRevisionApprovedBy(revision);
            RevisionDate = GetRevisionDate(revision);
        }

        private string GetRevisionApprovedBy(Revision revision)
        {
            if (revision == null)
                return string.Empty;

            try
            {
                if (!string.IsNullOrWhiteSpace(revision.IssuedBy))
                    return revision.IssuedBy;
            }
            catch
            {
            }

            string value = GetParameterString(
                revision,
                "Утвердил",
                "Issued By",
                "Approved By",
                "ApprovedBy",
                "Кем утвержден",
                "Кто утвердил");

            return value;
        }

        private string GetRevisionDate(Revision revision)
        {
            if (revision == null)
                return string.Empty;

            try
            {
                if (!string.IsNullOrWhiteSpace(revision.RevisionDate))
                    return revision.RevisionDate;
            }
            catch
            {
            }

            string value = GetParameterString(revision, "Дата", "Revision Date");
            return value;
        }

        private string GetParameterString(Element element, params string[] parameterNames)
        {
            if (element == null || parameterNames == null)
                return string.Empty;

            foreach (string name in parameterNames)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                Parameter p = element.LookupParameter(name);
                if (p == null)
                    continue;

                string value = GetParameterValueAsString(p);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return string.Empty;
        }

        private string GetParameterValueAsString(Parameter p)
        {
            if (p == null)
                return string.Empty;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                string valueString = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(valueString))
                    return valueString;

                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.Double)
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);

                if (p.StorageType == StorageType.ElementId)
                    return IDHelper.ElIdInt(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
            }

            return string.Empty;
        }

        protected void OnPropertyChanged(string propName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propName));
        }
    }

    internal class ElementIdEqualityComparer : IEqualityComparer<ElementId>
    {
        public bool Equals(ElementId x, ElementId y)
        {
            if (x == null && y == null)
                return true;

            if (x == null || y == null)
                return false;

            return IDHelper.ElIdInt(x) == IDHelper.ElIdInt(y);
        }

        public int GetHashCode(ElementId obj)
        {
            if (obj == null)
                return 0;

            return IDHelper.ElIdInt(obj);
        }
    }

    public partial class FrmIZMFilling : Window
    {
        private readonly Document _doc;
        private readonly FrmIZMFillingViewModel _vm;

        public FrmIZMFilling(Document doc)
        {
            InitializeComponent();

            _doc = doc;
            _vm = new FrmIZMFillingViewModel(_doc);
            _vm.SelectedSheetsChanged += Vm_SelectedSheetsChanged;
            _vm.DuplicateRevisionDetected += Vm_DuplicateRevisionDetected;

            DataContext = _vm;
        }

        private void Vm_SelectedSheetsChanged(object sender, EventArgs e)
        {
        }

        private void Vm_DuplicateRevisionDetected(object sender, EventArgs e)
        {
            MessageBox.Show(
                "На одном листе одно и то же изменение нельзя выбрать несколько раз.",
                "Предупреждение",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (_vm.SheetRows.Count == 0)
            {
                MessageBox.Show("Не выбраны листы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (SheetRowItem row in _vm.SheetRows)
            {
                if (row.HasDuplicateRevisions())
                {
                    MessageBox.Show(
                        "Обнаружены повторяющиеся ИЗМы в одном из листов. Исправьте таблицу и повторите.",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }

            using (Transaction t = new Transaction(_doc, "Заполнение ИЗМов"))
            {
                t.Start();

                foreach (SheetRowItem row in _vm.SheetRows)
                    row.ApplyToSheet();

                t.Commit();
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Изменения применены.");
            sb.AppendLine("Обработано листов: " + _vm.SheetRows.Count);

            MessageBox.Show(sb.ToString(), "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}