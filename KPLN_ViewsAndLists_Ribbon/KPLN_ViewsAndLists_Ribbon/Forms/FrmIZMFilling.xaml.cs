using Autodesk.Revit.DB;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    internal class FrmIZMFillingViewModel : INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly Dictionary<int, SheetRowItem> _sheetRowCache = new Dictionary<int, SheetRowItem>();
        private readonly List<SignatureComboItem> _signatureDefinitions = new List<SignatureComboItem>();

        public ObservableCollection<SheetTreeNode> RootNodes { get; private set; }
        public ObservableCollection<SheetRowItem> SheetRows { get; private set; }
        public ObservableCollection<RevisionComboItem> RevisionItems { get; private set; }
        public ObservableCollection<SheetStatusItem> StatusItems { get; private set; }

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
            {
                _signatureDefinitions.Add(new SignatureComboItem(symbol.Id, GetSignatureDisplayName(symbol)));
            }
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

            return symbol.Id.IntegerValue.ToString(CultureInfo.InvariantCulture);
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
                int key = sheet.Id.IntegerValue;
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
                    return p.AsElementId().IntegerValue.ToString(CultureInfo.InvariantCulture);
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
            IdValue = Id.IntegerValue;
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
            IdValue = Id.IntegerValue;
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
        private bool _isSorting;
        private bool _isResettingDuplicateRevision;
        private bool _isRebuildingRevisionUi;

        public ViewSheet Sheet { get; private set; }
        public string SheetNumber { get; private set; }
        public string SheetName { get; private set; }
        public ObservableCollection<SheetRevisionLine> Lines { get; private set; }

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

            Lines = new ObservableCollection<SheetRevisionLine>();
            Lines.CollectionChanged += Lines_CollectionChanged;

            for (int i = 1; i <= 4; i++)
                Lines.Add(new SheetRevisionLine(_doc, i, _revisionDefinitions, _signatureDefinitions));

            SubscribeLines(Lines, true);
            LoadFromSheet();
        }

        private void Lines_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SubscribeLines(e.NewItems, true);
            SubscribeLines(e.OldItems, false);
            RefreshDisplayNumbers();
            RebuildRevisionUi();
        }

        private void SubscribeLines(IList items, bool subscribe)
        {
            if (items == null)
                return;

            foreach (object item in items)
            {
                SheetRevisionLine line = item as SheetRevisionLine;
                if (line == null)
                    continue;

                if (subscribe)
                    line.PropertyChanged += Line_PropertyChanged;
                else
                    line.PropertyChanged -= Line_PropertyChanged;
            }
        }

        private void Line_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "RevisionIdValue")
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

        private void CommitAllRevisionSelections()
        {
            foreach (SheetRevisionLine line in Lines)
                line.CommitRevisionSelection();
        }

        private bool IsValidRevisionId(ElementId revisionId)
        {
            if (revisionId == null)
                return false;

            if (revisionId == ElementId.InvalidElementId)
                return false;

            if (revisionId.IntegerValue <= 0)
                return false;

            return _doc.GetElement(revisionId) is Revision;
        }

        public void LoadFromSheet()
        {
            List<Revision> revisions = Sheet.GetAdditionalRevisionIds()
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

        private void LoadStaticColumnsFromSheet()
        {
            int[] statusDigits = ParseStatusDigits(GetParameterString(Sheet, "Ш.ШифрСтатусЛиста"));
            FamilyInstance titleBlock = GetMainTitleBlock();

            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int mirroredSlotNumber = GetMirroredSlotNumber(rowIndex);

                line.QuantityText = GetParameterString(Sheet, GetQtyParamName(mirroredSlotNumber));
                line.StatusCode = GetStatusCodeBySlotNumber(statusDigits, mirroredSlotNumber);

                int displayNumber = line.DisplayNumber;
                line.SignatureTypeId = GetSignatureTypeIdFromTitleBlock(titleBlock, displayNumber);
            }
        }

        public void ApplyToSheet()
        {
            SortLinesByRevisionOrder();

            List<ElementId> selectedRevisionIds = GetSelectedRevisionIds();
            Sheet.SetAdditionalRevisionIds(selectedRevisionIds);

            // Кол. уч. не трогаем логикой сортировки.
            // Пишем его строго из текущей визуальной строки в соответствующий зеркальный слот.
            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int mirroredSlotNumber = GetMirroredSlotNumber(rowIndex);

                SetParameterString(Sheet, GetQtyParamName(mirroredSlotNumber), line.QuantityText);
            }

            // Лист (статус) тоже автономный.
            StringBuilder statusBuilder = new StringBuilder();
            for (int slotNumber = 1; slotNumber <= Lines.Count; slotNumber++)
            {
                SheetRevisionLine line = GetLineByMirroredSlotNumber(slotNumber);
                int status = line != null ? line.StatusCode : 0;
                statusBuilder.Append(status.ToString(CultureInfo.InvariantCulture));
            }

            SetParameterString(Sheet, "Ш.ШифрСтатусЛиста", statusBuilder.ToString());

            // Подпись автономная. Пишем строго в основную надпись по номеру строки 4/3/2/1.
            FamilyInstance titleBlock = GetMainTitleBlock();

            for (int rowIndex = 0; rowIndex < Lines.Count; rowIndex++)
            {
                SheetRevisionLine line = Lines[rowIndex];
                int displayNumber = line.DisplayNumber;
                SetSignatureTypeIdToTitleBlock(titleBlock, displayNumber, line.SignatureTypeId);
            }
        }

        public void SortLinesByRevisionOrder()
        {
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

                // ВАЖНО:
                // Переставляем ТОЛЬКО revision.
                // Quantity / Status / Signature не трогаем вообще.
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

        private void RebuildRevisionUi()
        {
            _isRebuildingRevisionUi = true;

            try
            {
                foreach (SheetRevisionLine line in Lines)
                    line.RebuildAvailableRevisionItems(_revisionDefinitions);

                foreach (SheetRevisionLine currentLine in Lines)
                {
                    HashSet<int> selectedByOtherLines = new HashSet<int>(
                        Lines.Where(x => !ReferenceEquals(x, currentLine) && IsValidRevisionId(x.RevisionId))
                             .Select(x => x.RevisionId.IntegerValue));

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
                return 0; // пустые сверху

            return 1;     // заполненные снизу
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

            Parameter p = titleBlock.LookupParameter(baseName);
            if (p != null)
                return p;

            p = titleBlock.LookupParameter(baseName + "<Типовые аннотации>");
            if (p != null)
                return p;

            return null;
        }

        private FamilyInstance GetMainTitleBlock()
        {
            List<FamilyInstance> titleBlocks = new FilteredElementCollector(_doc, Sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            if (titleBlocks.Count == 0)
                return null;

            FamilyInstance preferred = titleBlocks
                .FirstOrDefault(x => GetTitleBlockFamilyName(x).StartsWith("020_Основная надпись", StringComparison.CurrentCultureIgnoreCase));

            return preferred ?? titleBlocks.FirstOrDefault();
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

            return string.Empty;
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
                        .FirstOrDefault(x => x.IdValue == valueToSet.IntegerValue);

                    p.Set(item != null ? item.Name : string.Empty);
                }
            }
            catch
            {
            }
        }

        private int GetStatusCodeBySlotNumber(int[] digits, int slotNumber)
        {
            if (digits == null || digits.Length < 4)
                return 0;

            int index = slotNumber - 1;

            if (index < 0 || index >= digits.Length)
                return 0;

            return digits[index];
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

        public bool HasDuplicateRevisions()
        {
            List<int> ids = Lines
                .Where(x => IsValidRevisionId(x.RevisionId))
                .Select(x => x.RevisionId.IntegerValue)
                .ToList();

            return ids.Count != ids.Distinct().Count();
        }

        private int[] ParseStatusDigits(string input)
        {
            int[] result = new int[] { 0, 0, 0, 0 };

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
                    else
                        result[i] = 0;
                }
            }

            return result;
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
                    return p.AsElementId().IntegerValue.ToString(CultureInfo.InvariantCulture);
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
                    : ElementId.InvalidElementId.IntegerValue;

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
                    : ElementId.InvalidElementId.IntegerValue;

                ApplySignatureTypeIdInternal(newIdValue, true);
            }
        }

        public ElementId RevisionId
        {
            get { return _revisionId; }
            set
            {
                int newIdValue = value != null
                    ? value.IntegerValue
                    : ElementId.InvalidElementId.IntegerValue;

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
                    ? value.IntegerValue
                    : ElementId.InvalidElementId.IntegerValue;

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

        public SheetRevisionLine(
            Document doc,
            int displayNumber,
            List<RevisionComboItem> revisionDefinitions,
            List<SignatureComboItem> signatureDefinitions)
        {
            _doc = doc;
            _displayNumber = displayNumber;

            _revisionId = ElementId.InvalidElementId;
            _revisionIdValue = ElementId.InvalidElementId.IntegerValue;
            _lastCommittedRevisionIdValue = ElementId.InvalidElementId.IntegerValue;

            _signatureTypeId = ElementId.InvalidElementId;
            _signatureTypeIdValue = ElementId.InvalidElementId.IntegerValue;

            _quantityText = string.Empty;
            _statusCode = 0;
            _revisionDocNumber = string.Empty;
            _revisionDate = string.Empty;

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
                if (item.IdValue == ElementId.InvalidElementId.IntegerValue || item.IdValue <= 0)
                {
                    item.IsEnabled = true;
                    continue;
                }

                bool isCurrentSelection = _revisionId != null &&
                                          _revisionId != ElementId.InvalidElementId &&
                                          _revisionId.IntegerValue == item.IdValue;

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

            return AvailableSignatureItems.FirstOrDefault(x => x.IdValue == ElementId.InvalidElementId.IntegerValue);
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
            _revisionId = normalizedIdValue == ElementId.InvalidElementId.IntegerValue
                ? ElementId.InvalidElementId
                : new ElementId(normalizedIdValue);

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
            _signatureTypeId = normalizedIdValue == ElementId.InvalidElementId.IntegerValue
                ? ElementId.InvalidElementId
                : new ElementId(normalizedIdValue);

            SyncSelectedSignatureItem();

            if (raiseEvents)
            {
                OnPropertyChanged("SignatureTypeIdValue");
                OnPropertyChanged("SignatureTypeId");
            }
        }

        private int NormalizeRevisionIdValue(int revisionIdValue)
        {
            if (revisionIdValue <= 0)
                return ElementId.InvalidElementId.IntegerValue;

            if (revisionIdValue == ElementId.InvalidElementId.IntegerValue)
                return ElementId.InvalidElementId.IntegerValue;

            Revision revision = _doc.GetElement(new ElementId(revisionIdValue)) as Revision;
            if (revision == null)
                return ElementId.InvalidElementId.IntegerValue;

            return revisionIdValue;
        }

        private int NormalizeSignatureTypeIdValue(int signatureTypeIdValue)
        {
            if (signatureTypeIdValue <= 0)
                return ElementId.InvalidElementId.IntegerValue;

            if (signatureTypeIdValue == ElementId.InvalidElementId.IntegerValue)
                return ElementId.InvalidElementId.IntegerValue;

            Element e = _doc.GetElement(new ElementId(signatureTypeIdValue));
            if (e == null)
                return ElementId.InvalidElementId.IntegerValue;

            return signatureTypeIdValue;
        }

        private bool IsValidRevisionId(ElementId revisionId)
        {
            if (revisionId == null)
                return false;

            if (revisionId == ElementId.InvalidElementId)
                return false;

            if (revisionId.IntegerValue <= 0)
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
            _revisionIdValue = ElementId.InvalidElementId.IntegerValue;
            _lastCommittedRevisionIdValue = ElementId.InvalidElementId.IntegerValue;

            _signatureTypeId = ElementId.InvalidElementId;
            _signatureTypeIdValue = ElementId.InvalidElementId.IntegerValue;

            _quantityText = string.Empty;
            _statusCode = 0;
            _revisionDocNumber = string.Empty;
            _revisionDate = string.Empty;

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
        }

        public void ApplyRevisionState(ElementId revisionId)
        {
            int newIdValue = revisionId != null
                ? revisionId.IntegerValue
                : ElementId.InvalidElementId.IntegerValue;

            newIdValue = NormalizeRevisionIdValue(newIdValue);

            bool revisionChanged = _revisionIdValue != newIdValue;

            _revisionIdValue = newIdValue;
            _revisionId = newIdValue == ElementId.InvalidElementId.IntegerValue
                ? ElementId.InvalidElementId
                : new ElementId(newIdValue);

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
            string value = GetParameterString(
                revision,
                "Утвердил",
                "Approved By",
                "ApprovedBy",
                "Кем утвержден",
                "Кто утвердил");

            return value;
        }

        private string GetRevisionDate(Revision revision)
        {
            if (!string.IsNullOrWhiteSpace(revision.RevisionDate))
                return revision.RevisionDate;

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
                    return p.AsElementId().IntegerValue.ToString(CultureInfo.InvariantCulture);
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

            return x.IntegerValue == y.IntegerValue;
        }

        public int GetHashCode(ElementId obj)
        {
            if (obj == null)
                return 0;

            return obj.IntegerValue;
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
                    MessageBox.Show("Обнаружены повторяющиеся ИЗМы в одном из листов. Исправьте таблицу и повторите.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
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