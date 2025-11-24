using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для элемента модели
    /// </summary>
    public sealed class TreeElementEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isChecked;
        private bool _isExpanded;

        private TreeElementEntity() { }

        /// <summary>
        /// Ссылка на элемент
        /// </summary>
        public Element TEE_Element { get; private set; }

        /// <summary>
        /// Имя элемента дерева
        /// </summary>
        public string TEE_Name { get; private set; }

        /// <summary>
        /// Коллекция "детей"
        /// </summary>
        public List<TreeElementEntity> TEE_ChildrenColl { get; set; } = new List<TreeElementEntity>();

        /// <summary>
        /// Ссылка на родителя
        /// </summary>
        public TreeElementEntity TEE_Parent { get; set; }

        /// <summary>
        /// Спец формат подсчёта эл-в
        /// </summary>
        public string TEE_ChildrensCount
        {
            get
            {
                int count = 0;
                CountChildrens(this, ref count);
                if (count > 0)
                    return $": {count} эл.";

                return string.Empty;
            }
        }

        /// <summary>
        /// Группа развернута?
        /// </summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Группа/элемент выбран
        /// </summary>
        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value)
                    return;

                _isChecked = value;
                NotifyPropertyChanged();
                UpdateChildrenCheck(value);
                UpdateSubParentCheck();
            }
        }

        public static TreeElementEntity[] CreateTreeElEnt_ByCategory(IEnumerable<Element> userSelElems)
        {
            List<TreeElementEntity> result = new List<TreeElementEntity>();

            string noFamName = "<Без семейства>";
            // Группирую по категории
            var groupedByCat = userSelElems
                .GroupBy(el => el.Category.Name)
                .OrderBy(g => g.Key);


            foreach (IGrouping<string, Element> catGroup in groupedByCat)
            {
                // Уровень 1: КАТЕГОРИЯ
                TreeElementEntity catNode = new TreeElementEntity()
                {
                    TEE_Name = catGroup.Key,
                };

                // Группирую по семейству
                var groupedByFamily = catGroup
                    .GroupBy(el =>
                    {
                        string famName = null;
                        if (el is Family family)
                            famName = family.FamilyCategory.Name;
                        else if (el is ElementType elType)
                            famName = elType.FamilyName;
                        else if (el is FamilyInstance fi)
                            famName = fi.Symbol?.FamilyName;
                        else
                            famName = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsValueString();

                        if (string.IsNullOrEmpty(famName))
                            return noFamName;

                        return famName;
                    })
                    .OrderBy(g => g.Key);

                foreach (IGrouping<string, Element> famGroup in groupedByFamily)
                {
                    // Уровень 2: СЕМЕЙСТВА
                    TreeElementEntity famNode = new TreeElementEntity()
                    {
                        TEE_Name = famGroup.Key,
                        TEE_Parent = catNode
                    };

                    // Узровень 3: ЭЛЕМЕНТЫ
                    foreach (Element el in famGroup)
                    {
                        TreeElementEntity elNode = new TreeElementEntity()
                        {
                            TEE_Name = $"{el.Name}, id: {el.Id}",
                            TEE_Element = el,
                            TEE_Parent = famNode
                        };
                        famNode.TEE_ChildrenColl.Add(elNode);
                    }

                    catNode.TEE_ChildrenColl.Add(famNode);
                }

                result.Add(catNode);
            }

            return result.ToArray();
        }

        public static TreeElementEntity[] SortTreeElEnt_ByParameter(Document doc, IEnumerable<Element> elements, IEnumerable<SelectionByModelM_ParamM> paramMColl)
        {
            List<TreeElementEntity> result = new List<TreeElementEntity>();

            string noParamData = "<Параметр отсутсвует>";
            string emptyParamData = "<Параметр пустой>";
            ParamEntity[] paramEntities = paramMColl.Select(pm => pm.ParamM_SelectedParameter).ToArray();

            // Группировка по значению параметра
            var paramGroups = elements
                .GroupBy(e =>
                {
                    List<string> keyParts = new List<string>();

                    foreach (var pEntity in paramEntities)
                    {
                        string paramName = pEntity.RevitParamName;
                        Parameter p = e.LookupParameter(paramName);

                        if (p == null && doc.GetElement(e.GetTypeId()) is Element typeElem)
                            p = typeElem.LookupParameter(paramName);

                        string value;

                        if (p == null)
                            value = noParamData;
                        else if (!p.HasValue)
                            value = emptyParamData;
                        else
                            value = DocWorker.GetParamValueInSI(doc, p);

                        keyParts.Add($"{paramName}: {value}");
                    }

                    // Формируем ключ:
                    // Param1: X_Param2: Y_Param3: Z
                    return string.Join(" + ", keyParts);
                })
                .OrderBy(g => g.Key);

            // --- СТВАРЭННЕ ДРЭВА ---
            foreach (var pGroup in paramGroups)
            {
                var topNode = new TreeElementEntity
                {
                    TEE_Name = pGroup.Key,
                };

                // катэгорыі як падузлы
                var catTree = CreateTreeElEnt_ByCategory(pGroup);

                foreach (var catEntity in catTree)
                {
                    catEntity.TEE_Parent = topNode;
                    topNode.TEE_ChildrenColl.Add(catEntity);
                }

                result.Add(topNode);
            }

            return result.ToArray();
        }

        /// <summary>
        /// Забрать ВСЕ выбранные пользователем элементы модели
        /// </summary>
        /// <param name="treeElementEntities"></param>
        /// <param name="rElems"></param>
        public static void GetAllCheckedRevitElemsFromTreeElemColl(IEnumerable<TreeElementEntity> treeElementEntities, ref List<Element> rElems)
        {
            foreach(TreeElementEntity treeEnt in treeElementEntities)
            {
                if (treeEnt.TEE_ChildrenColl == null || treeEnt.TEE_ChildrenColl.Count() == 0)
                {
                    if (treeEnt.TEE_Element != null && treeEnt.IsChecked)
                        rElems.Add(treeEnt.TEE_Element);
                }
                else
                    GetAllCheckedRevitElemsFromTreeElemColl(treeEnt.TEE_ChildrenColl, ref rElems);
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>
        /// Рекурсивынй подсчёт всех "детей"
        /// </summary>
        /// <returns></returns>
        private int CountChildrens(TreeElementEntity entity, ref int count)
        {
            if (entity.TEE_ChildrenColl.Count() == 0)
                return 0;
            else
            {
                count += entity.TEE_ChildrenColl.Count(ch => ch.TEE_Element != null);
                foreach (var child in entity.TEE_ChildrenColl)
                {
                    CountChildrens(child, ref count);
                }
            }

            return count;
        }

        /// <summary>
        /// Передача чека на всех "детей"
        /// </summary>
        /// <param name="value"></param>
        private void UpdateChildrenCheck(bool value)
        {
            foreach (var child in TEE_ChildrenColl)
            {
                child._isChecked = value;
                child.NotifyPropertyChanged(nameof(IsChecked));
                child.UpdateChildrenCheck(value);
            }
        }

        /// <summary>
        /// Передача чека на всех "детей", у которых свои "дети"
        /// </summary>
        private void UpdateSubParentCheck()
        {
            if (TEE_Parent == null)
                return;

            bool allChecked = true;
            bool allUnchecked = true;
            foreach (var child in TEE_Parent.TEE_ChildrenColl)
            {
                if (child.IsChecked)
                    allUnchecked = false;
                else
                    allChecked = false;
            }

            if (allChecked)
                TEE_Parent._isChecked = true;
            else if (allUnchecked)
                TEE_Parent._isChecked = false;
            else
                return; // змешаны стан — нічога не мяняем

            TEE_Parent.NotifyPropertyChanged(nameof(IsChecked));
            TEE_Parent.UpdateSubParentCheck();
        }
    }
}
