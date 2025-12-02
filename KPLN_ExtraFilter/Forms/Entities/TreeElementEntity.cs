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

        /// <summary>
        /// Создаю дерево с группировкой Категория-Семейство-Тип
        /// </summary>
        public static TreeElementEntity[] CreateTreeElEnt_ByCatANDFamANDType(IEnumerable<Element> userSelElems)
        {
            List<TreeElementEntity> result = new List<TreeElementEntity>();

            string noFamName = "<Без семейства>";
            string noTypeName = "<Без типа>";

            // Уровень 1: КАТЕГОРИЯ
            var groupedByCat = userSelElems
                .Where(el => el.Category != null)
                .GroupBy(el => el.Category.Name)
                .OrderBy(g => g.Key);

            foreach (IGrouping<string, Element> catGroup in groupedByCat)
            {
                TreeElementEntity catNode = new TreeElementEntity
                {
                    TEE_Name = catGroup.Key
                };

                // Уровень 2: СЕМЕЙСТВО
                var groupedByFamily = catGroup
                    .GroupBy(el =>
                    {
                        string famName = null;

                        if (el is Family family)
                            famName = family.Name;
                        else if (el is ElementType elType)
                            famName = elType.FamilyName;
                        else if (el is FamilyInstance fi)
                            famName = fi.Symbol?.FamilyName;
                        else
                        {
                            // "Семейство : Тип"
                            string famAndType = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsValueString();
                            if (!string.IsNullOrEmpty(famAndType))
                            {
                                int idx = famAndType.IndexOf(':');
                                famName = idx > 0 ? famAndType.Substring(0, idx).Trim() : famAndType;
                            }
                        }

                        if (string.IsNullOrEmpty(famName))
                            famName = noFamName;

                        return famName;
                    })
                    .OrderBy(g => g.Key);

                foreach (IGrouping<string, Element> famGroup in groupedByFamily)
                {
                    TreeElementEntity famNode = new TreeElementEntity
                    {
                        TEE_Name = famGroup.Key,
                        TEE_Parent = catNode
                    };

                    // Уровень 3: ТИП
                    var groupedByType = famGroup
                        .GroupBy(el =>
                        {
                            string typeName = null;

                            if (el is ElementType elType)
                            {
                                typeName = elType.Name;
                            }
                            else if (el is FamilyInstance fi)
                            {
                                typeName = fi.Symbol?.Name;
                            }
                            else
                            {
                                // Пробуем взять имя типа параметрами
                                Parameter typeParam =
                                    el.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) ??
                                    el.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);

                                if (typeParam != null)
                                {
                                    string val = typeParam.AsValueString();
                                    if (!string.IsNullOrEmpty(val))
                                    {
                                        // Для "Семейство : Тип" откусываем часть после двоеточия
                                        int idx = val.IndexOf(':');
                                        typeName = idx >= 0 && idx < val.Length - 1
                                            ? val.Substring(idx + 1).Trim()
                                            : val;
                                    }
                                }
                            }

                            if (string.IsNullOrEmpty(typeName))
                                typeName = noTypeName;

                            return typeName;
                        })
                        .OrderBy(g => g.Key);

                    foreach (IGrouping<string, Element> typeGroup in groupedByType)
                    {
                        TreeElementEntity typeNode = new TreeElementEntity
                        {
                            TEE_Name = typeGroup.Key,
                            TEE_Parent = famNode
                        };

                        // Уровень 4: ЭЛЕМЕНТЫ
                        foreach (Element el in typeGroup)
                        {
                            TreeElementEntity elNode = new TreeElementEntity
                            {
                                TEE_Name = $"{el.Name}, id: {el.Id}",
                                TEE_Element = el,
                                TEE_Parent = typeNode
                            };

                            typeNode.TEE_ChildrenColl.Add(elNode);
                        }

                        famNode.TEE_ChildrenColl.Add(typeNode);
                    }

                    catNode.TEE_ChildrenColl.Add(famNode);
                }

                result.Add(catNode);
            }

            return result.ToArray();
        }

        /// <summary>
        /// Создаю дерево с группировкой Значение параметра-Категория-Семейство-Тип
        /// </summary>
        public static TreeElementEntity[] CreateTreeElEnt_ByParamANDCatANDFamANDType(Document doc, IEnumerable<Element> elements, IEnumerable<SelectionByModelM_ParamM> paramMColl)
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
                        Parameter param = e.LookupParameter(paramName);

                        if (param == null && doc.GetElement(e.GetTypeId()) is Element typeElem)
                            param = typeElem.LookupParameter(paramName);

                        string value;

                        if (param == null)
                            value = noParamData;
                        else if (!param.HasValue)
                            value = emptyParamData;
                        else
                            value = DocWorker.GetParamValueInSI(doc, param);

                        // Добивка по пустым значениям пар-в
                        if (string.IsNullOrEmpty(value))
                            value = emptyParamData;


                        keyParts.Add($"{paramName}: \"{value}\"");
                    }

                    // Формируем ключ:
                    // Param1: X_Param2: Y_Param3: Z
                    return string.Join(" + ", keyParts);
                })
                .OrderBy(g => g.Key);

            // Создаём дерево
            foreach (var pGroup in paramGroups)
            {
                var topNode = new TreeElementEntity
                {
                    TEE_Name = pGroup.Key,
                };

                // Категорийное дерево - это подкатегория
                var catTree = CreateTreeElEnt_ByCatANDFamANDType(pGroup);

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
