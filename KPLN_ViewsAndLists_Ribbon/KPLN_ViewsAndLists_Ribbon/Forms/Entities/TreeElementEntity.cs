using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ViewsAndLists_Ribbon.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для элемента модели
    /// </summary>
    public sealed class TreeElementEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isChecked;
        private bool _isExpanded;
        private bool _isVisible = true;

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

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible == value)
                    return;

                _isVisible = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Создаю дерево с группировкой Категория-Семейство
        /// </summary>
        public static TreeElementEntity[] CreateTreeElEnt_ByCatANDFam(IEnumerable<Element> userSelElems)
        {
            List<TreeElementEntity> result = new List<TreeElementEntity>();

            string noFamName = "<Без семейства>";

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


                    // Уровень 3: ЭЛЕМЕНТЫ
                    foreach (Element el in famGroup)
                    {
                        TreeElementEntity elNode = new TreeElementEntity
                        {
                            TEE_Name = $"{el.Name}, id: {el.Id}",
                            TEE_Element = el,
                        };
                        
                        famNode.TEE_ChildrenColl.Add(elNode);
                    }

                    catNode.TEE_ChildrenColl.Add(famNode);
                }

                result.Add(catNode);
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
            foreach (TreeElementEntity treeEnt in treeElementEntities)
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
