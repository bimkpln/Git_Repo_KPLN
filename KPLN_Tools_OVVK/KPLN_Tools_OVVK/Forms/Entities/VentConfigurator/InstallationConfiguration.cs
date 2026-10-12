using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class InstallationConfiguration
    {
        internal List<ConfigurationIssue> ReviewIssues = new List<ConfigurationIssue>();
        internal Dictionary<string, string> RecoveredInputs = new Dictionary<string, string>(StringComparer.Ordinal);

        internal HashSet<string> ActiveSharedParameters()
        {
            return new HashSet<string>(new[] { SharedParameters.FrameVisible, SharedParameters.ServiceRight, SharedParameters.ServiceDepth }
                .Concat(Blocks.SelectMany(b => SharedParameters.ForSection(b.Type))), StringComparer.Ordinal);
        }

        internal IEnumerable<string> EditableInputNames()
        {
            return Info.Assignments().Keys.Concat(Dimensions().Where(p => !p.Value.IsCalculated).Select(p => p.Key))
                .Concat(SharedBooleans.Where(p => ActiveSharedParameters().Contains(p.Key) && !p.Value.IsCalculated).Select(p => p.Key))
                .Concat(Enumerable.Range(0, Blocks.Count).Select(i => ParameterName(SlotAt(i))))
                .Concat(new[] { "Секции_Промежуточные_Количество", ValveControl.ParameterName });
        }

        internal object InputValue(string name)
        {
            name = FamilyParameterNames.Canonical(name);
            if (name == "Секции_Промежуточные_Количество")
                return IntermediateCount;
            if (name == ValveControl.ParameterName)
                return HasValve ? 1 : 0;
            string text;
            if (Info.Assignments().TryGetValue(name, out text))
                return text;
            DimensionValue dimension;
            if (Dimensions().TryGetValue(name, out dimension))
            {
                double value;
                return dimension.TryNumber(out value) ? (object)value : dimension.Text;
            }
            BooleanValue flag;
            if (SharedBooleans.TryGetValue(name, out flag))
                return flag.Value.HasValue ? (object)(flag.Value.Value ? 1 : 0) : null;
            for (int i = 0; i < Blocks.Count; i++)
                if (ParameterName(SlotAt(i)) == name)
                    return Blocks[i].Type;
            return null;
        }

        private string InputToken(string name)
        {
            var value = InputValue(name);
            return value is SectionTypeChoice ? ((SectionTypeChoice)value).Key : Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }

        internal void SetInputValue(string name, object value)
        {
            name = FamilyParameterNames.Canonical(name);
            if (name == "Секции_Промежуточные_Количество")
            {
                int count = (int)value;
                if (count < 1 || count > 9)
                    throw new ArgumentException("Допустимо от 1 до 9 промежуточных секций.");
                Resize(count + (HasValve ? 3 : 2));
                return;
            }
            if (name == ValveControl.ParameterName)
            {
                int flagValue = (int)value;
                if (flagValue != 0 && flagValue != 1)
                    throw new ArgumentException("Укажите наличие клапана: Да или Нет.");
                SetValve(flagValue == 1);
                return;
            }
            if (Info.Assignments().ContainsKey(name))
            {
                Info.SetParameterValue(name, value as string);
                return;
            }
            DimensionValue dimension;
            if (Dimensions().TryGetValue(name, out dimension))
            {
                dimension.ReadInput(value is double ? (double)value : double.NaN);
                return;
            }
            BooleanValue flag;
            if (SharedBooleans.TryGetValue(name, out flag))
            {
                if (value == null)
                {
                    flag.Value = null;
                    return;
                }
                int flagValue = (int)value;
                if (flagValue != 0 && flagValue != 1)
                    throw new ArgumentException("Значение должно быть Да или Нет.");
                flag.Value = flagValue == 1;
                return;
            }
            for (int i = 0; i < Blocks.Count; i++)
            {
                int slot = SlotAt(i);
                if (ParameterName(slot) != name)
                    continue;
                var choice = value as SectionTypeChoice;
                var matched = choice == null ? null : Catalog[slot].Choices.SingleOrDefault(c => c.Key == choice.Key);
                Blocks[i].Type = matched;
                if (choice != null && matched == null)
                    throw new ArgumentException("В основе нет элемента «" + choice.DisplayName + "». Выберите замену в составе установки.");
                return;
            }
        }

        private static SectionCatalog InputDefaults(SectionCatalog schema)
        {
            var defaults = schema.Copy();
            defaults.IntermediateCount = 1;
            defaults.HasValve = false;
            defaults.InstallationWidthMm = 500;
            defaults.InstallationHeightMm = 300;
            defaults.FrameHeightMm = 100;
            foreach (var block in defaults.Values)
            {
                block.SourceType = block.Choices.FirstOrDefault(c => c.IsEmpty);
                block.DefaultLengthMm = 1;
                block.DefaultWidthMm = 500;
                block.DefaultHeightMm = 300;
                block.DefaultOffsetXMm = block.DefaultOffsetYMm = 0;
            }
            foreach (string name in SharedParameters.LengthNames)
                defaults.SharedLengthsMm[name] = name == SharedParameters.ServiceDepth ? 1000 : FamilyParameterNames.IsOffset(name) ? 0 : double.NaN;
            foreach (string name in SharedParameters.BooleanNames)
                defaults.SharedFlags[name] = true;
            return defaults;
        }

        internal object SuggestedInputValue(string name)
        {
            name = FamilyParameterNames.Canonical(name);
            if (name == "Секции_Промежуточные_Количество")
                return 1;
            if (name == ValveControl.ParameterName)
                return 0;
            if (Info.Assignments().ContainsKey(name))
                return string.Empty;
            if (SharedParameters.BooleanNames.Contains(name))
                return 1;
            if (name == SharedParameters.ServiceDepth)
                return 1000.0;
            if (name == "Установка_Ширина")
                return 500.0;
            if (name == "Установка_Высота")
                return 300.0;
            if (name == "Основание_Рама_Высота")
                return 100.0;
            if (FamilyParameterNames.IsOffset(name))
                return 0.0;
            for (int i = 0; i < Blocks.Count; i++)
            {
                int slot = SlotAt(i);
                var block = Blocks[i];
                if (name == ParameterName(slot))
                    return Catalog[slot].Choices.FirstOrDefault(c => c.IsEmpty);
                if (name == DimensionParameterName(slot, "Длина"))
                    return block.Type?.IsEmpty == true ? 1.0 : block.IsValve ? 210.0 : block.IsConnector ? 100.0 : 500.0;
                if (block.IsConnector && name == DimensionParameterName(slot, "Ширина"))
                    return 500.0;
                if (block.IsConnector && name == DimensionParameterName(slot, "Высота"))
                    return 300.0;
            }
            // Диаметры патрубков не угадываем и не берём из заполненных типов на X.
            return null;
        }

        internal List<ConfigurationIssue> Problems()
        {
            var active = new HashSet<string>(EditableInputNames(), StringComparer.Ordinal);
            var result = ReviewIssues.Where(i => !i.Accepted && active.Contains(i.ParameterName)
                && i.InitialValue == InputToken(i.ParameterName)).ToList();
            Action<string, string> add = (name, message) =>
            {
                if (!result.Any(i => i.ParameterName == name))
                    result.Add(new ConfigurationIssue
                    {
                        ParameterName = name,
                        Message = message,
                        SuggestedValue = SuggestedInputValue(name)
                    });
            };
            foreach (var pair in Dimensions().Where(p => !p.Value.IsValid))
                add(pair.Key, "Укажите корректный размер. Смещения могут быть нулевыми или отрицательными; остальные размеры должны быть больше нуля.");
            foreach (var pair in SharedBooleans.Where(p => active.Contains(p.Key) && !p.Value.IsValid))
                add(pair.Key, "Выберите Да/Нет или сторону.");
            for (int i = 0; i < Blocks.Count; i++)
                if (Blocks[i].Type == null)
                    add(ParameterName(SlotAt(i)), "Выберите доступный тип элемента в составе установки.");
            return result;
        }

        internal void ApplySuggestion(ConfigurationIssue issue)
        {
            SetInputValue(issue.ParameterName, issue.SuggestedValue);
            issue.Accepted = true;
            RecalculateLocal();
        }

        internal Dictionary<string, object> EditableInputs()
        {
            return EditableInputNames().Distinct().ToDictionary(n => n, InputValue, StringComparer.Ordinal);
        }

        internal bool SameInputs(InstallationConfiguration other)
        {
            if (other == null)
                return false;
            foreach (var pair in EditableInputs())
            {
                var a = pair.Value;
                var b = other.InputValue(pair.Key);
                if (a is SectionTypeChoice && b is SectionTypeChoice)
                {
                    if (((SectionTypeChoice)a).Key != ((SectionTypeChoice)b).Key)
                        return false;
                }
                else if (a is double && b is double)
                {
                    if (!DimensionValue.IsFinite((double)b) || Math.Abs((double)a - (double)b) > 0.001)
                        return false;
                }
                else if (!Equals(a, b))
                    return false;
            }
            return true;
        }

        internal static InstallationConfiguration FromProjectInputs(SectionCatalog basis, IDictionary<string, object> values,
            IDictionary<string, string> readErrors, InstallationConfiguration savedType = null)
        {
            var result = CreateDraft(InputDefaults(basis));
            result.CopySourceTypeName = BaseFamilyTypeName;
            var errors = new Dictionary<string, string>(readErrors, StringComparer.Ordinal);
            Func<string, object, bool> usable = (name, value) =>
            {
                if (value == null)
                    return false;
                DimensionValue dimension;
                if (result.Dimensions().TryGetValue(name, out dimension))
                    return value is double && DimensionValue.IsFinite((double)value) && (dimension.AllowsSigned || (double)value > 0);
                if (name == "Секции_Промежуточные_Количество")
                    return value is int && (int)value >= 1 && (int)value <= 9;
                if (name == ValveControl.ParameterName || result.SharedBooleans.ContainsKey(name))
                    return value is int && ((int)value == 0 || (int)value == 1);
                if (result.Info.Assignments().ContainsKey(name))
                    return value is string;
                for (int i = 0; i < result.Blocks.Count; i++)
                    if (ParameterName(result.SlotAt(i)) == name)
                        return value is SectionTypeChoice && result.Catalog[result.SlotAt(i)].Choices.Any(c => c.Key == ((SectionTypeChoice)value).Key);
                return false;
            };
            Action<string> apply = name =>
            {
                object value;
                values.TryGetValue(name, out value);
                bool unread = value == null || errors.ContainsKey(name);
                DimensionValue dimension;
                bool emptySize = result.Dimensions().TryGetValue(name, out dimension) && !dimension.AllowsSigned
                    && value is double && (double)value == 0;
                if (unread || emptySize || name == SharedParameters.ServiceRight && !usable(name, value))
                {
                    object saved = savedType?.InputValue(name);
                    if (usable(name, saved))
                    {
                        value = saved;
                        errors.Remove(name);
                        result.RecoveredInputs[name] = "из одноимённого сохранённого типа";
                    }
                    else if (!errors.ContainsKey(name) || name == SharedParameters.ServiceRight)
                    {
                        object fallback = result.SuggestedInputValue(name);
                        if (name == SharedParameters.ServiceRight && fallback == null)
                            fallback = 1;
                        if (name == SharedParameters.ServiceDepth && !usable(name, fallback))
                            fallback = 1000.0;
                        for (int i = 0; i < result.Blocks.Count; i++)
                            if (name == DimensionParameterName(result.SlotAt(i), "Длина") && result.Blocks[i].Type?.IsEmpty == true)
                                fallback = 1.0;
                        if (usable(name, fallback))
                        {
                            value = fallback;
                            errors.Remove(name);
                            result.RecoveredInputs[name] = "значение по умолчанию конфигуратора";
                        }
                    }
                }
                if (errors.ContainsKey(name))
                    return;
                try
                {
                    result.SetInputValue(name, value);
                }
                catch (Exception ex) { errors[name] = ex.Message; }
            };
            apply(ValveControl.ParameterName);
            apply("Секции_Промежуточные_Количество");
            for (int i = 0; i < result.Blocks.Count; i++)
                apply(ParameterName(result.SlotAt(i)));
            var names = result.EditableInputNames().Distinct().ToList();
            foreach (string name in names)
            {
                if (name == ValveControl.ParameterName || name == "Секции_Промежуточные_Количество"
                    || Enumerable.Range(0, result.Blocks.Count).Any(i => ParameterName(result.SlotAt(i)) == name))
                    continue;
                apply(name);
            }
            foreach (var error in errors.Where(p => names.Contains(p.Key)))
            {
                if (result.Dimensions().ContainsKey(error.Key) || result.SharedBooleans.ContainsKey(error.Key)
                    || Enumerable.Range(0, result.Blocks.Count).Any(i => ParameterName(result.SlotAt(i)) == error.Key))
                    result.SetInputValue(error.Key, null);
                result.ReviewIssues.Add(new ConfigurationIssue
                {
                    ParameterName = error.Key,
                    Message = error.Value,
                    InitialValue = result.InputToken(error.Key),
                    SuggestedValue = result.SuggestedInputValue(error.Key)
                });
            }
            foreach (var dimension in result.Dimensions().Values.Where(d => d.IsCalculated))
                dimension.Invalidate();
            result.RecalculateLocal();
            return result;
        }

        private string _calculationInputSignature;
        internal List<SectionValue> Blocks = new List<SectionValue>();
        // Каталог содержит имена и исходные размеры, никаких ElementId из закрытого документа.
        internal SectionCatalog Catalog;
        internal string CopySourceTypeName;
        internal InstallationInfo Info = new InstallationInfo();
        // Один набор значений на тип установки: все подходящие секции используют те же объекты.
        internal Dictionary<string, DimensionValue> SharedDimensions = new Dictionary<string, DimensionValue>();
        internal Dictionary<string, BooleanValue> SharedBooleans = new Dictionary<string, BooleanValue>();
        internal DimensionValue InstallationWidth, InstallationHeight, FrameHeight;
        internal LocalGeometry Geometry = new LocalGeometry();

        internal bool CanChangeValve
        {
            get
            {
                return ValveControl.CanEdit(Catalog.Dependencies.Find(ValveControl.ParameterName));
            }
        }

        private SectionValue _valve;

        internal bool HasValve
        {
            get
            {
                return Blocks.Count > 1 && Blocks[1].IsValve;
            }
        }

        internal int Minimum
        {
            get
            {
                return HasValve ? 4 : 3;
            }
        }

        internal int Capacity
        {
            get
            {
                return HasValve ? 12 : 11;
            }
        }

        internal int SectionCount
        {
            get
            {
                return Blocks.Count;
            }
        }

        internal int IntermediateCount
        {
            get
            {
                return Blocks.Count(b => !b.IsFixed);
            }
        }

        internal int FirstSectionIndex
        {
            get
            {
                return HasValve ? 2 : 1;
            }
        }

        private static int SlotAt(int index, IList<SectionValue> blocks)
        {
            if (index == 0)
                return 1;
            if (index == blocks.Count - 1)
                return 12;
            if (blocks[index].IsValve)
                return 2;
            return index + (blocks[1].IsValve ? 1 : 2);
        }

        internal int SlotAt(int index)
        {
            return SlotAt(index, Blocks);
        }

        internal string BlockTitle(int index)
        {
            if (Blocks[index].IsConnector)
                return "Соединитель";
            if (Blocks[index].IsValve)
                return "Клапан";
            return "Секция " + (index - FirstSectionIndex + 1);
        }

        internal static string ParameterName(int slot)
        {
            if (slot < 1 || slot > 12)
                throw new ArgumentOutOfRangeException(nameof(slot));
            if (slot == 1 || slot == 12)
                return "Секция_" + slot + "_Соединитель_Тип";
            if (slot == 2)
                return "Секция_2_Соединитель_Клапан_Тип";
            return "Секция_" + slot + "_Промежуточная_" + (slot - 2) + "_Тип";
        }

        internal static string DimensionParameterName(int slot, string dimension)
        {
            if (dimension != "Длина" && dimension != "Ширина" && dimension != "Высота" && dimension != "Смещение по X" && dimension != "Смещение по Y")
                throw new ArgumentException("Неизвестный размер секции.");
            if (dimension != "Длина" && slot != 1 && slot != 12)
                throw new ArgumentException("Ширина, высота и смещения задаются только для соединителей.");
            string typeParameter = ParameterName(slot);
            return typeParameter.Substring(0, typeParameter.Length - "Тип".Length) + dimension;
        }

        internal static IEnumerable<string> InterfaceParameterNames()
        {
            foreach (string name in new InstallationInfo().Assignments().Keys)
                yield return name;
            foreach (string name in SharedParameters.LengthNames.Concat(SharedParameters.BooleanNames))
                yield return name;
            yield return "Секции_Промежуточные_Количество";
            yield return "Соединитель_Приточный_Клапан";
            yield return "Установка_Ширина";
            yield return "Установка_Высота";
            yield return "Основание_Рама_Высота";
            for (int slot = 1; slot <= 12; slot++)
            {
                yield return ParameterName(slot);
                yield return DimensionParameterName(slot, "Длина");
                if (slot != 1 && slot != 12)
                    continue;
                yield return DimensionParameterName(slot, "Ширина");
                yield return DimensionParameterName(slot, "Высота");
                yield return DimensionParameterName(slot, "Смещение по X");
                yield return DimensionParameterName(slot, "Смещение по Y");
            }
        }

        internal static InstallationConfiguration CreateDraft(SectionCatalog catalog, bool newTypeDefaults = false)
        {
            if (catalog == null || catalog.Count != 12 || !Enumerable.Range(1, 12).All(catalog.ContainsKey))
                throw new InvalidOperationException("Не загружен состав секций исходного семейства.");
            if (catalog.IntermediateCount < 1 || catalog.IntermediateCount > 9)
                throw new InvalidOperationException("В исходном типе некорректное количество промежуточных секций.");
            var configuration = new InstallationConfiguration
            {
                Catalog = catalog,
                InstallationWidth = catalog.Dimension("Установка_Ширина", catalog.InstallationWidthMm),
                InstallationHeight = catalog.Dimension("Установка_Высота", catalog.InstallationHeightMm),
                FrameHeight = catalog.Dimension("Основание_Рама_Высота", catalog.FrameHeightMm)
            };
            foreach (string name in SharedParameters.LengthNames)
            {
                double value;
                configuration.SharedDimensions.Add(name, catalog.Dimension(name, catalog.SharedLengthsMm.TryGetValue(name, out value) ? value : double.NaN));
            }
            foreach (string name in SharedParameters.BooleanNames)
            {
                bool? value;
                configuration.SharedBooleans.Add(name, new BooleanValue(catalog.SharedFlags.TryGetValue(name, out value) ? value : null,
                    catalog.Dependencies.IsCalculated(name)));
            }
            configuration._valve = configuration.NewBlock(2, valve: true, resetEmptyLength: false);
            configuration.Blocks.Add(configuration.NewBlock(1, connector: true, resetEmptyLength: false));
            if (catalog.HasValve)
                configuration.Blocks.Add(configuration._valve);
            for (int slot = 3; slot < 3 + catalog.IntermediateCount; slot++)
                configuration.Blocks.Add(configuration.NewBlock(slot, resetEmptyLength: false));
            configuration.Blocks.Add(configuration.NewBlock(12, connector: true, resetEmptyLength: false));
            configuration.RecalculateLocal();
            if (newTypeDefaults)
            {
                configuration.SharedDimensions[SharedParameters.ServiceDepth].Text = "1000";
                configuration.SharedBooleans[SharedParameters.ServiceRight].Value = true;
                foreach (var block in configuration.Blocks.Concat(new[] { configuration._valve }).Distinct())
                    block.ResetEmptyLength();
                configuration.RecalculateLocal();
            }
            return configuration;
        }

        private SectionValue NewBlock(int slot, bool connector = false, bool valve = false, bool resetEmptyLength = true)
        {
            var definition = Catalog[slot];
            var block = new SectionValue
            {
                IsConnector = connector,
                IsValve = valve,
                Type = definition.SourceType,
                Length = Catalog.Dimension(DimensionParameterName(slot, "Длина"), definition.DefaultLengthMm),
                Width = connector ? Catalog.Dimension(DimensionParameterName(slot, "Ширина"), definition.DefaultWidthMm) : null,
                Height = connector ? Catalog.Dimension(DimensionParameterName(slot, "Высота"), definition.DefaultHeightMm) : null,
                OffsetX = connector ? Catalog.Dimension(DimensionParameterName(slot, "Смещение по X"), definition.DefaultOffsetXMm) : null,
                OffsetY = connector ? Catalog.Dimension(DimensionParameterName(slot, "Смещение по Y"), definition.DefaultOffsetYMm) : null
            };
            if (resetEmptyLength)
                block.ResetEmptyLength();
            return block;
        }

        internal InstallationConfiguration Copy()
        {
            var copy = new InstallationConfiguration
            {
                Catalog = Catalog,
                CopySourceTypeName = CopySourceTypeName,
                Info = Info.Copy(),
                SharedDimensions = SharedDimensions.ToDictionary(p => p.Key, p => p.Value.Copy()),
                SharedBooleans = SharedBooleans.ToDictionary(p => p.Key, p => p.Value.Copy()),
                InstallationWidth = InstallationWidth.Copy(),
                InstallationHeight = InstallationHeight.Copy(),
                FrameHeight = FrameHeight.Copy(),
                Blocks = Blocks.Select(b => b.Copy()).ToList(),
                _calculationInputSignature = _calculationInputSignature,
                ReviewIssues = ReviewIssues.Select(i => i.Copy()).ToList(),
                RecoveredInputs = new Dictionary<string, string>(RecoveredInputs)
            };
            copy._valve = HasValve ? copy.Blocks[1] : _valve.Copy();
            copy.RecalculateLocal();
            return copy;
        }

        internal void SetValve(bool enabled)
        {
            if (enabled == HasValve)
                return;
            if (!CanChangeValve)
                throw new InvalidOperationException("Наличие клапана задано формулой или свойствами исходного семейства.\n"
                + Catalog.Dependencies.Tooltip("Соединитель_Приточный_Клапан"));
            var candidate = Blocks.ToList();
            if (enabled)
                candidate.Insert(1, _valve);
            else
                candidate.RemoveAt(1);
            ValidateChoices(candidate, allowUnselected: true);
            Blocks = candidate;
            RefreshDimensionModes();
            // При выключении сохраняем выбранный состав клапана для следующего включения.
        }

        internal void Resize(int count)
        {
            if (count < Minimum || count > Capacity)
                throw new ArgumentException("Количество блоков должно быть от " + Minimum + " до " + Capacity + ".");
            var candidate = Blocks.ToList();
            while (candidate.Count > count)
                candidate.RemoveAt(candidate.Count - 2);
            while (candidate.Count < count)
                candidate.Insert(candidate.Count - 1, NewBlock(3 + candidate.Count(b => !b.IsFixed)));
            ValidateChoices(candidate, allowUnselected: true);
            Blocks = candidate;
            RefreshDimensionModes();
        }

        internal void InsertAfter(int index)
        {
            if (SectionCount >= Capacity)
                throw new InvalidOperationException("Достигнут предел секций исходного семейства.");
            int target = Math.Max(FirstSectionIndex, Math.Min(index + 1, Blocks.Count - 1));
            var candidate = Blocks.ToList();
            candidate.Insert(target, NewBlock(target + (HasValve ? 1 : 2)));
            ValidateChoices(candidate, allowUnselected: true);
            Blocks = candidate;
            RefreshDimensionModes();
        }

        internal void Remove(int index)
        {
            if (index < FirstSectionIndex || index >= Blocks.Count - 1 || SectionCount <= Minimum)
                return;
            var candidate = Blocks.ToList();
            candidate.RemoveAt(index);
            ValidateChoices(candidate, allowUnselected: true);
            Blocks = candidate;
            RefreshDimensionModes();
        }

        internal void Move(int index, int target)
        {
            if (index < FirstSectionIndex || target < FirstSectionIndex || index >= Blocks.Count - 1 || target >= Blocks.Count - 1)
                return;
            var candidate = Blocks.ToList();
            var block = candidate[index];
            candidate.RemoveAt(index);
            candidate.Insert(target, block);
            ValidateChoices(candidate, allowUnselected: true);
            Blocks = candidate;
            RefreshDimensionModes();
        }

        private void ValidateChoices(IList<SectionValue> blocks, bool allowUnselected = false)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                int slot = SlotAt(i, blocks);
                if (blocks[i].Type == null && allowUnselected)
                    continue;
                if (blocks[i].Type == null || !Catalog[slot].Choices.Any(c => c.Key == blocks[i].Type.Key))
                    throw new InvalidOperationException("Выбранный состав недоступен для параметра «" + Catalog[slot].ParameterName + "».");
            }
        }

        internal void Validate()
        {
            if (SectionCount < Minimum || SectionCount > Capacity || IntermediateCount < 1 || IntermediateCount > 9
                || !Blocks[0].IsConnector || !Blocks.Last().IsConnector
                || Blocks.Count(b => b.IsConnector) != 2 || Blocks.Count(b => b.IsValve) != (HasValve ? 1 : 0))
                throw new ArgumentException("Проверьте количество секций и положение клапана.");
            ValidateChoices(Blocks);
            ValidateDimensions();
            foreach (var flag in SharedBooleans.Where(p => ActiveSharedParameters().Contains(p.Key)))
                if (!flag.Value.IsValid)
                    throw new ArgumentException("Задайте значение «" + SharedParameters.Label(flag.Key) + "».");
            if (!CanChangeValve && HasValve != LocalGeometry.SourceFlag(Catalog, "Соединитель_Приточный_Клапан", true))
                throw new InvalidOperationException("Состояние клапана не соответствует формуле исходного семейства.");
        }

        private void ValidateDimensions()
        {
            foreach (var dimension in Dimensions())
                if (!dimension.Value.IsCalculated)
                    dimension.Value.Millimeters(dimension.Key);
        }

        internal string Signature()
        {
            // Сравниваем сохраняемые входы, а не результаты формул или ссылки на объекты ComboBox.
            var values = Dimensions().Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal).Select(p =>
            {
                double value;
                return p.Key + "=" + (p.Value.TryNumber(out value) ? value.ToString("R", CultureInfo.InvariantCulture) : p.Value.Text);
            });
            var fields = new[] { HasValve.ToString() }.Concat(Blocks.Select(b => b.Type?.Key ?? ""))
                .Concat(values).Concat(SharedBooleans.Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                    .Select(p => p.Key + "=" + p.Value.Value))
                .Concat(Info.Assignments().OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Key + "=" + p.Value));
            return string.Concat(fields.Select(value => value.Length + ":" + value));
        }

        internal void ValidateForOutput()
        {
            var problems = Problems();
            if (problems.Count > 0)
                throw new InvalidOperationException("Проверьте параметры в конфигураторе:\n"
                + string.Join("\n", problems.Select(p => p.ParameterName + ": " + p.Message)));
            var missing = IncompleteFields();
            if (missing.Count > 0)
                throw new InvalidOperationException("Заполните параметры установки:\n" + string.Join("\n", missing));
            Validate();
        }

        internal List<string> IncompleteFields()
        {
            var fields = new List<string>();
            Action<string, DimensionValue> check = (label, value) => { if (value == null || !value.IsValid) fields.Add(label); };
            foreach (var dimension in SharedDimensions.Where(p => ActiveSharedParameters().Contains(p.Key)))
                check(SharedParameters.Label(dimension.Key), dimension.Value);
            foreach (var flag in SharedBooleans.Where(p => ActiveSharedParameters().Contains(p.Key)))
                if (!flag.Value.IsValid)
                    fields.Add(SharedParameters.Label(flag.Key));
            check("Ширина установки", InstallationWidth);
            check("Высота установки", InstallationHeight);
            check("Высота рамы", FrameHeight);
            for (int i = 0; i < Blocks.Count; i++)
            {
                var block = Blocks[i];
                string title = block.IsConnector ? (i == 0 ? "Первый соединитель" : "Последний соединитель") : BlockTitle(i);
                if (block.Type == null)
                    fields.Add(title + ": тип оборудования");
                check(title + ": длина", block.Length);
                if (!block.IsConnector)
                    continue;
                check(title + ": ширина", block.Width);
                check(title + ": высота", block.Height);
                check(title + ": смещение по X", block.OffsetX);
                check(title + ": смещение по Y", block.OffsetY);
            }
            return fields;
        }

        internal Dictionary<string, DimensionValue> Dimensions()
        {
            var result = new Dictionary<string, DimensionValue> { { "Установка_Ширина", InstallationWidth },
                    { "Установка_Высота", InstallationHeight }, { "Основание_Рама_Высота", FrameHeight } };
            foreach (var dimension in SharedDimensions.Where(p => ActiveSharedParameters().Contains(p.Key)))
                result.Add(dimension.Key, dimension.Value);
            for (int i = 0; i < Blocks.Count; i++)
            {
                var block = Blocks[i];
                int slot = SlotAt(i);
                result[DimensionParameterName(slot, "Длина")] = block.Length;
                if (!block.IsConnector)
                    continue;
                result[DimensionParameterName(slot, "Ширина")] = block.Width;
                result[DimensionParameterName(slot, "Высота")] = block.Height;
                result[DimensionParameterName(slot, "Смещение по X")] = block.OffsetX;
                result[DimensionParameterName(slot, "Смещение по Y")] = block.OffsetY;
            }
            return result;
        }

        internal void RefreshDimensionModes(Func<string, bool> isCalculated = null)
        {
            foreach (var dimension in Dimensions())
                dimension.Value.SetCalculatedMode(isCalculated == null ? Catalog.Dependencies.IsCalculated(dimension.Key) : isCalculated(dimension.Key));
        }

        internal void RefreshBooleanModes(Func<string, bool> isCalculated = null)
        {
            foreach (var flag in SharedBooleans)
                flag.Value.SetCalculatedMode(isCalculated == null ? Catalog.Dependencies.IsCalculated(flag.Key) : isCalculated(flag.Key));
        }

        internal Dictionary<string, int> BooleanAssignments()
        {
            return SharedBooleans.Where(p => ActiveSharedParameters().Contains(p.Key) && !p.Value.IsCalculated && p.Value.Value.HasValue)
            .ToDictionary(p => p.Key, p => p.Value.Value.Value ? 1 : 0);
        }

        internal void RecalculateLocal()
        {
            RefreshDimensionModes();
            RefreshBooleanModes();
            string inputs = HasValve + "|" + string.Join("|", Blocks.Select(b => b.Type?.Key ?? ""))
                + "|" + string.Join("|", Dimensions().Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                    .Select(p => p.Key + "=" + p.Value.Text))
                + "|" + string.Join("|", SharedBooleans.Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                    .Select(p => p.Key + "=" + p.Value.Value));
            // При чтении показываем результаты исходного типа. Они устаревают только после изменения входов.
            if (_calculationInputSignature != null && _calculationInputSignature != inputs)
                foreach (var dimension in Dimensions().Values)
                    dimension.Invalidate();
            _calculationInputSignature = inputs;
            Geometry = LocalGeometry.Calculate(this);
        }

        internal DimensionValue BlockWidth(int index)
        {
            return Blocks[index].IsConnector ? Blocks[index].Width : Blocks[index].IsValve ? Blocks[0].Width : InstallationWidth;
        }

        internal DimensionValue BlockHeight(int index)
        {
            return Blocks[index].IsConnector ? Blocks[index].Height : Blocks[index].IsValve ? Blocks[0].Height : InstallationHeight;
        }

        private IEnumerable<int> InactiveLengthSlots()
        {
            var active = new HashSet<int>(Enumerable.Range(0, Blocks.Count).Select(SlotAt));
            for (int slot = 1; slot <= 12; slot++)
                if (!active.Contains(slot) && !Catalog.Dependencies.IsCalculated(DimensionParameterName(slot, "Длина")))
                    yield return slot;
        }

        internal bool NeedsInactiveLengthReset
        {
            get
            {
                return InactiveLengthSlots().Any(slot => !DimensionValue.IsFinite(Catalog[slot].DefaultLengthMm)
                || Math.Abs(Catalog[slot].DefaultLengthMm - 1.0) > 0.000001);
            }
        }

        internal Dictionary<string, double> DimensionAssignments()
        {
            Validate();
            var result = new Dictionary<string, double>();
            foreach (var dimension in Dimensions())
                if (!dimension.Value.IsCalculated)
                    result[dimension.Key] = dimension.Value.Millimeters(dimension.Key);
                else
                    result.Remove(dimension.Key);
            // Отключённый клапан и скрытые промежуточные позиции имеют длину 1 мм.
            // Видимые блоки сохраняют заданную длину, в том числе вручную изменённый пустой блок.
            foreach (int slot in InactiveLengthSlots())
                result[DimensionParameterName(slot, "Длина")] = 1.0;
            return result;
        }

        internal Dictionary<int, SectionTypeChoice> SectionAssignments()
        {
            Validate();
            var result = new Dictionary<int, SectionTypeChoice>();
            for (int index = 0; index < Blocks.Count; index++)
                result[SlotAt(index)] = Blocks[index].Type;
            return result;
        }
    }
}
