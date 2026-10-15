using System;
using System.Collections.Generic;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class DepartmentPromptSettings
    {
        internal const string ArchitectureKey = "ExternalModel.DepartmentPrompt.AR";
        internal const string StructureKey = "ExternalModel.DepartmentPrompt.KR";
        internal const string EngineeringKey = "ExternalModel.DepartmentPrompt.IOS";

        public string Architecture { get; set; } = string.Empty;
        public string Structure { get; set; } = string.Empty;
        public string Engineering { get; set; } = string.Empty;

        public DepartmentPromptSettings Clone()
        {
            return new DepartmentPromptSettings
            {
                Architecture = Architecture, Structure = Structure, Engineering = Engineering
            };
        }

        public string GetText(string profile)
        {
            switch (profile)
            {
                case "AR": return Architecture ?? string.Empty;
                case "KR": return Structure ?? string.Empty;
                case "IOS": return Engineering ?? string.Empty;
                default: return string.Empty;
            }
        }

        public void SetText(string profile, string text)
        {
            switch (profile)
            {
                case "AR": Architecture = text ?? string.Empty; break;
                case "KR": Structure = text ?? string.Empty; break;
                case "IOS": Engineering = text ?? string.Empty; break;
                default: throw new ArgumentException("Unknown department prompt profile.", nameof(profile));
            }
        }

        public static string ResolveProfile(int subDepartmentId, Func<int, int?> getParentId = null)
        {
            var visited = new HashSet<int>();
            while (subDepartmentId > 0 && visited.Count < 16 && visited.Add(subDepartmentId))
            {
                switch (subDepartmentId)
                {
                    case 2: return "AR";
                    case 3: return "KR";
                    case 4: case 5: case 6: case 7:
                    case 41: case 51: case 71: case 72: return "IOS";
                    case 1: case 8: return null; // ALL and BIM have no department-specific instructions.
                }
                if (getParentId == null) return null;
                int? parent = getParentId(subDepartmentId);
                if (!parent.HasValue) return null;
                subDepartmentId = parent.Value;
            }
            return null;
        }

        public string BuildContext(string profile)
        {
            string text = GetText(profile);
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            string name = profile == "AR" ? "АР" : profile == "KR" ? "КР" : "ИОС";
            return "\r\n\r\n## Дополнительные инструкции отдела " + name
                + "\r\nОтдел пользователя: " + name + "."
                + "\r\nЭти инструкции уточняют общие правила. При противоречии приоритет имеют общие правила."
                + "\r\n\r\n" + text.Trim();
        }
    }
}
