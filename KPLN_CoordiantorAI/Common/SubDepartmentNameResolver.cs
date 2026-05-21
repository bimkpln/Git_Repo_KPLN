namespace KPLN_CoordiantorAI.Common
{
    public static class SubDepartmentNameResolver
    {
        public static SubDepartmentInfo[] GetKnownSubDepartments()
        {
            return new[]
            {
                new SubDepartmentInfo { Id = 2, Name = "АР" },
                new SubDepartmentInfo { Id = 3, Name = "КР" },
                new SubDepartmentInfo { Id = 4, Name = "ОВиК" },
                new SubDepartmentInfo { Id = 5, Name = "ВК" },
                new SubDepartmentInfo { Id = 6, Name = "ЭОМ" },
                new SubDepartmentInfo { Id = 7, Name = "СС" },
                new SubDepartmentInfo { Id = 8, Name = "BIM" },
                new SubDepartmentInfo { Id = 41, Name = "ИТП" },
                new SubDepartmentInfo { Id = 51, Name = "ПТ" },
                new SubDepartmentInfo { Id = 71, Name = "АВ" },
                new SubDepartmentInfo { Id = 72, Name = "СПС" }
            };
        }

        public static string GetName(int subDepartmentId)
        {
            switch (subDepartmentId)
            {
                case 1:
                    return "ALL";
                case 2:
                    return "АР";
                case 3:
                    return "КР";
                case 4:
                    return "ОВиК";
                case 5:
                    return "ВК";
                case 6:
                    return "ЭОМ";
                case 7:
                    return "СС";
                case 8:
                    return "BIM";
                case 41:
                    return "ИТП";
                case 51:
                    return "ПТ";
                case 71:
                    return "АВ";
                case 72:
                    return "СПС";
                default:
                    return subDepartmentId.ToString();
            }
        }
    }
}