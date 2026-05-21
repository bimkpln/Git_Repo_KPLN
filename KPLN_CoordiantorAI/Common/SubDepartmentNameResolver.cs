namespace KPLN_CoordiantorAI.Common
{
    public static class SubDepartmentNameResolver
    {
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