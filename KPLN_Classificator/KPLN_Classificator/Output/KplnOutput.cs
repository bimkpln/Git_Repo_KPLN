#if Revit2020 || Revit2018
using KPLN_Loader;
using static KPLN_Loader.Output.Output;
#endif
using System;
using System.Collections.Generic;
using System.Text;

namespace KPLN_Classificator
{
    public class KplnOutput : Output
    {
#if Revit2020 || Revit2018
        private Dictionary<OutputMessageType, Preferences.MessageType> typesMap
    = new Dictionary<OutputMessageType, Preferences.MessageType>();

        public KplnOutput()
        {
            typesMap.Add(OutputMessageType.Code, Preferences.MessageType.Code);
            typesMap.Add(OutputMessageType.Critical, Preferences.MessageType.Critical);
            typesMap.Add(OutputMessageType.Error, Preferences.MessageType.Error);
            typesMap.Add(OutputMessageType.Header, Preferences.MessageType.Header);
            typesMap.Add(OutputMessageType.Regular, Preferences.MessageType.Regular);
            typesMap.Add(OutputMessageType.Success, Preferences.MessageType.Success);
            typesMap.Add(OutputMessageType.System_OK, Preferences.MessageType.System_OK);
            typesMap.Add(OutputMessageType.System_Regular, Preferences.MessageType.System_Regular);
            typesMap.Add(OutputMessageType.Warning, Preferences.MessageType.Warning);
        }

        public override void PrintInfo(string value, OutputMessageType type)
        {
            Print(value, typesMap[type]);
            log.AppendLine(string.Format("[{0}] {1} {2}", type, DateTime.UtcNow, value));
        }

        public override void PrintDebug(string value, OutputMessageType type, bool check)
        {
            if (check)
            {
                PrintInfo(value, type);
            }
            else
            {
                log.AppendLine(string.Format("[{0}] {1} {2}", type, DateTime.UtcNow, value));
            }
        }

        public override void PrintErr(Exception e, string value)
        {
            PrintError(e, value);
            log.AppendLine(string.Format("[{0}] {1} {2}", typesMap[OutputMessageType.Error], DateTime.UtcNow, value));
        }

#endif
#if Revit2020Std || Revit2018Std
        public override void PrintDebug(string value, OutputMessageType type, bool check)
        {
            throw new NotImplementedException();
        }

        public override void PrintErr(Exception e, string value)
        {
            throw new NotImplementedException();
        }

        public override void PrintInfo(string value, OutputMessageType type)
        {
            throw new NotImplementedException();
        }

#endif
    }
}
