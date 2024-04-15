using System;
using System.Collections.Generic;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Classificator
{
    public class KplnOutput : Output
    {
        private Dictionary<OutputMessageType, MessageType> typesMap = new Dictionary<OutputMessageType, MessageType>();

        public KplnOutput()
        {
            typesMap.Add(OutputMessageType.Code, MessageType.Code);
            typesMap.Add(OutputMessageType.Critical, MessageType.Critical);
            typesMap.Add(OutputMessageType.Error, MessageType.Error);
            typesMap.Add(OutputMessageType.Header, MessageType.Header);
            typesMap.Add(OutputMessageType.Regular, MessageType.Regular);
            typesMap.Add(OutputMessageType.Success, MessageType.Success);
            typesMap.Add(OutputMessageType.System_OK, MessageType.System_OK);
            typesMap.Add(OutputMessageType.System_Regular, MessageType.System_Regular);
            typesMap.Add(OutputMessageType.Warning, MessageType.Warning);
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
    }
}
