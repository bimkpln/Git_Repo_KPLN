using System;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace KPLN_NavisMcpBridge
{
    /// <summary>
    /// Точка входа аддина. Navisworks добавляет кнопку на вкладку "Add-ins" —
    /// первый клик поднимает локальный HTTP-мост, повторный клик его останавливает.
    ///
    /// VERIFY: атрибут [Plugin] и базовый класс AddInPlugin — актуальны для
    /// managed API начиная с Navisworks ~2013 и не менялись годами, но стоит
    /// свериться с образцом из SDK 2020 (обычно лежит рядом с установкой или
    /// в отдельно ставящемся Navisworks 2020 Developer's Guide / SDK).
    /// </summary>
    [Plugin("KPLN_NavisMcpBridge.Bridge", "KPLN",
        DisplayName = "KPLN MCP Bridge",
        ToolTip = "Запустить/остановить локальный HTTP-мост для MCP")]
    public class BridgePlugin : AddInPlugin
    {
        private static HttpBridgeServer _server;

        public override int Execute(params string[] parameters)
        {
            try
            {
                if (_server == null)
                {
                    _server = new HttpBridgeServer(port: 8765);
                    _server.Start();
                    MessageBox.Show(
                        "MCP-мост запущен на http://127.0.0.1:8765",
                        "KPLN_NavisMcpBridge",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _server.Stop();
                    _server = null;
                    MessageBox.Show(
                        "MCP-мост остановлен",
                        "KPLN_NavisMcpBridge",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка моста: " + ex,
                    "KPLN_NavisMcpBridge",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }

            return 0;
        }
    }
}
