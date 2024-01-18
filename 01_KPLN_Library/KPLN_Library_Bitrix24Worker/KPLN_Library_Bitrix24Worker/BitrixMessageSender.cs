using System;
using System.Net.Http;
using System.Windows.Forms;

namespace KPLN_Library_Bitrix24Worker
{
    /// <summary>
    /// Сервис по отправке сообщений в Битрикс
    /// </summary>
    public class BitrixMessageSender
    {
        /// <summary>
        /// Отправить сообщение в чат бим-отдела от имени спец. пользователя
        /// </summary>
        /// <param name="msg"></param>
        /// <exception cref="Exception"></exception>
        public static async void SendErrorMsg_ToBIMChat(string msg)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client
                        .GetAsync(String.Format(@"https://kpln.bitrix24.ru/rest/1310/qd00y541wgy6wyyz/im.message.add.json?MESSAGE={0}&DIALOG_ID=chat4240", $"{msg}"));
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
