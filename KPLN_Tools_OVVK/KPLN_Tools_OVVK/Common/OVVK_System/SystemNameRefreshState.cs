using Autodesk.Revit.DB;
using KPLN_Library_DBWorker;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    /// <summary>Общая для модели метка обязательного перестроения имён систем.</summary>
    internal sealed class SystemNameRefreshState
    {
        private readonly Document _doc;
        private bool _pendingError;

        private SystemNameRefreshState(Document doc) { _doc = doc; }

        internal bool IsRequired { get; private set; }
        internal string Token { get; private set; }
        internal string StorageError { get; private set; }

        internal static SystemNameRefreshState ForDocument(Document doc) => new SystemNameRefreshState(doc);

        internal void Observe(bool hasAbbreviationErrors)
        {
            // Не теряем обнаруженную ошибку, если в момент проверки сервер был недоступен.
            _pendingError |= hasAbbreviationErrors;
            Access(state =>
            {
                if (_pendingError)
                {
                    // Новое обнаружение ошибки отменяет подтверждение из ранее начатого сеанса.
                    state.Required = true;
                    state.Token = Guid.NewGuid().ToString("N");
                    return true;
                }
                return false;
            });
        }

        internal void Complete(string refreshToken, bool hasAbbreviationErrors, bool refreshSucceeded)
        {
            if (hasAbbreviationErrors || !refreshSucceeded)
            {
                Observe(hasAbbreviationErrors);
                return;
            }

            Access(state =>
            {
                // Другой сеанс мог обнаружить новую ошибку во время обновления.
                if (!_pendingError && state.Required && !string.IsNullOrEmpty(refreshToken) && state.Token == refreshToken)
                {
                    state.Required = false;
                    return true;
                }
                return false;
            });
        }

        private void Access(Func<StoredState, bool> update)
        {
            try
            {
                SystemProjectConfig.Access(_doc, root =>
                {
                    var model = SystemProjectConfig.Model(root, SystemProjectConfig.ModelName(_doc));
                    var state = model["Refresh"].ToObject<StoredState>();
                    if (state == null || (state.Required && string.IsNullOrEmpty(state.Token)))
                        throw new InvalidDataException("Повреждён статус обновления систем.");
                    bool changed = update(state);
                    if (changed)
                        model["Refresh"] = Newtonsoft.Json.Linq.JObject.FromObject(state);
                    IsRequired = state.Required;
                    Token = state.Token;
                    return changed;
                });
                _pendingError = false;
                StorageError = null;
            }
            catch (Exception ex)
            {
                IsRequired |= _pendingError;
                StorageError = "Не удалось прочитать или сохранить метку обновления систем: " + ex.Message
                    + " Проверь доступ к конфигу и нажми «Обновить список».";
            }
        }

        private sealed class StoredState
        {
            public bool Required { get; set; }
            public string Token { get; set; }
        }
    }
}
