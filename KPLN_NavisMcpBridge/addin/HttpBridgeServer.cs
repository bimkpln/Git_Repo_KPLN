using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Newtonsoft.Json;

namespace KPLN_NavisMcpBridge
{
    /// <summary>
    /// Простой JSON/HTTP мост поверх HttpListener. Слушает только localhost —
    /// наружу не торчит. Каждый обработчик маршалит вызов Navisworks API на
    /// главный поток через MainThreadDispatcher.
    ///
    /// Роуты (MVP, Clash Detective). У COM-объектов Clash Detective в
    /// Navisworks 2020 нет стабильного Guid — тесты и результаты
    /// идентифицируются по имени (name), поэтому в путях используется имя
    /// (URL-encoded на стороне клиента).
    ///   GET  /clash/tests
    ///   GET  /clash/tests/{testName}/results
    ///         ?includePaths=false&status=Active&includeItemBounds=true
    ///          &includeSizes=true
    ///   POST /clash/tests/{testName}/run
    ///   POST /clash/results/{testName}/{resultName}/status   body: {status, comment}
    ///   GET  /clash/tests/{testName}/export?format=html
    /// </summary>
    internal sealed class HttpBridgeServer
    {
        private readonly int _port;
        private HttpListener _listener;
        private Thread _thread;
        private MainThreadDispatcher _dispatcher;
        private volatile bool _running;

        public HttpBridgeServer(int port)
        {
            _port = port;
        }

        public void Start()
        {
            // Диспетчер создаётся здесь — Start() вызывается из Execute(),
            // то есть уже на главном потоке Navisworks.
            _dispatcher = new MainThreadDispatcher();

            _listener = new HttpListener();
            _listener.Prefixes.Add($"http://127.0.0.1:{_port}/");
            _listener.Start();
            _running = true;

            _thread = new Thread(Loop) { IsBackground = true, Name = "KPLN_NavisMcpBridge-Http" };
            _thread.Start();
        }

        public void Stop()
        {
            _running = false;
            try { _listener?.Stop(); } catch { /* ignore */ }
            try { _dispatcher?.Dispose(); } catch { /* ignore */ }
        }

        private void Loop()
        {
            while (_running)
            {
                HttpListenerContext ctx;
                try
                {
                    ctx = _listener.GetContext();
                }
                catch (Exception)
                {
                    if (!_running) return;
                    continue;
                }

                // Каждый запрос — в своём потоке, чтобы долгие операции
                // (например RunTest) не блокировали приём новых соединений.
                ThreadPool.QueueUserWorkItem(_ => Handle(ctx));
            }
        }

        private void Handle(HttpListenerContext ctx)
        {
            var req = ctx.Request;
            var resp = ctx.Response;
            resp.ContentType = "application/json; charset=utf-8";
            resp.KeepAlive = false; // проще и надёжнее, чем разбираться с keep-alive у HttpListener

            // Сначала полностью вычисляем статус и payload, НЕ трогая resp —
            // так WriteJson ниже вызывается ровно один раз за запрос. Раньше
            // WriteJson могла быть вызвана из try, упасть на сериализации/
            // записи, и catch-блок пытался вызвать WriteJson ЕЩЁ раз на уже
            // тронутом resp — отсюда была ошибка "операция после передачи
            // отклика".
            int status;
            object payload;
            try
            {
                payload = Route(req);
                status = 200;
            }
            catch (InvalidOperationException ex)
            {
                status = 409;
                payload = new { error = ex.Message };
            }
            catch (NotImplementedException ex)
            {
                status = 501;
                payload = new { error = ex.Message };
            }
            catch (Exception ex)
            {
                status = 500;
                payload = new { error = ex.ToString() };
            }

            try
            {
                WriteJson(resp, status, payload);
            }
            catch
            {
                // Отклик уже не в валидном состоянии (клиент отвалился и
                // т.п.) — больше сделать ничего нельзя, просто не роняем
                // поток целиком.
            }
            finally
            {
                try { resp.OutputStream.Close(); } catch { /* ignore */ }
            }
        }

        private object Route(HttpListenerRequest req)
        {
            // ВАЖНО: используем req.RawUrl, а не req.Url.AbsolutePath.
            // Url.AbsolutePath в .NET Framework для не-ASCII сегментов пути
            // (кириллица в именах тестов/результатов) декодирует %XX
            // побайтово как отдельные code point'ы, а не как байты UTF-8 —
            // многобайтовые символы превращаются в мусор, и поиск теста по
            // имени падает с "не найден" (409), хотя тест точно есть.
            // RawUrl отдаёт путь как есть (percent-encoded, не тронутый
            // Uri-парсером) — декодируем его сами через DecodeUtf8Percent.
            var rawTarget = req.RawUrl ?? "/";
            var path = rawTarget.Split('?')[0];
            var method = req.HttpMethod;

            Match m;

            if (method == "GET" && path == "/clash/tests")
                return _dispatcher.Run(() => ClashService.ListTests());

            if (method == "GET" && (m = Regex.Match(path, @"^/clash/tests/([^/]+)/results$")).Success)
            {
                // includePaths=false пропускает самую дорогую часть —
                // резолв полного пути предков (AncestorsAndSelf) для обоих
                // элементов каждого результата. На тестах с тысячами
                // результатов это единственное, что укладывается в разумное
                // время; включайте пути, только если реально нужны имена
                // элементов, а не только статус/дистанция/имя результата.
                var includePaths = req.QueryString["includePaths"] != "false";
                // status — необязательный фильтр (New/Active/Approved/...):
                // на тестах с тысячами результатов почти всегда реально нужны
                // только "открытые" (напр. Active), а не все — фильтрация до
                // построения DTO/геометрии на порядок сокращает число
                // COM-вызовов и время ответа. Значения ASCII — decode не нужен.
                var status = req.QueryString["status"];
                // includeItemBounds — bounding box КАЖДОГО ЭЛЕМЕНТА пары
                // целиком (не зоны пересечения) — для приблизительной оценки
                // направления оси трубы (см. комментарий в ClashService).
                // Требует резолва ModelItem, как и includePaths — по
                // умолчанию выключено, чтобы не удорожать обычные запросы.
                var includeItemBounds = req.QueryString["includeItemBounds"] == "true";
                // includeSizes — номинальное сечение элемента из его свойств
                // ("ø200 мм-ø200 мм", "300x200-300x200"). По геометрии сечение
                // не берётся: габарит корпуса клапана или отвода систематически
                // больше сечения воздуховода.
                var includeSizes = req.QueryString["includeSizes"] == "true";
                return _dispatcher.Run(() => ClashService.GetResults(DecodeUtf8Percent(m.Groups[1].Value), includePaths, status, includeItemBounds, includeSizes));
            }

            if (method == "POST" && (m = Regex.Match(path, @"^/clash/tests/([^/]+)/run$")).Success)
            {
                _dispatcher.Run(() => ClashService.RunTest(DecodeUtf8Percent(m.Groups[1].Value)));
                return new { ok = true };
            }

            if (method == "POST" && (m = Regex.Match(path, @"^/clash/results/([^/]+)/([^/]+)/status$")).Success)
            {
                var body = ReadBody<StatusUpdateBody>(req);
                _dispatcher.Run(() => ClashService.SetResultStatus(
                    DecodeUtf8Percent(m.Groups[1].Value), DecodeUtf8Percent(m.Groups[2].Value),
                    body.status, body.comment));
                return new { ok = true };
            }

            if (method == "GET" && (m = Regex.Match(path, @"^/clash/tests/([^/]+)/export$")).Success)
            {
                var format = req.QueryString["format"] ?? "html";
                var pathOut = _dispatcher.Run(() => ClashService.ExportReport(DecodeUtf8Percent(m.Groups[1].Value), format));
                return new { path = pathOut };
            }

            throw new InvalidOperationException($"Неизвестный маршрут: {method} {path}");
        }

        /// <summary>
        /// Правильно декодирует percent-encoded сегмент пути как UTF-8:
        /// собирает сырые байты (ASCII-символы как есть, %XX — как байт) и
        /// декодирует весь буфер разом как UTF-8, чтобы многобайтовые
        /// последовательности (кириллица и т.п.) не разваливались на
        /// отдельные code point'ы, как это делает Uri.AbsolutePath.
        /// </summary>
        private static string DecodeUtf8Percent(string raw)
        {
            var bytes = new List<byte>(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                if (raw[i] == '%' && i + 2 < raw.Length)
                {
                    bytes.Add(Convert.ToByte(raw.Substring(i + 1, 2), 16));
                    i += 2;
                }
                else
                {
                    bytes.Add((byte)raw[i]);
                }
            }
            return Encoding.UTF8.GetString(bytes.ToArray());
        }

        private static T ReadBody<T>(HttpListenerRequest req)
        {
            using var reader = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8);
            var text = reader.ReadToEnd();
            return string.IsNullOrWhiteSpace(text) ? default : JsonConvert.DeserializeObject<T>(text);
        }

        private static void WriteJson(HttpListenerResponse resp, int status, object payload)
        {
            resp.StatusCode = status;
            var json = JsonConvert.SerializeObject(payload);
            var bytes = Encoding.UTF8.GetBytes(json);
            resp.ContentLength64 = bytes.Length;
            resp.OutputStream.Write(bytes, 0, bytes.Length);
        }

        private class StatusUpdateBody
        {
            public string status;
            public string comment;
        }
    }
}
