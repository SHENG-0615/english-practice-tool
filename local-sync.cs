using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

// A loopback-only companion for the two explicitly named workbooks beside the app.
// It does not browse the disk or accept paths from HTTP clients.
public static class IeltsLocalSync
{
    const string App = "english-practice-local-v1";
    const int MaxBody = 64 * 1024 * 1024;
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
    static readonly Dictionary<string, string> Names = new Dictionary<string, string> {
        { "dictation", "听写同步.xlsx" }, { "vocabulary", "词汇练习.xlsx" }
    };
    static string root, token, htmlPath, sourceHtmlPath;
    static int port;
    static bool stopping;
    static DateTime lastUse;
    class Request
    {
        public string Method, Path;
        public Dictionary<string, string> Headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public byte[] Body;
    }
    static string Hash(byte[] bytes)
    {
        using (var hash = SHA256.Create()) return "\"" + BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant() + "\"";
    }
    static string Hash(Stream stream)
    {
        using (var hash = SHA256.Create()) return "\"" + BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() + "\"";
    }
    static string FilePath(string key) { return Path.Combine(root, Names[key]); }
    static object Info(bool includeToken)
    {
        var files = new Dictionary<string, object>();
        foreach (var item in Names) if (File.Exists(FilePath(item.Key))) files[item.Key] = new { name = item.Value, path = FilePath(item.Key) };
        var data = new Dictionary<string, object> {
            { "app", App }, { "root", root }, { "url", "http://127.0.0.1:" + port },
            { "pagePath", htmlPath }, { "files", files }
        };
        if (includeToken) data["token"] = token;
        return data;
    }
    static bool Reuse(string statePath, bool openBrowser)
    {
        try {
            var data = Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(statePath, Utf8));
            if (!String.Equals((string)data["root"], root, StringComparison.OrdinalIgnoreCase)) return false;
            var address = (string)data["url"];
            if (!Regex.IsMatch(address, @"^http://127\.0\.0\.1:\d+$")) return false;
            var request = (HttpWebRequest)WebRequest.Create(address + "/api/ping");
            request.Proxy = null; request.Timeout = 600;
            request.Headers["X-IELTS-Token"] = (string)data["token"];
            using (var response = request.GetResponse()) using (var reader = new StreamReader(response.GetResponseStream(), Utf8)) {
                var info = Json.Deserialize<Dictionary<string, object>>(reader.ReadToEnd());
                if ((string)info["app"] != App || !String.Equals((string)info["root"], root, StringComparison.OrdinalIgnoreCase)) return false;
            }
            ConfigureHtml(Json.Serialize(data).Replace("<", "\\u003c"));
            if (openBrowser) OpenPage();
            return true;
        } catch { return false; }
    }
    static void OpenPage()
    {
        var start = new ProcessStartInfo(htmlPath) { UseShellExecute = true };
        Process.Start(start);
    }
    static void ConfigureHtml(string config)
    {
        string html = File.ReadAllText(sourceHtmlPath, Utf8);
        string pattern = @"/\* LOCAL_CONFIG_START \*/.*?/\* LOCAL_CONFIG_END \*/";
        if (!Regex.IsMatch(html, pattern, RegexOptions.Singleline)) throw new IOException("HTML 版本不匹配，请使用整套文件。");
        html = Regex.Replace(html, pattern, delegate(Match match) { return "/* LOCAL_CONFIG_START */" + config + "/* LOCAL_CONFIG_END */"; }, RegexOptions.Singleline);
        File.WriteAllText(htmlPath, html, Utf8);
    }
    public static void Run(string folder, bool openBrowser)
    {
        root = Path.GetFullPath(folder);
        if (root.Length > Path.GetPathRoot(root).Length) root = root.TrimEnd(Path.DirectorySeparatorChar);
        sourceHtmlPath = Path.Combine(root, "index.html");
        string runtimeFolder = Path.Combine(root, ".local-runtime");
        Directory.CreateDirectory(runtimeFolder);
        htmlPath = Path.Combine(runtimeFolder, "index.html");
        if (!File.Exists(sourceHtmlPath)) throw new FileNotFoundException("找不到同文件夹里的 index.html。");

        string statePath = Path.Combine(runtimeFolder, "connection.json");
        if (File.Exists(statePath) && Reuse(statePath, openBrowser)) return;
        token = Guid.NewGuid().ToString("N");
        byte[] digest;
        using (var hash = SHA256.Create()) digest = hash.ComputeHash(Utf8.GetBytes(root.ToLowerInvariant()));
        int firstPort = 39000 + BitConverter.ToUInt16(digest, 0) % 12000;
        TcpListener listener = null;
        for (int attempt = 0; attempt < 25; attempt++) {
            try { port = firstPort + attempt; listener = new TcpListener(IPAddress.Loopback, port); listener.Start(); break; }
            catch (SocketException) { if (listener != null) listener.Stop(); listener = null; }
        }
        if (listener == null) throw new IOException("本机端口暂时不可用，请稍后再试。");
        try {
            string config = Json.Serialize(Info(true)).Replace("<", "\\u003c");
            ConfigureHtml(config);
            File.WriteAllText(statePath, config, Utf8);
            lastUse = DateTime.UtcNow;
            if (openBrowser) OpenPage();
            while (!stopping && DateTime.UtcNow - lastUse < TimeSpan.FromHours(12)) {
                if (!listener.Pending()) { System.Threading.Thread.Sleep(100); continue; }
                using (var client = listener.AcceptTcpClient()) {
                    client.ReceiveTimeout = 10000; client.SendTimeout = 30000;
                    try { ProcessRequest(client.GetStream()); }
                    catch { /* A cancelled tab or invalid request must not end the companion. */ }
                }
            }
        } finally { listener.Stop(); }
    }
    static Request Read(NetworkStream stream)
    {
        var header = new MemoryStream();
        int state = 0;
        while (header.Length < 16384) {
            int b = stream.ReadByte();
            if (b < 0) throw new IOException("Incomplete request");
            header.WriteByte((byte)b);
            state = (state == 0 && b == 13) ? 1 : (state == 1 && b == 10) ? 2 : (state == 2 && b == 13) ? 3 : (state == 3 && b == 10) ? 4 : b == 13 ? 1 : 0;
            if (state == 4) break;
        }
        if (state != 4) throw new IOException("Header too large");
        var lines = Encoding.ASCII.GetString(header.ToArray()).Split(new[] { "\r\n" }, StringSplitOptions.None);
        var start = lines[0].Split(' ');
        if (start.Length != 3 || !start[1].StartsWith("/")) throw new IOException("Invalid request");
        var request = new Request { Method = start[0], Path = start[1].Split('?')[0] };
        for (int i = 1; i < lines.Length; i++) {
            int colon = lines[i].IndexOf(':');
            if (colon <= 0) continue;
            string name = lines[i].Substring(0, colon).Trim();
            if (request.Headers.ContainsKey(name)) throw new IOException("Duplicate header");
            request.Headers[name] = lines[i].Substring(colon + 1).Trim();
        }
        if (request.Headers.ContainsKey("Transfer-Encoding")) throw new IOException("Chunked requests are not supported");
        int length = 0;
        if (request.Headers.ContainsKey("Content-Length") && !Int32.TryParse(request.Headers["Content-Length"], out length)) throw new IOException("Invalid size");
        if (length < 0 || length > MaxBody) throw new IOException("Workbook too large");
        request.Body = new byte[length];
        for (int read = 0; read < length;) {
            int count = stream.Read(request.Body, read, Math.Min(65536, length - read));
            if (count == 0) throw new IOException("Incomplete workbook");
            read += count;
        }
        return request;
    }
    static void Reply(NetworkStream stream, int status, byte[] bytes, string contentType, string origin, Dictionary<string, string> extra)
    {
        var header = new StringBuilder("HTTP/1.1 " + status + " " + (status == 200 ? "OK" : status == 204 ? "No Content" : "Error") + "\r\n");
        header.Append("Connection: close\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\n");
        header.Append("Content-Type: " + contentType + "\r\nContent-Length: " + bytes.Length + "\r\n");
        if (origin == "null") header.Append("Access-Control-Allow-Origin: null\r\nVary: Origin\r\nAccess-Control-Expose-Headers: ETag, X-Modified-Time\r\n");
        if (extra != null) foreach (var item in extra) header.Append(item.Key + ": " + item.Value + "\r\n");
        header.Append("\r\n");
        byte[] head = Encoding.ASCII.GetBytes(header.ToString());
        stream.Write(head, 0, head.Length); stream.Write(bytes, 0, bytes.Length); stream.Flush();
    }
    static void SendJson(NetworkStream stream, int status, object value, string origin)
    {
        Reply(stream, status, Utf8.GetBytes(Json.Serialize(value)), "application/json; charset=utf-8", origin, null);
    }
    static void ProcessRequest(NetworkStream stream)
    {
        Request request = Read(stream);
        string origin = request.Headers.ContainsKey("Origin") ? request.Headers["Origin"] : "";
        if (origin != "" && origin != "null") { SendJson(stream, 403, new { error = "只允许本机 HTML 访问。" }, ""); return; }
        string host = request.Headers.ContainsKey("Host") ? request.Headers["Host"] : "";
        if (host != "127.0.0.1:" + port) { SendJson(stream, 403, new { error = "无效的本机地址。" }, ""); return; }
        if (request.Method == "OPTIONS") {
            Reply(stream, 204, new byte[0], "text/plain", origin, new Dictionary<string, string> {
                { "Access-Control-Allow-Methods", "GET, PUT, POST, OPTIONS" }, { "Access-Control-Allow-Headers", "X-IELTS-Token, If-Match, Content-Type" }, { "Access-Control-Max-Age", "600" }
            }); return;
        }
        string supplied = request.Headers.ContainsKey("X-IELTS-Token") ? request.Headers["X-IELTS-Token"] : "";
        if (supplied != token) { SendJson(stream, 403, new { error = "本机连接已更新，请从开始练习重新打开。" }, origin); return; }
        lastUse = DateTime.UtcNow;
        if (request.Method == "GET" && request.Path == "/api/ping") { SendJson(stream, 200, Info(false), origin); return; }
        if (request.Method == "POST" && request.Path == "/api/stop") { stopping = true; SendJson(stream, 200, new { stopped = true }, origin); return; }
        const string prefix = "/api/workbooks/";
        string key = request.Path.StartsWith(prefix) ? request.Path.Substring(prefix.Length) : "";
        if (!Names.ContainsKey(key)) { SendJson(stream, 404, new { error = "文件不在本机练习范围内。" }, origin); return; }
        string path = FilePath(key);
        try {
            if (request.Method == "GET") {
                byte[] bytes; long modified;
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    if (file.Length > MaxBody) throw new IOException("工作簿过大。");
                    using (var copy = new MemoryStream()) { file.CopyTo(copy); bytes = copy.ToArray(); }
                    modified = (File.GetLastWriteTimeUtc(path).Ticks - new DateTime(1970, 1, 1).Ticks) / TimeSpan.TicksPerMillisecond;
                }
                Reply(stream, 200, bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", origin, new Dictionary<string, string> { { "ETag", Hash(bytes) }, { "X-Modified-Time", modified.ToString() } });
                return;
            }
            if (request.Method != "PUT") { SendJson(stream, 405, new { error = "不支持此操作。" }, origin); return; }
            string expected = request.Headers.ContainsKey("If-Match") ? request.Headers["If-Match"] : "";
            if (expected.Length != 66 || request.Body.Length < 4 || request.Body[0] != 80 || request.Body[1] != 75) {
                SendJson(stream, 400, new { error = "工作簿内容或版本无效，未写入。" }, origin); return;
            }
            string temp = Path.Combine(root, ".sync-" + Guid.NewGuid().ToString("N") + ".tmp");
            try {
                using (var output = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { output.Write(request.Body, 0, request.Body.Length); output.Flush(true); }
                // Deny other readers/writers while allowing this process's atomic replacement.
                using (var current = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Delete)) {
                    if (Hash(current) != expected) { SendJson(stream, 409, new { error = "表格已被其他程序修改，已保留本次作答，请再同步。" }, origin); return; }
                    string backupFolder = Path.Combine(root, "同步备份");
                    Directory.CreateDirectory(backupFolder);
                    string backup = Path.Combine(backupFolder, "上次同步前_" + Names[key]);
                    File.Replace(temp, path, backup, true);
                }
            } finally { if (File.Exists(temp)) File.Delete(temp); }
            SendJson(stream, 200, new { saved = true, path = path }, origin);
        } catch (FileNotFoundException) { SendJson(stream, 404, new { error = "找不到表格，请把两份 XLSX 放回同一文件夹。" }, origin); }
        catch (UnauthorizedAccessException) { SendJson(stream, 423, new { error = "无法写入表格，请关闭 WPS 或 Excel 后再同步。" }, origin); }
        catch (IOException) { SendJson(stream, 423, new { error = "表格正在使用或无法写入，请关闭 WPS 或 Excel 后再同步。" }, origin); }
    }
}
