using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Aspose.Cells;
using Newtonsoft.Json.Linq;
using Microsoft.Web.WebView2.Core;
using System.Threading.Tasks;
using System.Net;
using Microsoft.Owin.FileSystems;
using Microsoft.Owin.StaticFiles;
using Microsoft.Owin.Hosting;
using Owin;

namespace Highcharts
{
    public partial class Form1 : Form
    {
        WebView2 WebBrowser_1 = new WebView2();
        // private LocalServer server;
        private OwinServer server;
        private string tempFilePath;

        public Form1()
        {
            InitializeComponent();
            InitializeWebView2();
        }

        private async void InitializeWebView2()
        {
            try
            {
                var webView2Environment = await CoreWebView2Environment.CreateAsync();
                await WebBrowser_1.EnsureCoreWebView2Async(webView2Environment);

                string baseFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
                string serverUrl = "http://localhost:8080/";

                // Создаем временную копию HTML файла
                string htmlFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Chart.html");
                tempFilePath = Path.Combine("files/", Path.GetRandomFileName() + ".html");
                File.Copy(htmlFilePath, tempFilePath, true);

                // Запускаем OWIN сервер
                server = new OwinServer();
                server.Start(baseFolder, serverUrl);

                string resourcesPath = Path.Combine(baseFolder, "libraryes");

                WebBrowser_1.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "local", resourcesPath, CoreWebView2HostResourceAccessKind.Allow);

                WebBrowser_1.Dock = DockStyle.Fill;
                this.Controls.Add(WebBrowser_1);

                // Устанавливаем источник для WebView2
                string htmlUri = serverUrl + tempFilePath;
                Console.WriteLine(htmlUri);
                WebBrowser_1.Source = new Uri(htmlUri);

                WebBrowser_1.CoreWebView2.DOMContentLoaded += CoreWebView2_DOMContentLoaded;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error initializing WebView2: " + ex.Message);
            }
        }

        private async void CoreWebView2_DOMContentLoaded(object sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            await excel_reader();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (server == null)
            {
                server.Stop();
            }
        }

        public class OwinServer
        {
            private IDisposable _webApp;

            public void Start(string baseFolder, string url)
            {
                _webApp = WebApp.Start(url, app =>
                {
                    var fileSystem = new PhysicalFileSystem(baseFolder);
                    var options = new FileServerOptions
                    {
                        EnableDirectoryBrowsing = false,
                        FileSystem = fileSystem
                    };
                    app.UseFileServer(options);
                });
            }

            public void Stop()
            {
                _webApp.Dispose();
            }
        }

        public class LocalServer
        {
            private HttpListener _listener;
            private string _baseFolder;

            public LocalServer(string baseFolder, int port = 8080)
            {
                _baseFolder = baseFolder;
                _listener = new HttpListener();
                _listener.Prefixes.Add("http://localhost:" + port + "/");
            }

            public void Start()
            {
                _listener.Start();
                Task.Run(() => HandleRequests());
            }

            private async Task HandleRequests()
            {
                while (_listener.IsListening)
                {
                    var context = await _listener.GetContextAsync();
                    var response = context.Response;

                    try
                    {
                        string filename = context.Request.Url.AbsolutePath.Substring(1);
                        string filepath = Path.Combine(_baseFolder, filename);
                        if (File.Exists(filepath))
                        {
                            byte[] buffer = File.ReadAllBytes(filepath);
                            response.ContentType = GetContentType(filepath);
                            response.ContentLength64 = buffer.Length;
                            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                        }
                        else
                        {
                            response.StatusCode = (int)HttpStatusCode.NotFound;
                            byte[] buffer = Encoding.UTF8.GetBytes("File not found");
                            response.ContentLength64 = buffer.Length;
                            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Server error: " + ex.Message);
                    }
                    finally
                    {
                        response.OutputStream.Close();
                    }
                }
            }

            private string GetContentType(string path)
            {
                switch (Path.GetExtension(path).ToLower())
                {
                    case ".html":
                        return "text/html";
                    case ".css":
                        return "text/css";
                    case ".js":
                        return "application/javascript";
                    case ".png":
                        return "image/png";
                    case ".jpg":
                    case ".jpeg":
                        return "image/jpeg";
                    default:
                        return "application/octet-stream";
                }
            }

            public void Stop()
            {
                _listener.Stop();
            }
        }

        // private async void excel_reader()
        private Task<string> excel_reader()
        {
            // Load Excel file
            Workbook wb = new Workbook(@"../../parameters.xlsx");

            // Get all worksheets
            WorksheetCollection collection = wb.Worksheets;

            int code = 0;
            string name = "";
            string script = "";

            JArray all_data = new JArray();

            // Loop through all the worksheets
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                // Get worksheet using its index
                Worksheet worksheet = collection[worksheetIndex];

                // Get number of rows and columns
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                // Loop through rows
                for (int i = 1; i <= rows; i++) // Начинаю перебор со второго значения, поскольку в первом лежит имя столбца
                {
                    // Сначала приводим европейский вид времени с запятой, к другому, потому что не конвертится
                    // date = worksheet.Cells[i, 0].Value.ToString().Replace(",",".");
                    // Затем парсим эту строку в datetime
                    // DateTime.Parse(date)
                    // Для оптимизации, не использовал переменные, а сделал все разом
                    int code_of_value = Convert.ToInt32(worksheet.Cells[i, 1].Value);
                    long milliseconds = (long)(DateTime.Parse(worksheet.Cells[i, 0].Value.ToString().Replace(",", ".")).ToUniversalTime().Subtract(
                        new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
                        ).TotalMilliseconds);

                    double value = Convert.ToDouble(worksheet.Cells[i, 2].Value);

                    // Итерация берется на 2 шага меньше, от значения реального файла
                    // То есть берем: i = 10000, а значение будет в file[10002]
                    // Все потому что в первом значении лежат названия столбцов, еще минус один,
                    // поскольку отсчет в массивах начинается с нуля

                    if (code != code_of_value || i == rows - 1) // Здесь херовая проверка, для добавления последнего элемента
                    {
                        switch (code)
                        { // Временный вариант для наименования кодов
                            case 0:
                                name = "T пр(°C)| #0";
                                break;
                            case 1:
                                name = "T обр(°C)| #1";
                                break;
                            case 3:
                                name = "G пр(т)| #3";
                                break;
                            case 4:
                                name = "G обр(т)| #4";
                                break;
                            case 11:
                                name = "W пр(ГКал)| #11";
                                break;
                        }

                        code = code_of_value;

                        script = "elevationData.push(" +
                            new JObject(
                            new JProperty("name", name),
                            new JProperty("data", all_data)
                        ) + ")";

                        WebBrowser_1.ExecuteScriptAsync(script);

                        all_data.Clear();
                    }

                    code = code_of_value;

                    all_data.Add(new JArray(
                        milliseconds,
                        value
                    ));
                }
            }

            script = @"for (let item of elevationData) {
                        chart.addSeries(item);
                    };";

            return WebBrowser_1.ExecuteScriptAsync(script);
        }

        private async void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
