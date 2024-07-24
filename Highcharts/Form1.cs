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
using System.Net.Http;

namespace Highcharts
{
    public partial class Form1 : Form
    {
        private OwinServer server;
        private string tempFilePath;
        private static IDisposable _webApp;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView;

        public Form1()
        {
            InitializeComponent();
            InitializeWebView2();
        }


        private async void webView_CoreWebView2InitializationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            if (webView != null && webView.CoreWebView2 != null)
            {
                webView.Dock = DockStyle.Fill;

                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "local", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "libraryes"), CoreWebView2HostResourceAccessKind.Allow);

                webView.CoreWebView2.Navigate("http://localhost:8000/" + AppDomain.CurrentDomain.BaseDirectory + "Chart.html");
                
                webView.CoreWebView2.DOMContentLoaded += CoreWebView2_DOMContentLoaded;
            }
        }

        private async void InitializeWebView2()
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), string.Format("{0}", Environment.UserName));

                await InitAsync(path);

                string baseFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
                string serverUrl = "http://localhost:8000/";

                // Создаем временную копию HTML файла
                string htmlFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Chart.html");
                tempFilePath = Path.Combine("files", Path.GetRandomFileName() + ".html");
                File.Copy(htmlFilePath, tempFilePath, true);

                // Запускаем OWIN сервер
                server = new OwinServer();
                server.Start(baseFolder, serverUrl);

                string resourcesPath = Path.Combine(baseFolder, "libraryes");

                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "local", resourcesPath, CoreWebView2HostResourceAccessKind.Allow);

                // Устанавливаем источник для WebView2
                string htmlUri = serverUrl + baseFolder + tempFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error initializing WebView2: " + ex.Message);
            }
        }

        private async Task InitAsync(string path) {
            var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(userDataFolder: path);
            await webView.EnsureCoreWebView2Async(env);
        }


        private async void CoreWebView2_DOMContentLoaded(object sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            await excel_reader();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            server.Stop();
        }

        public class OwinServer
        {
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
                if(_webApp != null) {
                    _webApp.Dispose();
                }
            }
        }

        private async Task<string> excel_reader()
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

                        await webView.ExecuteScriptAsync(script);

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

            return await webView.ExecuteScriptAsync(script);
        }
    }
}
