using CefSharp;
using CefSharp.OffScreen;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;

namespace SOFSInventoryDownloader
{
    class Program
    {
        public static bool Downloaded = false;
        public static string InventoryFilePath = @"C:\SOFS Inventory\SOFS_Inventory.xlsx";

        static void Main(string[] args)
        {
            string username = "user";
            string password = GetPasswordFromDatabase(username);

            AsyncContext.Run(async () =>
            {
                var settings = new CefSettings
                {
                    CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\\Cache")
                };

                bool cefInitialized = await Cef.InitializeAsync(settings, performDependencyCheck: true);
                if (!cefInitialized) throw new Exception("Failed to initialize CefSharp.");

                string loginUrl = "https://sofslogin.supplieroasis.com/adfs/ls/?wa=wsignin1.0&wtrealm=urn%3Agateway-ui-edge%3Aprod&wctx=https%3A%2F%2Fedge.supplieroasis.com%2Fdashboard";

                using (var browser = new ChromiumWebBrowser(loginUrl))
                {
                    browser.DownloadHandler = new InventoryDownloadHandler();

                    var response = await browser.WaitForInitialLoadAsync();
                    if (!response.Success)
                        throw new Exception($"Initial page load failed: {response.ErrorCode}");

                    Console.WriteLine("Logging in...");
                    string loginScript = $@"
                        document.querySelector('#userNameInput').value = '{username}';
                        document.querySelector('#passwordInput').value = '{password}';
                        document.querySelector('#submitButton').click();";
                    await browser.EvaluateScriptAsync(loginScript);

                    await Task.Delay(6000);
                    await browser.LoadUrlAsync("https://edge.supplieroasis.com/gateway/ng/app/index.html#/products");
                    await Task.Delay(10000);

                    string exportScript = @"document.querySelector('a.qa-export-products').click();";
                    Console.WriteLine("Initiating export...");
                    await browser.EvaluateScriptAsync(exportScript);
                    await Task.Delay(15000);

                    for (int attempt = 0; attempt < 3 && !Downloaded; attempt++)
                    {
                        Console.WriteLine($"Retrying download... Attempt {attempt + 1}");
                        await browser.EvaluateScriptAsync(exportScript);
                        await Task.Delay(15000);
                    }
                }

                if (Downloaded)
                {
                    Console.WriteLine("Parsing downloaded inventory...");
                    var inventory = InventoryParser.Parse(InventoryFilePath);
                    InventoryParser.SaveToDatabase(inventory);
                    InventoryParser.ArchiveFile(InventoryFilePath);
                }
                else
                {
                    Console.WriteLine("Download failed.");
                }

                Console.WriteLine("Process completed.");
                Cef.Shutdown();
            });
        }

        private static string GetPasswordFromDatabase(string username)
        {
            using var conn = new SqlConnection("Data Source=localhost;Initial Catalog=master;User ID=user;Password=password");
            conn.Open();

            using var cmd = new SqlCommand($"SELECT [password] FROM [SOFSPassword] WHERE [Username]=@username", conn);
            cmd.Parameters.AddWithValue("@username", username);

            return cmd.ExecuteScalar() as string;
        }
    }
}
