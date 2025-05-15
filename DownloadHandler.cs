using CefSharp;

namespace SOFSInventoryDownloader
{
    public class InventoryDownloadHandler : IDownloadHandler
    {
        public void OnBeforeDownload(IWebBrowser chromiumWebBrowser, IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
        {
            if (!callback.IsDisposed)
            {
                using (callback)
                {
                    Console.WriteLine("Downloading inventory file...");
                    callback.Continue(Program.InventoryFilePath, showDialog: false);
                    Program.Downloaded = true;
                }
            }
        }

        public void OnDownloadUpdated(IWebBrowser chromiumWebBrowser, IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
        {
            // Optional: Add progress log
        }
    }
}
