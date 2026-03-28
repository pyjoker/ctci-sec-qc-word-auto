using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordAutoTool;

public class MainForm : Form
{
    private readonly WebView2 _webView;
    private readonly int _port;

    public MainForm(int port)
    {
        _port = port;

        Text = "益鼎 QC 自動化工具";
        Width = 900;
        Height = 680;
        MinimumSize = new Size(720, 540);
        StartPosition = FormStartPosition.CenterScreen;

        // Set app icon (optional: skip if no icon file)
        // Icon = new Icon("app.ico");

        _webView = new WebView2
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(_webView);

        Load += async (_, _) => await InitWebViewAsync();
    }

    private async Task InitWebViewAsync()
    {
        var env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(Path.GetTempPath(), "益鼎QC_WebView2Cache"));
        await _webView.EnsureCoreWebView2Async(env);

        // Disable right-click context menu and dev tools shortcut for production feel
        _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
        _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
        // 關鍵代碼：關閉所有的預設對話框（包含信任警告）
        _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false; // 關閉右鍵選單
        _webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = false;      // 關閉密碼保存

        // Auto-grant File System Access (showDirectoryPicker / showOpenFilePicker)
        // so the "允許此網站檢視及複製檔案" dialog never appears.
        _webView.CoreWebView2.PermissionRequested += (_, args) =>
        {
            if (args.PermissionKind == CoreWebView2PermissionKind.FileReadWrite)
                args.State = CoreWebView2PermissionState.Allow;
        };

        _webView.CoreWebView2.Navigate($"http://localhost:{_port}/");
    }
}
