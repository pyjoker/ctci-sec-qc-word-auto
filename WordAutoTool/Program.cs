using System.Net;
using System.Net.Sockets;
using System.Reflection;
using Microsoft.Extensions.FileProviders;
using WordAutoTool;
using WordAutoTool.Services;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        int port = GetFreePort();
        var serverReady = new ManualResetEventSlim(false);

        // Build the web host
        var builder = WebApplication.CreateBuilder(args);
        builder.WebHost.UseUrls($"http://localhost:{port}");
        builder.WebHost.ConfigureKestrel(k =>
        {
            k.Limits.MaxRequestBodySize = 500 * 1024 * 1024; // 500 MB
        });
        builder.Services.AddControllers();
        builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o =>
        {
            o.MultipartBodyLengthLimit = 500 * 1024 * 1024; // 500 MB
        });
        builder.Services.AddSingleton<WordProcessingService>();
        builder.Services.AddSingleton<InspectService>();
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

        var app = builder.Build();

        // Serve wwwroot files from embedded resources (works in single-file exe)
        var embeddedProvider = new EmbeddedFileProvider(
            Assembly.GetExecutingAssembly(),
            "WordAutoTool.wwwroot");
        app.UseStaticFiles(new StaticFileOptions { FileProvider = embeddedProvider });

        app.MapControllers();

        // Fallback: serve index.html from embedded resources
        app.MapFallback(async ctx =>
        {
            var file = embeddedProvider.GetFileInfo("index.html");
            ctx.Response.ContentType = "text/html; charset=utf-8";
            await using var stream = file.CreateReadStream();
            await stream.CopyToAsync(ctx.Response.Body);
        });

        // Signal when Kestrel is ready
        app.Lifetime.ApplicationStarted.Register(() => serverReady.Set());

        // Start Kestrel on a background (MTA) thread — required: do NOT await here
        var cts = new CancellationTokenSource();
        Task.Run(() => app.RunAsync(cts.Token));

        // Wait until the server is actually listening before showing the window
        serverReady.Wait();

        // Run WinForms on the STA thread
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm(port));

        // Shut down Kestrel when the window closes
        cts.Cancel();
    }

    static int GetFreePort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }
}
