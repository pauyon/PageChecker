using PageChecker.Library;
using PageChecker.ConsoleApp;
using Serilog;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .WriteTo.File("logs/PageChecker.log", rollingInterval: RollingInterval.Day)
    .CreateLogger();

var host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((context, services) =>
    {
        // Add Serilog to the logging pipeline
        services.AddLogging(loggingBuilder =>
            loggingBuilder.AddSerilog(dispose: true));

        // Register the main application class
        services.AddTransient<App>();
        services.AddTransient<ConsoleUtility>();

        // Register other classes that require logging
        services.AddTransient<FileReaderBase>();
    })
    .UseSerilog() // Ensure Serilog is used as the logging provider
    .Build();

// Use the host to get the main application class and run it
var app = host.Services.GetRequiredService<App>();
app.Run();


