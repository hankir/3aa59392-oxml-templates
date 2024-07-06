using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Template;

var host = Host.CreateDefaultBuilder(args)
  .ConfigureServices(services =>
  {
    services.AddSingleton(new CommandLineArgs(args));
    services.AddHostedService<CommandLineHandlerService>();
  })
  .ConfigureLogging(loggingBuilder => loggingBuilder.ClearProviders());

await host.RunConsoleAsync();
