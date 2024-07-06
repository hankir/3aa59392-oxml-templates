using System.CommandLine;
using Microsoft.Extensions.Hosting;

namespace Template
{
  internal record CommandLineArgs(string[] Value) { }

  internal class CommandLineHandlerService : IHostedService
  {
    private readonly CommandLineArgs args;
    private readonly IHostApplicationLifetime lifeTime;

    public CommandLineHandlerService(CommandLineArgs args, IHostApplicationLifetime lifeTime)
    {
      this.args = args;
      this.lifeTime = lifeTime;
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
      var rootCommand = new RootCommand("Create, update or inspect documents from templates.");

      var documentFile = new Argument<FileInfo>("document", "Path to document file.");
      var templateFile = new Argument<FileInfo>("template", "Path to template file.");
      var outputFile = new Argument<FileInfo>("output", "Path to new file.");

      var createCommand = new Command("create", "Create new document from template.");
      createCommand.Add(templateFile);
      createCommand.Add(outputFile);
      createCommand.SetHandler((file, output) => TemplateEngine.CreateFromTemplate(file, output), templateFile, outputFile);

      var updateCommand = new Command("update", "Update document templated values.");
      updateCommand.Add(documentFile);
      updateCommand.SetHandler(file => TemplateEngine.Process(file), documentFile);

      var showCommand = new Command("show", "Show template mapping.");
      showCommand.Add(documentFile);
      showCommand.SetHandler(file => TemplateEngine.ShowTemplateMapping(file), documentFile);

      rootCommand.AddCommand(createCommand);
      rootCommand.AddCommand(updateCommand);
      rootCommand.AddCommand(showCommand);
      await rootCommand.InvokeAsync(this.args.Value);
      this.lifeTime.StopApplication();
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
  }
}
