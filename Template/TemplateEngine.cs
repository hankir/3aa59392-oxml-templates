using Tempate.Parsing;

namespace Template
{
  /// <summary>
  /// Движок для работы с шаблонами документов.
  /// </summary>
  internal static class TemplateEngine
  {
    /// <summary>
    /// Создать документ из шаблона.
    /// </summary>
    /// <param name="template">Шаблон документа.</param>
    /// <param name="output">Новый документ, который будет создан из шаблона.</param>
    public static void CreateFromTemplate(FileInfo template, FileInfo output)
    {
      if (!template.Exists)
      {
        Console.WriteLine("Template file not found.");
        return;
      }

      Mapping mapping;
      using (var templateStream = template.OpenRead())
      {
        // Получаем параметры из шаблона.
        // TODO: В проекте не реализована реальная привязка к данным,
        //   а сделано обновление данных шаблона из пользовательского ввода консоли.
        // Поэтому маппинг полученный из шаблона содержит пустые значение DataPath.
        mapping = Mapping.GetMappingFromTemplate(templateStream);

        // Копируем содержимое шаблона в новый документ.
        using (var outputStream = output.Create())
        {
          templateStream.CopyTo(outputStream);
          templateStream.Position = 0;
        }

        // Сохраняем связь "Параметр шаблона" -> "Свойства сущности" в новый документ, созданный из шаблона.
        using (var outputStream = output.Open(FileMode.Open, FileAccess.ReadWrite))
        {
          Mapping.StoreMapping(outputStream, mapping);
        }
      }

      Console.WriteLine($"Document '{output.FullName}' created from template '{template.FullName}'.");
    }

    /// <summary>
    /// Обновить значение параметров шаблона в документе.
    /// </summary>
    /// <param name="outputFile">Документ созданный из шаблона.</param>
    public static void Process(FileInfo outputFile)
    {
      if (!outputFile.Exists)
      {
        Console.WriteLine("Document file not found.");
        return;
      }

      using var output = outputFile.Open(FileMode.Open, FileAccess.ReadWrite);

      // Получим связь "Параметр шаблона" -> "Свойства сущности".
      // Эта связь хранится в самом теле документа.
      var mapping = Mapping.GetMapping(output);
      if (!mapping.Elements.Any())
      {
        // Без Mapping неизвестно, как получить значение из источника данных.
        Console.WriteLine("Template properties not found.");
        return;
      }

      using var visitor = new DocxDocumentVisitor();
      visitor.Init(output);
      Console.WriteLine("Fill template property values:");
      foreach (var parameter in visitor.GetParameters())
      {
        Console.Write($"  Set '{parameter.Name}': ");

        string? parameterValue = GetParameterValue(parameter.Name, mapping);

        if (!string.IsNullOrWhiteSpace(parameterValue))
          parameter.Value = parameterValue;
      }
      Console.WriteLine("Document stream has been processed");
    }

    /// <summary>
    /// Показать параметры шаблона.
    /// </summary>
    /// <param name="documentFile">Документ, созданный из шаблона.</param>
    public static void ShowTemplateMapping(FileInfo documentFile)
    {
      if (!documentFile.Exists)
      {
        Console.WriteLine("Document file not found.");
        return;
      }

      using var fileContent = documentFile.OpenRead();
      var map = Mapping.GetMapping(fileContent);
      if (map.Elements.Any())
      {
        Console.WriteLine("Found template properties:");
        foreach (var p in map.Elements)
          Console.WriteLine("  \"{0}\": \"{1}\"", p.ParameterName, p.DataPath);
      }
      else
        Console.WriteLine("Template properties not found.");
    }

    /// <summary>
    /// Получить значение параметра шаблона по связи "Параметр шаблона" -> "Свойства сущности".
    /// </summary>
    /// <param name="parameterName">Имя параметра.</param>
    /// <param name="mapping">Связка свойств сущности с параметрами шаблона.</param>
    /// <returns>Значение параметра.</returns>
    /// <remarks>
    /// Используя mapping, можно получить значение свойства сущности.
    /// Mapping хранит информацию о том, как брать значение свойства сущности по указанному mapping.Elements.DataPath.
    /// DataPath - это может быть функция, или путь через цепочку значений свойств и т.п.
    /// Например, Author -> Name -> Accusative
    /// </remarks>
    private static string? GetParameterValue(string parameterName, Mapping mapping)
    {
      // Для примера получаем значение свойства из консоли.
      return Console.ReadLine();
    }
  }
}
