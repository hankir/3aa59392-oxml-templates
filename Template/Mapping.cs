using System.Text;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using Tempate.Parsing;

namespace Template
{
  internal record MappingElement(string ParameterName, string? DataPath) { }

  /// <summary>
  /// Класс, представляющий маппинг параметров на свойства сущности.
  /// </summary>
  internal class Mapping
  {
    #region Поля и свойства

    private const string ParameterNamePrefix = "TPL_";

    private readonly List<MappingElement> elements = new List<MappingElement>();

    /// <summary>
    /// Коллекция элементов маппинга.
    /// </summary>
    public IEnumerable<MappingElement> Elements
    {
      get { return this.elements; }
    }

    #endregion

    #region Методы

    public static Mapping GetMappingFromTemplate(Stream templateBody)
    {
      if (!templateBody.CanRead)
        throw new ArgumentException("Stream must be readable and writable");

      try
      {
        var result = new Mapping();
        using (var visitor = new DocxDocumentVisitor())
        {
          visitor.Init(templateBody);
          var rules = new Dictionary<string, string?>();
          foreach (var parameter in visitor.GetParameters())
          {
            if (parameter != null && !string.IsNullOrWhiteSpace(parameter.Name))
            {
              // В проекте не реализована реальная привязка к данным, поэтому оставляем пустые значение DataPath.
              rules[parameter.Name] = " ";
            }
          }

          return GetMapping(rules);
        }
      }
      finally
      {
        templateBody.Position = 0;
      }
    }

    /// <summary>
    /// Получить маппинг из тела документа.
    /// </summary>
    /// <param name="body">Тело документа.</param>
    /// <returns>Экземпляр маппинга.</returns>
    public static Mapping GetMapping(Stream body)
    {
      try
      {
        using var wordDocument = WordprocessingDocument.Open(body, false);
        if (wordDocument == null)
          return new Mapping();

        var customProps = wordDocument.CustomFilePropertiesPart;
        var properties = customProps != null ? customProps.Properties.OfType<CustomDocumentProperty>()
          .Where(p => p.Name != null && p.Name.Value.StartsWith(ParameterNamePrefix, StringComparison.Ordinal)) : null;

        if (properties == null || !properties.Any())
          return new Mapping();

        var rules = new Dictionary<string, string>();
        foreach (var property in properties)
        {
          try
          {
            var dataPath = Encoding.UTF8.GetString(Convert.FromBase64String(property.InnerText));
            var parameterName = property.Name.Value.Split(new[] { ParameterNamePrefix }, StringSplitOptions.RemoveEmptyEntries).SingleOrDefault();
            if (parameterName != null)
              rules.Add(parameterName, dataPath);
          }
          catch (Exception)
          {
            Console.WriteLine(string.Format("Unable get parameter value for document parameter '{0}'", property.Name.Value));
          }
        }

        return GetMapping(rules);
      }
      finally
      {
        body.Position = 0;
      }
    }

    /// <summary>
    /// Сохранить маппинг в тело документа.
    /// </summary>
    /// <param name="body">Тело документа.</param>
    /// <param name="mapping">Маппинг.</param>
    /// <remarks>
    /// По мотивам https://msdn.microsoft.com/ru-ru/library/office/hh674468.aspx.
    /// </remarks>
    public static void StoreMapping(Stream body, Mapping mapping)
    {
      try
      {
        using var document = WordprocessingDocument.Open(body, body.CanWrite);
        var customProps = document.CustomFilePropertiesPart ?? document.AddCustomFilePropertiesPart();
        customProps.Properties = new Properties();
        foreach (var mappingElement in mapping.Elements)
        {
          var dataPath = mappingElement.DataPath;
          if (!string.IsNullOrEmpty(dataPath))
          {
            var newProp = new CustomDocumentProperty();
            newProp.VTLPWSTR = new VTLPWSTR(Convert.ToBase64String(Encoding.UTF8.GetBytes(dataPath)));
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = ParameterNamePrefix + mappingElement.ParameterName;
            customProps.Properties.AppendChild(newProp);
          }
        }

        int pid = 2;
        foreach (CustomDocumentProperty item in customProps.Properties)
          item.PropertyId = pid++;
        customProps.Properties.Save();
      }
      finally
      {
        body.Position = 0;
      }
    }

    /// <summary>
    /// Получить маппинг.
    /// </summary>
    /// <param name="rules">Параметры для автозаполнения.</param>
    /// <returns>Экземпляр маппинга.</returns>
    private static Mapping GetMapping(Dictionary<string, string?> rules)
    {
      var instance = new Mapping();
      if (rules == null || !rules.Any())
        return instance;

      instance.elements.AddRange(rules.Select(rule => new MappingElement(rule.Key, rule.Value)));
      return instance;
    }

    #endregion

    #region Конструкторы

    /// <summary>
    /// Конструктор.
    /// </summary>
    private Mapping() { }

    #endregion
  }
}
