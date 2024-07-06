using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tempate.Parsing
{
  /// <summary>
  /// Расширения для получения элементов управления содержимым.
  /// </summary>
  /// <remarks>По мотивам http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2011/04/11/137383.aspx .</remarks>
  internal static class ContentControlExtensions
  {
    /// <summary>
    /// Получить коллекцию контролов из XmlPart.
    /// </summary>
    /// <param name="part">XmlPart.</param>
    /// <returns>Коллекция контент контролов из заданной XmlPart.</returns>
    private static IEnumerable<OpenXmlElement> ContentControls(this OpenXmlPart part)
    {
      var resultCollection = new List<OpenXmlElement>();
      resultCollection.AddRange(part.RootElement.Descendants().Where(e => e is SdtBlock || e is SdtRun));
      var tables = part.RootElement.Descendants().Where(e => e is Table);
      resultCollection.AddRange(tables.SelectMany(table => table.Descendants<SdtCell>()));
      return resultCollection;
    }

    /// <summary>
    /// Получить коллекцию контролов из документа docx.
    /// </summary>
    /// <param name="doc">Шаблон документа.</param>
    /// <returns>Коллекция контролов в Xml-представлении.</returns>
    internal static IEnumerable<OpenXmlElement> ContentControls(this WordprocessingDocument doc)
    {
      foreach (var cc in doc.MainDocumentPart.ContentControls())
        yield return cc;
      foreach (var header in doc.MainDocumentPart.HeaderParts)
      {
        foreach (var cc in header.ContentControls())
          yield return cc;
      }
      foreach (var footer in doc.MainDocumentPart.FooterParts)
      {
        foreach (var cc in footer.ContentControls())
          yield return cc;
      }
      if (doc.MainDocumentPart.FootnotesPart != null)
      {
        foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
          yield return cc;
      }
      if (doc.MainDocumentPart.EndnotesPart != null)
      {
        foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
          yield return cc;
      }
    }
  }
}
