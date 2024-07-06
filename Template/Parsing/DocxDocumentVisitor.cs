using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tempate.Parsing
{
  /// <summary>
  /// Визитор параметров для docx документа.
  /// </summary>
  public class DocxDocumentVisitor : IDisposable
  {

    #region Поля и свойства

    /// <summary>
    /// Текущий документ.
    /// </summary>
    private WordprocessingDocument document;

    #endregion

    #region Методы

    /// <summary>
    /// Проверка валидности контроля.
    /// </summary>
    /// <param name="contentControl">Проверяемый контент контрол.</param>
    /// <returns>Признак валидность контроля.</returns>
    private bool IsValidControl(OpenXmlElement contentControl)
    {
      var props = contentControl.Elements<SdtProperties>().FirstOrDefault();
      if (props == null)
        return false;
      var alias = props.Elements<SdtAlias>().FirstOrDefault();
      return alias != null;
    }

    #endregion

    #region IDocumentVisitor

    public IEnumerable<IParameter> GetParameters()
    {
      foreach (var contentControl in this.document.ContentControls())
      {
        if (!this.IsValidControl(contentControl))
          continue;

        yield return new DocxParameter(contentControl);
      }
    }

    public void Init(Stream body)
    {
      this.document = WordprocessingDocument.Open(body, body.CanWrite);
    }

    #endregion

    #region IDisposable

    /// <summary>
    /// Признак освобожденности взятых ресурсов.
    /// </summary>
    private bool isDisposed;

    public void Dispose()
    {
      this.Dispose(true);
      GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
      if (!this.isDisposed)
      {
        if (disposing)
        {
          if (this.document != null)
            this.document.Dispose();
        }

        this.isDisposed = true;
      }
    }

    ~DocxDocumentVisitor()
    {
      this.Dispose(false);
    }

    #endregion
  }
}
