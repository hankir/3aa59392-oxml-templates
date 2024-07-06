using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Tempate.Parsing
{
  /// <summary>
  /// Docx-реализация API для работы с параметром.
  /// </summary>
  public class DocxParameter : IParameter
  {
    #region Поля и свойства

    /// <summary>
    /// XML-элемент, являющийся контент-контролем.
    /// </summary>
    protected OpenXmlElement ContentControl { get; private set; }

    /// <summary>
    /// Список элементов, содержащих текст.
    /// </summary>
    private readonly List<Text> openXmlTextElements;

    private Text displayValue;

    /// <summary>
    /// XML-элемент, содержащий отображаемое значение.
    /// </summary>
    protected Text DisplayValue
    {
      get
      {
        if (this.displayValue == null)
          this.displayValue = this.ContentControl.Descendants<Text>().FirstOrDefault();
        return this.displayValue;
      }
    }

    #endregion

    #region Методы

    /// <summary>
    /// Убрать текст во всех дочерних узлах.
    /// </summary>
    private void ClearAllDescendants()
    {
      foreach (var textElement in this.ContentControl.Descendants<Text>())
        textElement.Text = string.Empty;
    }

    /// <summary>
    /// Означить значение для поля документа.
    /// </summary>
    /// <param name="value">Значение.</param>
    protected virtual void AssignValue(string? value)
    {
      if (this.DisplayValue == null)
        return;

      this.DisplayValue.Text = value;
    }

    /// <summary>
    /// Записать сообщение в лог.
    /// </summary>
    /// <param name="message">Сообщение.</param>
    protected void WriteToLog(string message)
    {
      Console.WriteLine(string.Format("Can't assign value of parameter '{0}': {1}", this.Name, message));
    }

    /// <summary>
    /// Убираем тэг ShowingPlaceholde из контроля в созданном документе, чтобы значение контроля, после создания документа, можно было редактировать.
    /// </summary>
    private void RemoveShowingPlaceholder()
    {
      this.ContentControl.GetFirstChild<SdtProperties>().RemoveAllChildren<ShowingPlaceholder>();
    }

    #endregion

    #region IParameter

    public string Name { get; private set; }

    public string? Value
    {
      get
      {
        var sb = new StringBuilder();
        foreach (var element in this.openXmlTextElements)
          sb.Append(element.Text);
        return sb.ToString();
      }

      set
      {
        this.ClearAllDescendants();
        this.RemoveShowingPlaceholder();
        this.AssignValue(value);
      }
    }

    #endregion

    #region Конструкторы

    /// <summary>
    /// Конструктор.
    /// </summary>
    /// <param name="contentControl">XML-элемент, являющийся контент контролем.</param>
    public DocxParameter(OpenXmlElement contentControl)
    {
      this.ContentControl = contentControl;
      var alias = this.ContentControl.Elements<SdtProperties>().First().Elements<SdtAlias>().FirstOrDefault();
      this.Name = alias != null ? alias.Val.ToString() : string.Empty;
      this.openXmlTextElements = this.ContentControl.Descendants<Text>().ToList();
      if (this.openXmlTextElements == null || !this.openXmlTextElements.Any())
        this.openXmlTextElements = new List<Text> { new Text(string.Empty) };
    }

    #endregion
  }
}
