namespace Tempate.Parsing
{
  /// <summary>
  /// Интерфейс для работы с параметром в документе.
  /// </summary>
  public interface IParameter
  {
    /// <summary>
    /// Имя параметра.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Значение параметра.
    /// </summary>
    string? Value { get; set; }
  }
}
