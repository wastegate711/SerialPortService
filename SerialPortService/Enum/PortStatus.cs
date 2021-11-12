namespace SerialPortService.Enum
{
    /// <summary>
    /// Статус доступности порта.
    /// </summary>
    public enum PortStatus
    {


        /// <summary>
        /// Порт существует, но не доступен (может быть открыт для другой программы).
        /// </summary>
        Unavailable = 0,
        /// <summary>
        /// Доступен для использования.
        /// </summary>
        Available = 1,
        /// <summary>
        /// Порт не существует.
        /// </summary>
        Absent = -1

    }
}
