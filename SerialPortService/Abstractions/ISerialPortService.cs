using System;
using System.ComponentModel;

namespace SerialPortService.Abstractions
{
    public interface ISerialPortService
    {
        /// <summary>
        /// Статус порта.
        /// </summary>
        bool Online { get; }

        /// <summary>
        /// Содержит имя СОМ порта.
        /// </summary>
        [DefaultValue("COM1")]
        string PortName { get; set; }

        /// <summary>
        /// Содержит скорость соединения.
        /// </summary>
        [DefaultValue(2400)]
        int BaudRate { get; set; }

        /// <summary>
        /// Содержит количество передаваемых бит данных.
        /// </summary>
        [DefaultValue(8)] 
        int DataBit { get; set; }

        /// <summary>
        /// Закрывает СОМ порт.
        /// </summary>
        void Close();

        /// <summary>
        /// Возвращает список портов имеющихся в системе.
        /// </summary>
        string[] GetPortName();

        /// <summary>
        /// Открывает СОМ порт и устанавливает настройки.
        /// </summary>
        /// <returns>Если порт открыт вернет true иначе false.</returns>
        bool Open();

        /// <summary>
        ///  Записывает в порт данные.
        /// </summary>
        /// <param name="tosend">Данные для записи.</param>
        /// <returns>Возвращает количество байт записаных в порт.</returns>
        uint Write(byte[] tosend);

        /// <summary>
        ///  Записывает в порт данные.
        /// </summary>
        /// <param name="data">Данные для записи.</param>
        void Write(byte data);

        /// <summary>
        ///  Записывает в порт данные.
        /// </summary>
        /// <param name="data">Данные для записи.</param>
        /// <returns>Возвращает количество байт записаных в порт.</returns>
        uint Write(string data);

        /// <summary>
        /// Считывает имеющиесяя данные из порта.
        /// </summary>
        /// <param name="data">На вход потаются данные из порта.</param>
        abstract void Read(byte[] data);

        /// <summary>
        /// Срабатывает когда в порт поступили данные.
        /// </summary>
        event Action<byte[]> DataReceived;

        void Dispose();
    }
}
