using SerialPortService.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SerialPortService.Abstractions
{
    public interface ISerialPortService
    {
        /// <summary>
        /// Закрывает СОМ порт.
        /// </summary>
        void Close();

        /// <summary>
        /// Возвращает список портов имеющихся в системе.
        /// </summary>
        string[] GetPortName();

        /// <summary>
        /// Проверяет доступность порта.
        /// </summary>
        /// <param name="portName">Имя порта "COM?".</param>
        /// <returns>Статус доступности порта.</returns>
        PortStatus IsPortAvailable(string portName);

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
    }
}
