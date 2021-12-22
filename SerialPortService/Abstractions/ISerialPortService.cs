﻿using SerialPortService.Enum;
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
        public event Action<byte[]> DataReceived;
    }
}
