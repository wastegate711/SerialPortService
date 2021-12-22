using System;
using SerialPortService.Abstractions;
using SerialPortService.Services;

namespace ConsoleTestApp
{
    class Program
    {
        private static ISerialPortService _serialPortService;

        private static void Main(string[] args)
        {
            byte[] data = new byte[10];

            for (int i = 0; i < 10; i++)
            {
                data[i] = (byte)i;
            }

            _serialPortService = new Serial_Port("com3", 115200);
            _serialPortService.DataReceived += SerialPortService_DataReceived;
            _serialPortService.Open();
            _serialPortService.Write(data);
            while (true)
            {
                //var t = _serialPortService.Write(Console.ReadKey().KeyChar.ToString());
                var t=Console.ReadLine();

                if (t.Length > 1)
                {
                    _serialPortService.Write(data);
                }
            }
        }

        private static void SerialPortService_DataReceived(byte[] obj)
        {
            for (int i = 0; i < obj.Length; i++)
            {
                Console.Write(obj[i]);
            }
        }
    }
}
