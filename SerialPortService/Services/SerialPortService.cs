using Microsoft.Win32;
using SerialPortService.Abstractions;
using SerialPortService.Enum;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Threading;

namespace SerialPortService.Services
{
    public class Serial_Port : IDisposable, ISerialPortService
    {
        private BaseSettings _settings;
        private Exception _rxException;
        /// <summary>
        /// Поток считывающий входящие данные.
        /// </summary>
        private Thread _rxThread;
        private CancellationTokenSource _cts = new CancellationTokenSource();
        private ManualResetEvent TransFlag = new ManualResetEvent(true);
        private ManualResetEvent writeEvent = new ManualResetEvent(false);
        private ManualResetEvent startEvent = new ManualResetEvent(false);
        private bool online;
        private bool auto;
        private IntPtr hPort;
        private bool rxExceptionReported = false;
        private int writeCount = 0;
        private bool[] empty = new bool[1];
        private IntPtr ptrUWO;
        private int stateRTS = 2;
        private int stateDTR = 2;
        private int stateBRK = 2;
        private bool checkSends = true;
        private bool dataQueued = false;
        private ASCII RxTerm;
        private ASCII[] TxTerm;
        private ASCII[] RxFilter;
        private uint RxBufferP = 0;
        private byte[] RxBuffer;
        private string RxString = "";

        /// <summary>
        /// Размер буфера входящих данных.
        /// </summary>
        private const int LenRxBuffer = 255;
        /// <summary>
        /// Срабатывает когда в порт поступили данные.
        /// </summary>
        public event Action<byte[]> DataReceived;

        #region Свойства.
        /// <summary>
        /// Статус порта.
        /// </summary>
        public bool Online
        {
            get
            {
                if (!online)
                {
                    return false;
                }
                else
                {
                    return CheckOnline();
                }
            }
        }

        /// <summary>
        /// Содержит имя СОМ порта.
        /// </summary>
        [DefaultValue("COM1")]
        public string PortName
        {
            get => _settings != null ? _settings.Port : string.Empty;
            set
            {
                if (_settings != null)
                    _settings.Port = value;
                else
                    throw new CommPortException($"_settings=null при попытке изменить PortName.");

            }
        }

        /// <summary>
        /// Содержит скорость соединения.
        /// </summary>
        [DefaultValue(2400)]
        public int BaudRate
        {
            get => _settings?.BaudRate ?? 0;
            set
            {
                if (_settings != null)
                    _settings.BaudRate = value;
                else
                    throw new CommPortException($"_settings=null при попытке установить новое значение BaudRate.");
            }
        }

        /// <summary>
        /// Содержит количество передаваемых бит данных.
        /// </summary>
        [DefaultValue(8)]
        public int DataBit
        {
            get => _settings?.dataBits ?? 0;
            set
            {
                if (_settings != null)
                    _settings.dataBits = value;
                else
                    throw new CommPortException($"_settings=null при попытке установить значение DataBit.");
            }
        }
        #endregion
        #region Конструктор.
        /// <summary>
        /// Базовый коструктор, по умолчанию инициализирует "СОМ1" baudrate=2400
        /// </summary>
        public Serial_Port()
        {
            _settings = new BaseSettings();
        }

        /// <summary>
        /// Инициализирует порт с указаным именем порта.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        public Serial_Port(string portName) : this()
        {
            _settings.Port = portName;
        }

        /// <summary>
        /// Инициализирует порт с указанным именем и скоростью.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        /// <param name="speed">Скорость порта.</param>
        public Serial_Port(string portName, int speed)
            : this()
        {
            _settings.Port = portName;
            _settings.BaudRate = speed;
        }

        /// <summary>
        /// Инициализирует порт с указаным именем, скоростью и количеством бит данных.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        /// <param name="speed">Скорость порта.</param>
        /// <param name="dataBits">Количество бит данных.</param>
        public Serial_Port(string portName, int speed, int dataBits)
            : this()
        {
            _settings.Port = portName;
            _settings.BaudRate = speed;
            _settings.dataBits = dataBits;
        }

        /// <summary>
        /// Инициализирует порт с указаным именем, скоростью и количеством бит данных и parity.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        /// <param name="speed">Скорость.</param>
        /// <param name="dataBit">Количество бит данных.</param>
        /// <param name="parity">Parity.</param>
        public Serial_Port(string portName, int speed, int dataBit, Parity parity)
            : this()
        {
            _settings.Port = portName;
            _settings.BaudRate = speed;
            _settings.dataBits = dataBit;
            _settings.ParityMode = parity;
        }

        /// <summary>
        /// Инициализирует порт с указаным именем, скоростью, количеством бит данных, parity и количеством стоповых бит.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        /// <param name="speed">Скорость.</param>
        /// <param name="dataBit">Количество бит данных.</param>
        /// <param name="parity">Parity.</param>
        /// <param name="stopBits">Количество стоп бит.</param>
        public Serial_Port(string portName, int speed, int dataBit, Parity parity, StopBits stopBits)
            : this()
        {
            _settings.Port = portName;
            _settings.BaudRate = speed;
            _settings.dataBits = dataBit;
            _settings.ParityMode = parity;
            _settings.stopBits = stopBits;
        }

        /// <summary>
        /// Инициализирует порт с указаным именем, скоростью, количеством бит данных,
        /// parity, количеством стоповых бит и аппаратным контролем.
        /// </summary>
        /// <param name="portName">Имя порта.</param>
        /// <param name="speed">Скорость.</param>
        /// <param name="dataBit">Количество бит данных.</param>
        /// <param name="parity">Parity.</param>
        /// <param name="stopBits">Количество стоп бит.</param>
        /// <param name="handshake">Аппаратный контроль.</param>
        public Serial_Port(
            string portName,
            int speed,
            int dataBit,
            Parity parity,
            StopBits stopBits,
            Handshake handshake) : this()
        {
            _settings.Port = portName;
            _settings.BaudRate = speed;
            _settings.dataBits = dataBit;
            _settings.ParityMode = parity;
            _settings.stopBits = stopBits;
        }
        #endregion
        #region Деструктор.
        public void Dispose()
        {
            Close();
        }

        ~Serial_Port()
        {
            Close();
        }
        #endregion

        /// <inheritdoc/>
        public string[] GetPortName()
        {
            string[] ports = null;
            RegistryKey registryKey1 = null;
            RegistryKey registryKey2 = null;

            try
            {   // Получаем доступ к ключам реестра.
                registryKey1 = Registry.LocalMachine;
                registryKey2 = registryKey1.OpenSubKey("HARDWARE\\DEVICEMAP\\SERIALCOMM", false);
                new RegistryPermission(RegistryPermissionAccess.Read, "HKEY_LOCAL_MACHINE\\HARDWARE\\DEVICEMAP\\SERIALCOMM").Assert();

                if (registryKey2 != null)
                {
                    string[] valueNames = registryKey2.GetValueNames(); // читаем имена разделов.
                    ports = new string[valueNames.Length];

                    for (Int32 i = 0; i < valueNames.Length; i++)
                    {
                        ports[i] = (string)registryKey2.GetValue(valueNames[i]); // получаем значения разделов реестра.
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                registryKey1?.Close();
                registryKey2?.Close();
                CodeAccessPermission.RevertAssert();
            }

            return ports ?? Array.Empty<string>();
            //return SerialPort.GetPortNames();
        }

        /// <summary>
        /// Проверяет доступность порта.
        /// </summary>
        /// <param name="portName">Имя порта "COM?".</param>
        /// <returns>Статус доступности порта.</returns>
        public static PortStatus IsPortAvailable(string portName)
        {
            IntPtr hPort = WinAPI.CreateFile(
                portName,
                WinAPI.GENERIC_READ | WinAPI.GENERIC_WRITE,
                0,
                IntPtr.Zero,
                WinAPI.OPEN_EXISTING,
                WinAPI.FILE_FLAG_OVERLAPPED,
                IntPtr.Zero);

            if (hPort == (IntPtr)WinAPI.INVALID_HANDLE_VALUE)
            {
                if (Marshal.GetLastWin32Error() == WinAPI.ERROR_ACCESS_DENIED)
                {
                    return PortStatus.Unavailable;
                }
                else
                {
                    //JH 1.3: Automatically try AltName if supplied name fails:
                    hPort = WinAPI.CreateFile(
                        portName,
                        WinAPI.GENERIC_READ | WinAPI.GENERIC_WRITE,
                        0,
                        IntPtr.Zero,
                        WinAPI.OPEN_EXISTING,
                        WinAPI.FILE_FLAG_OVERLAPPED,
                        IntPtr.Zero);

                    if (hPort == (IntPtr)WinAPI.INVALID_HANDLE_VALUE)
                    {
                        return Marshal.GetLastWin32Error() == WinAPI.ERROR_ACCESS_DENIED
                            ? PortStatus.Unavailable
                            : PortStatus.Absent;
                    }
                }
            }

            WinAPI.CloseHandle(hPort);
            return PortStatus.Available;
        }

        /// <inheritdoc/>
        public void Close()
        {
            if (online)
            {
                auto = false;
                InternalClose();
                _rxException = null;
            }
        }

        /// <inheritdoc/>        
        public bool Open()
        {
            WinAPI.DCB PortDCB = new WinAPI.DCB();
            WinAPI.COMMTIMEOUTS CommTimeouts = new WinAPI.COMMTIMEOUTS();
            //CommBaseSettings commSettings;
            WinAPI.OVERLAPPED wo = new WinAPI.OVERLAPPED();
            WinAPI.COMMPROP cp;

            //если порт открыт выходим.
            if (Online)
            {
                return false;
            }

            //commSettings = CommSettings();
            //_settings.BaudRate = BaudRate;
            hPort = WinAPI.CreateFile(
                _settings.Port,
                WinAPI.GENERIC_READ | WinAPI.GENERIC_WRITE,
                0,
                IntPtr.Zero,
                WinAPI.OPEN_EXISTING,
                WinAPI.FILE_FLAG_OVERLAPPED,
                IntPtr.Zero);

            if (hPort == (IntPtr)WinAPI.INVALID_HANDLE_VALUE)
            {
                if (Marshal.GetLastWin32Error() == WinAPI.ERROR_ACCESS_DENIED)
                {
                    return false;
                }
                else
                {
                    //Попытка открыть порт с альтернативным именем.
                    hPort = WinAPI.CreateFile(
                        AltName(_settings.Port),
                        WinAPI.GENERIC_READ | WinAPI.GENERIC_WRITE,
                        0,
                        IntPtr.Zero,
                        WinAPI.OPEN_EXISTING,
                        WinAPI.FILE_FLAG_OVERLAPPED,
                        IntPtr.Zero);

                    if (hPort == (IntPtr)WinAPI.INVALID_HANDLE_VALUE)
                    {
                        if (Marshal.GetLastWin32Error() == WinAPI.ERROR_ACCESS_DENIED)
                        {
                            return false;
                        }
                        else
                        {
                            throw new CommPortException("Открытие порта с альтернативным именем не возможно.");
                        }
                    }
                }
            }

            online = true;

            //JH1.1: Changed from 0 to "magic number" to give instant return on ReadFile:
            CommTimeouts.ReadIntervalTimeout = WinAPI.MAXDWORD;
            CommTimeouts.ReadTotalTimeoutConstant = 0;
            CommTimeouts.ReadTotalTimeoutMultiplier = 0;

            //JH1.2: 0 does not seem to mean infinite on non-NT platforms, so default it to 10
            //seconds per byte which should be enough for anyone.
            if (_settings.sendTimeoutMultiplier == 0)
            {
                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    CommTimeouts.WriteTotalTimeoutMultiplier = 0;
                }
                else
                {
                    CommTimeouts.WriteTotalTimeoutMultiplier = 10000;
                }
            }
            else
            {
                CommTimeouts.WriteTotalTimeoutMultiplier = _settings.sendTimeoutMultiplier;
            }
            CommTimeouts.WriteTotalTimeoutConstant = _settings.sendTimeoutConstant;

            PortDCB.Init(
                ((_settings.ParityMode == Parity.Odd) || (_settings.ParityMode == Parity.Even)),
                _settings.txFlowCTS,
                _settings.txFlowDSR,
                (int)_settings.useDTR,
                _settings.rxGateDSR,
                !_settings.txWhenRxXoff,
                _settings.txFlowX,
                _settings.rxFlowX,
                (int)_settings.useRTS);

            PortDCB.BaudRate = _settings.BaudRate;
            PortDCB.ByteSize = (byte)_settings.dataBits;
            PortDCB.Parity = (byte)_settings.ParityMode;
            PortDCB.StopBits = (byte)_settings.stopBits;
            PortDCB.XoffChar = (byte)_settings.XoffChar;
            PortDCB.XonChar = (byte)_settings.XonChar;

            if ((_settings.rxQueue != 0) || (_settings.txQueue != 0))
            {
                if (!WinAPI.SetupComm(hPort, (uint)_settings.rxQueue, (uint)_settings.txQueue))
                {
                    ThrowException("Bad queue settings");
                }
            }

            //JH 1.2: Defaulting mechanism for handshake thresholds - prevents problems of setting specific
            //defaults which may violate the size of the actually granted queue. If the user specifically sets
            //these values, it's their problem!
            if ((_settings.rxLowWater == 0) || (_settings.rxHighWater == 0))
            {
                if (!WinAPI.GetCommProperties(hPort, out cp))
                {
                    cp.dwCurrentRxQueue = 0;
                }

                if (cp.dwCurrentRxQueue > 0)
                {
                    //If we can determine the queue size, default to 1/10th, 8/10ths, 1/10th.
                    //Note that HighWater is measured from top of queue.
                    PortDCB.XoffLim = PortDCB.XonLim = (short)((int)cp.dwCurrentRxQueue / 10);
                }
                else
                {
                    //If we do not know the queue size, set very low defaults for safety.
                    PortDCB.XoffLim = PortDCB.XonLim = 8;
                }
            }
            else
            {
                PortDCB.XoffLim = (short)_settings.rxHighWater;
                PortDCB.XonLim = (short)_settings.rxLowWater;
            }

            if (!WinAPI.SetCommState(hPort, ref PortDCB))
            {
                ThrowException("Bad com settings");
            }

            if (!WinAPI.SetCommTimeouts(hPort, ref CommTimeouts))
            {
                ThrowException("Bad timeout settings");
            }

            stateBRK = 0;

            if (_settings.useDTR == HSOutput.None)
            {
                stateDTR = 0;
            }

            if (_settings.useDTR == HSOutput.Online)
            {
                stateDTR = 1;
            }

            if (_settings.useRTS == HSOutput.None)
            {
                stateRTS = 0;
            }

            if (_settings.useRTS == HSOutput.Online)
            {
                stateRTS = 1;
            }

            checkSends = _settings.checkAllSends;
            wo.Offset = 0;
            wo.OffsetHigh = 0;

            if (checkSends)
            {
                wo.hEvent = writeEvent.Handle;
            }
            else
            {
                wo.hEvent = IntPtr.Zero;
            }

            ptrUWO = Marshal.AllocHGlobal(Marshal.SizeOf(wo));

            Marshal.StructureToPtr(wo, ptrUWO, true);
            writeCount = 0;
            //JH1.3:
            empty[0] = true;
            dataQueued = false;
            
            _rxException = null;
            rxExceptionReported = false;
            _rxThread = new Thread(ReceiveThread);
            _rxThread.Name = "SerialPortReceiveData";
            _rxThread.Priority = ThreadPriority.AboveNormal;
            _rxThread.Start(_cts.Token);

            //JH1.2: More robust thread start-up wait.
            startEvent.WaitOne(500, false);
            auto = false;
            auto = _settings.autoReopen;
            return true;
        }

        /// <inheritdoc/> 
        public virtual void Read(byte[] data)
        {
            DataReceived?.Invoke(data);
        }

        /// <inheritdoc/>
        public uint Write(byte[] tosend)
        {
            uint sent = 0;

            if (!CheckOnline()) // проверка, что порт открыт.
            {
                return 0;
            }

            CheckResult(); //
            writeCount = tosend.GetLength(0);
            var result = WinAPI.WriteFile(hPort, tosend, (uint)writeCount, out sent, ptrUWO);

            if (result)
            {
                writeCount -= (int)sent;
            }
            else
            {
                if (Marshal.GetLastWin32Error() != WinAPI.ERROR_IO_PENDING)
                {
                    ThrowException("Ошибка передачи.");
                }
                //JH1.3:
                dataQueued = true;

                return 0;
            }

            return sent;
        }

        /// <inheritdoc/>
        public void Write(byte data)
        {
            byte[] buf = new byte[1];
            buf[0] = data;
            Write(buf);
        }

        public uint Write(string data)
        {
            return Write(Encoding.ASCII.GetBytes(data));
        }

        private void CheckResult()
        {
            uint sent = 0;

            //JH 1.3: Fixed a number of problems working with checkSends == false. Byte counting was unreliable because
            //occasionally GetOverlappedResult would return true with a completion having missed one or more previous
            //completions. The test for ERROR_IO_INCOMPLETE was incorrectly for ERROR_IO_PENDING instead.

            if (writeCount > 0)
            {
                if (WinAPI.GetOverlappedResult(hPort, ptrUWO, out sent, checkSends))
                {
                    if (checkSends)
                    {
                        writeCount -= (int)sent;

                        if (writeCount != 0)
                        {
                            ThrowException("Send Timeout");
                        }

                        writeCount = 0;
                    }
                }
                else
                {
                    if (Marshal.GetLastWin32Error() != WinAPI.ERROR_IO_INCOMPLETE)
                    {
                        ThrowException("Write Error");
                    }
                }
            }
        }

        /// <summary>
        /// Метод выполняется в отдельном потоке и считывает входящие данные.
        /// </summary>
        private void ReceiveThread(object cancel)
        {
            CancellationToken cancellationToken = (CancellationToken)cancel;

            if (cancellationToken == null)
            {
                throw new CommPortException($"cancel=null не удачная передача токена отмены в принимающий поток.");
            }

            byte[] buf = new byte[LenRxBuffer];
            uint gotbytes;
            bool starting;

            starting = true;
            AutoResetEvent sg = new AutoResetEvent(false);
            WinAPI.OVERLAPPED ov = new WinAPI.OVERLAPPED();

            IntPtr unmanagedOv;
            IntPtr uMask;
            uint eventMask = 0;
            unmanagedOv = Marshal.AllocHGlobal(Marshal.SizeOf(ov));
            uMask = Marshal.AllocHGlobal(Marshal.SizeOf(eventMask));

            ov.Offset = 0;
            ov.OffsetHigh = 0;
            ov.hEvent = sg.Handle;
            Marshal.StructureToPtr(ov, unmanagedOv, true);

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    if (!WinAPI.SetCommMask(hPort, WinAPI.EV_RXCHAR | WinAPI.EV_TXEMPTY | WinAPI.EV_CTS | WinAPI.EV_DSR
                        | WinAPI.EV_BREAK | WinAPI.EV_RLSD | WinAPI.EV_RING | WinAPI.EV_ERR))
                    {
                        throw new CommPortException("IO Error [001]");
                    }

                    Marshal.WriteInt32(uMask, 0);
                    //JH 1.2: Tells the main thread that this thread is ready for action.
                    if (starting)
                    {
                        startEvent.Set();
                        starting = false;
                    }
                    if (!WinAPI.WaitCommEvent(hPort, uMask, unmanagedOv))
                    {
                        if (Marshal.GetLastWin32Error() == WinAPI.ERROR_IO_PENDING)
                        {
                            sg.WaitOne();
                        }
                        else
                        {
                            throw new CommPortException("Не удалось установить событие в неуправляемой памяти. [002]");
                        }
                    }

                    eventMask = (uint)Marshal.ReadInt32(uMask);

                    if ((eventMask & WinAPI.EV_ERR) != 0)
                    {
                        UInt32 errs;

                        if (WinAPI.ClearCommError(hPort, out errs, IntPtr.Zero))
                        {
                            //JH 1.2: BREAK condition has an error flag and and an event flag. Not sure if both
                            //are always raised, so if CE_BREAK is only error flag ignore it and set the EV_BREAK
                            //flag for normal handling. Also made more robust by handling case were no recognised
                            //error was present in the flags. (Thanks to Fred Pittroff for finding this problem!)
                            int ec = 0;
                            StringBuilder s = new StringBuilder("UART Error: ", 40);

                            if ((errs & WinAPI.CE_FRAME) != 0)
                            {
                                s = s.Append("Framing,");
                                ec++;
                            }
                            if ((errs & WinAPI.CE_IOE) != 0)
                            {
                                s = s.Append("IO,");
                                ec++;
                            }
                            if ((errs & WinAPI.CE_OVERRUN) != 0)
                            {
                                s = s.Append("Overrun,");
                                ec++;
                            }
                            if ((errs & WinAPI.CE_RXOVER) != 0)
                            {
                                s = s.Append("Receive Cverflow,");
                                ec++;
                            }
                            if ((errs & WinAPI.CE_RXPARITY) != 0)
                            {
                                s = s.Append("Parity,");
                                ec++;
                            }
                            if ((errs & WinAPI.CE_TXFULL) != 0)
                            {
                                s = s.Append("Transmit Overflow,");
                                ec++;
                            }
                            if (ec > 0)
                            {
                                s.Length = s.Length - 1;
                                throw new CommPortException(s.ToString());
                            }
                            else
                            {
                                if (errs == WinAPI.CE_BREAK)
                                {
                                    eventMask |= WinAPI.EV_BREAK;
                                }
                                else
                                {
                                    throw new CommPortException("IO Error [003]");
                                }
                            }
                        }
                        else
                        {
                            throw new CommPortException("IO Error [003]");
                        }
                    }
                    if ((eventMask & WinAPI.EV_RXCHAR) != 0)
                    {
                        do
                        {
                            gotbytes = 0;

                            if (!WinAPI.ReadFile(hPort, buf, (uint)buf.Length, out gotbytes, unmanagedOv))
                            {
                                //JH 1.1: Removed ERROR_IO_PENDING handling as comm timeouts have now
                                //been set so ReadFile returns immediately. This avoids use of CancelIo
                                //which was causing loss of data. Thanks to Daniel Moth for suggesting this
                                //might be a problem, and to many others for reporting that it was!

                                int x = Marshal.GetLastWin32Error();

                                throw new CommPortException("IO Error [004]");
                            }
                            if (gotbytes > 0)
                            {
                                //OnRxChar(buf[0]);
                                byte[] recievData = new byte[gotbytes];

                                //нужно передать только прешедшие данные, а не весь массив целиком.
                                for (int k = 0; k < gotbytes; k++)
                                {
                                    recievData[k] = buf[k];
                                }

                                Read(recievData);
                            }

                        } while (gotbytes > 0);
                    }

                    if ((eventMask & WinAPI.EV_TXEMPTY) != 0)
                    {
                        //JH1.3:
                        lock (empty)
                        {
                            empty[0] = true;
                        }
                    }

                    if ((eventMask & WinAPI.EV_BREAK) != 0)
                    {
                        //OnBreak();
                    }

                    uint i = 0;

                    if ((eventMask & WinAPI.EV_CTS) != 0)
                    {
                        i |= WinAPI.MS_CTS_ON;
                    }

                    if ((eventMask & WinAPI.EV_DSR) != 0)
                    {
                        i |= WinAPI.MS_DSR_ON;
                    }
                    if ((eventMask & WinAPI.EV_RLSD) != 0)
                    {
                        i |= WinAPI.MS_RLSD_ON;
                    }

                    if ((eventMask & WinAPI.EV_RING) != 0)
                    {
                        i |= WinAPI.MS_RING_ON;
                    }

                    if (i != 0)
                    {
                        uint f;

                        if (!WinAPI.GetCommModemStatus(hPort, out f))
                        {
                            throw new CommPortException("IO Error [005]");
                        }

                        //OnStatusChange(new ModemStatus(i), new ModemStatus(f));
                    }
                }
            }
            catch (Exception e)
            {
                //JH 1.3: Added for shutdown robustness (Thanks to Fred Pittroff, Mark Behner and Kevin Williamson!), .
                WinAPI.CancelIo(hPort);

                if (uMask != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(uMask);
                }

                if (unmanagedOv != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(unmanagedOv);
                }

                if (!(e is ThreadAbortException))
                {
                    _rxException = e;
                    //OnRxException(e);
                }
            }
        }

        protected void OnRxChar(byte ch)
        {
            ASCII ca = (ASCII)ch;

            if ((ca == RxTerm) || (RxBufferP > RxBuffer.GetUpperBound(0)))
            {
                //JH 1.1: Use static encoder for efficiency. Thanks to Prof. Dr. Peter Jesorsky!
                lock (RxString)
                {
                    RxString = Encoding.ASCII.GetString(RxBuffer, 0, (int)RxBufferP);
                }

                RxBufferP = 0;

                if (TransFlag.WaitOne(0, false))
                {
                    //OnRxLine(RxString);
                }
                else
                {
                    TransFlag.Set();
                }
            }
            else
            {
                bool wr = true;

                if (RxFilter != null)
                {
                    for (int i = 0; i <= RxFilter.GetUpperBound(0); i++)
                    {
                        if (RxFilter[i] == ca)
                        {
                            wr = false;
                        }
                    }
                }
                if (wr)
                {
                    RxBuffer[RxBufferP] = ch;
                    RxBufferP++;
                }
            }
        }

        /// <summary>
        /// Проверяет открыт порт или нет.
        /// </summary>
        /// <returns>Если порт открыт возвращает true.</returns>
        private bool CheckOnline()
        {
            if ((_rxException != null) && (!rxExceptionReported))
            {
                rxExceptionReported = true;
                ThrowException("rx");
            }
            if (online)
            {
                //JH 1.1: Avoid use of GetHandleInformation for W98 compatability.
                if (hPort != (IntPtr)WinAPI.INVALID_HANDLE_VALUE)
                {
                    return true;
                }

                ThrowException("Offline");
                return false;
            }
            else
            {
                if (auto)
                {
                    if (Open())
                    {
                        return true;
                    }
                }

                ThrowException("Offline");
                return false;
            }
        }
        /// <summary>
        /// Останавливает ввод/вывод, закрывает хендел и высвобождает память.
        /// </summary>
        private void InternalClose()
        {
            WinAPI.CancelIo(hPort);

            if (_rxThread != null)
            {
                try
                {
                    _cts.Cancel();
                    //JH 1.3: Improve robustness of Close in case were followed by Open:
                    _rxThread.Join(100);
                    _cts.Dispose();
                    _rxThread = null;
                }
                catch (ThreadAbortException abort)
                {

                }
                catch (PlatformNotSupportedException ex)
                {

                }
                catch (Exception ex)
                {
                    throw new CommPortException(ex.Message);
                }
            }

            WinAPI.CloseHandle(hPort);

            if (ptrUWO != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(ptrUWO);
            }

            stateRTS = 2;
            stateDTR = 2;
            stateBRK = 2;
            online = false;
        }

        /// <summary>
        /// Use this to throw exceptions in derived classes. Correctly handles threading issues
        /// and closes the port if necessary.
        /// </summary>
        /// <param name="reason">Description of fault</param>
        private void ThrowException(string reason)
        {
            if (Thread.CurrentThread == _rxThread)
            {
                throw new CommPortException(reason);
            }
            else
            {
                if (online)
                {
                    InternalClose();
                }
                if (_rxException == null)
                {
                    throw new CommPortException(reason);
                }
                else
                {
                    throw new CommPortException(_rxException);
                }
            }
        }

        /// <summary>
        /// Возвращает альтернативное название порта \\.\COM1 for COM1:
        /// Некоторые системы требуют этой формы для двойного или большего количества чисел COM-порта цифры.
        /// </summary>
        /// <param name="s">Имя порта COM1 или COM1:</param>
        /// <returns>Имя в виде \\.\COM1</returns>
        private string AltName(string s)
        {
            if (s.EndsWith(":"))
            {
                s = s.Substring(0, s.Length - 1);
            }

            if (s.StartsWith(@"\"))
            {
                return s;
            }

            return @"\\.\" + s;
        }
    }
}
