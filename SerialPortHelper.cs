using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;

namespace WeldCheckerLogger
{
    public class SerialPortUser
    {
        /// <summary>
        /// 波特率
        /// </summary>
        public int BaudRate { get; set; }
        /// <summary>
        /// 数据位
        /// </summary>
        public int DataBits { get; set; }
        /// <summary>
        /// 停止位
        /// </summary>
        public StopBits StopBits { get; set; }
        /// <summary>
        /// 奇偶校验
        /// </summary>
        public Parity Parity { get; set; }
        /// <summary>
        /// 端口号
        /// </summary>
        public string PortName;
    }

    public delegate void DelegateReceiveData(Object user, string msg);

    public class SerialPortHelper
    {
        private bool _bOpen = false;
        /// <summary>
        /// 串口打开状态
        /// </summary>
        public bool bOpen { get { return _bOpen; } set { _bOpen = value; } }

        /// <summary>
        /// 接收数据委托对象
        /// </summary>
        public DelegateReceiveData DelegateReceive;

        private SerialPort sp;

        private StringBuilder builder;

        public SerialPortHelper()
        {
            sp = new SerialPort();
            builder = new StringBuilder();
            sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(DataReceived);
        }

        private string OpenSerialPort(SerialPortUser port)
        {
            try
            {
                sp.BaudRate = port.BaudRate;
                sp.DataBits = port.DataBits;
                sp.StopBits = port.StopBits;
                sp.Parity = port.Parity;
                sp.PortName = port.PortName;

                sp.Open();
            }
            catch (System.Exception ex)
            {
                return ex.Message.ToString();
            }
            return string.Empty;
        }

        private void CloseSerialPort()
        {
            if (sp.IsOpen)
                sp.Close();
        }

        private bool SendData(string msg)
        {
            if (ConnectStatu())
            {
                sp.Write(msg);
                return true;
            }
            return false;
        }

        private bool SendData(byte[] bData)
        {
            if (ConnectStatu())
            {
                sp.Write(bData, 0, bData.Length);
                return true;
            }
            return false;
        }

        private bool SendData(char[] bData)
        {
            if (ConnectStatu())
            {
                sp.Write(bData, 0, bData.Length);
                return true;
            }
            return false;
        }

        private bool ConnectStatu()
        {
            return sp.IsOpen;
        }

        private void DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(25);
            string receiveString = string.Empty;
            int length = sp.BytesToRead;
            //byte[] data = new byte[length];
            char[] data = new char[length];
            if (length > 0)
            {
                sp.Read(data, 0, length);
                sp.DiscardInBuffer();

                StringBuilder sb = new StringBuilder();
                foreach (char c in data)
                {
                    sb.Append(c);
                }

                if (DelegateReceive != null)
                    DelegateReceive(sender, sb.ToString());
            }
        }

        public bool Open(SerialPortUser param)
        {
            if (OpenSerialPort(param) == string.Empty)
            {
                return _bOpen = true;
            }
            return false;
        }

        public bool Send(byte[] byData)
        {
            bool temp = SendData(byData);
            return temp;
        }

        public bool Send(char[] cData)
        {
            bool temp = SendData(cData);
            return temp;
        }

        public bool Send(string msg)
        {
            bool temp = SendData(msg);
            return temp;
        }

        public bool Close()
        {
            try
            {
                CloseSerialPort();
                _bOpen = false;
            }
            catch
            {
                return false;
            }
            return true;
        }

        public string[] GetPortNumber()
        {
            return SerialPort.GetPortNames();
        }
    }
}
