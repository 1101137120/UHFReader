using System;
using System.Collections.Generic;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;

namespace exceltxt1119
{
    class Tools
    {
        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:X2}", b);
            return hex.ToString();
        }

        public static byte[] StringToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }

        public static string IntToHex(int iValue, int len)
        {
            string data = null;
            if (len == 1)
            {
                data = iValue.ToString("X");
            }
            else
            if (len == 2)
            {
                data = iValue.ToString("X2");
            }
            else if (len == 3)
            {
                data = iValue.ToString("X3");
            }
            else if (len == 4)
            {
                data = iValue.ToString("X4");
            }
            else if (len == 5)
            {
                data = iValue.ToString("X5");
            }
            else if (len == 6)
            {
                data = iValue.ToString("X6");
            }

            Console.WriteLine("INTTOHEX:"+ data);
            return data;
        }
        public static byte[] iCheckSum(byte[] data)
        {
            byte[] bytes = new byte[2];
            int intValue = 0;
            for (int i = 0; i < data.Length; i++)
            {
                intValue = intValue + (int)data[i];
            }
            byte[] intBytes = BitConverter.GetBytes(intValue);
            Array.Reverse(intBytes);
            bytes[0] = intBytes[intBytes.Length - 2];
            bytes[1] = intBytes[intBytes.Length - 1];
            //Console.WriteLine(ByteArrayToString(bytes) + "");
            return bytes;
        }


        //字串 轉 hex
        public static string ConvertStringToHex(string text)
        {
            byte[] byt = System.Text.UnicodeEncoding.UTF8.GetBytes(text);
            return ByteArrayToString(byt);
        }

        //hex 轉 字串
        public static string ConvertHexToString(byte[] HexValue)
        {
            return UnicodeEncoding.UTF8.GetString(HexValue);
        }


        //2 to 16
        public static string ConvertBinaryToHex(string strBinary)
        {
            string strHex = Convert.ToInt32(strBinary, 2).ToString("x8");
            return strHex;
        }


        public static int ConvertHexToInt(string hex)
        {
            int num = Int32.Parse(hex, System.Globalization.NumberStyles.HexNumber);
           // Console.WriteLine(num);

            return num;
        }

        public static string ConvertHexToString(string HexValue)
        {
            string StrValue = "";
            while (HexValue.Length > 0)
            {
                StrValue += System.Convert.ToChar(System.Convert.ToUInt32(HexValue.Substring(0, 2), 16)).ToString();
                HexValue = HexValue.Substring(2, HexValue.Length - 2);
            }
            return StrValue;
        }



        //------------------------------------------------------------
        public event EventHandler onApScanEvent;
        public class ApScanEventArgs : EventArgs //SMCEslReceiveEventArgs
        {
            public List<AP_Information> data;
        }
        public class AP_Information
        {
            public string AP_IP = "";
            public string AP_MAC_Address = "";
            public string AP_Name = "";
        }
        public void SNC_GetAP_Info()
        {
            List<AP_Information> old = new List<AP_Information> { };
            byte[] data = new byte[4]; //broadcast data
            data[0] = 0xff;
            data[1] = 0x01;
            data[2] = 0x01;
            data[3] = 0x02;

            IPEndPoint ip = new IPEndPoint(IPAddress.Broadcast, 1500); //braodcast IP address, and corresponding port
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces(); //get all network interfaces of the computer
            foreach (NetworkInterface adapter in nics)
            {
                // Only select interfaces that are Ethernet type and support IPv4 (important to minimize waiting time)
                if (adapter.NetworkInterfaceType != NetworkInterfaceType.Ethernet) { continue; }
                if (adapter.Supports(NetworkInterfaceComponent.IPv4) == false) { continue; }
                try
                {
                    IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                    foreach (var ua in adapterProperties.UnicastAddresses)
                    {
                        if (ua.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            //SEND BROADCAST IN THE ADAPTER
                            Socket bcSocket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp); //broadcast socket
                            bcSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, 1);
                            bcSocket.ReceiveTimeout = 200; //receive timout 200ms
                            IPEndPoint myLocalEndPoint = new IPEndPoint(ua.Address, 20058);
                            bcSocket.Bind(myLocalEndPoint);
                            bcSocket.SendTo(data, ip);
    
                            //RECEIVE BROADCAST IN THE ADAPTER
                            int BUFFER_SIZE_ANSWER = 1024;
                            byte[] bufferAnswer = new byte[BUFFER_SIZE_ANSWER];
                            do
                            {
                                try
                                {
                                    var recv = bcSocket.Receive(bufferAnswer);
                                    var redata = new byte[recv];
                                    Array.Copy(bufferAnswer, 0, redata, 0, recv);

                                    if (recv == 36)
                                    {
                                        string dataS = Tools.ByteArrayToString(redata);

                                        int ip1 = Tools.ConvertHexToInt(dataS.Substring(10, 2));
                                        int ip2 = Tools.ConvertHexToInt(dataS.Substring(12, 2));
                                        int ip3 = Tools.ConvertHexToInt(dataS.Substring(14, 2));
                                        int ip4 = Tools.ConvertHexToInt(dataS.Substring(16, 2));

                                        string AP_IP = ip1 + "." + ip2 + "." + ip3 + "." + ip4;
                                        string MacAddress = dataS.Substring(18, 12);
                                        string DeviceName = dataS.Substring(38, 32);

                                        AP_Information mAP_Information = new AP_Information();
                                       mAP_Information.AP_IP = AP_IP;
                                        mAP_Information.AP_MAC_Address = MacAddress;
                                        mAP_Information.AP_Name = Tools.ConvertHexToString( DeviceName) ;

                                        old.Add(mAP_Information);
                                    }
                                }
                                catch (Exception e)
                                {
                                   // Console.Write(e.ToString() + Environment.NewLine);
                                   // bcSocket.Close();
                                    break;
                                }
                            } while (bcSocket.ReceiveTimeout != 0); //fixed receive timeout for each adapter that supports our broadcast
                            bcSocket.Close();
                        }
                    }
                }

                catch { }
            }
            ApScanEventArgs obj = new ApScanEventArgs();
            obj.data = old;
            onApScanEvent(this, obj);
        }

        //-------------------------------------------------------

    }
}
