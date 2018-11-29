using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.IO.Ports;
using Microsoft.Office.Interop.Excel;

namespace excelctxt1119
{
    public partial class Form1 : Form
    {

        int writecount = 0;
        int readcount = 0;
        byte[] scanID = null;
        List<SerialPort> port = new List<SerialPort>();
        public Form1()
        {
            InitializeComponent();
            LoadComport();
            button4.Visible = false;
            button5.Visible = false;
            

        }

        private void button1_Click(object sender, EventArgs e)
        {

            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

             //   Workbook workbook = new Workbook();
           //    workbook.LoadFromFile(@"D:\michelle\e-iceblue\Spire.XLS\Demos\Data\dataexport.xls", ExcelVersion.Version97to2003);
               // Worksheet sheet = workbook.Worksheets[0];
                  string tableName = "[Sheet1$]";//在頁簽名稱後加$，再用中括號[]包起來
                  string sql = "select * from " + tableName;//SQL查詢
                System.Data.DataTable dt = GetExcelDataTable(openFileDialog1.FileName, sql);
                  dataGridView1.DataSource = dt;

                  DataGridViewColumn dgvc = new DataGridViewTextBoxColumn();
                  dgvc.Width = 60;
                  dgvc.Name = "写入状态";
                  dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                  DataGridViewColumn dgvc2 = new DataGridViewTextBoxColumn();
                  dgvc2.Width = 60;
                  dgvc2.Name = "读取状态";
                  dgvc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                  this.dataGridView1.Columns.Insert(1, dgvc);
                  this.dataGridView1.Columns.Insert(2, dgvc2);
            }
        }


        private System.Data.DataTable GetExcelDataTable(string filePath, string sql)
        {
            //Office 2003
           // OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;Readonly=0'");

            //Office 2007
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");

            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dt.TableName = "tmp";
            conn.Close();
            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = "exporttxt";
            saveFileDialog1.Filter = "txt Files|*.txt;";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                TextWriter writer = new StreamWriter(saveFileDialog1.FileName);
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if(j==0)
                            if(dataGridView1.Rows[i].Cells[j].Value!=null)
                         writer.Write(dataGridView1.Rows[i].Cells[j].Value.ToString());
                    }
                    writer.WriteLine();
                    // writer.WriteLine("-----------------------------------------------------");
                }
                writer.Close();
                MessageBox.Show("Data Exported");
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine("KeyPress" + e.KeyChar);
            DialogResult dialogResult = new DialogResult();

            if ((e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar.CompareTo('0') > 0 || e.KeyChar.CompareTo('9') < 0)
            {
                Console.WriteLine("ININKeyPress" + e.KeyChar);
            }
            else
            {
           //     scanstate.Text = "輸入法請切換英文。";
           //     scanstate.ForeColor = Color.Red;
                return;
            }

            if (e.KeyChar == 13)
            {
                Console.WriteLine(textBox1.Text);
                CheckSyntaxAndReport();
                textBox1.Text = "";
            }
        }

        private void CheckSyntaxAndReport()
        {
            bool isNull = true;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                
                if (dr.Cells[0].Value!=null&&textBox1.Text == dr.Cells[0].Value.ToString())
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[dr.Index].Cells[0];
                    dataGridView1.Rows[dr.Index].Cells[1].Value = DBNull.Value;
                    dataGridView1.Rows[dr.Index].Cells[2].Value = DBNull.Value;
                    isNull = false;
                    byte[] id = TagIDtoByte(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
                    scanID = id;
                    writeTagID(id);
                    break;
                    
                }
            }

            if(isNull)
                MessageBox.Show("data is null");
        }


        //取得Comport
        private void LoadComport()
        {
            List<string> port = Load_AllComPortName();
            string s = button1.Text;
            ComPortList.Items.Clear();
            if (port.Count == 0)
                return;

            int index = 0;
            for (int i = 0; i < port.Count; i++)
            {
                ComPortList.Items.Add(port[i]);
                if (port[i] == s) index = i;
            }

            ComPortList.SelectedIndex = index;
        }


        private void readTagID()
        {
            System.Threading.Thread.Sleep(100);
            byte[] commEnter = new byte[] { 0x43, 0x4D, 0x26, 0x00, 0x07, 0x00, 0x01, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00};
            
            byte[] package = iCheckSum(commEnter);
            for (int i = 0; i < port.Count; i++)
            {
                port[i].Write(package, 0, package.Length);
            }
            readcount++;
        }

        private void writeTagID(byte[] id)
        {
            System.Threading.Thread.Sleep(100);
            byte[] commEnter = new byte[] { 0x43, 0x4D, 0x27, 0x00, 0x13, 0x00, 0x01, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00 };
            byte[] newArray = new byte[commEnter.Length+id.Length];
            commEnter.CopyTo(newArray, 0);
            id.CopyTo(newArray, commEnter.Length);
            byte[] package = iCheckSum(newArray);
            for (int i = 0; i < port.Count; i++)
            {
                port[i].Write(package, 0, package.Length);
            }
            writecount++;
        }


        private byte[] TagIDtoByte(string id)
        {
            byte[] newArray = new byte[id.Length/2];
            int count = 0;
            for (int i = 0; i < id.Length; i=i+2)
            {
                newArray[count]= Convert.ToByte(id.Substring(i, 2), 16);
                count++;
            }
            
          
            return newArray;

        }

        private byte BCC(byte[]  b, int size)
        {
            byte c = 0;
            for (int i = 0; i < size; i++)
            {
                c ^= b[i];
            }
            return c;
        }


        public static byte[] iCheckSum(byte[] data)
        {
            byte[] bytes = new byte[data.Length + 1];
            byte intValue = 0;
            for (int i = 0; i < data.Length; i++)
            {
                if(i>4)
                    intValue = (byte)(intValue ^ data[i]);

                bytes[i] = data[i];
            }
           // byte[] intBytes = BitConverter.GetBytes(intValue);
            //   Array.Reverse(intBytes);
            bytes[data.Length] = intValue;
          //  bytes[data.Length] = (byte)((intBytes[intBytes.Length - 1] ^ (byte)0xff)+(byte)0x01);
            return bytes;
        }

        public List<string> Load_AllComPortName()
        {
            List<string> strs = new List<string>();
            string[] portNames = SerialPort.GetPortNames();
            for (int i = 0; i < (int)portNames.Length; i++)
            {
                strs.Add(portNames[i]);
            }
            return strs;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (port.Count == 0)
                {
                    // port = null;
                    List<string> portList = Load_AllComPortName();
                    for (int sd = 0; sd < portList.Count; sd++)
                    {
                        SerialPort ddd = new SerialPort(portList[sd].ToString(), 115200, Parity.None, 8, StopBits.One);
                        ddd.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);
                        port.Add(ddd);
                    }
                    //port = new SerialPort(ComPortList.Text, Convert.ToInt32(ComPortBaudrate.Text), Parity.None, 8, StopBits.Two);
                    //  port.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);


                }
                else if (!port[0].IsOpen)
                {
                    port.Clear();
                    List<string> portList = Load_AllComPortName();
                    for (int sd = 0; sd < portList.Count; sd++)
                    {
                        SerialPort ddd = new SerialPort(portList[sd], 115200, Parity.None, 8, StopBits.One);

                        ddd.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);
                        port.Add(ddd);
                    }
                    // port = null;
                    // port = new SerialPort(ComPortList.Text, Convert.ToInt32(ComPortBaudrate.Text), Parity.None, 8, StopBits.Two);
                    //port.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);
                }


                if (!port[0].IsOpen)
                {
                    try
                    {
                        for (int sd = 0; sd < port.Count; sd++)
                        {
                            port[sd].Open();
                        }

                    }
                    catch (Exception ex)
                    {
                    //    autoLoadTagID.Stop();
                      //  autoRead.Text = autoLoad;
                        //autoRead.ForeColor = Color.Black;
                        //autoReadclick = false;
                        Console.WriteLine("ex" + ex);
                        for (int sd = 0; sd < port.Count; sd++)
                        {
                            port[sd].Dispose();
                            port[sd].Close();
                        }

                        // MessageBox.Show("串口出問題請重新啟動程式");
                    }
                }
                else
                {
                   // autoLoadTagID.Stop();
                    //autoRead.Text = autoLoad;
                    //autoRead.ForeColor = Color.Black;
                    //autoReadclick = false;
                    for (int sd = 0; sd < port.Count; sd++)
                    {
                        port[sd].Dispose();
                        port[sd].Close();
                    }
                }
                if (port[0].IsOpen == true)
                {
                    ConnectStatus.Text = "连接成功";
                    //button2.Text = ConnectButtonNL;
                    ConnectStatus.ForeColor = Color.Green;
                    //isConnect = true;
                    Console.WriteLine("success");
                }
                else
                {
                    ConnectStatus.Text ="连接失败";
                    //button2.Text = ConnectButtonL;
                    ConnectStatus.ForeColor = Color.Red;
                    ///isConnect = false;
                    //autoLoadTagID.Stop();
                    //autoRead.Text = autoLoad;
                    //autoRead.ForeColor = Color.Black;
                    //autoReadclick = false;
                    Console.WriteLine("fail");
                    for (int sd = 0; sd < port.Count; sd++)
                    {
                        port[sd].Dispose();
                        port[sd].Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please confirm that the device is connected");
            }

        }


        private void port1_DataReceived(object sender, EventArgs e)
        {
            Console.WriteLine("port1_DataReceived");

            List<byte> packet = new List<byte>();
            for (int i = 0; i < port.Count; i++)
            {
                while (port[i].BytesToRead != 0)
                {

                    packet.Add((byte)port[i].ReadByte());


                }
            }


            byte[] bArrary = packet.ToArray();

            if (bArrary.Length<7||bArrary[0] != (byte)0x43|| bArrary[1] != (byte)0x4D)
                return;


            for (int i = 0; i < bArrary.Length; i++)
            {
                Console.WriteLine(bArrary[i].ToString("X2"));
            }

            if (bArrary[2] == (byte)0x26 && bArrary[6] != (byte)0x80 && bArrary[4] != (byte)0x01)
            {
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Value = "读取成功";
                readcount = 0;
               // writeTagID();

            }


             if (bArrary[2] == (byte)0x26 && bArrary[6] == (byte)0x80 && bArrary[4] == (byte)0x01)
            {
                if (readcount < 5)
                    readTagID();
                else
                {
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Value = "读取失败";
                    readcount = 0;
                }
               
            }


            if (bArrary[2] == (byte)0x27 && bArrary[6] == (byte)0x00 && bArrary[4] == (byte)0x01)
            {
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value = "写入成功";

                //System.Threading.Thread.Sleep(200);
                readTagID();

                writecount = 0; 
                //  writeTagID();

            }


            if (bArrary[2] == (byte)0x27 && bArrary[6] == (byte)0x80 && bArrary[4] == (byte)0x01)
            {
                if (writecount < 5)
                    writeTagID(scanID);
                else
                {
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value = "写入失败";
                    writecount = 0;
                }
               

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            byte[] commEnter = new byte[] { 0x43, 0x4D, 0x25,  0x00, 0x00, 0x00, 0x00 };

            byte[] package = iCheckSum(commEnter);
            for (int i = 0; i < port.Count; i++)
            {
                port[i].Write(package, 0, package.Length);
                Console.WriteLine("++++++++++++++++++++");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

          //  byte[] id = new byte[] { 0x87, 0x88, 0x77, 0x55, 0x66, 0x53, 0x22, 0x11, 0x53, 0x88, 0x99, 0x00};
            byte[] id=  TagIDtoByte("029990000000000000002016");
            writeTagID(id);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
