using System;
using System.Drawing;
using System.IO.Ports;
using System.Text;
using System.Windows.Forms;
using Jitbit.Utils;
using SimpleTCP;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using OfficeOpenXml;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net.Sockets;
using System.Net;
using System.Net.NetworkInformation;
using System.Windows.Ink;

namespace WeldCheckerLogger
{
    public partial class Form1 : Form
    {
        CsvExport myExport = new CsvExport();
        private SerialPortHelper serialPortHelper;
        private Header header;
        private Measurement measurement;
        private string fileName232;
        private string[] measurementArray;
        string me;
        string con;
        string me_cu;
        string con_cu;
        string thu_3;
        int j = 0;
        int k = 0;
        Byte[] bytes;
        Socket connection;
        int type;

        public Form1()
        {
            InitializeComponent();
            LoadDataGridView1();
            LoadDataGridView2();
            bytes = new Byte[1024];
        }   

        private void Form1_Load(object sender, EventArgs e)
        {
            header = new Header();
            measurement = new Measurement();
        }

        #region DataReceive 232
        private void ReceiveData(Object user, string msg)
        {
            SetLabReceive(msg);
            //FormatHeader(msg, "232");
            // FormatMeasurements(msg.Substring(28), "232");
        }

        private delegate void DelegateLabReceive(string Text);
        private void SetLabReceive(string Text)
        {
            try
            {
                if (tbStatus232.InvokeRequired)
                {
                    DelegateLabReceive stcb = new DelegateLabReceive(SetLabReceive);
                    tbStatus232.Invoke(stcb, new object[] { Text });
                }
                else
                {
                    if (type == 1)
                    {
                        tbStatus232.Text = Text;

                        string[] barcode;

                        barcode = Text.Split('\r');

                        for (int i = 0; i < barcode.Length; i++)
                        {
                            //  MessageBox.Show("Phan tu thu " + i + " =" + barcode[i]);
                            // newbarcode[i] = barcode[i].Replace("\r","");

                            if (barcode[i].Length == 10 || barcode[i].Length == 9)
                            {
                                string mame = barcode[i];
                                string mame1 = mame.Trim();
                                me = mame1;
                            }
                            if (barcode[i].Length == 15 || barcode[i].Length == 16)
                            {
                                string macon = barcode[i];
                                string macon1 = macon.Trim();
                                con = macon1;
                            }

                        }

                        if (me != null && con != con_cu && con != null)
                        {
                            j++;
                            dataTable.Rows.Add(j, me, con);
                            btn_Status_Code.BackColor = Color.Green;
                            con_cu = con;
                            tab232.BackColor = Color.Beige;

                        }
                        else if (con == con_cu && con != null)
                        {
                            tbStatus232.Text = "Warning: This code has been read, please read with another code";
                            btn_Status_Code.BackColor = Color.Yellow;
                            tab232.BackColor = Color.Yellow;
                        }
                        else if (con == null || me == null)
                        {
                            tbStatus232.Text = "Warning: The code is unreadable, change the distance and press capture the code again.";
                            btn_Status_Code.BackColor = Color.Red;
                            tab232.BackColor = Color.Red;
                        }
                        // me_cu = me;

                        me = null;
                        con = null;
                    }
                    if (type == 2)
                    {
                        tbStatus232.Text = Text;

                        string[] barcode;

                        barcode = Text.Split('\r');

                        for (int i = 0; i < barcode.Length; i++)
                        {
                            //  MessageBox.Show("Phan tu thu " + i + " =" + barcode[i]);
                            // newbarcode[i] = barcode[i].Replace("\r","");

                            if (barcode[i].Length == 10 || barcode[i].Length == 11)
                            {
                                string mame = barcode[i];
                                string mame1 = mame.Trim();
                                me = mame1;
                            }
                            if (barcode[i].Length == 15 || barcode[i].Length == 16)
                            {
                                string macon = barcode[i];
                                string macon1 = macon.Trim();
                                con = macon1;
                            }
                            if (barcode[i].Length == 2 || barcode[i].Length == 3 || barcode[i].Length == 4 || barcode[i].Length == 5 || barcode[i].Length == 6)
                            {
                                string ma3 = barcode[i];
                                string ma3_1 = ma3.Trim();
                                thu_3 = ma3_1;
                            }

                        }

                        if (me != null && con != con_cu && con != null)
                        {
                            k++;
                            dataTable_Type2.Rows.Add(k, me, con, thu_3);
                            btn_Status_Code.BackColor = Color.Green;
                            con_cu = con;
                            tab232.BackColor = Color.Beige;

                        }
                        else if (con == con_cu && con != null)
                        {
                            tbStatus232.Text = "Warning: This code has been read, please read with another code";
                            btn_Status_Code.BackColor = Color.Yellow;
                            tab232.BackColor = Color.Yellow;
                        }
                        else if (con == null || me == null)
                        {
                            tbStatus232.Text = "Warning: The code is unreadable, change the distance and press capture the code again.";
                            btn_Status_Code.BackColor = Color.Red;
                            tab232.BackColor = Color.Red;
                        }
                        // me_cu = me;

                        me = null;
                        con = null;
                    }


                }
            }
            catch { }
        }

        #endregion

        #region DataReceive TCP
     
        #endregion

        #region Data Format
       

        private void FormatMeasurements(string inputMeasurements, string protocol)
        {
            char delimiter = ',';
            measurementArray = inputMeasurements.Split(delimiter);

            for(int i = 0; i < 10; i++)
            {
                CSVExportMeasurement(i, measurementArray[i*3+0], measurementArray[i*3+1], measurementArray[i*3+2], protocol);
            }
            Array.Clear(measurementArray, 0, 30);
        }
        #endregion
        #region CSV Export
       
        private void CSVExportMeasurement(int index, string mic, string judgment, string measUnit, string protocol)
        {
            try
            {
                myExport["Measurement item code" + index.ToString()] = mic;
                myExport["Judgment" + index.ToString()] = judgment;
                myExport["Value" + index.ToString()] = measUnit;

                if (protocol == "232")
                {
                   // myExport.ExportToFile(savePath + fileName232);
                }
                else if (protocol == "tcp")
                {
                  //  myExport.ExportToFile(savePath + fileNameTCP);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
        #endregion

        #region Buttons RS232
        private void butComOpen_Click(object sender, EventArgs e)
        {
            if (!serialPortHelper.bOpen)
            {
                if (cbSerial.Text == "")
                {
                    MessageBox.Show("COM cannot be empty!");
                    return;
                }
                if (cbBaudRate.Text == "")
                {
                    MessageBox.Show("BaudRate cannot be empty!");
                    return;
                }

                SerialPortUser param = new SerialPortUser();
                param.PortName = cbSerial.Text;
                param.BaudRate = Convert.ToInt32(cbBaudRate.Text);
                param.DataBits = 8;
                param.Parity = Parity.None;
                param.StopBits = StopBits.One;
                if (serialPortHelper.Open(param))
                {
                    butComOpen.Text = "Close COM";
                    panColor.BackColor = Color.Green;
                    ControlEnabled(false);
                }
                else
                {
                    butComOpen.Text = "Open COM";
                    panColor.BackColor = Color.Red;
                }
            }
            else
            {
                serialPortHelper.Close();
                butComOpen.Text = "Open COM";
                panColor.BackColor = Color.Yellow;
                ControlEnabled(true);
            }
        }

        private void ControlEnabled(bool bEnabled)
        {
            cbSerial.Enabled = bEnabled;
            cbBaudRate.Enabled = bEnabled;
        }

      
        private void btnInit232_Click(object sender, EventArgs e)
        {
            serialPortHelper = new SerialPortHelper();
            serialPortHelper.DelegateReceive = ReceiveData;

            #region SerialPort Param
            string[] strPortNuber = serialPortHelper.GetPortNumber();
            for (int i = 0; i < strPortNuber.Length; i++)
            {
                cbSerial.Items.Add(strPortNuber[i]);
            }
            if (cbSerial.Items.Count != 0)
            {
                cbSerial.SelectedIndex = 0;
            }

            cbBaudRate.Text = "19200";
            #endregion

            btnInit232.BackColor = Color.Green;
            btnInit232.Enabled = false;

            fileName232 = DateTime.Now.ToString("RS232_" + "yy-MM-dd_HH-mm-ss") + ".csv";
        }

        #endregion

        #region Buttons TCP/IP
        

        #endregion

        #region Them
        void LoadDataGridView1()
        {
            // dataTable.Columns.Add("Stt", typeof(int));
        
                dataTable.Columns.Add("NO.", typeof(string));
                dataTable.Columns.Add("CODE 1", typeof(string));
                dataTable.Columns.Add("CODE 2", typeof(string));
                dataGridView1.DataSource = dataTable;
            
           
        }

        void LoadDataGridView2()
        {
            // dataTable.Columns.Add("Stt", typeof(int));
            
         
            
                dataTable_Type2.Columns.Add("NO.", typeof(string));
                dataTable_Type2.Columns.Add("CODE 1", typeof(string));
                dataTable_Type2.Columns.Add("CODE 2", typeof(string));
                dataTable_Type2.Columns.Add("CODE 3", typeof(string));
                dataGridView2.DataSource = dataTable_Type2;
            
        }
        private void tbSavePath_TextChanged(object sender, EventArgs e)
        {

        }
        private DataTable dataTable = new DataTable();
        private DataTable dataTable_Type2= new DataTable();
       // private object openFileDialog1;

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

      
        private void button2_Click(object sender, EventArgs e)
        {
            string[] barcode;
            barcode = tbStatus232.Text.Split('\r');
            for (int i = 0; i < barcode.Length; i++)
            {
                 //   MessageBox.Show("Phan tu thu " + i + " =" + barcode[i]);
                //dataTable.Rows.Add(barcode[i]);
                //  dataGridView1.DataSource = dataTable;
                
            }
            List<Madoccode> list = new List<Madoccode>()
            {
                new Madoccode{ MACODE_ME = barcode[0]},
                new Madoccode{ MACODE_ME = barcode[1]},
                new Madoccode{ MACODE_ME = barcode[2]},
                new Madoccode{ MACODE_ME = barcode[3]},
                new Madoccode{ MACODE_ME = barcode[4]},
                new Madoccode{ MACODE_ME = barcode[5]},
                new Madoccode{ MACODE_ME = barcode[6]},
                new Madoccode{ MACODE_ME = barcode[7]},
                new Madoccode{ MACODE_ME = barcode[8]},
            };

            // khởi tạo wb rỗng
            XSSFWorkbook wb = new XSSFWorkbook();

            // Tạo ra 1 sheet
            ISheet sheet = wb.CreateSheet();
            // Tạo row
            var row0 = sheet.CreateRow(0);
            // Merge lại row đầu 3 cột
            row0.CreateCell(0); // tạo ra cell trc khi merge
            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 2);
            sheet.AddMergedRegion(cellMerge);
            row0.GetCell(0).SetCellValue("Đầu đọc code COGNEX");
            // Ghi tên cột ở row 1
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("Mã code");
            // bắt đầu duyệt mảng và ghi tiếp tục
            int rowIndex = 2;
            foreach (var item in list)
            {
                // tao row mới
                var newRow = sheet.CreateRow(rowIndex);

                // set giá trị
                newRow.CreateCell(0).SetCellValue(item.MACODE_ME);
               
                // tăng index
                rowIndex++;
            }

            // xong hết thì save file lại
            FileStream fs = new FileStream(@"D:\Millenium3.xlsx", FileMode.Create);
            wb.Write(fs);
            fs.Close();
        }
        #endregion
        #region Export Excel
        private void ExportExcel(string path)
        {
            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    application.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            application.Columns.AutoFit();
            application.ActiveWorkbook.SaveCopyAs(path);
            application.ActiveWorkbook.Saved = true;
        }
        #endregion

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void btn_Export_excel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 2003 (*.xls)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("Xuất file thành công");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xuất file không thành công\n" + ex.Message);
                }
            }
        }

       

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panColor_Paint(object sender, PaintEventArgs e)
        {

        }

        private void cbSerial_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void tab232_Click(object sender, EventArgs e)
        {

        }
      
        

       

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            if (type == 1)
            {
                dataTable.Columns.Clear();
                dataTable.Clear();
                dataGridView1.Enabled = true;
                dataGridView1.Controls.Clear();
                j = 0;
                con_cu = null;
                tbStatus232.Enabled = true;
                tbStatus232.Clear();
                //LoadDataGridView1();
                // tbStatus232.Enabled = true;
                // tbStatus232.Controls.Clear();
            }
            if (type == 2)
            {
                dataTable_Type2.Columns.Clear();
                dataTable_Type2.Clear();
                dataGridView2.Enabled = true;
                dataGridView2.Controls.Clear();
                k = 0;
                con_cu = null;
                tbStatus232.Enabled = true;
                tbStatus232.Clear();
            }    
            

        }

        private void btn_Type1_Click(object sender, EventArgs e)
        {
           type = 1;
           btn_Type1.BackColor = Color.Green;
           btn_Type2.BackColor= Color.Beige;
          // dataTable_Type2.Columns.Clear();
         //  dataTable.Clear();
         //  dataGridView1.Enabled = true;
          // dataGridView1.Controls.Clear();
         //  j = 0;
           con_cu = null;
           tbStatus232.Enabled = true;
           tbStatus232.Clear();
           dataGridView2.Hide();
           dataGridView1.Show();
            // LoadDataGridView1();
        }

        private void btn_Type2_Click(object sender, EventArgs e)
        {
            type = 2;
            btn_Type2.BackColor = Color.Green;
            btn_Type1.BackColor = Color.Beige;
           // dataTable.Columns.Clear();
           // dataTable.Clear();
           // dataGridView1.Enabled = true;
          //  dataGridView1.Controls.Clear();
          //  j = 0;
            con_cu = null;
            tbStatus232.Enabled = true;
            tbStatus232.Clear();
            dataGridView1.Hide();
            dataGridView2.Show();
            
            //LoadDataGridView2();
        }
    }
}
