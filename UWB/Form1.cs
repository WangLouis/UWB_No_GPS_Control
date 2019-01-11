using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


using System.Collections;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;

namespace hello
{
    public partial class Form1 : Form
    {

        System.IO.Ports.SerialPort serialport = new System.IO.Ports.SerialPort();

        int RX_Counter = 0;
        int Log_Counter = 0;
        bool send_flag = false;
        bool log_flag = false;
        bool log_flag_set = false;
        byte[] Response = new byte[1024];
        byte[] TX_Data = new byte[10];

        public const int I_UWB_LPS_TAG_DATAFRAME0_LENGTH = 128;

        private Boolean receiving;
        private SerialPort comport;
        private Int32 totalLength = 0;
        private Thread t;
        delegate void Display(Byte[] buffer);
        double Anchor_x, Anchor_y, Anchor_z;

        static DateTime gps_epoch = new DateTime(1980, 1, 6, 0, 0, 0, DateTimeKind.Utc);
        static DateTime unix_epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        static DateTime my_start = DateTime.UtcNow;
        //static uint GPS_LEAPSECONDS_MILLIS = 18000;

        string pathFile = @"C:\UWBtest\UWBtest2.xlsx";
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;

        
        public Form1()
        {
            InitializeComponent();    
            foreach (string com in System.IO.Ports.SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(com);
            }

            excelApp = new Excel.Application();    // 開啟一個新的應用程式
            //excelApp.Visible = true;             // 讓Excel文件可見
            excelApp.DisplayAlerts = false;        // 停用警告訊息
            excelApp.Workbooks.Add(Type.Missing);  // 加入新的活頁簿
            wBook = excelApp.Workbooks[1];         // 引用第一個活頁簿
            wBook.Activate();                      // 設定活頁簿焦點
        }

        public class DroneData
        {
            public IPEndPoint ep;
            public float lastPosX = 0;
            public float lastPosY = 0;
            public float lastPosZ = 0;
            public DateTime lastTime = DateTime.MaxValue;
            public int lost_count = 0;
            public int gps_skip_count = 0;
            public DroneData(string ip, int port)
            {
                ep = new IPEndPoint(IPAddress.Parse(ip), port);
            }
        }

        private static MAVLink.MavlinkParse mavlinkParse = new MAVLink.MavlinkParse();
        private static Socket mavSock = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
        private static Dictionary<string, DroneData> drones = new Dictionary<string, DroneData>(5);
        private static Stopwatch stopwatch;

        public void Main()
        {
            stopwatch = new Stopwatch();
            stopwatch.Start();

            drones.Add("bebop2", new DroneData("192.168.42.1", 20000));
            
            MAVLink.mavlink_system_time_t cmd = new MAVLink.mavlink_system_time_t();
            cmd.time_boot_ms = 0;
            cmd.time_unix_usec = (ulong)((DateTime.UtcNow - unix_epoch).TotalMilliseconds * 1000);
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SYSTEM_TIME, cmd);
            foreach (KeyValuePair<string, DroneData> drone in drones)
            {
                mavSock.SendTo(pkt, drone.Value.ep);
            }
            send_flag = true;

        }
        /*
        public void processFrameData()
        {
            MAVLink.mavlink_att_pos_mocap_t att_pos = new MAVLink.mavlink_att_pos_mocap_t();
            att_pos.time_usec = (ulong)(((DateTime.UtcNow - unix_epoch).TotalMilliseconds - 10) * 1000);
            att_pos.x = Anchor_y; //north Anchor_y
            att_pos.y = Anchor_x; //east Anchor_x
            att_pos.z = Anchor_z; //down
                                  //att_pos.q = new float[4] { rbData.qw, rbData.qx, rbData.qz, -rbData.qy };

            DroneData drone = drones;
            drone.lost_count = 0;
            byte[] pkt;
            pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.ATT_POS_MOCAP, att_pos);
            mavSock.SendTo(pkt, drones.ep);

        }*/
        

        private void btnDate_Click(object sender, EventArgs e)
        {
            /*
            int xx = 1, yy = 2;
  
            label14.Text = Convert.ToString(xx * Math.Cos(radian) + yy * Math.Sin(radian));
            label36.Text = Convert.ToString((-xx * Math.Sin(radian)) + yy * Math.Cos(radian));
            */
            send_flag = false;
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void DisplayText(Byte[] buffer)
        {
            //textBox1.Text += String.Format("{0}{1}", BitConverter.ToString(buffer), Environment.NewLine);
            label9.Text = buffer[0].ToString("X2");
            label12.Text = Convert.ToString(buffer.Length);
            totalLength = totalLength + buffer.Length;
            label10.Text = totalLength.ToString();
            /*for (int i = 0; i < 896; i++)
            {
                Response[RX_Counter] = buffer[i];
                if (RX_Counter < 896) RX_Counter++;
                else RX_Counter = 0;
            }*/
            check_uwb(0x55, 0x00, 128);
        }

        private void DoReceive()
        {
            /*
            Byte[] buffer = new Byte[1024];
            while (receiving)
            {
                if (comport.BytesToRead > 0)
                {
                    Int32 length = comport.Read(buffer, 0, buffer.Length);
                    Array.Resize(ref buffer, length);
                    Display d = new Display(DisplayText);
                    this.Invoke(d, new Object[] { buffer });
                    Array.Resize(ref buffer, 1024);
                }
                Thread.Sleep(16);
            }*/

            Boolean readingFromBuffer;
            Int32 count = 0;
            Byte[] buffer = new Byte[896];
            while (receiving)
            {
                readingFromBuffer = true;
                while (comport.BytesToRead < buffer.Length && count < 501)
                {
                    Thread.Sleep(16);
                    count++;
                    if (count > 500) //|| (buffer[0] != 0x55)
                    {
                        readingFromBuffer = false;
                    }
                }
                count = 0;

                if (readingFromBuffer)
                {
                    //Int32 length = comport.Read(buffer, 0, buffer.Length);
                    comport.Read(buffer, 0, buffer.Length);
                    Display d = new Display(DisplayText);
                    this.Invoke(d, new Object[] { buffer });
              
                }
                else
                {
                    comport.DiscardInBuffer();
                }

                if (buffer[0] != 0x55 && buffer[895] != 0xee)
                {
                    comport.DiscardInBuffer();
                }
                else
                {
                    for (int i = 0; i < buffer.Length; i++)
                    {
                        Response[RX_Counter] = buffer[i];
                        if (RX_Counter < 896) RX_Counter++;
                        else RX_Counter = 0;
                    }
                }
                //Thread.Sleep(16);
            }

        }
        float theta = 35 + 90;//角度值 4f->64
        
        private void check_uwb(byte Response1, byte Response2, int Check)//02,10,F3,02,4F,4B,E5,
        {
            RX_Counter = 0;
            theta = Convert.ToInt32(textBox2.Text) + 90;
            label32.Text = Convert.ToString(theta);
            double radian = theta * Math.PI / 180;//轉換弧度值
            //Thread.Sleep(1300);
            //label10.Text = Response[0].ToString("X2") + " " + Response[1].ToString("X2") + " ";
            if (Response[0] == Response1 && Response[1] == Response2)
            {
                label9.Text = "UWB OK";
                /*
                label10.Text = Response[0].ToString("X2") + " " + Response[24].ToString("X2") + " " + Response[25].ToString("X2") + " " + Response[26].ToString("X2") + " ";
                float Tag_dis1 = ((Response[26] << 16) | (Response[25] << 8) | (Response[24] << 0)) / 256000.0f;
                label12.Text = Convert.ToString(Tag_dis1);*/

                Anchor_x = ((Response[6] << 24) | (Response[5] << 16) | (Response[4] << 8) | (0x00 << 0)) / 256000.0f;
                label26.Text = Convert.ToString(Anchor_x);

                Anchor_y = ((Response[9] << 24) | (Response[8] << 16) | (Response[7] << 8) | (0x00 << 0)) / 256000.0f;
                label27.Text = Convert.ToString(Anchor_y);

                Anchor_z = ((Response[12] << 24) | (Response[11] << 16) | (Response[10] << 8) | (0x00 << 0)) / 256000.0f;
                //Anchor_z = Anchor_z + 0.3;
                label28.Text = Convert.ToString(Anchor_z);
                if (Anchor_z < 1)
                    Anchor_z = Anchor_z + 0.2;  //4f-> 0.3
                //processFrameData();
                if (send_flag)
                {
                    long cur_ms = stopwatch.ElapsedMilliseconds;
                    MAVLink.mavlink_att_pos_mocap_t att_pos = new MAVLink.mavlink_att_pos_mocap_t();
                    //att_pos.time_usec = (ulong)(((DateTime.UtcNow - unix_epoch).TotalMilliseconds - 10) * 1000);
                    att_pos.time_usec = (ulong)(cur_ms * 1000);
                    //att_pos.x = Anchor_y; //north Anchor_y
                    //att_pos.y = Anchor_x; //east Anchor_x
                    att_pos.x = (float)(Anchor_x * Math.Cos(radian) + Anchor_y * Math.Sin(radian)); //north
                    att_pos.y = (float)-((-Anchor_x * Math.Sin(radian)) + Anchor_y * Math.Cos(radian)); //east
                    att_pos.z = (float)-Anchor_z; //down
                    Anchor_x = att_pos.x;
                    Anchor_y = att_pos.y;
                    Anchor_z = att_pos.z;
                    label7.Text = Convert.ToString(att_pos.x);
                    label13.Text = Convert.ToString(att_pos.y);
                    label22.Text = Convert.ToString(att_pos.z);

                    //att_pos.q = new float[4] { rbData.qw, rbData.qx, rbData.qz, -rbData.qy };
                    DroneData drone = drones["bebop2"];
                    drone.lost_count = 0;

                    byte[] pkt;
                    pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.ATT_POS_MOCAP, att_pos);
                    mavSock.SendTo(pkt, drone.ep);
                }
                // for (int i = 0; i < 23; i++ )
                //     richTextBox1.Text += Response[i].ToString("X2") + " "; 

                if (log_flag)
                {
                    Log_Counter = Log_Counter+1;
                    try
                    {
                        wSheet = (Excel._Worksheet)wBook.Worksheets[1];   // 引用第一個工作表
                        wSheet.Name = "UWB Sensor Value Log";   // 命名工作表的名稱
                        wSheet.Activate();  // 設定工作表焦點  

                        //excelApp.Cells[1, 1] = "Excel測試";

                        excelApp.Cells[Log_Counter, 1] = Anchor_x; //att_pos.x
                        excelApp.Cells[Log_Counter, 2] = Anchor_y; //att_pos.y
                        excelApp.Cells[Log_Counter, 3] = Anchor_z; //att_pos.z
  
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
                    }

                    if (log_flag_set)
                    {
                        try
                        {
                            //另存活頁簿
                            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                        }

                        wBook.Close(false, Type.Missing, Type.Missing);   //關閉活頁簿
                        excelApp.Quit();  //關閉Excel
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);  //釋放Excel資源
                        wBook = null;
                        wSheet = null;
                        excelApp = null;
                        GC.Collect();
                        Console.Read();
                        log_flag = false;
                        log_flag_set = false;
                        Log_Counter = 0;
                    }
                }
            }    
        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            //設定連接埠為9600、n、8、1、n
            /*
            serialport.PortName = comboBox1.Text;
            serialport.BaudRate = 460800;
            serialport.DataBits = 8;                   
            serialport.StopBits = System.IO.Ports.StopBits.One;
            serialport.Parity = System.IO.Ports.Parity.None;
            serialport.Handshake = System.IO.Ports.Handshake.None;
            serialport.Encoding = Encoding.Default;//傳輸編碼方式
            serialport.DataReceived += new SerialDataReceivedEventHandler(ReceiveMessage);

            try
            {
                serialport.Open();
                label21.Text = "connect ok";
                textBox1.Text = "Initial text contents of the TextBox.";
            }
            catch (UnauthorizedAccessException uae)
            {
                serialport.Close();
                serialport.Dispose();
                label21.Text = "connect fall";
            }
            //button2.Text = "連線";
            */

            comport = new SerialPort(comboBox1.Text, 460800, Parity.None, 8, StopBits.One);
            if (!comport.IsOpen)
            {
                comport.Open();
                receiving = true;
                t = new Thread(DoReceive);
                t.IsBackground = true;
                t.Start();
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void button5_Click(object sender, EventArgs e)
        {
            log_flag = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            log_flag_set = true;
        }


        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        //DroneData drone = drones["bebop2"];

        private void button4_Click_1(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 9;
            cmd.target_system = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);

        }

        private void button7_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 3;
            cmd.target_system = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_command_long_t cmd = new MAVLink.mavlink_command_long_t();
            cmd.command = (ushort)MAVLink.MAV_CMD.COMPONENT_ARM_DISARM;
            cmd.target_system = 1;
            cmd.param1 = 0;
            cmd.param2 = 21196;
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.COMMAND_LONG, cmd);
            DroneData drone = drones["bebop2"];
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_command_long_t cmd = new MAVLink.mavlink_command_long_t();
            cmd.command = (ushort)MAVLink.MAV_CMD.TAKEOFF;
            cmd.target_system = 1;
            //cmd.target_component = 250;
            cmd.param7 = 1.2f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.COMMAND_LONG, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_position_target_local_ned_t cmd = new MAVLink.mavlink_set_position_target_local_ned_t();
            cmd.target_system = 1;
            cmd.coordinate_frame = (byte)MAVLink.MAV_FRAME.LOCAL_NED;
            cmd.type_mask = 0xff8;
            cmd.x = 0.0f + float.Parse(textBox3.Text);
            label33.Text = Convert.ToString(cmd.x);
            cmd.y = 0.0f + float.Parse(textBox4.Text);
            label34.Text = Convert.ToString(cmd.y);
            cmd.z = -1.2f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_POSITION_TARGET_LOCAL_NED, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 4;
            cmd.target_system = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_gps_global_origin_t cmd = new MAVLink.mavlink_set_gps_global_origin_t();
            cmd.target_system = 1;
            cmd.latitude = (int)(24.7733321 * 10000000);
            cmd.longitude = (int)(121.0449535 * 10000000);
            cmd.altitude = 100;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_GPS_GLOBAL_ORIGIN, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button6_Click(object sender, EventArgs e)
        {
                Main();
                lblShow.Text = "System start";
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }


    }

    
}
