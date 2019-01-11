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


namespace hello
{
    public partial class Form1 : Form
    {

        System.IO.Ports.SerialPort serialport = new System.IO.Ports.SerialPort();
        byte[] CMD_SET_Channel = new byte[9];
        byte[] Assoc_List = new byte[5];
        byte[] Short_Address = new byte[12];
        byte[] MSG_Send = new byte[9];
        byte[] Temperature_Read_Value = new byte[7];
        byte[] ADC_Read_Value = new byte[9];

        byte FCS;
        int TX_Counter, RX_Counter,kkk=0;
        byte[] Response = new byte[100];
        byte[] TX_Data = new byte[10];
        float[] CO1 = new float[202];
        float[] CO2 = new float[202];
        float[] CO3 = new float[202];
        float[] COtow1 = new float[1003];
        float[] COtow2 = new float[1003];
        float[] COtow3 = new float[1003];
        float[] FT1 = new float[10];
        float[] FT2 = new float[10];
        float[] FT3 = new float[10];
        float[] Output1 = new float[102];
        float[] Output2 = new float[102];
        float[] Output3 = new float[102];
        float[] Output4 = new float[102];
        float[] Output5 = new float[102];
        float[] Output = new float[102];
        float[] tem4 = new float[100];


        public Form1()
        {
            InitializeComponent();

            foreach (string com in System.IO.Ports.SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(com);
            }
        }

        private float fun1(float i, float a, float b)
        {
           float t;
           if(i<=a)
              t=1;
           else if(i>a&&i<=b)
               t=(b-i)/(b-a);
           else
           t=0;
        return t;
        }

        private float tri(float i, float a, float b, float c, float d)
        {
            float t;
            if (i < a)
                t = 0;
            else if (a <= i && i <= b)
                t = (i - a) / (b - a);
            else if (b < i && i <= c)
                t = 1;
            else if (c < i && i <= d)
                t = (d - i) / (d - c);
            else
                t = 0;
            return t;
        }

        private float fun2(float i, float a, float b)
        {
            float t;
            if (i <= a)
                t = 0;
            else if (i > a && i <= b)
                t = (i - a) / (b - a);
            else
                t = 1;
            return t;
        }

        private float risk_level(int CO, int CO22, int T)
        {


            float tow1 = 0, tow2 = 0, tow3 = 0, tow4 = 0, tow5 = 0, tow6 = 0, tow7 = 0, tow8 = 0, tow9 = 0, tow10 = 0, tow11 = 0, tow12 = 0, tow13 = 0, tow14 = 0, tow15 = 0, tow16 = 0;
            float tow17 = 0, tow18 = 0, tow19 = 0, tow20 = 0, tow21 = 0, tow22 = 0, tow23 = 0, tow24 = 0, tow25 = 0, tow26 = 0, tow27 = 0;
            float a = 0, b = 0, y = 0;

            if (CO < 100 && CO22 < 700 && T < 4)                           //rule1
            {
                tow1 = Math.Min(CO1[CO], COtow1[CO22]);
                tow1 = Math.Min(tow1, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output1[x] > tow1)
                    {
                        Output1[x] = tow1;
                    }
                }
            }

            if (CO < 100 && CO22 < 700 && T > 2 && T < 6)                  //rule2
            {
                tow2 = Math.Min(CO1[CO], COtow1[CO22]);
                tow2 = Math.Min(tow2, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output2[x] > tow2)
                    {
                        Output2[x] = tow2;
                    }
                }
            }

            if (CO < 100 && CO22 > 550 && CO22 < 800 && T < 4)             //rule3
            {
                tow3 = Math.Min(CO1[CO], COtow2[CO22]);
                tow3 = Math.Min(tow3, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output2[x] > tow3)
                    {
                        Output2[x] = tow3;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 < 700 && T < 4)               //rule4
            {
                tow4 = Math.Min(CO2[CO], COtow1[CO22]);
                tow4 = Math.Min(tow4, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output2[x] > tow4)
                    {
                        Output2[x] = tow4;
                    }
                }
            }

            if (CO < 100 && CO22 > 550 && CO22 < 800 && T > 2 && T < 6)       //rule5
            {
                tow5 = Math.Min(CO1[CO], COtow2[CO22]);
                tow5 = Math.Min(tow5, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output3[x] > tow5)
                    {
                        Output3[x] = tow5;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 < 700 && T > 2 && T < 6)            //rule6
            {
                tow6 = Math.Min(CO2[CO], COtow1[CO22]);
                tow6 = Math.Min(tow6, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output3[x] > tow6)
                    {
                        Output3[x] = tow6;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 550 && CO22 < 800 && T < 4)        //rule7
            {
                tow7 = Math.Min(CO2[CO], COtow2[CO22]);
                tow7 = Math.Min(tow7, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output3[x] > tow7)
                    {
                        Output3[x] = tow7;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 550 && CO22 < 800 && T > 2 && T < 6)     //rule8
            {
                tow8 = Math.Min(CO2[CO], COtow2[CO22]);
                tow8 = Math.Min(tow8, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output3[x] > tow8)
                    {
                        Output3[x] = tow8;
                    }
                }
            }

            if (CO < 100 && CO22 < 700 && T > 4)                                //rule9
            {
                tow9 = Math.Min(CO1[CO], COtow1[CO22]);
                tow9 = Math.Min(tow9, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow9)
                    {
                        Output4[x] = tow9;
                    }
                }
            }

            if (CO < 100 && CO22 > 550 && CO22 < 800 && T > 4)                    //rule10
            {
                tow10 = Math.Min(CO1[CO], COtow2[CO22]);
                tow10 = Math.Min(tow10, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow10)
                    {
                        Output4[x] = tow10;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 < 700 && T > 4)                     //rule11
            {
                tow11 = Math.Min(CO2[CO], COtow1[CO22]);
                tow11 = Math.Min(tow11, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow11)
                    {
                        Output4[x] = tow11;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 550 && CO22 < 800 && T > 4)          //rule12
            {
                tow12 = Math.Min(CO2[CO], COtow2[CO22]);
                tow12 = Math.Min(tow12, FT3[T]);

                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow12)
                    {
                        Output4[x] = tow12;
                    }
                }
            }

            if (CO < 100 && CO22 > 700 && T < 4)                             //rule13
            {
                tow13 = Math.Min(CO1[CO], COtow3[CO22]);
                tow13 = Math.Min(tow13, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow13)
                    {
                        Output4[x] = tow13;
                    }
                }
            }

            if (CO < 100 && CO22 > 700 && T > 2 && T < 6)                     //rule14
            {
                tow14 = Math.Min(CO1[CO], COtow3[CO22]);
                tow14 = Math.Min(tow14, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow14)
                    {
                        Output4[x] = tow14;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 700 && T < 4)                  //rule15
            {
                tow15 = Math.Min(CO2[CO], COtow3[CO22]);
                tow15 = Math.Min(tow15, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow15)
                    {
                        Output4[x] = tow15;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 700 && T > 2 && T < 6)           //rule16
            {
                tow16 = Math.Min(CO2[CO], COtow3[CO22]);
                tow16 = Math.Min(tow16, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow16)
                    {
                        Output4[x] = tow16;
                    }
                }
            }

            if (CO > 100 && CO22 < 700 && T < 4)                            //rule17
            {
                tow17 = Math.Min(CO3[CO], COtow1[CO22]);
                tow17 = Math.Min(tow17, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow17)
                    {
                        Output4[x] = tow17;
                    }
                }
            }

            if (CO > 100 && CO22 < 700 && T > 2 && T < 6)                    //rule18
            {
                tow18 = Math.Min(CO3[CO], COtow1[CO22]);
                tow18 = Math.Min(tow18, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow18)
                    {
                        Output4[x] = tow18;
                    }
                }
            }

            if (CO > 100 && CO22 > 550 && CO22 < 800 && T < 4)              //rule19
            {
                tow19 = Math.Min(CO3[CO], COtow2[CO22]);
                tow19 = Math.Min(tow19, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow19)
                    {
                        Output4[x] = tow19;
                    }
                }
            }

            if (CO > 100 && CO22 > 550 && CO22 < 800 && T > 2 && T < 6)       //rule20
            {
                tow20 = Math.Min(CO3[CO], COtow2[CO22]);
                tow20 = Math.Min(tow20, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output4[x] > tow20)
                    {
                        Output4[x] = tow20;
                    }
                }
            }

            if (CO < 100 && CO22 > 700 && T > 4)                           //rule21
            {
                tow21 = Math.Min(CO1[CO], COtow3[CO22]);
                tow21 = Math.Min(tow21, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow21)
                    {
                        Output5[x] = tow21;
                    }
                }
            }

            if (CO > 35 && CO < 150 && CO22 > 700 && T > 4)               //rule22
            {
                tow22 = Math.Min(CO2[CO], COtow3[CO22]);
                tow22 = Math.Min(tow22, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow22)
                    {
                        Output5[x] = tow22;
                    }
                }
            }

            if (CO > 100 && CO22 > 700 && T < 4)                       //rule23
            {
                tow23 = Math.Min(CO3[CO], COtow3[CO22]);
                tow23 = Math.Min(tow23, FT1[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow23)
                    {
                        Output5[x] = tow23;
                    }
                }
            }

            if (CO > 100 && CO22 > 700 && T > 2 && T < 6)                //rule24
            {
                tow24 = Math.Min(CO3[CO], COtow3[CO22]);
                tow24 = Math.Min(tow24, FT2[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow24)
                    {
                        Output5[x] = tow24;
                    }
                }
            }

            if (CO > 100 && CO22 < 700 && T > 4)                     //rule25
            {
                tow25 = Math.Min(CO3[CO], COtow1[CO22]);
                tow25 = Math.Min(tow25, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow25)
                    {
                        Output5[x] = tow25;
                    }
                }
            }

            if (CO > 100 && CO22 > 550 && CO22 < 800 && T > 4)          //rule26
            {
                tow26 = Math.Min(CO3[CO], COtow2[CO22]);
                tow26 = Math.Min(tow26, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow26)
                    {
                        Output5[x] = tow26;
                    }
                }
            }

            if (CO > 100 && CO22 > 700 && T > 4)                     //rule27
            {
                tow27 = Math.Min(CO3[CO], COtow3[CO22]);
                tow27 = Math.Min(tow27, FT3[T]);
                for (int x = 0; x <= 100; x++)
                {
                    if (Output5[x] > tow27)
                    {
                        Output5[x] = tow27;
                    }
                }
            }

            float k1 = 0, k2 = 0, k3 = 0, k4 = 0, k5 = 0;
            k1 = tow1;
            k2 = tow2 + tow3 + tow4;
            k3 = tow5 + tow6 + tow7 + tow8;
            k4 = tow9 + tow10 + tow11 + tow12 + tow13 + tow14 + tow15 + tow16 + tow17 + tow18 + tow19 + tow20;
            k5 = tow21 + tow22 + tow23 + tow24 + tow25 + tow26 + tow27;

            for (int x = 0; x <= 100; x++)                           //Output
            {
                if (k1 > 0 && k2 == 0)
                    Output[x] = Output1[x];
                else if (k1 > 0 && k2 > 0 && k3 == 0)
                    Output[x] = Math.Max(Output1[x], Output2[x]);
                else if (k1 > 0 && k2 > 0 && k3 > 0 && k4 == 0)
                {
                    Output[x] = Math.Max(Output1[x], Output2[x]);
                    Output[x] = Math.Max(Output[x],Output3[x]);
                }
                else if (k1 > 0 && k2 > 0 && k3 > 0 && k4 > 0 && k5 == 0)
                {
                    Output[x] = Math.Max(Output1[x], Output2[x]);
                    Output[x] = Math.Max(Output[x], Output3[x]);
                    Output[x] = Math.Max(Output[x], Output4[x]);
                }
                else if (k1 > 0 && k2 > 0 && k3 > 0 && k4 > 0 && k5 > 0)
                {
                    Output[x] = Math.Max(Output1[x], Output2[x]);
                    Output[x] = Math.Max(Output[x], Output3[x]);
                    Output[x] = Math.Max(Output[x], Output4[x]);
                    Output[x] = Math.Max(Output[x], Output5[x]);

                }
                else if (k1 == 0 && k2 > 0 && k3 == 0)
                    Output[x] = Output2[x];
                else if (k1 == 0 && k2 > 0 && k3 > 0 && k4 == 0)
                    Output[x] = Math.Max(Output2[x], Output3[x]);
                else if (k1 == 0 && k2 > 0 && k3 > 0 && k4 > 0 && k5 == 0)
                {
                    Output[x] = Math.Max(Output2[x], Output3[x]);
                    Output[x] = Math.Max(Output[x], Output4[x]);
                }
                else if (k1 == 0 && k2 > 0 && k3 > 0 && k4 > 0 && k5 > 0)
                {
                    Output[x] = Math.Max(Output2[x], Output3[x]);
                    Output[x] = Math.Max(Output[x], Output4[x]);
                    Output[x] = Math.Max(Output[x], Output5[x]);
                }
                else if (k2 == 0 && k3 > 0 && k4 == 0)
                    Output[x] = Output3[x];
                else if (k2 == 0 && k3 > 0 && k4 > 0 && k5 == 0)
                    Output[x] = Math.Max(Output3[x], Output4[x]);
                else if (k2 == 0 && k3 > 0 && k4 > 0 && k5 > 0)
                {
                    Output[x] = Math.Max(Output3[x], Output4[x]);
                    Output[x] = Math.Max(Output[x], Output5[x]);
                }
                else if (k3 == 0 && k4 > 0 && k5 == 0)
                    Output[x] = Output4[x];
                else if (k3 == 0 && k4 > 0 && k5 > 0)
                    Output[x] = Math.Max(Output4[x], Output5[x]);
                else if (k4 == 0 && k5 > 0)
                    Output[x] = Output5[x];
            }

            for (int x = 0; x <= 100; x++)                           //Output
            {
                b = b + Output[x];
            }
            for (int x = 0; x <= 100; x++)                           //Output
            {
                a = a + (x * Output[x]);
            }
            y = a / b;
            //label26.Text = CO1[CO] + "";
            //label27.Text = CO2[CO] + "";
            //label27.Text = CO3[CO] + "";
            label26.Text = Convert.ToString(Math.Round(Convert.ToDouble(CO1[CO]), 2, MidpointRounding.AwayFromZero));
            label27.Text = Convert.ToString(Math.Round(Convert.ToDouble(CO2[CO]), 2, MidpointRounding.AwayFromZero));
            label28.Text = Convert.ToString(Math.Round(Convert.ToDouble(CO3[CO]), 2, MidpointRounding.AwayFromZero));

            //label29.Text = COtow1[CO22] + "";
            //label30.Text = COtow2[CO22] + "";
            //label31.Text = COtow3[CO22] + "";
            label29.Text = Convert.ToString(Math.Round(Convert.ToDouble(COtow1[CO22]), 2, MidpointRounding.AwayFromZero));
            label30.Text = Convert.ToString(Math.Round(Convert.ToDouble(COtow2[CO22]), 2, MidpointRounding.AwayFromZero));
            label31.Text = Convert.ToString(Math.Round(Convert.ToDouble(COtow3[CO22]), 2, MidpointRounding.AwayFromZero));

            //label32.Text = FT1[T] + "";
            //label33.Text = FT2[T] + "";
            //label34.Text = FT3[T] + "";
            label32.Text = Convert.ToString(Math.Round(Convert.ToDouble(FT1[T]), 2, MidpointRounding.AwayFromZero));
            label33.Text = Convert.ToString(Math.Round(Convert.ToDouble(FT2[T]), 2, MidpointRounding.AwayFromZero));
            label34.Text = Convert.ToString(Math.Round(Convert.ToDouble(FT3[T]), 2, MidpointRounding.AwayFromZero));

            //if (CO2[CO] > 0.5 || CO3[CO] > 0 || COtow1[CO22] > 0.5 || COtow2[CO22] > 0 && FT2[T] < 0.5 && FT3[T] == 0)
            //{
            //label38.Text = "請開啟窗戶，保持通風";
            //}
            //else if (CO2[CO] < 0.5 && CO3[CO] == 0 && COtow1[CO22] < 0.5 && COtow2[CO22] ==0 && FT2[T] > 0.5 || FT3[T] > 0.1)
            //{
            //label38.Text = "注意溫度!";
            //}
            // else if (CO2[CO] > 0.5 && CO3[CO] > 0.1 && COtow1[CO22] > 0.5 && COtow2[CO22] > 0.1 && FT2[T] > 0.5 && FT3[T] > 0.1)
            //{
            //label38.Text = "請開啟窗戶，保持通風，並注意溫度!";
            //}
            //else
            //{
            // label38.Text = "無";
            //}


            return y;
        }



        private double Grey(int T1, int T2, int T3, int T4)
        {

            int i,j,k,p,s;
            double[] Tem = new double[100];
            double[,] YN = new double[5,5];
            double[,] B = new double[10,10];
            double[,] BT = new double[10,10];
            double[,] BTB = new double[10, 10];
            double[,] invBTB = new double[10, 10];
            double[,] invBTBBT = new double[10, 10];
            double[,] ahat = new double[10, 10];
            double a, b, forecastT;


            Tem[1]=T1;
            Tem[2]=T2;
            Tem[3]=T3;
            Tem[4]=T4;


            for(k=1;k<4;k++)   
            {                
            B[k,2]=1;
            }
            B[1,1]=-1*(Tem[1]+Tem[2]/2);
            B[2,1]=-1*(Tem[1]+Tem[2]+Tem[3]/2);
            B[3,1]=-1*(Tem[1]+Tem[2]+Tem[3]+Tem[4]/2);
            for(s=1;s<4;s++)    
            {
            BT[2,s] = 1;
            }
            BT[1,1] = B[1,1];
            BT[1,2] = B[2,1];
            BT[1,3] = B[3,1];
            for (k=1;k<4;k++)  
            {
            YN[k,1] = Tem[k+1];
            }
            BTB[1,1] = BT[1,1] * B[1,1] + BT[1,2] * B[2,1] + BT[1,3] * B[3,1];
            BTB[2,1] = BT[2,1] * B[1,1] + BT[2,2] * B[2,1] + BT[2,3] * B[3,1];
            BTB[1,2] = BT[1,1] * B[1,2] + BT[1,2] * B[2,2] + BT[1,3] * B[3,2];
            BTB[2,2] = BT[2,1] * B[1,2] + BT[2,2] * B[2,2] + BT[2,3] * B[3,2];
            invBTB[1, 1] = BTB[2, 2] / (BTB[1, 1] * BTB[2, 2] - BTB[1, 2] * BTB[2, 1]);
            invBTB[2, 1] = -1 * BTB[1, 2] / (BTB[1, 1] * BTB[2, 2] - BTB[1, 2] * BTB[2, 1]);
            invBTB[1, 2] = -1 * BTB[2, 1] / (BTB[1, 1] * BTB[2, 2] - BTB[1, 2] * BTB[2, 1]);
            invBTB[2, 2] = BTB[1, 1] / (BTB[1, 1] * BTB[2, 2] - BTB[1, 2] * BTB[2, 1]);
            invBTBBT[1,1] = invBTB[1,1] * BT[1,1] + invBTB[1,2] * BT[2,1];
            invBTBBT[1,2] = invBTB[1,1] * BT[1,2] + invBTB[1,2] * BT[2,2];
            invBTBBT[1,3] = invBTB[1,1] * BT[1,3] + invBTB[1,2] * BT[2,3];
            invBTBBT[2,1] = invBTB[2,1] * BT[1,1] + invBTB[2,2] * BT[2,1];
            invBTBBT[2,2] = invBTB[2,1] * BT[1,2] + invBTB[2,2] * BT[2,2];
            invBTBBT[2,3] = invBTB[2,1] * BT[1,3] + invBTB[2,2] * BT[2,3];
            ahat[1,1]=invBTBBT[1,1]*YN[1,1] + invBTBBT[1,2]*YN[2,1] + invBTBBT[1,3]*YN[3,1];
            ahat[2,1]=invBTBBT[2,1]*YN[1,1] + invBTBBT[2,2]*YN[2,1] + invBTBBT[2,3]*YN[3,1];

            a = ahat[1,1];
            b = ahat[2,1];

            Double expa1 = Math.Pow(2.71828, a);
            Double expa2 = Math.Pow(2.71828, a*-4);

            forecastT = (1 - expa1) * (Tem[1] - b / a) * expa2;
            forecastT = (int)forecastT;

            //for (i = 0; i <= Tem.GetUpperBound(0); i++)
            //{
            //    Console.Write("Tem1[{0}]={1} ", i, Tem[i]);
            //}
            

            lblShow.BackColor = Color.Aqua;
            //lblShow.Text="你好!";
            //lblShow.Text = forecastT.ToString();
            label12.Text = forecastT.ToString() + "℃";

            return forecastT;
        }

        private void btnDate_Click(object sender, EventArgs e)
        {
            //Assoc_List[5]={0x02,0x00,0x13,0x00,0x00};		       		   	   搜尋模組
            Assoc_List[0] = 0x02;
            Assoc_List[1] = 0x00;
            Assoc_List[2] = 0x13;
            Assoc_List[3] = 0x00;
            Assoc_List[4] = 0x00;
            

            //讀取Device的Short Address
            TX_Counter = 5;
            RX_Counter = 0;

            for (int x = 0; x < TX_Counter - 1; x++)
            {
                TX_Data[x] = Assoc_List[x];
            }
            while (RX_Counter == 0 || Response[9] != FCS)   //  7     11     13
            {
                zb_data_send();
                check_zigbee(0x10, 0x13, 10);                //  8     12     14
                Thread.Sleep(3000);
            }
                                                            //  1      3      4

            Short_Address[0] = Response[5];          // 二氧化碳感測器
            Short_Address[1] = Response[6];

            Short_Address[2] = Response[7];          // 一氧化碳感測器
            Short_Address[3] = Response[8];

            //Short_Address[4] = Response[9];          // 風扇
            //Short_Address[5] = Response[10];

            //Short_Address[6] = Response[11];         //  氧化碳感測器
            //Short_Address[7] = Response[12];

            //Short_Address[8] = Response[13];         //  氧化碳感測器
            //Short_Address[9] = Response[14];

            lblShow.Text = "node search ok!";


            //lblShow.BackColor = Color.Chocolate;
            //lblShow.Text = DateTime.Now.ToString();

            //MSG_Send[9]={0x02,0x00,0x43,0x04,0x00,0x00,0x01,0x06,0x00};
            

        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void check_zigbee(byte Response1, byte Response2, int Check)//02,10,F3,02,4F,4B,E5,
        {
            RX_Counter = 0;
            Thread.Sleep(1300);
            if (Response[1] == Response1 && Response[2] == Response2 && RX_Counter == Check)
            {
                FCS = Response[1];
                for (int i = 2; i < Check - 1; i++)
                {
                    FCS ^= Response[i];
                }
            }
           // for (int i = 0; i < 23; i++ )
           //     richTextBox1.Text += Response[i].ToString("X2") + " ";     
        }

        private void zb_data_send()
        {
            int i = 2;
            FCS = TX_Data[1];
            for (i = 2; i < TX_Counter - 1; i++)
            {
                FCS ^= TX_Data[i];
            }
            TX_Data[i] = FCS;
            serialport.Write(TX_Data, 0, TX_Counter);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            // CMD_SET_Channel[9]={0x02,0x00,0xF3,0x04,0x00,0x00,0x00,0x00,0x00};

            CMD_SET_Channel[0] = 0x02;
            CMD_SET_Channel[1] = 0x00;
            CMD_SET_Channel[2] = 0xf3;
            CMD_SET_Channel[3] = 0x04;
            CMD_SET_Channel[4] = 0x00;
            CMD_SET_Channel[5] = 0x00;  
            CMD_SET_Channel[6] = 0x80;  
            CMD_SET_Channel[7] = 0x00;
            CMD_SET_Channel[8] = 0x00;

            TX_Counter = 9;

            for (int x = 0; x < TX_Counter - 1; x++)
            {
                TX_Data[x] = CMD_SET_Channel[x];
            }

            while (RX_Counter == 0 || Response[6] != FCS)
            {
                zb_data_send();
                check_zigbee(0x10, 0xF3, 7);
            }
            RX_Counter = 0;
            lblShow.BackColor = Color.Aqua;
            // lblShow.Text = "Channel set ok!";

        }

        private void button2_Click(object sender, EventArgs e)
        {
                    //設定連接埠為9600、n、8、1、n
                    serialport.PortName = comboBox1.Text;
                    serialport.BaudRate = 9600;
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
                    }
                    catch (UnauthorizedAccessException uae)
                    {
                        serialport.Close();
                        serialport.Dispose();
                        label21.Text = "connect fall";
                    }
                    //button2.Text = "連線";
        }

        private void ReceiveMessage(object sender, SerialDataReceivedEventArgs e)
        {
            int bytes = serialport.BytesToRead;

            byte[] buffer = new byte[bytes];
            serialport.Read(buffer, 0, bytes);

            for (int i = 0; i < bytes; i++)
            {
                Response[RX_Counter] = buffer[i];
                RX_Counter++;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private double Read_Temperature()
        {
            //Temperature_Read_Value[7]={0x02,0x00,0x4C,0x02,0x00,0x00};
            int tem, tem2, tem3 = 0, T1 = 0, T2 = 0, T3 = 0, T4 = 0;
            Temperature_Read_Value[0] = 0x02;
            Temperature_Read_Value[1] = 0x00;
            Temperature_Read_Value[2] = 0x4C;  //0x4C
            Temperature_Read_Value[3] = 0x02;
            Temperature_Read_Value[4] = 0x00;
            Temperature_Read_Value[5] = 0x00;
            Temperature_Read_Value[6] = 0x00;
            int countk = 0;  
            RX_Counter = 0;
            TX_Counter = 7;
            for (int x = 0; x < TX_Counter - 1; x++)
            {
                TX_Data[x] = Temperature_Read_Value[x];
            }

                zb_data_send();
                check_zigbee(0x10, 0x4C, 9);

            
            string tem1;
            tem = (int) Response[7];
            tem1 = Convert.ToString(tem, 10);
            tem2 = Convert.ToInt16(tem1);
            label11.Text = Convert.ToString(tem, 10) + "℃";
            if (kkk < 4)
            {
                tem = (int)tem4[kkk];
                kkk++;
                if (kkk == 4)
                    kkk = 1;
            }
            T1 = (int)tem4[1];
            T2 = (int)tem4[2];
            T3 = (int)tem4[3];
            T4 = tem;
            countk = countk + 1 ;
            if (countk == 240)
            {
                countk = 0;
                tem3 = (int)Grey(T1, T2, T3, T4);
            }
   
            return tem3;
        }

         private double Read_CO2()
         {
            //TX_Counter = 9; //7
            //RX_Counter = 0;
            //for (int x = 0; x < TX_Counter - 1; x++)
            //{
            //    TX_Data[x] = 0;
            //}

            TX_Counter = 7; //7
            RX_Counter = 0;
            ADC_Read_Value[0] = 0x02;
            ADC_Read_Value[1] = 0x00;
            ADC_Read_Value[2] = 0x4D;
            //ADC_Read_Value[3] = 0x02;
            //ADC_Read_Value[4] = Short_Address[0];    // 6
            //ADC_Read_Value[5] = Short_Address[1];    // 7
            ADC_Read_Value[3] = 0x02;  //04
            ADC_Read_Value[4] = Short_Address[0];
            ADC_Read_Value[5] = Short_Address[1];
 


            for (int x = 0; x < TX_Counter - 1; x++)
            {
                TX_Data[x] = ADC_Read_Value[x];
            }


            //while (RX_Counter == 0 || Response[8] != FCS)
            //{
            zb_data_send();
            check_zigbee(0x10, 0x4D, 9);
            //}
            //RX_Counter = 0;

            int CO2value1, CO2value2;
            double CO2valuex;
            CO2value1 = (int)Response[6];     //16進轉10進   
            CO2value2 = (int)Response[7];     //16進轉10進
            CO2value1 = (int)CO2value1 * 16 * 16;
            CO2valuex = CO2value1 + CO2value2;
            CO2valuex = CO2valuex / 2047 * 1000 *3.3;
            CO2valuex = (int)CO2valuex;
            CO2valuex = (1280 - CO2valuex) * 5.4 ;  //2.5   3   3.4
            CO2valuex = (int)CO2valuex + 370;

            if (CO2valuex < 370)
            {
                CO2valuex = 370;
            }

            if (CO2valuex > 1000)
            {
                CO2valuex = 1000;
            }

            //lblShow.Text = "CO2 ok!";
            label10.Text = CO2valuex + "ppm";

            RX_Counter = 0;
            return CO2valuex;
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }


        private double Read_CO()
        {
            TX_Counter = 7;
            RX_Counter = 0;
            ADC_Read_Value[0] = 0x02;
            ADC_Read_Value[1] = 0x00;
            ADC_Read_Value[2] = 0x4D;
            ADC_Read_Value[3] = 0x02;
            ADC_Read_Value[4] = Short_Address[2];    // 6
            ADC_Read_Value[5] = Short_Address[3];    // 7


            for (int x = 0; x < TX_Counter - 1; x++)
            {
                TX_Data[x] = ADC_Read_Value[x];
            }


            //while (RX_Counter == 0 || Response[8] != FCS)
            //{
            zb_data_send();
            check_zigbee(0x10, 0x4D, 9);
            //}
            //RX_Counter = 0;

            int COvalue1, COvalue2;
            double COvaluex;
            COvalue1 = (int)Response[6];     //16進轉10進   
            COvalue2 = (int)Response[7];     //16進轉10進
            COvalue1 = (int)COvalue1 * 16 * 16;
            COvaluex = COvalue1 + COvalue2;
            COvaluex = COvaluex / 2047 * 1000;
            COvaluex = (int)COvaluex;
            COvaluex = (435 - COvaluex) * 8;  //742   2.5
            COvaluex = (int)COvaluex;

            if (COvaluex < 5)
                COvaluex = 5;


            if (COvaluex > 200)
                COvaluex = 200;


            //lblShow.Text = "CO ok!";
            label9.Text = COvaluex + "ppm";

            RX_Counter = 0;
            return COvaluex;

        }

        private void button5_Click(object sender, EventArgs e)
        {
                
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //for (int z = 0; z <= 10; z++)
            //{
                lblShow.Text = "Please Choose Button";

                for (int x = 0; x <= 200; x++)      //CO      membership
                {
                    CO1[x] = fun1(x, 35, 100);
                    CO2[x] = tri(x, 35, 100, 100, 150);
                    CO3[x] = fun2(x, 100, 150);
                }

                for (int x = 0; x <= 1000; x++)     //COtow   membership
                {
                    COtow1[x] = fun1(x, 500, 700);
                    COtow2[x] = tri(x, 550, 700, 700, 800);
                    COtow3[x] = fun2(x, 700, 800);
                }

                for (int x = 0; x <= 9; x++)         //FTem   membership
                {
                    FT1[x] = fun1(x, 2, 4);
                    FT2[x] = tri(x, 2, 4, 4, 6);
                    FT3[x] = fun2(x, 4, 6);
                }

                for (int x = 0; x <= 100; x++)       //Output membership
                {
                    Output1[x] = fun1(x, 15, 30);
                    Output2[x] = tri(x, 15, 30, 30, 50);
                    Output3[x] = tri(x, 30, 50, 50, 70);
                    Output4[x] = tri(x, 50, 70, 70, 85);
                    Output5[x] = fun2(x, 70, 85);
                }
                
                int CO, CO22, T;
                float level;
                CO = (int)Read_CO();
                Thread.Sleep(1000);
                CO22 = (int)Read_CO2();
                Thread.Sleep(1000);
                T = (int)Read_Temperature();
                level = risk_level(CO, CO22, T);
                label9.Text = CO + "ppm";
                label10.Text = CO22 + "ppm";
                label12.Text = T + "℃";

                label36.Text = Convert.ToString(Math.Round(Convert.ToDouble(level), 2, MidpointRounding.AwayFromZero));
                //label36.Text = " " + level;

                if ((int)level < 20)
                {
                    label14.Text = "Safe";
                    this.label14.ForeColor = System.Drawing.Color.Green;
                    label38.Text = "The current status is okay.";
                }
                else if ((int)level < 40)
                {
                    label14.Text = "Caution";
                    this.label14.ForeColor = System.Drawing.Color.Blue;
                    label38.Text = "Need to pay attention to the current environment.";
                }
                else if ((int)level < 60)
                {
                    this.label14.Font = new System.Drawing.Font("Times New Roman", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    label14.Text = "Low danger !";
                    this.label14.ForeColor = System.Drawing.Color.Red;
                    label38.Text = "Might result in death, serious injury, or damage!";
                }
                else if ((int)level < 80)
                {
                    this.label14.Font = new System.Drawing.Font("Times New Roman", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    label14.Text = "Medium danger ! !";
                    this.label14.ForeColor = System.Drawing.Color.Red;
                    label38.Text = "May result in death, serious injury, or damage!";
                }
                else
                {
                    this.label14.Font = new System.Drawing.Font("Times New Roman", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    label14.Text = "High danger ! ! !";
                    this.label14.ForeColor = System.Drawing.Color.Red;
                    label38.Text = "Could result in death, serious injury, or damage!";
                }

           // Thread.Sleep(6000);
            //}

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }


    }
}
