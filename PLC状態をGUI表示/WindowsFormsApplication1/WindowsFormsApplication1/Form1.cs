using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ActProgTypeLib;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        ActProgType ap;
        int portnum;
        string portnumfn="portnum.txt";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string st;
            try
            {
                st = System.IO.File.ReadAllText(portnumfn);
                portnum = Int32.Parse(st);
            }
            catch
            {
                portnum = 3;
            }
        }
        private int searchport()
        {
            int ret,i;
            int[] testport = new int[100];

            for (i = 0; i < 100; i += 2)
            {
                testport[i] = portnum + i / 2;
                testport[i + 1] = portnum - i / 2;
            }
            for (i = 0; i < 100; i++)
            {
                if(testport[i]<3) continue;
                ap.ActPortNumber=testport[i];
                label18.Text = ap.ActPortNumber.ToString();
                ret = ap.Open();
                if (ret == 0) break;
            }
            if ((i > 1)&&(i<100))
            {
                System.IO.File.WriteAllText(portnumfn, testport[i].ToString());
            }
            if(i<100){
                return 0;
            }
            else{
                return 1;
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            ap = new ActProgType(); 
            
            //ap設定
            ap.ActBaudRate = ActDefine.BAUDRATE_9600;
            ap.ActDataBits = ActDefine.DATABIT_8;
            ap.ActControl = ActDefine.TRC_DTR_OR_RTS;
            ap.ActStopBits = ActDefine.STOPBIT_ONE;
            ap.ActCpuType = ActDefine.CPU_FX2NCPU;
       //     ap.ActCpuType = ActDefine.CPU_FX1NCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit = ActDefine.NO_PARRITY;
            ap.ActIONumber = 1023;
            ap.ActPortNumber = portnum;
            ap.ActProtocolType = ActDefine.PROTOCOL_SERIAL;
            ap.ActTimeOut = 100;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_FXCPU;


            int ret;
            int d10, d11,pos11=0,pos12=0, pos21=0, pos22=0;
            int[] val = new int[2000];
            int num1=0, num2=0, way2=0;
            ret = searchport();
            if (ret == 0)
            {
                ret|=ap.ReadDeviceRandom("d10", 1, out d10);
                pos11 = d10 & 0x3ff;
                num1 = (d10 >> 12) & 0x3;
                ret|=ap.ReadDeviceRandom("d11", 1, out d11);
                pos12 = d11 & 0x3ff;
                num2 = (d11 >> 12) & 0x3;
                way2 = (d11 >> 14) & 0x3;
                ret|=ap.ReadDeviceRandom("d12", 1, out pos21);
                ret|=ap.ReadDeviceRandom("d13", 1, out pos22);

                ap.Close();

            }
            if(ret==0){
                label3.Text = pos11.ToString();
                label4.Text = pos12.ToString();
                label11.Text = num1.ToString();

                label8.Text = pos21.ToString();
                label6.Text = pos22.ToString();
                label13.Text = num2.ToString();
                label15.Text = way2.ToString();
            }
            else{
                label3.Text = "??";
                label4.Text = "??";
                label11.Text = "??";

                label8.Text = "??";
                label6.Text = "??";
                label13.Text = "??";
                label15.Text = "??";
            }

        }

 

 
    }
}
