using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using ACTPCUSBLib;



namespace ConsoleApplication1
{
    class Program
    {
        static ManagementObjectSearcher Osch = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
        static ManagementObjectCollection Ocl;
 //       ActProgType ap;

        static void Main(string[] args)
        {
        //args0 読む接点 '@'区切りで複数入れる
        //args1 IPアドレス末尾
//            selectusb();
            string plctype;
            int rsl;
            Console.WriteLine("0");
            disable_network();
            disable_allfirewall();
            enable_network();
            Console.WriteLine("1");
            System.Threading.Thread.Sleep(3500);

            Console.WriteLine("2");
            disable_network();
            pass_firewall(args[1]);
            enable_network();
            Console.WriteLine("3");


            plctype = detectPLCtype(args[1]);
            rsl=usbread(plctype, args[0],args[1].PadLeft(3, '0') + "-");
        }
        static void enable_network()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "netsh";
            p.StartInfo.Arguments = "interface set interface \"ローカル エリア接続 3\" enable";
            p.Start();
            p.WaitForExit();
        }
        static void disable_network()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "netsh";
            p.StartInfo.Arguments = " interface set interface \"ローカル エリア接続 3\" disable";
            p.Start();
            p.WaitForExit();

        }
        static void restart_usbserver()
        {
            disable_network();
            kill_usbserver();
            exec_usbserver();
            enable_network();
        }
        static void kill_usbserver()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
             p.StartInfo.Verb = "Runas";
             p.StartInfo.FileName = "taskkill";
             p.StartInfo.Arguments = "/im connect.exe /t /f";
             p.Start();
             p.WaitForExit();
        }
        static void exec_usbserver()
        {
            System.Diagnostics.Process.Start("C:\\Program Files\\silex technology\\SX Virtual Link\\Connect.exe");
        }

        static void disable_allfirewall()
        {
            string[] devlist = new string[1];
            readlist("devlist.txt", ref devlist);
            foreach (string dev in devlist)
            {
                if (dev!=null)
                {
                    var ipl = dev.Split('\t');

                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.Verb = "Runas";
                    p.StartInfo.FileName = "netsh";
                    p.StartInfo.Arguments = "advfirewall firewall set rule name=\"usbsvr";
                    p.StartInfo.Arguments += ipl[0].PadLeft(3,'0')+ "s\" new enable=yes"; 
                    p.Start();
                    p.WaitForExit();
                }
            }

        }
        static void pass_firewall(string ipadr)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "netsh";
            p.StartInfo.Arguments = "advfirewall firewall set rule name=\"usbsvr";
            p.StartInfo.Arguments += ipadr.PadLeft(3, '0') + "s\" new enable=no";
            p.Start();
            p.WaitForExit();
        }

        static string detectPLCtype(string nospvid)
        {
            string[] idlist= new string[1];
            readlist("devlist.txt",ref idlist);
            var vid = nospvid;

            foreach (string nospplcid in idlist)
            {
                var plcid = nospplcid.Split('\t');
                if (vid == plcid[0]) return plcid[1];
            }
            return "error";
            
        }
        static void readlist(string filename,ref string[] cnt){
            StreamReader sr = new StreamReader(@filename);
            while (sr.Peek() != -1)
            {
                cnt[cnt.Length - 1] = sr.ReadLine();
                Array.Resize(ref cnt, cnt.Length + 1);
            }
        }
        static void selectCPU_Q13UDV(ref ActQCPUQUSBClass ap)
              {
                  ap.ActCpuType = ActDefine.CPU_Q13UDVCPU;
                  ap.ActTimeOut = 5000;

              }
      static void selectCPU_Q13UDH(ref ActQCPUQUSBClass ap)
      {
          ap.ActCpuType = ActDefine.CPU_Q13UDHCPU;
          ap.ActTimeOut = 5000;

      }
      static void selectCPU_Q06UDV(ref ActQCPUQUSBClass ap)
      {
          ap.ActCpuType = ActDefine.CPU_Q06UDVCPU;
          ap.ActTimeOut = 5000;

      }
      static void selectCPU_Q03UDV(ref ActQCPUQUSBClass ap)
      {
          ap.ActCpuType = ActDefine.CPU_Q03UDVCPU;
          ap.ActTimeOut = 5000;

      }
      static void selectCPU_Q03UDE(ref ActQCPUQUSBClass ap)
      {
          ap.ActCpuType = ActDefine.CPU_Q03UDECPU;
          ap.ActTimeOut = 5000;
      }
      static void selectCPU_Q04UDV(ref ActQCPUQUSBClass ap)
      {
          ap.ActCpuType = ActDefine.CPU_Q04UDVCPU;
          ap.ActTimeOut = 5000;
      }


        static void selectCPU_FX3G(ref ActFXCPUUSBClass ap)
        {
            ap.ActCpuType = ActDefine.CPU_FX3GCPU;
            ap.ActTimeOut = 5000;

        }
        static int usbread(string cpu, string cntct,string fn)
        {
            
            string[] cn=cntct.Split('@');
            int data1=0;

            ActQCPUQUSBClass ap = new ActQCPUQUSBClass();
            switch (cpu)
            {
    //            case "uFX3G": selectCPU_FX3G(ref ap); break;
                case "uQ13UDV": selectCPU_Q13UDV(ref ap); break;
                case "uQ13UDH": selectCPU_Q13UDH(ref ap); break;
                case "uQ06UDV": selectCPU_Q06UDV(ref ap); break;
                case "uQ03UDV": selectCPU_Q03UDV(ref ap); break;
                case "uQ03UDE": selectCPU_Q03UDE(ref ap); break;
                case "uQ04UDV": selectCPU_Q04UDV(ref ap); break;




            }
            if (cpu != "error")
            {
                int i,ret;
                ret = 1;
                for(i=1;i<100;++i)
                {
                    ret = ap.Open();
                    System.Threading.Thread.Sleep(3000);
                    if(ret==0) break;
                    if(i%20==0) restart_usbserver();

                }
                if(i<99) foreach (string s in cn)
                {
                    ap.ReadDeviceRandom(s, 1, out data1);
                    writefile(fn, s, data1);
                }
                ap.Close();
            }
            return data1;

        }
 /*       static int usbread(string cpu, string cntct, string fn)
        {

            string[] cn = cntct.Split('@');
            int data1 = 0;

            ActQCPUQUSBClass ap = new ActFXCPUUSBClass();
            switch (cpu)
            {
                case "uFX3G": selectCPU_FX3G(ref ap); break;
                case "uQ13UDV": selectCPU_Q13UDV(ref ap); break;
                case "uQ13UDH": selectCPU_Q13UDH(ref ap); break;
            }
            if (cpu != "error")
            {
                int ret;
                ret = 1;
                while (ret != 0)
                {
                    ret = ap.Open();
                    System.Threading.Thread.Sleep(3000);

                }
                foreach (string s in cn)
                {
                    ap.ReadDeviceRandom(s, 1, out data1);
                    writefile(fn, s, data1);
                }
                ap.Close();
            }
            return data1;

        }*/
        static void writefile(string fn,string cnt, int rsl)
        {
            string fl;
            string str;
            DateTime dt;
            dt = DateTime.Now;
            fl = ".\\rec\\" + fn + +dt.Month + "-" + dt.Day + ".txt";
            str = dt.Hour + "-" + dt.Minute +" "+ cnt + " " + rsl;

            File.AppendAllText(fl, str + Environment.NewLine);

        }
    }
}
