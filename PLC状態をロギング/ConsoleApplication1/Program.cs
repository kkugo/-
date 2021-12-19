using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using ActProgTypeLib;



namespace ConsoleApplication1
{
    class Program
    {
        static ManagementObjectSearcher Osch = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
        static ManagementObjectCollection Ocl;
        static ActProgType ap;

        class dbuffer
        {
            int[] count = new int[2000];
            int[] val = new int[2000];
            string[] btop = new string[2000];
            int n = 0;
            scontcts lastwritten = new scontcts { };

            public void setbuf(int lgt)
            {
                Array.Resize(ref count, lgt);
                Array.Resize(ref   val, lgt);
                Array.Resize(ref  btop, lgt);
            }
            public void addring(scontcts dt)
            {
                if (!dt.ifsame(lastwritten))
                {
                    int i, j, ad;
                    ad = 0;
                    for (i = 0; i < dt.blocks; ++i)
                    {
                        for (j = 0; j < dt.blen[i]; ++j)
                        {
                            count[n] = dt.readcnt;
                            val[n] = dt.data[ad];
                            btop[n] = dt.btop[i] + "+" + j;
                            ad++;
                            n++;
                            if (n >= count.Length) n = 0;
                        }
                    }
                    dt.deepcopyto(ref lastwritten);
                }
            }
            public void add(scontcts dt)
            {
                int i, j, ad;
                ad = 0;
                for (i = 0; i < dt.blocks; ++i)
                {
                    for (j = 0; j < dt.blen[i]; ++j)
                    {
                        count[n] = dt.readcnt;
                        val[n] = dt.data[ad];
                        btop[n] = dt.btop[i] + "+" + j;
                        ad++;
                        n++;
                    }
                }
                if (n > count.Length - 1000) Array.Resize(ref count, count.Length + 1000);
                if (n > val.Length - 1000) Array.Resize(ref   val, val.Length + 1000);
                if (n > btop.Length - 1000) Array.Resize(ref  btop, btop.Length + 1000);

            }
            public void write()
            {
                string fl = "out.txt";
                string st;
                int i;
                st = "";
                for (i = 0; i < count.Length; ++i)
                {
                    st += count[i] + " " + btop[i] + " " + val[i];
                    st += Environment.NewLine;
                    if(i%5000==0){
                        File.AppendAllText(fl, st);
                        st = "";
                        Console.Write("#");
                    }
                }
                File.AppendAllText(fl, st);

            }
        }
        class scontcts
        {
            public int blocks;
            public int readcnt;
            public int[] blen = new int[1];
            public string[] btop = new string[1];
            public int[] data = new int[1];
            public int[] filter = new int[1];
            public int ocdxhx;
            public string[] cntname=new string[1];

            public void deepcopyto(ref scontcts dst)
            {
                dst.blocks = blocks;
                Array.Resize(ref dst.blen, blocks );
                Array.Resize(ref dst.btop, blocks );

                int i,j;
                j = 0;
                for (i = 0; i < blocks; ++i)
                {
                    dst.blen[i] = blen[i];
                    dst.btop[i] = btop[i];
                    j += blen[i];
                }
                Array.Resize(ref dst.data, j);
                for (i = 0; i < j; ++i) dst.data[i] = data[i];
 
            }
            string make_cntlist()
            {
                string val="";
                int i,j,k;
                string tpcn,st;
                int tpofs,tpval;
                for (j = 0; j < blocks; ++j)
                {
                    for (i = 0; i < blen[j]; ++i)
                    {
                        tpcn = btop[j];
                        tpofs = i * 16;
                        tpcn = (tpcn.Split('+'))[0];
                        st = tpcn.Substring(1, tpcn.Length - 1);
                        tpcn = tpcn.Substring(0, 1);
                        tpval = Convert.ToInt32(st, ocdxhx);
                        for (k = 0; k < 16; ++k)
                        {
                            val += tpcn + Convert.ToString(tpval + tpofs + k, ocdxhx) + ",";
                        }
                    }
                }
                return val;
            }
            public bool ifsame(scontcts cnts)
            {
                int i;
                bool rsl=true;
                int change;
                if (cnts == null) return false;
                if (cnts.data.Length != data.Length) return false;
                for(i=0;i<data.Length;++i){
                    change=(cnts.data[i] & filter[i])^(data[i] & filter[i]);
                    if ( change!=0 )
                    {
                        int k,a, b;
                        rsl = false;
                        for (k = 0; k < 16; ++k)
                        {
                            b = 1;
                            for (a = 0; a < k; ++a) b *= 2;
                            if ((b & change) != 0)
                            {
                                Console.Write("<" + cntname[i * 16 + k] + ">");
                            }
                        }
                    }
                }
                return rsl;
            }

            public void init_cnt(int i_ocdxhx)
            {
                string[] devlist = new string[1];
                readlist("cntc.txt", ref devlist);

                ocdxhx = i_ocdxhx;
                blocks = 0;
                foreach (string dev in devlist)
                {
                    string st;
                    if (dev != null)
                    {
                        var spl = dev.Split(' ');
 

                        Array.Resize(ref blen, blocks + 1);
                        Array.Resize(ref btop, blocks + 1);

                        st = spl[0];
                        blen[blocks] = int.Parse(spl[1]);
                        btop[blocks] = spl[0];

                        blocks++;
                    }
                }
                int i; int l;
                l = 0;
                for (i = 0; i < blocks; ++i)
                {
                    l += blen[i];
                }
                Array.Resize(ref data, l);
                Array.Resize(ref filter, l);
                for (i = 0; i < l; ++i) filter[i] = ~0;

                readcnt=0;
                string cl;
                cl = make_cntlist();
                cntname=cl.Substring(0,cl.Length-1).Split(',');
            }
            public void make_excpfiler(string fn)
            {
                string[] fbuf=new string[1];
                readlist(fn,ref fbuf);

                int j, i,k,n;
                for (k = 0; k < fbuf.Length; ++k)
                {
                    for (j = 0; j < cntname.Length; ++j)
                    {
                        n = j / 16;
                        i = j % 16;
                        if (cntname[j] == fbuf[k])
                        {
                            int a, b;
                            b = 1;
                            for (a = 0; a < i; ++a) b = b * 2;
                            filter[n] &= ~b;
                        }
                    }
                }

            }
            public int read()
            {
                int i;int ad;
                int ret=0;
                ad=0;
                for(i=0;i<blocks;++i){
                    if (ap.ReadDeviceBlock(btop[i], blen[i], out data[ad]) != 0)
                    {
                        ret = 1;
                        break;
                    }
                    ad+= blen[i];
                }
                readcnt++;
                return ret;
            }
            public int read_nobreak()
            {
                int i; int ad;
                int ret = 0;
                ad = 0;
                for (i = 0; i < blocks; ++i)
                {
                    if (ap.ReadDeviceBlock(btop[i], blen[i], out data[ad]) != 0)
                    {
                        ret = 1;
                    }
                    ad += blen[i];
                }
                readcnt++;
                return ret;
            }
            public void write()
            {
                string fl="out.txt";
                string st;
                int i,j,ad;
                ad=0;
                st = "";
                for (i = 0; i < blocks; ++i)
                {
                    for (j = 0; j < blen[i]; ++j)
                    {
                        st += readcnt + " " + btop[i] + "+" + j + " " + data[ad];
                        st += Environment.NewLine;
                        ad++;
                    }
                }
                File.AppendAllText(fl, st);
            }

        }
        static dbuffer dbf=new dbuffer();
        static scontcts cnt=new scontcts();

        static void Main(string[] args)
        {
        //args0 CPUtype
           int rsl;
            int data1;
//            plctype = detectPLCtype(args[0]);
            File.WriteAllText("stop.txt", "0");
            cnt.init_cnt(Convert.ToInt32(args[1]));
            cnt.make_excpfiler("excps.txt");
            dbf.setbuf(20000);
            rsl=devopen(args[0]);
            if (rsl < 100)
            {
                do
                {
                    int ret;
                    ret=cnt.read();
                    if (cnt.readcnt % 10 == 0) Console.Write("*");
                    if (ret == 0)
                    {
                        dbf.addring(cnt);
                    }
                    else
                    {
                        Console.Write("x");
                    }
                }
                while (!icheck());
                ap.Close();
                dbf.write();
            }
        }
        static bool icheck()
        {
            int scs;
            string st="0";            
            scs=0;
            while(scs==0){
                scs = 1;
                try
                {
                    StreamReader sr = new StreamReader("stop.txt");
                    st = sr.ReadLine();
                    sr.Close();
                }
                catch
                {
                    scs = 0;
                }
            }
            return st.Contains("1");
        }
        static void enable_network()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "netsh.exe";
            p.StartInfo.Arguments = "interface set interface \"ローカル エリア接続\" enable";
            p.Start();
            p.WaitForExit();
        }
        static void disable_network()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "netsh.exe";
            p.StartInfo.Arguments = " interface set interface \"ローカル エリア接続\" disable";

   
            p.Start();
            p.WaitForExit();

        }
        static void disableallmitusbishi()
        {
            int CountDevice = Osch.Get().Count;
            Ocl = Osch.Get();
            string[] Array_DeviceID;//取得ID分解用配列
            foreach (ManagementObject obj in Ocl)
            {
                var Value_DeviceID = obj["DeviceID"].ToString();
                var splitid = Value_DeviceID.Split('\\');
                if (splitid[0].Contains("USB")){
                    if( splitid[1].Contains("VID_06D3&PID_2870")
                      ||splitid[1].Contains("VID_06D3&PID_1800") )
                    {

                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.Verb = "Runas";
                    p.StartInfo.FileName = "c:\\obj\\vctest\\devcon.exe";
                    p.StartInfo.Arguments = "disable \"@" + Value_DeviceID  +"\"";
                    p.Start();
                    p.WaitForExit();
                    }
     
                }
            }

        }
        static void enableusb(string id)
        //デバイスインスタンスパス = idのデバイスを有効にする
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "c:\\obj\\vctest\\devcon.exe";
            p.StartInfo.Arguments = "enable \"@"+id+"\"";
            p.Start();
            p.WaitForExit();
        }
        static void disableusb(string id)
        //デバイスインスタンスパス = idのデバイスを無効にする
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "c:\\obj\\vctest\\devcon.exe";
            p.StartInfo.Arguments = "disable \"@" + id + "\"";
            p.Start();
            p.WaitForExit();
        }
        static void disable_allfirewall()
        {
            string[] devlist = new string[1];
            readlist("devlist.txt", ref devlist);
            foreach (string dev in devlist)
            {
                if (dev!=null)
                {
                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.Verb = "Runas";
                    p.StartInfo.FileName = "c:\\obj\\vctest\\devcon.exe";
                    p.StartInfo.Arguments = "advfirewall firewall set rule name=\"usbsvr";
                    p.StartInfo.Arguments += dev.PadLeft(3,'0') +" new enable"; 
                    p.Start();
                    p.WaitForExit();
                }
            }

        }
        static void pass_firewall(string ipadr)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.Verb = "Runas";
            p.StartInfo.FileName = "c:\\obj\\vctest\\devcon.exe";
            p.StartInfo.Arguments = "advfirewall firewall set rule name=\"usbsvr";
            p.StartInfo.Arguments += ipadr.PadLeft(3, '0') + " new disable";
            p.Start();
            p.WaitForExit();
        }
        static void contout(ref ActProgType ap0,string cnt)
        {
            int data;
            string fl;
            DateTime dt;
            dt=DateTime.Now;
            fl = "C:\\obj\\out\\d"+dt.Hour+"-"+dt.Minute+".txt";
            ap0.ReadDeviceRandom(cnt, 1, out data);
            File.AppendAllText(fl, dt.Second + " " + cnt + " " + data+Environment.NewLine);

        }
        static string detectPLCtype(string nospvid)
        {
            string[] idlist= new string[1];
            readlist("idlist.txt",ref idlist);
            var vid = nospvid.Split('\\');

            foreach (string nospplcid in idlist)
            {
                var plcid = nospplcid.Split(' ');
                if (vid[1] == plcid[0]) return plcid[1];
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
        static void selectCPU_Q06UDV(ref ActProgType ap)
        {
            ap.ActCpuType = ActDefine.CPU_Q06UDVCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit = 1;
            ap.ActDsidPropertyBit = 1;
            ap.ActIntelligentPreferenceBit = 0;
            ap.ActIONumber = 1023;
            ap.ActMultiDropChannelNumber = 0;
            ap.ActNetworkNumber = 0;
            ap.ActProtocolType = ActDefine.PROTOCOL_USB;
            ap.ActStationNumber = 255;
            ap.ActThroughNetworkType = 0;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_QNUSB;
        }
        static void selectCPU_Q03UDE(ref ActProgType ap)
        {
            ap.ActCpuType = ActDefine.CPU_Q03UDECPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit = 1;
            ap.ActDsidPropertyBit = 1;
            ap.ActIntelligentPreferenceBit = 0;
            ap.ActIONumber = 1023;
            ap.ActMultiDropChannelNumber = 0;
            ap.ActNetworkNumber = 0;
            ap.ActProtocolType = ActDefine.PROTOCOL_USB;
            ap.ActStationNumber = 255;
            ap.ActThroughNetworkType = 0;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_QNUSB;
        }
        static void selectCPU_Q13UDV(ref ActProgType ap)
        {
            ap.ActCpuType = ActDefine.CPU_Q13UDVCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit=1;
            ap.ActDsidPropertyBit=1;
            ap.ActIntelligentPreferenceBit=0;
            ap.ActIONumber = 1023;
            ap.ActMultiDropChannelNumber=0;
            ap.ActNetworkNumber=0;
            ap.ActProtocolType = ActDefine.PROTOCOL_USB;
            ap.ActStationNumber=255;
            ap.ActThroughNetworkType=0;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_QNUSB;
        }
        static void selectCPU_FX3G(ref ActProgType ap)
        {
            ap.ActCpuType = ActDefine.CPU_FX3GCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActIONumber = 0;
            ap.ActProtocolType = ActDefine.PROTOCOL_USB;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_FXCPU;
        }
        static void selectCPU_FX3UC(ref ActProgType ap)
        {
            ap.ActCpuType = ActDefine.CPU_FX3UCCPU;
            ap.ActBaudRate = ActDefine.BAUDRATE_115200;
            ap.ActControl = ActDefine.TRC_DTR_OR_RTS;
            ap.ActDataBits = ActDefine.DATABIT_8;
            ap.ActDestinationIONumber = 0;
            ap.ActStopBits = ActDefine.STOPBIT_ONE;
            ap.ActDidPropertyBit = ActDefine.NO_PARRITY;
            ap.ActIONumber = 1023;
            ap.ActPortNumber = 4;
            ap.ActProtocolType = ActDefine.PROTOCOL_SERIAL;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_FXCPU;
        }

        static void selectCPU_FX2N(ref ActProgType ap)
        {
/*            ap.ActBaudRate = ActDefine.BAUDRATE_9600;
            ap.ActControl = ActDefine.TRC_DTR_OR_RTS;
            ap.ActCpuType = ActDefine.CPU_FX2NCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit = 1;
            ap.ActIONumber = 1023;
            ap.ActPortNumber = 4;
            ap.ActProtocolType = ActDefine.PROTOCOL_SERIAL;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_FXCPU;
*/
            ap.ActBaudRate = ActDefine.BAUDRATE_19200;
            ap.ActDataBits = ActDefine.DATABIT_8;
            ap.ActControl = ActDefine.TRC_DTR_OR_RTS;
            ap.ActStopBits = ActDefine.STOPBIT_ONE;
            ap.ActCpuType = ActDefine.CPU_FX2NCPU;
            ap.ActDestinationIONumber = 0;
            ap.ActDidPropertyBit =  ActDefine.NO_PARRITY;
            ap.ActIONumber = 1023;
            ap.ActPortNumber = 4;
            ap.ActProtocolType = ActDefine.PROTOCOL_SERIAL;
            ap.ActTimeOut = 5000;
            ap.ActUnitNumber = 0;
            ap.ActUnitType = ActDefine.UNIT_FXCPU;

        }
        static int devopen(string cpu)
        {
            int i;
            ap = new ActProgType();
            i = 101;
            switch (cpu)
            {
                case "FX3G": selectCPU_FX3G(ref ap); break;
                case "FX3UC": selectCPU_FX3UC(ref ap); break;
                case "FX2N": selectCPU_FX2N(ref ap); break;
                case "Q03UDE": selectCPU_Q03UDE(ref ap); break;
                case "Q06UDV": selectCPU_Q06UDV(ref ap); break;
                case "Q13UDV": selectCPU_Q13UDV(ref ap); break;
            }
            if (cpu != "error")
            {
                int ret;
                ret = 1;
                i = 0;
                while ((ret != 0)&&(i<100))
                {
                    Console.Write(" openning port " + ap.ActPortNumber);
                    ret = ap.Open();
                    System.Threading.Thread.Sleep(1000);
                    if ((ret == 0x1808008) || (ret == 0x1808009)) Console.Write(" port");
                    if ((ret != 0)) Console.Write(" failed");
                    ++i;

                }
            }
            return i;

        }
        static int usbread(string cpu, string cntct)
        {
            int data1=0;

            ap = new ActProgType();
            switch (cpu)
            {
                case "FX3G": selectCPU_FX3G(ref ap); break;
                case "FX2N": selectCPU_FX2N(ref ap); break;
                case "Q13UDV": selectCPU_Q13UDV(ref ap); break;
            }
            if (cpu != "error")
            {
                int ret;
                ret = 1;
                while (ret != 0)
                {
                    ret = ap.Open();
                    System.Threading.Thread.Sleep(1000);

                }
                ap.ReadDeviceRandom(cntct, 1, out data1);
                ap.Close();
            }
            return data1;

        }
        static void writefile(string fn, int rsl)
        {
            string fl;
            string cnt;
            DateTime dt;
            dt = DateTime.Now;
            fl = ".\\rec\\" + fn + +dt.Month + "-" + dt.Day + ".txt";
            cnt = dt.Hour + "-" + dt.Minute + " " + rsl;

            File.AppendAllText(fl, cnt + Environment.NewLine);

        }
    }
}
