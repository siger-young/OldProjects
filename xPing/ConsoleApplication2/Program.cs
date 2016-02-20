using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Net;
using System.Net.Sockets;
using System.Net.NetworkInformation;

namespace xPing
{
    public class xPing
    {
        const int GOOD = 90;
        const int XTREME = 50;
        const int TIMES = 5;
        private static AutoResetEvent flag = new AutoResetEvent(false);
        private static List<string> IpList = new List<string>();
        private static string[] Ip;
        private static int PingTimes = 5;
        private static int PingCount = 0;
        private static string[] IpStatus;
        private static int[,] IpData;
        private static int[] IpAverage;
        private static string[] IpAverageDown;
        private static int Good;
        private static int Xtreme;
        //Main
        static void Main(string[] args)
        {
            InitConsole();
            InitConfig();
            StreamReader SrIp = new StreamReader("ip.txt");
            while (true)
            {
                string IpLine = SrIp.ReadLine();
                //string PartIp0,PartIp1;
                if (String.IsNullOrEmpty(IpLine))
                    break;
                if (IpLine.Contains("*"))
                {
                    //PartIp0 = IpLine.Split('*')[0];
                    //PartIp1 = IpLine.Split('*')[1];
                    for (int i = 0; i < 256; i++)
                    {
                        IpList.Add(IpLine.Replace("*", i.ToString()));
                        PingCount++;
                    }
                }
                else
                {
                    IpList.Add(IpLine);
                    PingCount++;
                }
            }
            Ip = IpList.ToArray();
            //DebugPrintIpList();
            IpData = new int[PingCount, PingTimes];
            IpAverage = new int[PingCount];
            IpAverageDown = new string[PingCount];
            IpStatus = new string[PingCount];
            Console.WriteLine(" Total: " + PingCount);
            Ping();
            GetAverage();
            PrintAverage();
            ShowHelp();
            Console.ReadKey();
        }

        private static void InitConsole()
        {
            //WriteFile("GetIdByIp.txt", "", false, true);
            //Console.BackgroundColor = ConsoleColor.DarkBlue;
            //Console.ForegroundColor = ConsoleColor.Gray;
            Console.WindowWidth = 80;
            Console.WindowHeight = 25;
            Console.Clear();
            if (File.Exists("ip.txt") == false)
            {
                ShowColor(ConsoleColor.DarkRed, "�޷��ҵ�ip.txt!");
                Console.ReadKey();
                System.Environment.Exit(0);
            }
            Console.Title = "BetaWorld xPing";
            Console.WriteLine(" BetaWorld xPing");
            Console.WriteLine("�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T");
            //Console.WriteLine("BetaWorld xPing");
            //Console.WriteLine("====================");
            Console.WriteLine();
            Console.WriteLine(" ������ּ�ڰ�������BetaWorldվ�����Ӧ�ٶȡ�(�����������������ֱ�ӻس�����)");
        }

        private static void DebugPrintIpList()
        {
            WriteFile("iplist.log", "Ip List", false, true);
            for (int i = 0; i < PingCount; i++)
            {
                WriteFile("iplist.log", Ip[i], true, true);
            }
        }


        private static void DebugPrintInfo()
        {
            WriteFile("info.log", "Ping Count: " + PingCount, false, true);
            WriteFile("info.log", "Ping Times: " + PingTimes, true, true);
        }

        private static void ShowHelp()
        {
            Console.WriteLine();
            Console.Write(" ���������Ѿ��������,���Բ鿴");
            ShowColor(ConsoleColor.Magenta, "debug.log");
            Console.WriteLine("��ȡ������־��");
            Console.WriteLine(" ��ȫ��Timed Out,�Ǿ�����RP������!");
            Console.Write(" �鿴");
            ShowColor(ConsoleColor.Magenta, "good.log");
            Console.WriteLine("��ȡ�Ϻýڵ���嵥");
            Console.Write(" �鿴");
            ShowColor(ConsoleColor.Magenta, "xtreme.log");
            Console.WriteLine("��ȡ���ѽڵ���嵥");
            Console.WriteLine(" ����Mike Gao��������,�ҵ�����㲻��Ҫ֪�� =D");
            ShowColor(ConsoleColor.Black," BetaWorld����Incapsula��ΪCDN, Incapsula��һ�����CDN");
            ShowColor(ConsoleColor.Black, " �����㷢���˱�������,��ô��һ���ɳ�Ϊһ���ڿ�!\n");
            Console.WriteLine(" ��������˳�...");
        }

        private static bool IsNumeric(string str)
        {
            foreach (char c in str)
            {
                if (!Char.IsNumber(c))
                {
                    return false;
                }
            }
            return true;
        }

        private static void PrintMaxThree()
        {
            for (int i = 0; i < PingCount; i++)
            {
                string CurrentIp = IpAverageDown[i].Split(' ')[0];
                string CurrentId = IpAverageDown[i].Split(' ')[1];
            }
        }

        private static void InitConfig()
        {
            InitTimes();
            InitXtreme();
            InitGood();
        }

        private static void ShowColor(ConsoleColor Color, string Message)
        {
            Console.ForegroundColor = Color;
            Console.Write(Message);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        private static void InitTimes()
        {
            while (true)
            {
                Console.Write(" ����������Ҫ�Ե����ڵ���ԵĴ���(times) [5]:");
                string Read = Console.ReadLine();
                if (Read == null || Read.Length == 0)
                {
                    PingTimes = TIMES;
                    return;
                }
                else if (IsNumeric(Read))
                {
                    if (Convert.ToInt32(Read) == 0)
                    {
                        ShowColor(ConsoleColor.DarkRed, " ��������Ϊ0!\n");
                    }
                    else
                    {
                        PingTimes = Convert.ToInt32(Read);
                        return;
                    }
                }
            }
        }

        private static void InitXtreme()
        {
            while (true)
            {
                Console.Write(" ����������Ϊһ��");
                ShowColor(ConsoleColor.DarkYellow, "����");
                Console.Write("�ڵ��ƽ����Ӧ(ms) [50]:");
                string Read = Console.ReadLine();
                if (Read == null || Read.Length == 0)
                {
                    Xtreme = XTREME;
                    break;
                }
                else if (IsNumeric(Read))
                {
                    if (Convert.ToInt32(Read) == 0)
                    {
                        ShowColor(ConsoleColor.DarkRed, " ��Ӧʱ�䲻��Ϊ0ms!\n");
                    }
                    else
                    {
                        Xtreme = Convert.ToInt32(Read);
                        break;
                    }
                }
            }
        }

        private static void InitGood()
        {
            while (true)
            {
                Console.Write(" ����������Ϊһ��");
                ShowColor(ConsoleColor.DarkGray, "�Ϻ�");
                Console.Write("�ڵ��ƽ����Ӧ(ms) [90]:");
                string Read = Console.ReadLine();
                if (Read == null || Read.Length == 0)
                {
                    Good = GOOD;
                    break;
                }
                else if (IsNumeric(Read))
                {
                    if (Convert.ToInt32(Read) == 0)
                    {
                        ShowColor(ConsoleColor.DarkRed, " ��Ӧʱ�䲻��Ϊ0ms!\n");
                    }
                    else if (Convert.ToInt32(Read) > Xtreme)
                    {
                        Good = Convert.ToInt32(Read);
                        break;
                    }
                    else if (Convert.ToInt32(Read) <= Xtreme)
                    {
                        ShowColor(ConsoleColor.DarkRed, " �Ϻ���Ӧʱ�������ڼ�����Ӧʱ��!\n");
                    }
                }
            }
        }

        private static void Ping()
        {
            WriteFile("debug.log", "Pinging data", false, true);
            for (int i = 0; i < PingCount; i++)
            {
                Console.WriteLine(" " + i + "\t" + Ip[i]);
                PingFunc(Ip[i]);
            }
        }

        private static void GetAverage()
        {
            for (int i = 0; i < PingCount; i++)
            {
                int Length = 0;
                for (int j = 0; j < PingTimes; j++)
                {
                    if (IpData[i, j] != 0)
                    {
                        Length++;
                    }
                }
                int Sum = GetSumById(i);
                if (Sum != 0 && Length != 0)
                {
                    IpAverage[i] = Sum / Length;
                }
                else
                {
                    IpAverage[i] = 0;
                }
                IpAverageDown[i] = IpAverage[i].ToString() + " " + i;
            }
        }

        private static void PrintAverage()
        {
            WriteFile("debug.log", "Average data", true, true);
            WriteFile("xtreme.log", "eXtreme data", false, true);
            WriteFile("good.log", "Good data", false, true);
            for (int i = 0; i < PingCount; i++)
            {
                string CurrentIp = Ip[i];
                string CurrentAverage = IpAverage[i].ToString();
                WriteFile("debug.log", " " + i + "\t" + String8ByteFormat(CurrentIp) + "\t" + CurrentAverage, true, true);
                if (IpStatus[i] == "Success" && IpAverage[i] >= 0 && IpAverage[i] <= Xtreme)
                {

                    WriteFile("xtreme.log", " " + i + "\t" + String8ByteFormat(CurrentIp) + "\t" + CurrentAverage, true, true);
                }
                else if (Xtreme < IpAverage[i] && IpAverage[i] <= Good)
                {
                    WriteFile("good.log", " " + i + "\t" + String8ByteFormat(CurrentIp) + "\t" + CurrentAverage, true, true);
                }
            }
        }

        //public static int[] BubbleSort(int[] Array)
        //{
        //    int[] list = Array;
        //    for (int i = 0; i < list.Length; i++)
        //    {
        //        for (int j = i; j < list.Length; j++)
        //        {
        //            if (list[i] < list[j])
        //            {
        //                int temp = list[i];
        //                list[i] = list[j];
        //                list[j] = temp;
        //            }
        //        }
        //    }
        //}

        private static int GetSumById(int id)
        {
            int Sum = 0;
            for (int i = 0; i < PingTimes; i++)
            {
                Sum += IpData[id, i];
            }
            return Sum;
        }

        private static int GetIdByIp(string IpAddress)
        {
            //WriteFile("GetIdByIp.txt", IpAddress, true, true);
            for (int i = 0; i < PingCount; i++)
            {
                if (Ip[i].Trim() == IpAddress)
                {
                    return i;
                }
            }
            return -1;
        }

        private static string GetIpById(int id)
        {
            return Ip[id];
        }

        private static void PingFunc(string str)
        {
            Ping p = new Ping();
            IPAddress IpAddress;
            str = str.Trim();
            if (IPAddress.TryParse(str, out IpAddress) && str != "0.0.0.0" && str.Split('.').Length == 4)
            {
                for (int i = 0; i < PingTimes; i++)
                {
                    p = new Ping();
                    p.PingCompleted += p_PingCompleted;
                    p.SendAsync(str, 500, str);
                    flag.WaitOne();
                }
            }
            else
            {
                ShowColor(ConsoleColor.DarkRed, "\t�Ƿ�IP!\n");
            }
        }

        private static void WriteFile(string FileName, string Content, bool Append, bool AutoCrLf)
        {
            StreamWriter File = new StreamWriter(FileName, Append);
            if (AutoCrLf)
            {
                File.WriteLine(Content);
            }
            else
            {
                File.Write(Content);
            }
            File.Close();
        }

        private static string String8ByteFormat(string str)
        {
            int Len = str.Length;
            string IpString = str;
            if (Len < 8)
            {
                for (int i = 0; i < 8 - Len; i++)
                {
                    IpString += " ";
                }
            }
            return IpString;
        }

        private static void p_PingCompleted(object sender, PingCompletedEventArgs e)
        {
            if (e.Reply != null)
            {
                lock (sender)
                {
                    string CurrentIp = e.UserState.ToString();
                    int CurrentId = GetIdByIp(CurrentIp);
                    int CurrentTime = Convert.ToInt32(e.Reply.RoundtripTime);
                    string Text = " " + CurrentId + "\t" + String8ByteFormat(CurrentIp) + "\t" + e.Reply.RoundtripTime + "\tStatus: " + e.Reply.Status;
                    WriteFile("debug.log", Text, true, true);
                    Console.WriteLine(Text);
                    for (int i = 0; i < PingTimes; i++)
                    {
                        if (IpStatus[CurrentId] != IPStatus.Success.ToString())
                        {
                            if (e.Reply.Status.ToString() == IPStatus.Success.ToString())
                                IpStatus[CurrentId] = IPStatus.Success.ToString();
                        }
                        if (IpData[CurrentId, i] == 0)
                        {
                            IpData[CurrentId, i] = CurrentTime;
                            break;
                        }
                    }
                    flag.Set();
                }
            }
        }
    }
}