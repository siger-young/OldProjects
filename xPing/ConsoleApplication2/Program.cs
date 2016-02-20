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
                ShowColor(ConsoleColor.DarkRed, "无法找到ip.txt!");
                Console.ReadKey();
                System.Environment.Exit(0);
            }
            Console.Title = "BetaWorld xPing";
            Console.WriteLine(" BetaWorld xPing");
            Console.WriteLine("════════════════════");
            //Console.WriteLine("BetaWorld xPing");
            //Console.WriteLine("====================");
            Console.WriteLine();
            Console.WriteLine(" 本工具旨在帮助提升BetaWorld站点的响应速度。(以下设置若不清楚请直接回车跳过)");
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
            Console.Write(" 整个流程已经运行完毕,可以查看");
            ShowColor(ConsoleColor.Magenta, "debug.log");
            Console.WriteLine("获取调试日志。");
            Console.WriteLine(" 若全部Timed Out,那就是你RP问题辣!");
            Console.Write(" 查看");
            ShowColor(ConsoleColor.Magenta, "good.log");
            Console.WriteLine("获取较好节点的清单");
            Console.Write(" 查看");
            ShowColor(ConsoleColor.Magenta, "xtreme.log");
            Console.WriteLine("获取极佳节点的清单");
            Console.WriteLine(" 可向Mike Gao反馈数据,我的身份你不需要知道 =D");
            ShowColor(ConsoleColor.Black," BetaWorld采用Incapsula作为CDN, Incapsula是一款免费CDN");
            ShowColor(ConsoleColor.Black, " 倘若你发现了本段文字,那么你一定可成为一名黑客!\n");
            Console.WriteLine(" 按任意键退出...");
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
                Console.Write(" 请输入您想要对单个节点测试的次数(times) [5]:");
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
                        ShowColor(ConsoleColor.DarkRed, " 次数不能为0!\n");
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
                Console.Write(" 请输入您认为一个");
                ShowColor(ConsoleColor.DarkYellow, "极佳");
                Console.Write("节点的平均响应(ms) [50]:");
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
                        ShowColor(ConsoleColor.DarkRed, " 响应时间不能为0ms!\n");
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
                Console.Write(" 请输入您认为一个");
                ShowColor(ConsoleColor.DarkGray, "较好");
                Console.Write("节点的平均响应(ms) [90]:");
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
                        ShowColor(ConsoleColor.DarkRed, " 响应时间不能为0ms!\n");
                    }
                    else if (Convert.ToInt32(Read) > Xtreme)
                    {
                        Good = Convert.ToInt32(Read);
                        break;
                    }
                    else if (Convert.ToInt32(Read) <= Xtreme)
                    {
                        ShowColor(ConsoleColor.DarkRed, " 较好响应时间必须大于极佳响应时间!\n");
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
                ShowColor(ConsoleColor.DarkRed, "\t非法IP!\n");
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