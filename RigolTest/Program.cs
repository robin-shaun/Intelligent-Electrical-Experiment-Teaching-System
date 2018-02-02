using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NationalInstruments.VisaNS;
using System.Windows.Forms;

namespace RigolTest
{
    /*class Program
    {
        static void Main(string[] args)
        {
            string m_strResourceName = null; //仪器资源名

            CVisaOpt m_VisaOpt = new CVisaOpt(); 

            string[] InstrResourceArray = m_VisaOpt.FindResource("?*INSTR"); //查找资源

            if (InstrResourceArray[0] == "未能找到可用资源!")
            {
                
            }
            else
            {
                //示例，选取DSG800系列仪器作为选中仪器
                for (int i = 0; i < InstrResourceArray.Length;i++ )
                {
                    
                    if (InstrResourceArray[i].Contains("DSG8"))
                    {
                        m_strResourceName = InstrResourceArray[i];
                    }
                }
               
            }
            //如果没有找到指定仪器直接退出
            if (m_strResourceName == null)
            {
                return;
            }
            //打开指定资源
            m_VisaOpt.OpenResource(m_strResourceName);
            //发送命令
            m_VisaOpt.Write("*IDN?");
            //读取命令
            string strback = m_VisaOpt.Read();
            //设置操作命令 1GHz频率 -10dBm幅度 打开RF输出开关
            m_VisaOpt.Write(":SOURce:FREQuency 1GHz");
            m_VisaOpt.Write(":SOURce:LEVel -10dBm");
            m_VisaOpt.Write(":OUTPut:STATe ON");
            //显示读取内容
            Console.Write(strback);
            
            //是否设备资源
            m_VisaOpt.Release();

            while (true) ;
        }
    }*/
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
