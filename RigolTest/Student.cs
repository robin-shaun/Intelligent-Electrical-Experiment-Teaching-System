using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Drawing;

namespace RigolTest
{
    class Student
    {
        //用于保存与每个客户相关信息：套接字与接收缓存
        public Socket socket;
        public byte[] Rcvbuffer;
        public String name;
        public String number;
        public String desknum;
        public int Count;
        public String grade;
        public Image Snapshot;
        public String RigolIP;
        public int imageCount = 0;   //截图数量

        //时间属性
        public DateTime dt_1;      //做完第一个实验分项的时刻
        public DateTime dt_2;      //第二个
        public DateTime dt_3;      //第三个，即结束时间

        //截图相关
        public Image[] snapShot = new Image[30];      //最多保存三十张截图
        public string[] description = new string[30]; //截图对应的描述


        //试验数据
        //进度变量
        public int process = 0;
        //静态工作点
        public double Beta;
        public double period1_Ub;
        public double period1_Uc;
        public double period1_Ue;
        public Image image_Q;
        //动态测量
        //不加反馈，接入无穷或5.1k电阻
        public double period2_Ui_1_infinity;
        public double period2_Uo_1_infinity;
        public double period2_Us_1_infinity;
        public Image image_1_infinity;
        public double period2_Ui_1_fivedotonek;
        public double period2_Uo_1_fivedotonek;
        public double period2_Us_1_fivedotonek;
        public Image image_1_fivedotonek;
        //加电压负反馈
        public double period2_Ui_2_infinity;
        public double period2_Uo_2_infinity;
        public double period2_Us_2_infinity;
        public Image image_2_infinity;
        public double period2_Ui_2_fivedotonek;
        public double period2_Uo_2_fivedotonek;
        public double period2_Us_2_fivedotonek;      
        public Image image_2_fivedotonek;
        //加电压负反馈，Rs为0
        public double period2_Ui_3_infinity;
        public double period2_Uo_3_infinity;
        public double period2_Us_3_infinity;
        public Image image_3_infinity;
        public double period2_Ui_3_fivedotonek;
        public double period2_Uo_3_fivedotonek;
        public double period2_Us_3_fivedotonek;
        public Image image_3_fivedotonek;
        //加电流负反馈
        public double period2_Ui_4_infinity;
        public double period2_Uo_4_infinity;
        public double period2_Us_4_infinity;
        public Image image_4_infinity;
        public double period2_Ui_4_fivedotonek;
        public double period2_Uo_4_fivedotonek;
        public double period2_Us_4_fivedotonek;
        public Image image_4_fivedotonek;
        //加电流负反馈，Rs为0
        public double period2_Ui_5_infinity;
        public double period2_Uo_5_infinity;
        public double period2_Us_5_infinity;
        public Image image_5_infinity;
        public double period2_Ui_5_fivedotonek;
        public double period2_Uo_5_fivedotonek;
        public double period2_Us_5_fivedotonek;
        public Image image_5_fivedotonek;
        //三个失真点
        //正常失真
        public Image image_period3_1;
        public double period3_Ue_1;
        public double period3_Ub_1;
        public double period3_Uc_1;
        //截止失真
        public Image image_period3_2;
        public double period3_Ue_2;
        public double period3_Ub_2;
        public double period3_Uc_2;
        //正常失真
        public Image image_period3_3;
        public double period3_Ue_3;
        public double period3_Ub_3;
        public double period3_Uc_3;
        public Boolean underchecked;

        //实例化方法
        public Student(String Name,String Number,String Desknum,String Riolip,int count)
        {
            name = Name;
            number = Number;
            desknum = Desknum;
            RigolIP = Riolip;
            Count = count;
            underchecked = false;
        }
        public Student(Socket s)
        {
            socket = s;
        }
        //清空接受缓存，在每一次新的接收之前都要调用该方法
        public void ClearBuffer()
        {
            Rcvbuffer = new byte[1024];
        }
        //
        public void Dispose()
        {
            try
            {
                socket.Shutdown(SocketShutdown.Both);
                socket.Close();
            }
            finally
            {
                socket = null;
                Rcvbuffer = null;
            }
        }
    }
}
