using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Drawing.Imaging;
using System.Windows.Media.Imaging;
using System.Collections;
using System.Net.Sockets;
using System.Net;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace RigolTest
{
    public partial class Form1 : Form
    {
        public System.Timers.Timer[] timers = new System.Timers.Timer[15];
        //public System.Timers.Timer timer2 = new System.Timers.Timer();
        private static int[] inTimers = new int[15];
        //private static int inTimer2 = 0;
        public static int[] timeSequence = new int[15];        //在最后,数组中存储的是学生完成实验的顺序
        
        Queue<int> HandInSequence = new Queue<int>();            //HandInSequence--用来存储学生提交顺序的队列；即每次处于待批改状态，stu.underchecked=1的学生只能有一个
        bool isThereUnderchecked = false;                        //逻辑：学生发送‘submit’，PC端收到后查找students中是否有underchecked=1者，没有，往下执行，有，进队；
                                                                 //当老师完成批改时，检查队中是否有项，有，出队，执行批改操作
        //保存与学生相关的信息列表
        static int i = 0;
        static int num = 15;
        Student[] students = new Student[num];
        //负责监听的套接字
        TcpListener listener;
        //指示是否启动了监听
        bool IsStart = false;
        private string Ω;
        //Picturebox数组
        PictureBox[] pictureboxes = new PictureBox[num];
        //连接label数组
        Label[] connectLabels = new Label[num];
        //提交label数组
        Label[] submitLabels = new Label[num];

        public Form1()
        {
            for (int i = 0; i < 15; i++)
            {
                timers[i] = new System.Timers.Timer();
                inTimers[i] = 0;
            }
            //用于排序的数组赋初值
            for (int i = 0; i < 15; i++)
                timeSequence[i] = i;
            InitializeComponent();
            students[0] = new Student("沙行安", "15051155", "1", "USB0::0x1AB1::0x0588::DS1ET181705183::INSTR", 0); 
            students[1] = new Student("苗子琛", "15071129", "2", "TCPIP::169.254.205.57::INSTR", 0);
            students[2] = new Student("肖昆", "14051166", "3", "TCPIP::169.254.151.57::INSTR", 0); 
            students[3] = new Student("钱瑞", "15031208", "4", "TCPIP::169.254.25.59::INSTR", 0);
            students[4] = new Student("苗津毓", "15031216", "5", "TCPIP::169.254.229.57::INSTR", 0);
            students[5] = new Student("张家奇", "15031204", "6", "TCPIP::169.254.174.57::INSTR", 0);
            students[6] = new Student("陈禹昕", "15031203", "7", "TCPIP::169.254.121.57::INSTR", 0);
            students[7] = new Student("刘畅", "15031199", "8", "TCPIP::169.254.203.57::INSTR", 0);
            students[8] = new Student("刘华华", "15031198", "9", " TCPIP::169.254.232.57::INSTR", 0);
            students[9] = new Student("钱于杰", "15031194", "10", "TCPIP::169.254.180.57::INSTR", 0);
            students[10] = new Student("赵俊贺", "15031191", "11", "TCPIP::169.254.201.57::INSTR", 0);
            students[11] = new Student("李艺涵", "15031190", "12", "TCPIP::192.168.1.104::INSTR", 0);
            students[12] = new Student("李欣瑞", "15031189", "13", "TCPIP::192.168.1.109::INSTR", 0);
            students[13] = new Student("李思奇", "15031184", "14", "TCPIP::192.168.1.108::INSTR", 0);
            students[14] = new Student("李泽贤", "15031165", "15", "TCPIP::192.168.1.114::INSTR", 0);
            for(int i=0;i<15;i++)
            {
                students[i].dt_1 = Convert.ToDateTime("2099/9/9 10:10:10");
                students[i].dt_2 = Convert.ToDateTime("2099/9/9 10:10:10");
                students[i].dt_3 = Convert.ToDateTime("2099/9/9 10:10:10");
            }
            //pictureboxes数组连接相应的picturebox
            pictureboxes[0] = pictureBox1;
            pictureboxes[1] = pictureBox2;
            pictureboxes[2] = pictureBox3;
            pictureboxes[3] = pictureBox4;
            pictureboxes[4] = pictureBox5;
            pictureboxes[5] = pictureBox6;
            pictureboxes[6] = pictureBox7;
            pictureboxes[7] = pictureBox8;
            pictureboxes[8] = pictureBox9;
            pictureboxes[9] = pictureBox10;
            pictureboxes[10] = pictureBox11;
            pictureboxes[11] = pictureBox12;
            pictureboxes[12] = pictureBox13;
            pictureboxes[13] = pictureBox14;
            pictureboxes[14] = pictureBox15;
            //连接label数组连接相应的label
            connectLabels[0] = label2;
            connectLabels[1] = label4;
            connectLabels[2] = label20;
            connectLabels[3] = label21;
            connectLabels[4] = label22;
            connectLabels[5] = label23;
            connectLabels[6] = label24;
            connectLabels[7] = label25;
            connectLabels[8] = label26;
            connectLabels[9] = label27;
            connectLabels[10] = label28;
            connectLabels[11] = label29;
            connectLabels[12] = label30;
            connectLabels[13] = label31;
            connectLabels[14] = label32;
            //提交label数组连接相应的label
            submitLabels[0] = label1;
            submitLabels[1] = label3;
            submitLabels[2] = label7;
            submitLabels[3] = label8;
            submitLabels[4] = label9;
            submitLabels[5] = label10;
            submitLabels[6] = label11;
            submitLabels[7] = label12;
            submitLabels[8] = label13;
            submitLabels[9] = label14;
            submitLabels[10] = label15;
            submitLabels[11] = label16;
            submitLabels[12] = label17;
            submitLabels[13] = label18;
            submitLabels[14] = label19;

            // 初始化计时器：间隔为1秒，绑定Elapsed事件，开启异步线程
            //设置timer可用
            for (int i = 0; i < 15; i++)
            {
                timers[i].Enabled = true;
                timers[i].Interval = 1000;
                timers[i].AutoReset = true;
            }
            timers[0].Enabled = true;
            //设置timer
            //timer1.Interval = 1000;
            //timer2.Interval = 1000;
            //timer2.AutoReset = true;
            //设置是否重复计时，如果该属性设为False,则只执行timer_Elapsed方法一次。
            //timer1.AutoReset = true;
            timers[1 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer1_Elapsed);
            timers[2 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer2_Elapsed);
            timers[3 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer3_Elapsed);
            timers[4 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer4_Elapsed);
            timers[5 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer5_Elapsed);
            timers[6 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer6_Elapsed);
            timers[7 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer7_Elapsed);
            timers[8 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer8_Elapsed);
            timers[9 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer9_Elapsed);
            timers[10 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer10_Elapsed);
            timers[11 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer11_Elapsed);
            timers[12 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer12_Elapsed);
            timers[13 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer13_Elapsed);
            timers[14 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer14_Elapsed);
            timers[15 - 1].Elapsed += new System.Timers.ElapsedEventHandler(timer15_Elapsed);
        }


        //对控件进行调用委托类型和委托方法
        //在列表中写字符串
        //delegate void AppendDelegate(string str);
        //AppendDelegate AppendString;
        //在建立列表时，向下拉列表中添加客户信息
        //delegate void AddDelegate(Student stu);
        //AddDelegate Addfriend;
        //在断开连接时，从下拉列表中删除客户信息
        //delegate void RemoveDelegate(Student stu);
        //RemoveDelegate Removefriend;

        //在列表中写字符串的委托方法
        /*private void AppendMethod(string str)
        {
            listBoxStatu.Items.Add(str);
            listBoxStatu.SelectedIndex = listBoxStatu.Items.Count - 1;
            listBoxStatu.ClearSelected();
        }

        //向下拉列表中添加信息的委托方法
        private void AddMethod(Student stu)
        {
            lock (students)
            {
                students.Add(stu);
            }
            //comboBoxClient.Items.Add(stu.socket.RemoteEndPoint.ToString());
        }

        //从下拉列表中删除信息的委托方法
        private void RemoveMethod(Student stu)
        {
            int i = students.IndexOf(stu);
            //comboBoxClient.Items.RemoveAt(i);
            lock (students)
            {
                students.Remove(stu);
            }
            stu.Dispose();
        }*/

        //加载服务器的方法
        private void FormServer_Load(object sender, EventArgs e)
        {
            //实例化委托对象，与委托方法关联
            //AppendString = new AppendDelegate(AppendMethod);
            //Addfriend = new AddDelegate(AddMethod);
            //Removefriend = new RemoveDelegate(RemoveMethod);

            //获取本机IPv4地址
            string localIp = string.Empty;
            string hostName = Dns.GetHostName();
            IPAddress[] addressList = Dns.GetHostEntry(hostName).AddressList;
            foreach (IPAddress ipAddress in addressList)
            {
                if (ipAddress.AddressFamily == AddressFamily.InterNetwork)
                {
                    localIp = ipAddress.ToString();
                }
            }
            int port = 2000;
            string host = localIp;
            label5.Text = host + ":" + port;
            IPEndPoint localep = new IPEndPoint(IPAddress.Parse(host), port);
            listener = new TcpListener(localep);
            listener.Start(10);
            IsStart = true;

            //接受连接请求的异步调用
            AsyncCallback callback = new AsyncCallback(AcceptCallBack);
            listener.BeginAcceptSocket(callback, listener);
        }


        private void AcceptCallBack(IAsyncResult ar)
        {
            try
            {
                //完成异步接收连接请求的异步调用
                //将连接信息添加到列表和下拉列表中
                Socket handle = listener.EndAcceptSocket(ar);
                Student stu = new Student(handle);
                AsyncCallback callback;
                //继续调用异步方法接收连接请求
                if (IsStart)
                {
                    callback = new AsyncCallback(AcceptCallBack);
                    listener.BeginAcceptSocket(callback, listener);
                }
                //开始在连接上进行异步的数据接收
                stu.ClearBuffer();
                callback = new AsyncCallback(ReceiveCallback);
                stu.socket.BeginReceive(stu.Rcvbuffer, 0, stu.Rcvbuffer.Length, SocketFlags.None, callback, stu);
            }
            catch
            {
                //在调用EndAcceptSocket方法时可能引发异常
                //套接字Listener被关闭，则设置为未启动侦听状态
                IsStart = false;
            }
        }

        private void ReceiveCallback(IAsyncResult ar)
        {
            Student stu = (Student)ar.AsyncState;
            try
            {
                int i = stu.socket.EndReceive(ar);
                if (i == 0)
                {
                    if (connectLabels[Convert.ToInt16(stu.desknum)-1].InvokeRequired)
                    {
                        Action<bool> actionDelegate = (x) => { connectLabels[Convert.ToInt16(stu.desknum)-1].Visible = x; };
                        this.Invoke(actionDelegate, false);
                    }
                    else
                    {
                        connectLabels[Convert.ToInt16(stu.desknum)-1].Visible = false;
                        return;
                    }
                    if (submitLabels[Convert.ToInt16(stu.desknum)-1].InvokeRequired)
                    {
                        Action<bool> actionDelegate = (x) => { submitLabels[Convert.ToInt16(stu.desknum)-1].Visible = x; };
                        this.Invoke(actionDelegate, false);
                    }
                    else
                    {
                        submitLabels[Convert.ToInt16(stu.desknum)-1].Visible = false;
                        return;
                    }
                    //若学生在批改或等待批改的过程中掉线：
                    //1、判断其是否处于待批改状态：是，改成false，出队一个进行窗口处理，isThereUnderchecked改成false；学生的underchecked改成false
                    //否，进行2：将队列全部出队，判断是不是掉线学生，不是，存入一临时队列，是，不入队，向学生发送掉线信息
                    //3、出队然后入队

                }
                else
                {
                    string data = Encoding.UTF8.GetString(stu.Rcvbuffer, 0, i);
                    AsyncCallback callback;
                    
                    if (data.StartsWith("submit1"))
                    {
                        if (isThereUnderchecked == false)      //没有处于待批改状态的学生
                        {
                            //指示现在有处于待批改状态的学生
                            isThereUnderchecked = true;
                            //指示学生处于待批改状态
                            stu.underchecked = true;
                            //对学生对应的label进行调整
                            if (submitLabels[Convert.ToInt32(stu.desknum)-1].InvokeRequired)
                            {
                                Action<bool> actionDelegate = (x) => { submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = x; };
                                this.Invoke(actionDelegate, true);
                            }
                            else
                            {
                                submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = true;
                            }
                            
                        }
                        else HandInSequence.Enqueue(Convert.ToInt32(stu.desknum));         //有处于待批改状态的学生；现在这个进队
                        //向学生发送消息：待批改
                        SendData(stu, "2");

                        //但是进队的学生同样进行数据保存的操作
                        int l = data.Length;
                        int al = Convert.ToInt32(data[l - 4]) - 48;
                        int bl = Convert.ToInt32(data[l - 3]) - 48;
                        int cl = Convert.ToInt32(data[l - 2]) - 48;
                        int dl = Convert.ToInt32(data[l - 1]) - 48;
                        stu.Beta = Convert.ToDouble(data.Substring(7, al));
                        stu.period1_Ue = Convert.ToDouble(data.Substring(7 + al, bl));
                        stu.period1_Ub = Convert.ToDouble(data.Substring(7 + al + bl, cl));
                        stu.period1_Uc = Convert.ToDouble(data.Substring(7 + al + bl + cl, dl));
                        //stu.image_Q=(Bitmap)Image.FromFile("E:\\Study\\Electronic Experiment\\示波器远程控制\\1号同学原始记录\\1.bmp", true);
                        stu.image_Q = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                        
                    }
                    if(data.StartsWith("Image_Q")&&stu.process==0)
                    {
                        sendImage(stu, stu.image_Q);
                    }
                    if (data.StartsWith("submit2"))
                    {
                        if (data.StartsWith("submit21"))
                        {
                            int l = data.Length;
                            int al = Convert.ToInt32(data[l - 3]) - 48;
                            int bl = Convert.ToInt32(data[l - 2]) - 48;
                            int cl = Convert.ToInt32(data[l - 1]) - 48;
                            stu.period2_Ui_1_infinity = Convert.ToDouble(data.Substring(8, al));
                            stu.period2_Us_1_infinity = Convert.ToDouble(data.Substring(8 + al, bl));
                            stu.period2_Uo_1_infinity = Convert.ToDouble(data.Substring(8 + al + bl, cl));
                            //stu.image_1_infinity = (Bitmap)Image.FromFile("E:\\Study\\Electronic Experiment\\示波器远程控制\\1号同学原始记录\\1.bmp", true);
                            stu.image_1_infinity = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                            SendData(stu, "6");
                        }
                        else if (data.StartsWith("submit22"))
                        {
                            int l = data.Length;
                            int al = Convert.ToInt32(data[l - 3]) - 48;
                            int bl = Convert.ToInt32(data[l - 2]) - 48;
                            int cl = Convert.ToInt32(data[l - 1]) - 48;
                            stu.period2_Ui_1_fivedotonek = Convert.ToDouble(data.Substring(8, al));
                            stu.period2_Us_1_fivedotonek = Convert.ToDouble(data.Substring(8 + al, bl));
                            stu.period2_Uo_1_fivedotonek = Convert.ToDouble(data.Substring(8 + al + bl, cl));
                            //stu.image_1_fivedotonek = (Bitmap)Image.FromFile("E:\\Study\\Electronic Experiment\\示波器远程控制\\1号同学原始记录\\1.bmp", true);
                            stu.image_1_fivedotonek = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                            SendData(stu, "7");
                        }
                        else if (data == "submit2")
                        {
                            if (isThereUnderchecked == false)      //没有处于待批改状态的学生
                            {
                                //指示现在有处于待批改状态的学生
                                isThereUnderchecked = true;
                                //指示学生处于待批改状态
                                stu.underchecked = true;
                                //对学生对应的label进行调整
                                if (submitLabels[Convert.ToInt32(stu.desknum)-1].InvokeRequired)
                                {
                                    Action<bool> actionDelegate = (x) => { submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = x; };
                                    this.Invoke(actionDelegate, true);
                                }
                                else
                                {
                                    submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = true;
                                };
                                
                            }
                            else HandInSequence.Enqueue(Convert.ToInt32(stu.desknum));         //有处于待批改状态的学生；现在这个进队
                            //向学生发送消息：待批改
                            SendData(stu, "2");
                        }
                    }
                    if (data.StartsWith("Image_1_Infinity"))
                    {
                        sendImage(stu, stu.image_1_infinity);
                    }
                    if (data.StartsWith("Image_1_Fivedotonek"))
                    {
                        sendImage(stu, stu.image_1_fivedotonek);
                    }
                    if (data.StartsWith("submit3"))
                    {
                        if (data.StartsWith("submit31"))
                        {
                            int l = data.Length;
                            int al = Convert.ToInt32(data[l - 3]) - 48;
                            int bl = Convert.ToInt32(data[l - 2]) - 48;
                            int cl = Convert.ToInt32(data[l - 1]) - 48;
                            stu.period3_Ue_1 = Convert.ToDouble(data.Substring(8, al));
                            stu.period3_Ub_1 = Convert.ToDouble(data.Substring(8 + al, bl));
                            stu.period3_Uc_1 = Convert.ToDouble(data.Substring(8 + al + bl, cl));
                            stu.image_period3_1 = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                            SendData(stu, "6");
                        }
                        else if (data.StartsWith("submit32"))
                        {
                            int l = data.Length;
                            int al = Convert.ToInt32(data[l - 3]) - 48;
                            int bl = Convert.ToInt32(data[l - 2]) - 48;
                            int cl = Convert.ToInt32(data[l - 1]) - 48;
                            stu.period3_Ue_2 = Convert.ToDouble(data.Substring(8, al));
                            stu.period3_Ub_2 = Convert.ToDouble(data.Substring(8 + al, bl));
                            stu.period3_Uc_2 = Convert.ToDouble(data.Substring(8 + al + bl, cl));
                            stu.image_period3_2 = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                            SendData(stu, "7");
                        }
                        else if (data.StartsWith("submit33"))
                        {
                            int l = data.Length;
                            int al = Convert.ToInt32(data[l - 3]) - 48;
                            int bl = Convert.ToInt32(data[l - 2]) - 48;
                            int cl = Convert.ToInt32(data[l - 1]) - 48;
                            stu.period3_Ue_3 = Convert.ToDouble(data.Substring(8, al));
                            stu.period3_Ub_3 = Convert.ToDouble(data.Substring(8 + al, bl));
                            stu.period3_Uc_3 = Convert.ToDouble(data.Substring(8 + al + bl, cl));
                            stu.image_period3_3 = pictureboxes[Convert.ToInt32(stu.desknum) - 1].Image;
                            SendData(stu, "8");
                        }
                        else if (data == "submit3")
                        {
                            if (isThereUnderchecked == false)      //没有处于待批改状态的学生
                            {
                                //指示现在有处于待批改状态的学生
                                isThereUnderchecked = true;
                                //指示学生处于待批改状态
                                stu.underchecked = true;
                                //对学生对应的label进行调整
                                if (submitLabels[Convert.ToInt32(stu.desknum)-1].InvokeRequired)
                                {
                                    Action<bool> actionDelegate = (x) => { submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = x; };
                                    this.Invoke(actionDelegate, true);
                                }
                                else
                                {
                                    submitLabels[Convert.ToInt32(stu.desknum)-1].Visible = true;
                                };
                                
                            }
                            else HandInSequence.Enqueue(Convert.ToInt32(stu.desknum));         //有处于待批改状态的学生；现在这个进队
                            //向学生发送消息：待批改
                            SendData(stu, "2");
                        }
                    }
                    if (data.StartsWith("Image_Period3_1"))
                    {
                        sendImage(stu, stu.image_period3_1);
                    }
                    if (data.StartsWith("Image_Period3_2"))
                    {
                        sendImage(stu, stu.image_period3_2);
                    }
                    if (data.StartsWith("Image_Period3_3"))
                    {
                        sendImage(stu, stu.image_period3_3);
                    }
                    if (data == "finish")
                    {
                        BuildOriginalRecord(stu);
                        BuildSnapshotRecord(stu);
                        //SendFileData(stu, "screenshot");
                        //SendFileData(stu, "report");
                        //SendFileData(stu, "wavedata");

                    }
                    switch (data)
                    {
                        case "1":
                            if (this.label2.InvokeRequired)
                            {
                                students[0].socket = stu.socket;
                                Action<bool> actionDelegate = (x) => { this.label2.Visible = x; };
                                this.Invoke(actionDelegate, true);
                                stu = students[0];
                                SendData(stu, "1");
                            }
                            else
                            {
                                label2.Visible = true;
                            };
                            break;
                        case "2":
                            if (this.label4.InvokeRequired)
                            {
                                students[1].socket = stu.socket;
                                Action<bool> actionDelegate = (x) => { this.label4.Visible = x; };
                                this.Invoke(actionDelegate, true);
                                stu = students[1];
                                SendData(stu, "1");
                            }
                            else
                            {
                                label4.Visible = true;
                            };
                            break;
                        default:
                            break;
                    }

                    if (data.StartsWith("save"))
                    {
                        //先将截图及描述保存在student对象中
                        stu.snapShot[stu.imageCount] = pictureboxes[Convert.ToInt16(stu.desknum) - 1].Image;
                        stu.imageCount++;                      
                        //pictureboxes[Convert.ToInt16(stu.desknum) - 1].Image.Save("E:\\Study\\Electronic Experiment\\示波器远程控制\\" + stu.desknum.ToString() + "号同学原始记录\\temp" + stu.imageCount.ToString() + ".bmp");
                        char n = data[4];
                        switch (n)
                        {
                            case '1':
                                BuildWaveData1(stu, 1);
                                break;
                            case '2':
                                BuildWaveData1(stu, 2);
                                break;
                            default:
                                break;

                        }
                    }
                    //data = string.Format("From[{0}]:{1}", stu.socket.RemoteEndPoint.ToString(), data);
                    //listBoxStatu.Invoke(AppendString, data);
                    stu.ClearBuffer();
                    callback = new AsyncCallback(ReceiveCallback);
                    stu.socket.BeginReceive(stu.Rcvbuffer, 0, stu.Rcvbuffer.Length, SocketFlags.None, callback, stu);

                    if(data.StartsWith("remark"))
                    {
                        stu.description[stu.imageCount-1] = data.Substring(6);//学生输入的描述
                    }
                }
            }
            catch
            {
                //students.Remove(stu);
            }
        
        
        
        }
        private void SendFileData(Student stu,String flag)
        {
            String base64Str;
            String path=null;
            AsyncCallback callback;
            FileInfo EzoneFile;
            FileStream EzoneStream=null;
            int PacketSize;
            int PacketCount;
            int LastDataPacket;
            byte[] data;
            switch (flag)
            {
                case "screenshot":
                    path = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学截图记录\\desert.jpg";
                    EzoneFile = new FileInfo(path);
                    EzoneStream = EzoneFile.OpenRead();
                    PacketSize = 100000;
                    PacketCount = (int)(EzoneStream.Length / ((long)PacketSize));
                    LastDataPacket = (int)(EzoneStream.Length - ((long)(PacketSize * PacketCount)));
                    data = new byte[PacketSize];
                    for (int i = 0; i < PacketCount; i++)
                    {
                        EzoneStream.Read(data, 0, data.Length);
 //                       base64Str = Convert.ToBase64String(data);
//                        data = Encoding.Default.GetBytes(base64Str);
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);

                    }

                    if (LastDataPacket != 0)
                    {
                        data = new byte[LastDataPacket];
                        EzoneStream.Read(data, 0, data.Length);
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);
                    }
                    break;

                case "report":
                    path= "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\OriginalRecord.doc";
                    EzoneFile = new FileInfo(path);
                    EzoneStream = EzoneFile.OpenRead();
                    PacketSize = 100000;
                    PacketCount = (int)(EzoneStream.Length / ((long)PacketSize));
                    LastDataPacket = (int)(EzoneStream.Length - ((long)(PacketSize * PacketCount)));
                    data = new byte[PacketSize];
                    for (int i = 0; i < PacketCount; i++)
                    {
                        EzoneStream.Read(data, 0, data.Length);
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);

                    }

                    if (LastDataPacket != 0)
                    {
                        data = new byte[LastDataPacket];
                        EzoneStream.Read(data, 0, data.Length);
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);
                    }
                    break;
                case "wavedata":
                    for (int j = 1; j < stu.imageCount + 1; j++)
                    {
                        path = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学截图记录\\WaveData" +j + ".txt";
                        EzoneFile = new FileInfo(path);
                        EzoneStream = EzoneFile.OpenRead();
                        PacketSize = 100000;
                        PacketCount = (int)(EzoneStream.Length / ((long)PacketSize));
                        LastDataPacket = (int)(EzoneStream.Length - ((long)(PacketSize * PacketCount)));
                        data = new byte[PacketSize];
                        for (int i = 0; i < PacketCount; i++)
                        {
                            EzoneStream.Read(data, 0, data.Length);
                            callback = new AsyncCallback(SendCallback);
                            stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);

                        }

                        if (LastDataPacket != 0)
                        {
                            data = new byte[LastDataPacket];
                            EzoneStream.Read(data, 0, data.Length);
                            callback = new AsyncCallback(SendCallback);
                            stu.socket.BeginSend(data, 0, data.Length, SocketFlags.None, callback, stu);
                        }
                    }
                    break;
                default:
                    break;

            }
            EzoneStream.Close();
        }
    
        private void SendData(Student stu,String flag)
        {
            try
            {
                byte[] msg  ;
                AsyncCallback callback;
                switch (flag)
                {
                    case "1":
                        String space;
                        if (stu.name.Length == 2)
                            space = "      ";
                        else if (stu.name.Length == 3)
                            space = "     ";
                        else
                            space = "    ";
                        msg = Encoding.UTF8.GetBytes("已连接" +stu.name+space + stu.number);
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    //data = string.Format("To[{0}]:{1}", stu.socket.RemoteEndPoint.ToString(), data);
                    case "2":
                        msg = Encoding.UTF8.GetBytes("待批改");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "3":
                        msg = Encoding.UTF8.GetBytes("通过");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "4":
                        msg = Encoding.UTF8.GetBytes("未通过");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "5":
                        msg = Encoding.UTF8.GetBytes("Image");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "6":
                        msg = Encoding.UTF8.GetBytes("确认一");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "7":
                        msg = Encoding.UTF8.GetBytes("确认二");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;
                    case "8":
                        msg = Encoding.UTF8.GetBytes("确认三");
                        callback = new AsyncCallback(SendCallback);
                        stu.socket.BeginSend(msg, 0, msg.Length, SocketFlags.None, callback, stu);
                        break;

                }
            }
            catch
            {
                //students.Remove(stu);
            }
        }
        private void sendImage(Student student,Image image)
        {
            byte[] imagesend = ImageToBytes(image);
            AsyncCallback callback = new AsyncCallback(SendCallback);
            int len = imagesend.Length;
            student.socket.BeginSend(imagesend, 0, imagesend.Length, SocketFlags.None, callback, student);
            
        }
        private void SendCallback(IAsyncResult ar)
        {
            Student stu = (Student)ar.AsyncState;
            try
            {
                stu.socket.EndSend(ar);
            }
            catch
            {
                //students.Remove(stu);
            }
        }
        //private delegate void RefreshImage(); //代理



        //事件：多线程计时器请求图片
        public void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[1-1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[1-1].RigolIP; //仪器资源名
                                                                          //打开指定资源           
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    /*if (m_VisaOpt.mbSession == null)
                    {
                        timer.Stop();
                        return;
                    }*/
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage1(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[1-1], 0);
            }
        }
        private void ShowImage1(Image img)
        {
            if (this.pictureBox1.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox1.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox1.Image = img;
            }
        }

        public void timer2_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[2-1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[2 - 1].RigolIP; //仪器资源名
                                                                          //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage2(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[2-1], 0);
            }
        }
        private void ShowImage2(Image img)
        {
            if (this.pictureBox2.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox2.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox2.Image = img;
            }
        }
        public void timer3_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[3 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[3 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage3(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[3 - 1], 0);
            }
        }
        private void ShowImage3(Image img)
        {
            if (this.pictureBox3.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox3.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox3.Image = img;
            }
        }
        public void timer4_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[4 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[4 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage4(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[4 - 1], 0);
            }
        }
        private void ShowImage4(Image img)
        {
            if (this.pictureBox4.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox4.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox4.Image = img;
            }
        }
        public void timer5_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[5 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[5 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage5(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[5 - 1], 0);
            }
        }
        private void ShowImage5(Image img)
        {
            if (this.pictureBox5.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox5.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox5.Image = img;
            }
        }
        public void timer6_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[6 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[6 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage6(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[6 - 1], 0);
            }
        }
        private void ShowImage6(Image img)
        {
            if (this.pictureBox6.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox6.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox6.Image = img;
            }
        }
        public void timer7_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[7 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[7 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage7(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[7 - 1], 0);
            }
        }
        private void ShowImage7(Image img)
        {
            if (this.pictureBox7.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox7.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox7.Image = img;
            }
        }
        public void timer8_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[8 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[8 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage8(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[8 - 1], 0);
            }
        }
        private void ShowImage8(Image img)
        {
            if (this.pictureBox8.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox8.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox8.Image = img;
            }
        }
        public void timer9_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[9 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[9 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage9(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[9 - 1], 0);
            }
        }
        private void ShowImage9(Image img)
        {
            if (this.pictureBox9.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox9.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox9.Image = img;
            }
        }
        public void timer10_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[10 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[10 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage10(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[10 - 1], 0);
            }
        }
        private void ShowImage10(Image img)
        {
            if (this.pictureBox10.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox10.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox10.Image = img;
            }
        }
        public void timer11_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[11 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[11 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage11(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[11 - 1], 0);
            }
        }
        private void ShowImage11(Image img)
        {
            if (this.pictureBox11.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox11.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox11.Image = img;
            }
        }
        public void timer12_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[12 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[12 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage12(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[12 - 1], 0);
            }
        }
        private void ShowImage12(Image img)
        {
            if (this.pictureBox12.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox12.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox12.Image = img;
            }
        }
        public void timer13_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[13 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[13 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage13(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[13 - 1], 0);
            }
        }
        private void ShowImage13(Image img)
        {
            if (this.pictureBox13.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox13.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox13.Image = img;
            }
        }
        public void timer14_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[14 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[14 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage14(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[14 - 1], 0);
            }
        }
        private void ShowImage14(Image img)
        {
            if (this.pictureBox14.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox14.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox14.Image = img;
            }
        }
        public void timer15_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref inTimers[15 - 1], 1) == 0)
            {
                //inTimer = 1;
                CVisaOpt m_VisaOpt = new CVisaOpt();
                string m_strResourceName = students[15 - 1].RigolIP; //仪器资源名
                //打开指定资源
                m_VisaOpt.OpenResource(m_strResourceName);
                //发送命令
                m_VisaOpt.Write(":DISPlay:DATA?");
                //读取图片位图数据流
                byte[] bmp = new byte[1152067];
                int cnt = 0;
                byte[] a = new byte[16384];
                for (int j = 0; j < 71; j++)
                {
                    a = m_VisaOpt.ReadByte();
                    int length = a.Length;
                    for (int i = 0; i < length; i++)
                        bmp[16384 * cnt + i] = a[i];
                    cnt++;
                }
                //数据流转换成图片
                byte[] Bmp = new byte[1152000 + 54];
                for (int i = 0; i < 1152000; i++)
                {
                    Bmp[i] = bmp[i + 11];
                }
                int d = Bmp[1152000 - 1 + 54];
                Image image = BytesToImage(Bmp);
                ImageFormat format = image.RawFormat;
                //显示图片
                ShowImage15(image);
                m_VisaOpt.Release();
                //inTimer = 0;
                Interlocked.Exchange(ref inTimers[15 - 1], 0);
            }
        }
        private void ShowImage15(Image img)
        {
            if (this.pictureBox15.InvokeRequired)
            {
                Action<Image> actionDelegate = (x) => { this.pictureBox15.Image = x; };
                this.Invoke(actionDelegate, img);
            }
            else
            {
                pictureBox15.Image = img;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //调用cmd启动apache
            Process p = new Process();
            p.StartInfo.FileName = @"cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
            p.StandardInput.WriteLine("cd /d D:\\Apache\\Apache24\\bin");
            p.StandardInput.AutoFlush = true;
            Thread.Sleep(3000);
            p.StandardInput.WriteLine("httpd -k start");
            Thread.Sleep(3000);
            p.StandardInput.WriteLine("exit");
            p.Close();

            FormServer_Load(sender,e);
            //timers[1-1].Start();
            //timers[2-1].Stop();
        }

        public static byte[] ImageToBytes(Image image)
        {
            ImageFormat format = image.RawFormat;
            using (MemoryStream ms = new MemoryStream())
            {
                if (format.Equals(ImageFormat.Jpeg))
                {
                    image.Save(ms, ImageFormat.Jpeg);
                }
                else if (format.Equals(ImageFormat.Png))
                {
                    image.Save(ms, ImageFormat.Png);
                }
                else if (format.Equals(ImageFormat.Bmp))
                {
                    image.Save(ms, ImageFormat.Bmp);
                }
                else if (format.Equals(ImageFormat.Gif))
                {
                    image.Save(ms, ImageFormat.Gif);
                }
                else if (format.Equals(ImageFormat.Icon))
                {
                    image.Save(ms, ImageFormat.Icon);
                }
                byte[] buffer = new byte[ms.Length];
                //Image.Save()会改变MemoryStream的Position，需要重新Seek到Begin
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buffer, 0, buffer.Length);
                return buffer;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

           

        }
        public static Image BytesToImage(byte[] buffer)
        {
            MemoryStream ms = new MemoryStream(buffer);
            Image image = System.Drawing.Image.FromStream(ms, false);
            return image;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            BuildSequenceRecord();
            //BuildOriginalRecord(students[0]);
            //SendFileData(students[0], "screenshot");
            //关闭timer1，对示波器进行访问，再打开timer1
            //先测试关闭timer1,特别注意要等待elapsd事件完成
            //timers[1 - 1].Stop();
            // while (Interlocked.CompareExchange(ref inTimers[1-1], 0, 1) == 0)
            //{
            //   Thread.Sleep(0);
            // }

            //生成原始点记录
            //访问示波器
            /*CVisaOpt m_VisaOpt = new CVisaOpt();
            string m_strResourceName = students[0].RigolIP; //仪器资源名
            //打开指m定资源
            m_VisaOpt.OpenResource(m_strResourceName);
            //发送命令
            //读通道一纵坐标
            m_VisaOpt.Write(":WAV:SOUR CHAN1");
            m_VisaOpt.Write(":WAV:MODE NORM");
            m_VisaOpt.Write(":WAV:FORM BYTE");
            m_VisaOpt.Write(":WAV:DATA?");

            //读取并显示
            byte[] bytes = m_VisaOpt.ReadByte();
            string[] datay = new string[1400];
            string[] datax = new string[1400];

            m_VisaOpt.Write(":WAV:YREF?");
            string YREF = m_VisaOpt.Read();
            m_VisaOpt.Write(":CHAN1:OFFS?");
            string OFFS1 = m_VisaOpt.Read();
            m_VisaOpt.Write(":CHAN1l:SCAL?");
            string SCALE1 = m_VisaOpt.Read();
            double scale1 = Convert.ToDouble(SCALE1.Substring(0, 8));
            int multi = Convert.ToInt16(SCALE1.Substring(10, 2));
            string sign = SCALE1.Substring(9, 1);
            if (sign == "+") scale1 = (Math.Pow(10, multi)) * scale1;
            else if (sign == "-") scale1 = (Math.Pow(10, 0 - multi)) * scale1;

            //转换
            for (int i = 0; i < 1400; i++)
            {
                //datay[i] = ((Convert.ToInt16(bytes[11 + i]) - 127) / 127 * 5 * scale1).ToString();
                double temp = ((Convert.ToInt16(bytes[11 + i])) - 127);
                datay[i] = (temp / 128 * 5 * scale1).ToString();
            }

            //读横坐标间隔
            m_VisaOpt.Write(":WAV:XINC?");
            string xincrement = m_VisaOpt.Read();
            double xIncrement = Convert.ToDouble(xincrement.Substring(0, 8));
            multi = Convert.ToInt16(xincrement.Substring(10, 2));
            sign = xincrement.Substring(9, 1);
            if (sign == "+") xIncrement = (Math.Pow(10, multi)) * xIncrement;
            else if (sign == "-") xIncrement = (Math.Pow(10, 0 - multi)) * xIncrement;
            for (int i = 0; i < 1400; i++)
            {
                datax[i] = (xIncrement * (i)).ToString();
            }
            //生成txt文档
            FileStream fs1 = new FileStream("E:\\WaveData.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
            StreamWriter sw = new StreamWriter(fs1);
            for (int i = 0; i < 1400; i++)
            {
                sw.WriteLine(datax[i] + "," + datay[i]);
                //sw.WriteLine("\n");//开始写入值
            }
            sw.Close();
            fs1.Close();

            //释放会话
            m_VisaOpt.Release();*/


            //定时器再次打开
            //timers[1 - 1].Start();       

            //BuildOriginalRecord(students[1]);
            //查找仪器资源
            /*string m_strResourceName = null; //仪器资源名
            CVisaOpt m_VisaOpt = new CVisaOpt();
            string[] InstrResourceArray = m_VisaOpt.FindResource("*INSTR?"); //查找资源
            if (InstrResourceArray[0] == "未能找到可用资源!")
            {
                m_strResourceName = "void";
            }
            else
            {
                //示例，选取DSG800系列仪器作为选中仪器
                for (int i = 0; i < InstrResourceArray.Length; i++)
                {

                    if (InstrResourceArray[i].Contains("DS2102A"))
                    {
                        m_strResourceName = InstrResourceArray[i];
                    }
                }
            }
            textBox1.Text = m_strResourceName;*/
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            //Form2 form2 = new Form2();
            //form2.pictureBox1.Image = pictureBox1.Image;

            //打开窗体
            //form2.ShowDialog();
        }

        private void pictureBox1_Layout(object sender, LayoutEventArgs e)
        {

        }

        public void InitializeByProcess(int desknumber, Form2 formToSet)
        {
            switch (students[desknumber].process)
            {
                case 0:
                    {
                        formToSet.label16.Text = "静态工作点测量拍照：";
                        formToSet.pictureBox1.Image = students[desknumber].image_Q;
                        formToSet.label8.Text = "测量数值：";
                        formToSet.label8.Visible = true;
                        formToSet.label9.Text = "Beta:";
                        formToSet.label9.Visible = true;
                        formToSet.label13.Text = students[desknumber].Beta.ToString();
                        formToSet.label13.Visible = true;
                        formToSet.label10.Text = "Ue/V:";
                        formToSet.label10.Visible = true;
                        formToSet.label14.Text = students[desknumber].period1_Ue.ToString();
                        formToSet.label14.Visible = true;
                        formToSet.label11.Text = "Ub/V:";
                        formToSet.label11.Visible = true;
                        formToSet.label15.Text = students[desknumber].period1_Ub.ToString();
                        formToSet.label15.Visible = true;
                        formToSet.label23.Text = "Uc/V:";
                        formToSet.label23.Visible = true;
                        formToSet.label24.Text = students[desknumber].period1_Uc.ToString();
                        formToSet.label24.Visible = true;
                        break;
                    }
                case 1:
                    {
                        formToSet.tempImage[0] = students[desknumber].image_1_infinity;
                        formToSet.tempImage[1] = students[desknumber].image_1_fivedotonek;
                        formToSet.button2.Text = "RL为无穷";
                        formToSet.button2.Visible = true;
                        formToSet.button4.Text = "RL为5.1k欧";
                        formToSet.button4.Visible = true;
                        formToSet.label16.Text = "图像:";
                        formToSet.label8.Text = "RL为无穷的测量值：";
                        formToSet.label8.Visible = true;
                        formToSet.label9.Text = "Ui/mV:";
                        formToSet.label9.Visible = true;
                        formToSet.label13.Text = students[desknumber].period2_Ui_1_infinity.ToString();
                        formToSet.label13.Visible = true;
                        formToSet.label10.Text = "Us/mV:";
                        formToSet.label10.Visible = true;
                        formToSet.label14.Text = students[desknumber].period2_Us_1_infinity.ToString();
                        formToSet.label14.Visible = true;
                        formToSet.label11.Text = "Uo/mV:";
                        formToSet.label11.Visible = true;
                        formToSet.label15.Text = students[desknumber].period2_Uo_1_infinity.ToString();
                        formToSet.label15.Visible = true;
                        formToSet.label12.Text = "RL为5.1k欧的测量值：";
                        formToSet.label12.Visible = true;
                        formToSet.label17.Text = "Ui/mV:";
                        formToSet.label17.Visible = true;
                        formToSet.label18.Text = students[desknumber].period2_Ui_1_fivedotonek.ToString();
                        formToSet.label18.Visible = true;
                        formToSet.label19.Text = "Us/mV:";
                        formToSet.label19.Visible = true;
                        formToSet.label20.Text = students[desknumber].period2_Us_1_fivedotonek.ToString();
                        formToSet.label20.Visible = true;
                        formToSet.label21.Text = "Uo/mV:";
                        formToSet.label21.Visible = true;
                        formToSet.label22.Text = students[desknumber].period2_Uo_1_fivedotonek.ToString();
                        formToSet.label22.Visible = true;
                        break;
                    }
                case 2:
                    {
                        formToSet.tempImage[0] = students[desknumber].image_period3_1;
                        formToSet.tempImage[1] = students[desknumber].image_period3_2;
                        formToSet.tempImage[2] = students[desknumber].image_period3_3;
                        formToSet.button2.Text = "正常失真";
                        formToSet.button2.Visible = true;
                        formToSet.button4.Text = "截止失真";
                        formToSet.button4.Visible = true;
                        formToSet.button5.Text = "饱和失真";
                        formToSet.button5.Visible = true;
                        break;
                    }
            }
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox1.Image;
            form2.label2.Text = students[0].name;
            form2.label4.Text = students[0].number;
            form2.label6.Text = students[0].desknum;

            

            //判断进度，显示相应界面
            if (students[0].underchecked == true) InitializeByProcess(1 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void Form2_sendMessage(string message,String deskNum)
        {
            if (message == "通过")
            {
                //让学生的进度条自加“1”
                students[Convert.ToInt16(deskNum) - 1].process += 1;
                //保存当前学生完成时间
                switch (students[Convert.ToInt16(deskNum) - 1].process)
                {
                    case 1: students[Convert.ToInt16(deskNum) - 1].dt_1 = DateTime.Now;
                        break;
                    case 2: students[Convert.ToInt16(deskNum) - 1].dt_2 = DateTime.Now;
                        break;
                    case 3: students[Convert.ToInt16(deskNum) - 1].dt_3 = DateTime.Now;
                        break;
                }               
                message = "3";
            }
            else if (message == "不通过")
                message = "4";
            //发送数据
            SendData(students[Convert.ToInt16(deskNum) - 1], message);
            //不处于待批改状态
            students[Convert.ToInt16(deskNum) - 1].underchecked = false;
            //将相应的红色提示符去掉
            switch (deskNum)
            {
                case "1": label1.Visible = false; break;
                case "2": label3.Visible = false; break;
                case "3": label7.Visible = false; break;
                case "4": label8.Visible = false; break;
                case "5": label9.Visible = false; break;
                case "6": label10.Visible = false; break;
                case "7": label11.Visible = false; break;
                case "8": label12.Visible = false; break;
                case "9": label13.Visible = false; break;
                case "10": label14.Visible = false; break;
                case "11": label15.Visible = false; break;
                case "12": label16.Visible = false; break;
                case "13": label17.Visible = false; break;
                case "14": label18.Visible = false; break;
                case "15": label19.Visible = false; break;
            }
            //判断队中是否有元素，有，将出队进行处理；
            if (HandInSequence.Count != 0)
            {
                int DESKNUM = HandInSequence.Dequeue();  //出队一个学生
                //指示现在有处于待批改状态的学生
                isThereUnderchecked = true;
                //指示学生处于待批改状态
                students[DESKNUM - 1].underchecked = true;
                //对学生对应的label进行调整
                if (submitLabels[DESKNUM - 1].InvokeRequired)
                {
                    Action<bool> actionDelegate = (x) => { submitLabels[DESKNUM - 1].Visible = x; };
                    this.Invoke(actionDelegate, true);
                }
                else
                {
                    submitLabels[DESKNUM - 1].Visible = true;
                };
               
            }
            else isThereUnderchecked = false;          //没有，将isThereUnderchecked改成false
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox2.Image;
            form2.label2.Text = students[2-1].name;
            form2.label4.Text = students[2-1].number;
            form2.label6.Text = students[2-1].desknum;



            //判断进度，显示相应界面
            if (students[2-1].underchecked == true) InitializeByProcess(2 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox6.Image;
            form2.label2.Text = students[6-1].name;
            form2.label4.Text = students[6 - 1].number;
            form2.label6.Text = students[6 - 1].desknum;



            //判断进度，显示相应界面
            if (students[6 - 1].underchecked == true) InitializeByProcess(6 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox3.Image;
            form2.label2.Text = students[3 - 1].name;
            form2.label4.Text = students[3 - 1].number;
            form2.label6.Text = students[3 - 1].desknum;



            //判断进度，显示相应界面
            if (students[3 - 1].underchecked == true) InitializeByProcess(3 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox4.Image;
            form2.label2.Text = students[4 - 1].name;
            form2.label4.Text = students[4 - 1].number;
            form2.label6.Text = students[4 - 1].desknum;

            //判断进度，显示相应界面
            if (students[4 - 1].underchecked == true) InitializeByProcess(4 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox5.Image;
            form2.label2.Text = students[5 - 1].name;
            form2.label4.Text = students[5 - 1].number;
            form2.label6.Text = students[5 - 1].desknum;

            //判断进度，显示相应界面
            if (students[5 - 1].underchecked == true) InitializeByProcess(5 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox7.Image;
            form2.label2.Text = students[7 - 1].name;
            form2.label4.Text = students[7 - 1].number;
            form2.label6.Text = students[7 - 1].desknum;

            //判断进度，显示相应界面
            if (students[7 - 1].underchecked == true) InitializeByProcess(7 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox8.Image;
            form2.label2.Text = students[8 - 1].name;
            form2.label4.Text = students[8 - 1].number;
            form2.label6.Text = students[8 - 1].desknum;

            //判断进度，显示相应界面
            if (students[8 - 1].underchecked == true) InitializeByProcess(8 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox9.Image;
            form2.label2.Text = students[9 - 1].name;
            form2.label4.Text = students[9 - 1].number;
            form2.label6.Text = students[9 - 1].desknum;

            //判断进度，显示相应界面
            if (students[9 - 1].underchecked == true) InitializeByProcess(9 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox10.Image;
            form2.label2.Text = students[10 - 1].name;
            form2.label4.Text = students[10 - 1].number;
            form2.label6.Text = students[10 - 1].desknum;

            //判断进度，显示相应界面
            if (students[10 - 1].underchecked == true) InitializeByProcess(10 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox11.Image;
            form2.label2.Text = students[11 - 1].name;
            form2.label4.Text = students[11 - 1].number;
            form2.label6.Text = students[11 - 1].desknum;

            //判断进度，显示相应界面
            if (students[11 - 1].underchecked == true) InitializeByProcess(11 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox12.Image;
            form2.label2.Text = students[12 - 1].name;
            form2.label4.Text = students[12 - 1].number;
            form2.label6.Text = students[12 - 1].desknum;

            //判断进度，显示相应界面
            if (students[12 - 1].underchecked == true) InitializeByProcess(12 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox13.Image;
            form2.label2.Text = students[13 - 1].name;
            form2.label4.Text = students[13 - 1].number;
            form2.label6.Text = students[13 - 1].desknum;

            //判断进度，显示相应界面
            if (students[13 - 1].underchecked == true) InitializeByProcess(13 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox14.Image;
            form2.label2.Text = students[14 - 1].name;
            form2.label4.Text = students[14 - 1].number;
            form2.label6.Text = students[14 - 1].desknum;

            //判断进度，显示相应界面
            if (students[14 - 1].underchecked == true) InitializeByProcess(14 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }

        private void pictureBox15_Click(object sender, EventArgs e)
        {
            //初始化：将图片等信息显示在另一个窗体上;将form1上的红框去掉
            Form2 form2 = new Form2();
            form2.pictureBox1.Image = pictureBox15.Image;
            form2.label2.Text = students[15 - 1].name;
            form2.label4.Text = students[15 - 1].number;
            form2.label6.Text = students[15 - 1].desknum;

            //判断进度，显示相应界面
            if (students[15 - 1].underchecked == true) InitializeByProcess(15 - 1, form2);

            //打开窗体
            form2.sendMessage += Form2_sendMessage;
            form2.ShowDialog();
        }
        
        //生成word  （原始数据记录）
        private void BuildOriginalRecord(Student stu)
        {
            object path;//文件路径
            string strContent;//文件内容
            MSWord.Application wordApp;//Word应用程序变量
            MSWord.Document wordDoc;//Word文档变量
            path = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\"+stu.desknum.ToString()+"号同学原始记录\\OriginalRecord.doc";//保存为Word2003文档
            // path = "d:\\myWord.docx";//保存为Word2007文档
            wordApp = new MSWord.Application();//初始化
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            //由于使用的是COM 库，因此有许多变量需要用Missing.Value 代替
            Object Nothing = Missing.Value;
            //新建一个word对象
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //页面设置
            wordDoc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;//设置纸张样式
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;//排列方式为垂直方向
            wordDoc.PageSetup.TopMargin = 57.0f;
            wordDoc.PageSetup.BottomMargin = 57.0f;
            wordDoc.PageSetup.LeftMargin = 57.0f;
            wordDoc.PageSetup.RightMargin = 57.0f;

            //写入文字
            wordApp.Selection.ParagraphFormat.LineSpacing = 15f;//设置文档的行间距
            object unite = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "试验分项一：静态工作点的测量: ";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;

            //
            //移动光标文档末尾
            object count = wordDoc.Paragraphs.Count;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
            wordApp.Selection.TypeParagraph();//插入段落

            //插入表格 2行4列
            int tableRow = 2;
            int tableColumn = 4;
            //定义一个word中的表格对象
            MSWord.Table table = wordDoc.Tables.Add(wordApp.Selection.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
            //填充表格
            table.Cell(1, 1).Range.Text = "Beta";
            table.Cell(1, 2).Range.Text = "Ue/V";
            table.Cell(1, 3).Range.Text = "Ub/V";
            table.Cell(1, 4).Range.Text = "Uc/V";
            table.Cell(2, 1).Range.Text = stu.Beta.ToString();
            table.Cell(2, 2).Range.Text = stu.period1_Ue.ToString();
            table.Cell(2, 3).Range.Text = stu.period1_Ub.ToString();
            table.Cell(2, 4).Range.Text = stu.period1_Uc.ToString();

            //设置表格
            table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            table.Columns[1].Width = 80f;
            table.Columns[2].Width = 80f;
            table.Columns[3].Width = 80f;
            table.Columns[4].Width = 80f;
            wordDoc.Content.InsertAfter("\n");

            //移动光标文档末尾
            //count = wordDoc.Paragraphs.Count;
            //WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            //wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
            //wordApp.Selection.TypeParagraph();//插入段落

            //添加文字2
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "试验分项二：动态参数的测量: ";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;
            wordDoc.Content.InsertAfter("\n");

            //添加表格2 11行5列
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            tableRow = 11;
            tableColumn = 5;
            //定义一个word中的表格对象
            MSWord.Table table_2 = wordDoc.Tables.Add(wordApp.Selection.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
            //填充表格
            for (i = 1; i <= 5; i++)
                for (int j = 1; j <= 11; j++)
                    table_2.Cell(j, i).Range.Font.Size = 12;
            for (i = 2; i <= 11; i++)
            {
                if (i % 2 == 1) table_2.Cell(i, 2).Range.Text = "5.1k";
                else table_2.Cell(i, 2).Range.Text = "∞";
            }
            for (i = 1; i <= 5; i++)
            {
                table_2.Cell(2 * i, 1).Merge(table_2.Cell((2 * i + 1), 1));
            }
            table_2.Cell(2, 1).Range.Text = "f=1kHz";
            table_2.Cell(4, 1).Range.Text = "f=1kHz 电压负反馈";
            table_2.Cell(6, 1).Range.Text = "f=1kHz 电压负反馈 Rs=0";
            table_2.Cell(8, 1).Range.Text = "f=1kHz 电流负反馈";
            table_2.Cell(10, 1).Range.Text = "f=1kHz 电流负反馈 Rs=0";
            table_2.Cell(1, 1).Range.Text = "状态";
            table_2.Cell(1, 2).Range.Text = "RL/Ω";
            table_2.Cell(1, 3).Range.Text = "Ui(max)/mV";
            table_2.Cell(1, 4).Range.Text = "Us(max)/mV";
            table_2.Cell(1, 5).Range.Text = "Uo(max)/V";
            //实验数据
            table_2.Cell(2, 3).Range.Text = stu.period2_Ui_1_infinity.ToString();
            table_2.Cell(2, 4).Range.Text = stu.period2_Us_1_infinity.ToString();
            table_2.Cell(2, 5).Range.Text = stu.period2_Uo_1_infinity.ToString();
            table_2.Cell(3, 3).Range.Text = stu.period2_Ui_1_fivedotonek.ToString();
            table_2.Cell(3, 4).Range.Text = stu.period2_Us_1_fivedotonek.ToString();
            table_2.Cell(3, 5).Range.Text = stu.period2_Uo_1_fivedotonek.ToString();

            //设置表格
            table_2.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table_2.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            table_2.Columns[1].Width = 80f;
            table_2.Columns[2].Width = 80f;
            table_2.Columns[3].Width = 80f;
            table_2.Columns[4].Width = 80f;
            table_2.Columns[5].Width = 80f;
            wordDoc.Content.InsertAfter("\n");

            //插入文字3 插入图片1
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "试验分项三：研究输出波形与静态工作点之间的关系:\n";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;

            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "1、输入信号过大失真的波形:";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;
            wordDoc.Content.InsertAfter("\n");

            //图片
            stu.image_period3_1.Save("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\1.bmp");
            string filename = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\1.bmp";
            //定义要向文档中插入图片的位置
            //移动光标文档末尾
            count = wordDoc.Paragraphs.Count;
            WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
            wordApp.Selection.TypeParagraph();//插入段落
            //定义该图片是否为外部链接
            object linkToFile = false;//默认
            //定义插入的图片是否随word一起保存
            object saveWithDocument = true;
            //向word中写入图片
            object Anchor = wordDoc.Application.Selection.Range;
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref Anchor);
            wordDoc.InlineShapes[1].Height = 240;
            wordDoc.InlineShapes[1].Width = 400;
            wordDoc.Content.InsertAfter("\n");
            
            //文字
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "2、截止失真的波形:";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;
            wordDoc.Content.InsertAfter("\n");
            //图片
            stu.image_period3_2.Save("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\2.bmp");
            filename = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\2.bmp";
            //定义要向文档中插入图片的位置
            //移动光标文档末尾
            count = wordDoc.Paragraphs.Count;
            WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
            wordApp.Selection.TypeParagraph();//插入段落
            //定义该图片是否为外部链接
            linkToFile = false;//默认
            //定义插入的图片是否随word一起保存
            saveWithDocument = true;
            //向word中写入图片
            Anchor = wordDoc.Application.Selection.Range;
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref Anchor);
            wordDoc.InlineShapes[2].Height = 240;
            wordDoc.InlineShapes[2].Width = 400;
            wordDoc.Content.InsertAfter("\n");

            //文字
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "3、饱和失真的波形:";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;
            wordDoc.Content.InsertAfter("\n");
            //图片
            stu.image_period3_3.Save("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\3.bmp");
            filename = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学原始记录\\3.bmp";
            //定义要向文档中插入图片的位置
            //移动光标文档末尾
            count = wordDoc.Paragraphs.Count;
            WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
            wordApp.Selection.TypeParagraph();//插入段落
            //定义该图片是否为外部链接
            linkToFile = false;//默认
            //定义插入的图片是否随word一起保存
            saveWithDocument = true;
            //向word中写入图片
            Anchor = wordDoc.Application.Selection.Range;
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref Anchor);
            wordDoc.InlineShapes[3].Height = 240;
            wordDoc.InlineShapes[3].Width = 400;
            wordDoc.Content.InsertAfter("\n");

            //文字
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
            strContent = "4、相应工作点:";//在文本中使用'\n'换行
            wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
            wordDoc.Paragraphs.Last.Range.Font.Size = 15;
            wordDoc.Paragraphs.Last.Range.Text = strContent;
            wordDoc.Content.InsertAfter("\n");
            //表格
            //插入表格 2行4列
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            tableRow = 4;
            tableColumn = 4;
            //定义一个word中的表格对象
            MSWord.Table table_3 = wordDoc.Tables.Add(wordApp.Selection.Range, tableRow, tableColumn, ref Nothing, ref Nothing);
            //填充表格
            table_3.Cell(1, 1).Range.Text = "失真类型";
            table_3.Cell(2, 1).Range.Text = "正常失真";
            table_3.Cell(3, 1).Range.Text = "截止失真";
            table_3.Cell(4, 1).Range.Text = "饱和失真";
            table_3.Cell(1, 2).Range.Text = "Ue/V";
            table_3.Cell(1, 3).Range.Text = "Ub/V";
            table_3.Cell(1, 4).Range.Text = "Uc/V";
            table_3.Cell(2, 2).Range.Text = stu.period3_Ue_1.ToString();
            table_3.Cell(2, 3).Range.Text = stu.period3_Ub_1.ToString();
            table_3.Cell(2, 4).Range.Text = stu.period3_Uc_1.ToString();
            table_3.Cell(3, 2).Range.Text = stu.period3_Ue_2.ToString();
            table_3.Cell(3, 3).Range.Text = stu.period3_Ub_2.ToString();
            table_3.Cell(3, 4).Range.Text = stu.period3_Uc_2.ToString();
            table_3.Cell(4, 2).Range.Text = stu.period3_Ue_3.ToString();
            table_3.Cell(4, 3).Range.Text = stu.period3_Ub_3.ToString();
            table_3.Cell(4, 4).Range.Text = stu.period3_Uc_3.ToString();
            //设置表格
            table_3.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
            table_3.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            table_3.Columns[1].Width = 80f;
            table_3.Columns[2].Width = 80f;
            table_3.Columns[3].Width = 80f;
            table_3.Columns[4].Width = 80f;
            wordDoc.Content.InsertAfter("\n");

            //WdSaveDocument为Word2003文档的保存格式(文档后缀.doc)\wdFormatDocumentDefault为Word2007的保存格式(文档后缀.docx)
            object format = MSWord.WdSaveFormat.wdFormatDocument;
            //将wordDoc 文档对象的内容保存为DOC 文档,并保存到path指定的路径
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭wordDoc文档
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        //通道1/2的原始数据
        private void BuildWaveData1(Student stu,int chan)
        {
            //生成原始点记录
            //访问示波器
            CVisaOpt m_VisaOpt = new CVisaOpt();
            string m_strResourceName = stu.RigolIP; //仪器资源名
            //打开指m定资源
            m_VisaOpt.OpenResource(m_strResourceName);
            //发送命令
            //读通道纵坐标
            m_VisaOpt.Write(":WAV:SOUR CHAN"+chan.ToString());  //选取通道
            m_VisaOpt.Write(":WAV:MODE NORM");
            m_VisaOpt.Write(":WAV:FORM BYTE");
            m_VisaOpt.Write(":WAV:DATA?");

            //读取并显示
            byte[] bytes = m_VisaOpt.ReadByte();
            string[] datay = new string[1400];
            string[] datax = new string[1400];

            m_VisaOpt.Write(":WAV:YREF?");
            string YREF = m_VisaOpt.Read();
            m_VisaOpt.Write(":CHAN1:OFFS?");
            string OFFS1 = m_VisaOpt.Read();
            m_VisaOpt.Write(":CHAN1l:SCAL?");
            string SCALE1 = m_VisaOpt.Read();
            double scale1 = Convert.ToDouble(SCALE1.Substring(0, 8));
            int multi = Convert.ToInt16(SCALE1.Substring(10, 2));
            string sign = SCALE1.Substring(9, 1);
            if (sign == "+") scale1 = (Math.Pow(10, multi)) * scale1;
            else if (sign == "-") scale1 = (Math.Pow(10, 0 - multi)) * scale1;

            //转换
            for (int i = 0; i < 1400; i++)
            {
                //datay[i] = ((Convert.ToInt16(bytes[11 + i]) - 127) / 127 * 5 * scale1).ToString();
                double temp = ((Convert.ToInt16(bytes[11 + i])) - 127);
                datay[i] = (temp / 128 * 5 * scale1).ToString();
            }

            //读横坐标间隔
            m_VisaOpt.Write(":WAV:XINC?");
            string xincrement = m_VisaOpt.Read();
            double xIncrement = Convert.ToDouble(xincrement.Substring(0, 8));
            multi = Convert.ToInt16(xincrement.Substring(10, 2));
            sign = xincrement.Substring(9, 1);
            if (sign == "+") xIncrement = (Math.Pow(10, multi)) * xIncrement;
            else if (sign == "-") xIncrement = (Math.Pow(10, 0 - multi)) * xIncrement;
            for (int i = 0; i < 1400; i++)
            {
                datax[i] = (xIncrement * (i)).ToString();
            }
            //生成txt文档
            FileStream fs1 = new FileStream("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学截图记录\\WaveData"+stu.imageCount.ToString()+".txt", FileMode.Create, FileAccess.Write);//创建写入文件 
            StreamWriter sw = new StreamWriter(fs1);
            for (int i = 0; i < 1400; i++)
            {
                sw.WriteLine(datax[i] + "," + datay[i]);
                //sw.WriteLine("\n");//开始写入值
            }
            sw.Close();
            fs1.Close();

            //释放会话
            m_VisaOpt.Release();
        }

        //保存截图及相关描述,生成word文档
        private void BuildSnapshotRecord(Student stu)
        {
            object path;//文件路径
            string strContent;//文件内容
            MSWord.Application wordApp;//Word应用程序变量
            MSWord.Document wordDoc;//Word文档变量
            path = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString()+"号同学截图记录\\SnapshotRecord.doc";//保存为Word2003文档
            // path = "d:\\myWord.docx";//保存为Word2007文档
            wordApp = new MSWord.Application();//初始化
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            //由于使用的是COM 库，因此有许多变量需要用Missing.Value 代替
            Object Nothing = Missing.Value;
            //新建一个word对象
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            //页面设置
            wordDoc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;//设置纸张样式
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;//排列方式为垂直方向
            wordDoc.PageSetup.TopMargin = 57.0f;
            wordDoc.PageSetup.BottomMargin = 57.0f;
            wordDoc.PageSetup.LeftMargin = 57.0f;
            wordDoc.PageSetup.RightMargin = 57.0f;
            object unite = Microsoft.Office.Interop.Word.WdUnits.wdStory;

            for (int i = 0; i < stu.imageCount; i++)
            {
                //图片
                stu.snapShot[i].Save("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学截图记录\\Snapshot"+(i+1)+".bmp"); //先将第i个截图保存到本地
                string filename = "D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\" + stu.desknum.ToString() + "号同学截图记录\\Snapshot" + (i+1)+ ".bmp"; //再将其读出来，插入到word文档中
                //定义要向文档中插入图片的位置
                //移动光标文档末尾
                object count = wordDoc.Paragraphs.Count;
                object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
                wordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
                wordApp.Selection.TypeParagraph();//插入段落
                //定义该图片是否为外部链接
                object linkToFile = false;//默认
                //定义插入的图片是否随word一起保存
                object saveWithDocument = true;
                //向word中写入图片
                object Anchor = wordDoc.Application.Selection.Range;
                wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref Anchor);
                //wordDoc.InlineShapes[1].Height = 240;
                //wordDoc.InlineShapes[1].Width = 400;
                wordDoc.Content.InsertAfter("\n");

                //文字
                wordApp.Selection.EndKey(ref unite, ref Nothing);
                wordApp.Selection.ParagraphFormat.FirstLineIndent = 0;//取消首行缩进的长度
                strContent = "图"+(i+1)+"."+stu.description[i];//第i次的截图描述
                //wordDoc.Paragraphs.Last.Range.Font.Name = "黑体";
                wordDoc.Paragraphs.Last.Range.Font.Size = 15;
                wordDoc.Paragraphs.Last.Range.Text = strContent;
                wordDoc.Content.InsertAfter("\n");
            }
            //WdSaveDocument为Word2003文档的保存格式(文档后缀.doc)\wdFormatDocumentDefault为Word2007的保存格式(文档后缀.docx)
            object format = MSWord.WdSaveFormat.wdFormatDocument;
            //将wordDoc 文档对象的内容保存为DOC 文档,并保存到path指定的路径
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭wordDoc文档
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

        }

        //生成完成时间记录先做完的在txt文档前面，排序依据为最后一个实验分项完成时间，这里是dt_3
        private void BuildSequenceRecord()
        {
            int temp;
            //简单的冒泡排序
            for (int i = 14; i > 0; i--)
                for (int j = 0; j < i;j++) {
                    if (DateTime.Compare(students[timeSequence[j]].dt_3, students[timeSequence[j + 1]].dt_3) > 0)    //将完成时间晚的向后调整
                    {
                        temp = timeSequence[j+1];
                        timeSequence[j + 1] = timeSequence[j];
                        timeSequence[j] = temp;
                    }            
                }
            //此时timeSequence数组中存放的是从快到慢完成实验的学生的index

            //生成txt文档
            FileStream fs1 = new FileStream("D:\\Apache\\Apache24\\htdocs\\北航电气实践智能教学系统\\完成时间记录.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
            StreamWriter sw = new StreamWriter(fs1);
            sw.WriteLine("此文档记录学生完成实验的时间点，并对学生完成快慢进行排序，快者在前");
            sw.WriteLine("\n");
            for (int i = 0; i < 15; i++)
            {
                sw.WriteLine((i+1)+"、 "+students[timeSequence[i]].name+","+students[timeSequence[i]].number+","+students[timeSequence[i]].dt_1.ToString()+","+students[timeSequence[i]].dt_2.ToString()+","+students[timeSequence[i]].dt_3.ToString());
                sw.WriteLine("\n");//开始写入值
            }
            sw.Close();
            fs1.Close();
        }
    }
}

