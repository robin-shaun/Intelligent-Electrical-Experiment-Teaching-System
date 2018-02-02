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

namespace RigolTest
{
    public delegate void SendMessge(string message,String deskNum);
    public partial class Form2 : Form
    {
        public event SendMessge sendMessage;
        public Form2()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //传消息
            sendMessage(comboBox1.Text,label6.Text);
            //关闭窗体
            Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = tempImage[0];
            label16.Text = button2.Text + "时的示波器图像：";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = tempImage[1];
            label16.Text = button4.Text + "时的示波器图像：";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = tempImage[2];
            label16.Text = button5.Text + "时的示波器图像：";
        }
    }
}
