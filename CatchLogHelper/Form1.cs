using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;
using System.Threading;

namespace CatchLogHelper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region 创建全局变量

        //创建数据存放地址相关变量
        public string diskname = "D:\\";
        public string filesavefolder = "";
        public string filesavefolder1 = "";
        public string filesavefolder2 = "";
        public string foldermarkname = "";
        public string filesavepath = "";
        public string currentdate = DateTime.Today.ToString("yyyyMMdd");

        //创建catch tool自己的log地址
        string logfilename = "";

        //创建压缩后的文件目录接收变量
        string filePath = "";

        //创建radioButton对象
        RadioButton room, station, chamber, winrar;

        //创建弹框提示对象,用来接收弹框返回值
        DialogResult dr;

        //创建WinRAR程序地址变量
        string winrar_address1 = @"C:\Program Files\WinRAR";
        string winrar_address2 = @"C:\Program Files (x86)\WinRAR";
        string winrar_address = @"C:\Program Files (x86)\WinRAR";

        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            //获取计算机名称
            string machineName = GetMachineName();
            tssLabel_pc_value.Text = machineName;

            //设置初始车间代号和站别代号
            /*
            for (int i = 0; i < ComputerNameList.Length / 4; i++)
            {
                if (machineName == ComputerNameList[i, 0])
                {

                }
            }
            */
            //将字符串data转换成枚举类型
            try
            {
                enumComputerName enumOne = (enumComputerName)Enum.Parse(typeof(enumComputerName), machineName.Substring(0, 10));

                matchingComputerName(enumOne);//匹配结果
            }
            catch
            {
                //MessageBox.Show("截取计算机名称异常","提示");
            }
            
            


            //获取系统时间
            tssLabel_date_value.Text = System.DateTime.Now.ToString();

            //获取选择的车间
        }

        #region 匹配默认的车间代号和站名

        //使用数组来存储默认的代号
        /* 
        public string[,] ComputerNameList = new string[,] {
         { "C177002030-PC", "701023", "Fiorano", "NearField"},
         { "C17A000383-PC", "701037", "Fiorano", "NearField"},
         { "C181001167-PC", "701050", "Fiorano", "NearField"},
         { "C181001166-PC", "701049", "Fiorano", "NearField"},
         { "C177002028-PC", "701025", "Fiorano", "SFR50CM"},
         { "C177002029-PC", "701041", "Fiorano", "SFR300CM"},
         { "C181001165-PC", "701052", "Fiorano", "SFR300CM"},
		 
         { "C183001067-PC", "702215", "F812", "NearField"},
         { "C183001066-PC", "702214", "F812", "NearField"},
         { "C181001372-PC", "702216", "F812", "NearField"},
         { "C181001371-PC", "702218", "F812", "SFR300CM"},
         { "C183001065-PC", "702217", "F812", "SFR300CM"}};
        */

        //使用枚举类型来记录计算机名称,来匹配默认的代号
        public enum enumComputerName
        {
            //[Fiorano]

			C177002030 = 0,
			C17A000383 = 1,
			C181001167 = 2,
			C181001166 = 3,
			
			C177002028 = 4,
			C177002029 = 5,
			
			C181001165 = 6,
			
			//[F812]
			
			C183001067 = 7,
			C183001066 = 8,
			C181001372 = 9,
			
			C181001371 = 10,
			C183001065 = 11

            //[IQC]

        };

        public void matchingComputerName(enumComputerName pcname)
        {
            switch (pcname)
            {
                case enumComputerName.C177002030:
                    radioButton1.Checked = true;//Fiorano
                    radioButton4.Checked = true;//NearField
                    radioButton7.Checked = true;//1#
                    break;
                case enumComputerName.C17A000383:
                    radioButton1.Checked = true;//Fiorano
                    radioButton4.Checked = true;//NearField
                    radioButton8.Checked = true;//2#
                    break;
                case enumComputerName.C181001167:
                    radioButton1.Checked = true;//Fiorano
                    radioButton4.Checked = true;//NearField
                    radioButton9.Checked = true;//3#
                    break;
                case enumComputerName.C181001166:
                    radioButton1.Checked = true;//Fiorano
                    radioButton4.Checked = true;//NearField
                    radioButton10.Checked = true;//4#
                    break;
                case enumComputerName.C177002028:
                    radioButton1.Checked = true;//Fiorano
                    radioButton5.Checked = true;//SFR50CM
                    radioButton11.Checked = true;//5#
                    break;
                case enumComputerName.C177002029:
                    radioButton1.Checked = true;//Fiorano
                    radioButton6.Checked = true;//SFR300CM
                    radioButton12.Checked = true;//6#
                    break;
                case enumComputerName.C181001165:
                    radioButton1.Checked = true;//Fiorano
                    radioButton6.Checked = true;//SFR300CM
                    radioButton13.Checked = true;//7#
                    break;
                case enumComputerName.C183001067:
                    radioButton2.Checked = true;//F812
                    radioButton4.Checked = true;//NearField
                    radioButton7.Checked = true;//1#
                    break;
                case enumComputerName.C183001066:
                    radioButton2.Checked = true;//F812
                    radioButton4.Checked = true;//NearField
                    radioButton8.Checked = true;//2#
                    break;
                case enumComputerName.C181001372:
                    radioButton2.Checked = true;//F812
                    radioButton4.Checked = true;//NearField
                    radioButton9.Checked = true;//3#
                    break;
                case enumComputerName.C181001371:
                    radioButton2.Checked = true;//F812
                    radioButton6.Checked = true;//SFR300CM
                    radioButton10.Checked = true;//4#
                    break;
                case enumComputerName.C183001065:
                    radioButton2.Checked = true;//F812
                    radioButton6.Checked = true;//SFR300CM
                    radioButton11.Checked = true;//5#
                    break;
                default:
                    break;
            }  
        }
        #endregion


        #region 获取基本信息:计算机名

        //获取计算机名称
        public static string GetMachineName()
        {
            try
            {
                return System.Environment.MachineName;
            }
            catch (Exception e)
            {
                return "uMnNk";
            }
        }

        //获取当前时间
        //定时器1:刷新时间 (每隔一秒钟就把当前时间赋值给label)
        private void timer1_Tick(object sender, EventArgs e)
        {
            tssLabel_date_value.Text = DateTime.Now.ToString();
        }


        #endregion


        #region 选择车间代号、站别、设备编号

        //选择车间1
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked) room = radioButton1;
            tssLabel_room_value.Text = room.Text;
        }
        //选择车间2
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked) room = radioButton2;
            tssLabel_room_value.Text = room.Text;
        }
        //选择车间3
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked) room = radioButton3;
            tssLabel_room_value.Text = room.Text;
        }

        //选择站别1
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked) station = radioButton4;
            tssLabel_station_value.Text = station.Text;
        }
        //选择站别2
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked) station = radioButton5;
            tssLabel_station_value.Text = station.Text;
        }
        //选择站别3
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked) station = radioButton6;
            tssLabel_station_value.Text = station.Text;
        }

        //选择设备编号1
        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked) chamber = radioButton7;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号2
        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked) chamber = radioButton8;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号3
        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked) chamber = radioButton9;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号4
        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked) chamber = radioButton10;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号5
        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton11.Checked) chamber = radioButton11;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号6
        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton12.Checked) chamber = radioButton12;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号7
        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked) chamber = radioButton13;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号8
        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton14.Checked) chamber = radioButton14;
            tssLabel_chamber_value.Text = chamber.Text;
        }
        //选择设备编号9
        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton15.Checked) chamber = radioButton15;
            tssLabel_chamber_value.Text = chamber.Text;
        }

        #endregion


        #region 备份数据前的准备工作

        //判断车间和站别是否已经选取
        public bool IsChooseRoomAndStation()
        {
            if (station == null || room == null || chamber == null)
            {
                dr = MessageBox.Show("请选择车间代号、站别、设备编号！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                return false;
            }
            else
            {
                //创建文件夹
                CreateFolder();
                return true;                
            }
        }

        //点击图片按钮文本框变为只读
        private void pictureBox_disknamebtn_Click(object sender, EventArgs e)
        {
            if (!textBox_diskname.Enabled)
            {
                textBox_diskname.Enabled = true;
                textBox_diskname.BackColor = Color.MediumAquamarine;
                pictureBox_disknamebtn.Image = Properties.Resources.lock2;

                
            }
            else
            {
                //保存前先判断是否含有特殊字符
                string str = textBox_diskname.Text;
                if (str.Contains('*') ||
                    str.Contains('?') ||
                    str.Contains('"') ||
                    str.Contains('<') ||
                    str.Contains('>') ||
                    str.Contains('|')
                    )
                {
                    MessageBox.Show("文件名不能包含以下字符:\r\n" + "*  ?  \"  <  >  |", "提示");
                }
                else
                {
                    textBox_diskname.Enabled = false;
                    textBox_diskname.BackColor = Control.DefaultBackColor;
                    pictureBox_disknamebtn.Image = Properties.Resources.lock1;
                    //文本更改后改变盘符地址
                    diskname = textBox_diskname.Text;
                }


            }
        }


        //创建文件保存地址
        public void CreateFolder()
        {
            //拼接文件夹名称
            filesavefolder = room.Text + "_" + chamber.Text;
            filesavefolder1 = room.Text + "_" + chamber.Text + "_" + currentdate + "_" + foldermarkname;
            filesavefolder2 = room.Text + "_" + chamber.Text + "_" + currentdate + "_" + station.Text;
            filesavepath = @"" + diskname.Trim() + filesavefolder + "\\" + filesavefolder1 + "\\";

            if (Directory.Exists(filesavepath) == false)//如果不存在就创建file文件夹
            {
                try
                {
                    Directory.CreateDirectory(filesavepath);
                }
                catch
                {
                    MessageBox.Show("盘符书写错误或不存在!", "提示");
                }
                
            }
        }

        //选择第三方程序WinRAR安装的路径
        private void radioButton_winraraddress1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_winraraddress1.Checked)
            {
                winrar_address = winrar_address1;
            }
        }
        private void radioButton_winraraddress2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_winraraddress2.Checked)
            {
                winrar_address = winrar_address2;
            }

        }


        #endregion


        #region 执行操作

        //备份按钮
        private void button_backup_Click(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;

            Thread tr = new Thread(testc);
            tr.Start();
        }

        private void testc()
        {
            //设置文件夹名称中的字符
            foldermarkname = "Backup";

            if (IsChooseRoomAndStation())
            {

                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + filesavefolder2 + "_" + @"TestManager.rar C:\TestManager\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteNotes();
                WriteCmdLog();

                //CreateFolder();
            }
            else
            {

            }
        }
        //一键抓取全部7份Log
        private void button_all_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                button1_Click(sender, e);
                button2_Click(sender, e);
                button3_Click(sender, e);
                button4_Click(sender, e);
                button5_Click(sender, e);
                button6_Click(sender, e);
                button7_Click(sender, e);
            }
            else
            {
                if (dr == DialogResult.OK)
                {
                    label_room.BackColor = Color.PaleVioletRed;
                    label_station.BackColor = Color.Plum;
                    label_chamber.BackColor = Color.Pink;
                }
                else
                {
                    return;
                }

            }

        }

        //第01个Log
        private void button1_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"01_SFISlogs.rar C:\TestManager\TM\UI\SFC\SFISlogs\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {
                
            }

            filePath = filesavepath + @"01_SFISlogs.rar";

            if (File.Exists(filePath))
	        {
		        label1.Text = Path.GetFullPath(filePath);
	        }
            else
            {
                label1.Text = "-----";
            }
            
        }

        //第02个Log
        private void button2_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"02_TesterUILog.rar C:\TestManager\TM\UI\TesterUILog\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"02_TesterUILog.rar";

            if (File.Exists(filePath))
            {
                label2.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label2.Text = "-----";
            }
        }

        //第03个Log
        private void button3_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"03_sequencerlog.rar C:\TestManager\TM\log\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"03_sequencerlog.rar";

            if (File.Exists(filePath))
            {
                label3.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label3.Text = "-----";
            }
        }

        //第04个Log
        private void button4_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"04_AutoMachinelog.rar C:\TestManager\log\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"04_AutoMachinelog.rar";

            if (File.Exists(filePath))
            {
                label4.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label4.Text = "-----";
            }
        }

        //第05个Log
        private void button5_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"05_ntlog.rar C:\TestManager\TM\ntlog.log";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"05_ntlog.rar";

            if (File.Exists(filePath))
            {
                label5.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label5.Text = "-----";
            }
        }

        //第06个Log
        private void button6_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"06_Intelli_Log.rar C:\vault\Intelli_Log\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"06_Intelli_Log.rar";

            if (File.Exists(filePath))
            {
                label6.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label6.Text = "-----";
            }
        }

        //第07个Log
        private void button7_Click(object sender, EventArgs e)
        {
            //设置文件夹名称中的字符
            foldermarkname = "Log";

            if (IsChooseRoomAndStation())
            {
                textBox1.Text += DateTime.Now.ToString() + "\r\n----------------------------------\r\n";

                string str1 = "C:";
                string str2 = "cd " + winrar_address;
                string str3 = @"winrar a -ep1 -o+ -inul -r -iback " + filesavepath + @"07_takePhoto_log.rar C:\TestManager\TM\takePhoto\log\";
                string str4 = "exit";
                //textBox1.Text = str1 + "\r\n" + str2 + "\r\n" + str3;
                DoDos(str1, str2, str3, str4);

                WriteCmdLog();
                //CreateFolder();
            }
            else
            {

            }

            filePath = filesavepath + @"07_takePhoto_log.rar";

            if (File.Exists(filePath))
            {
                label7.Text = Path.GetFullPath(filePath);
            }
            else
            {
                label7.Text = "-----";
            }
        }

        #endregion


        //执行cmd命令
        public void DoDos(string comd1, string comd2, string comd3, string comd4)
        {
            string output = null;
            Process p = new Process();//创建进程对象 
            p.StartInfo.FileName = "cmd.exe";//设定需要执行的命令 
            // startInfo.Arguments = "/C " + command;//“/C”表示执行完命令后马上退出  
            p.StartInfo.UseShellExecute = false;//不使用系统外壳程序启动 
            p.StartInfo.RedirectStandardInput = true;//可以重定向输入  
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = false;//不创建窗口 
            p.Start();
            // string comStr = comd1 + "&" + comd2 + "&" + comd3;
            p.StandardInput.WriteLine(comd1);
            p.StandardInput.WriteLine(comd2);
            p.StandardInput.WriteLine(comd3);
            p.StandardInput.WriteLine(comd4);
            output = p.StandardOutput.ReadToEnd();
            textBox1.Text += output;
            textBox1.Text += "\r\n";
            if (p != null)
            {
                p.Close();
            }
            // return output;
        }


        #region 记录和查看

        //保存DOS执行日志到txt
        public void WriteCmdLog()
        {
            string txt = "[" + GetMachineName() + "]" + "\r\n\r\n" + textBox1.Text;
            logfilename = @"" + diskname.Trim() + filesavefolder + "\\" + "catchloglist_" + currentdate + ".log";
            //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    filename = saveFileDialog1.FileName;
            //}
            System.IO.StreamWriter sw = new System.IO.StreamWriter(logfilename);
            sw.Write(txt);
            sw.Flush();
            sw.Close();
        }

        //为程序的备份添加备份说明notes.txt
        public void WriteNotes()
        {
            string notes = "备份说明:(TestManager)" + "\r\n\r\n"
                         + System.DateTime.Now.ToString() + " " + room.Text + " " + chamber.Text + " " + station.Text + "\r\n\r\n"
                         + "备份原因:" + "\r\n\r\n"
                         + "当前状态:" + "\r\n\r\n"
                         + "来自电脑:" + GetMachineName() + "\r\n\r\n"
                         + "byzqy";

            string notesfilepath = filesavepath + "notes.txt";

            System.IO.StreamWriter sw = new System.IO.StreamWriter(notesfilepath);;
            sw.Write(notes);
            sw.Flush();
            sw.Close();
        }

        //打开目标文件夹,查看tool log
        private void button_look_Click(object sender, EventArgs e)
        {

            //System.Diagnostics.Process.Start(@"D:\ZQY");
            //定义一个ProcessStartInfo实例
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //设置启动进程的初始目录
            info.WorkingDirectory = Application.StartupPath;
            //设置启动进程的应用程序或文档名
            info.FileName = logfilename;
            //设置启动进程的参数
            info.Arguments = "";
            //启动由包含进程启动信息的进程资源
            try
            {
                if (File.Exists(logfilename) == true)
                {
                    System.Diagnostics.Process.Start(info);
                }
                else
                {
                    MessageBox.Show("暂无记录!", "提示");
                    return;
                }
            }
            catch (System.ComponentModel.Win32Exception we)
            {
                MessageBox.Show(this, we.Message);
                return;
            }

        }

        //打开备份文件所在文件夹
        private void button_openfolder_Click(object sender, EventArgs e)
        {
            //获取“收藏夹”文件路径
            //string myFavoritesPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
            //string myFavoritesPath = @"D:\";
            //启动进程
            if (Directory.Exists(filesavepath) == true)
            {
                try
                {
                    //打开文件夹
                    System.Diagnostics.Process.Start(filesavepath);
                }
                catch
                {
                    MessageBox.Show("打开文件夹出现异常,请手动打开!", "提示");
                }
            }
            else
            {
                MessageBox.Show("文件夹尚未生成,请执行完备份动作后再打开!", "提示");
            }
        }


        #endregion
        

        #region 窗口控制

        //重启程序按钮(自己重新启动自己)
        private void label_restart_Click(object sender, EventArgs e)
        {
            Application.ExitThread();
            Application.Exit();
            Application.Restart();
            Process.GetCurrentProcess().Kill();
        }
        //关闭程序按钮
        private void label_close_Click(object sender, EventArgs e)
        {
            Application.ExitThread();
            Application.Exit();
        }

        //按下鼠标拖动窗体1
        private void label_titlename_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0);
        }

        //按下鼠标拖动窗体2
        private void panel_titlebar_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0);
        }

        //引用方法:移动(拖动)无边框窗体
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MOVE = 0xF010;
        public const int HTCAPTION = 0x0002;

        //点击任务栏实现最小化与还原
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_MINIMIZEBOX = 0x00020000;
                CreateParams cp = base.CreateParams;
                cp.Style = cp.Style | WS_MINIMIZEBOX;   // 允许最小化操作    
                return cp;
            }
        }

        //添加Panel的Paint事件，用颜色填充Panel区域
        private void panel_titlebar_Paint(object sender, PaintEventArgs e)
        {
            GradientColor(e);
        }

        //抽取成一个方法实现渐变色,在Paint中引用
        private void GradientColor(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Color FColor = Color.Green;
            Color TColor = Color.Yellow;

            Brush b = new LinearGradientBrush(this.ClientRectangle, FColor, TColor, LinearGradientMode.ForwardDiagonal);

            g.FillRectangle(b, this.ClientRectangle);
        }

        //光标进入按钮区域时显示小提示
        private void button_backup_MouseEnter(object sender, EventArgs e)
        {
            ToolTip p = new ToolTip();
            p.ShowAlways = true;
            p.SetToolTip(this.button_backup, "备份前请确认:\r\nTestManager 已经完全关闭!");
        }

        //光标进入Panel上方显示提示
        private void panel_winraraddress_MouseEnter(object sender, EventArgs e)
        {
            ToolTip p = new ToolTip();
            p.ShowAlways = true;
            p.SetToolTip(this.panel_winraraddress, "如果不能执行,\r\n尝试切换到旁边的按钮吧!");
        }

        //为控件倒圆角Panel
        private void panel_winraraddress_Paint(object sender, PaintEventArgs e)
        {
            //Draw(e.ClipRectangle, e.Graphics, 18, false, Color.FromArgb(0, 250, 154), Color.FromArgb(154, 205, 50));
            Draw(e.ClipRectangle, e.Graphics, 18, false, Color.FromArgb(255, 255, 255), Color.FromArgb(224, 224, 224));
            base.OnPaint(e);
            Graphics g = e.Graphics;
            //g.DrawString("Panel", new Font("微软雅黑", 9, FontStyle.Regular), new SolidBrush(Color.White), new PointF(10, 10));
        }
        //为控件倒圆角需要用到的方法
        private void Draw(Rectangle rectangle, Graphics g, int _radius, bool cusp, Color begin_color, Color end_color)
        {
            int span = 2;
            //抗锯齿
            g.SmoothingMode = SmoothingMode.AntiAlias;
            //渐变填充
            LinearGradientBrush myLinearGradientBrush = new LinearGradientBrush(rectangle, begin_color, end_color, LinearGradientMode.Vertical);
            //画尖角
            if (cusp)
            {
                span = 10;
                PointF p1 = new PointF(rectangle.Width - 12, rectangle.Y + 10);
                PointF p2 = new PointF(rectangle.Width - 12, rectangle.Y + 30);
                PointF p3 = new PointF(rectangle.Width, rectangle.Y + 20);
                PointF[] ptsArray = { p1, p2, p3 };
                g.FillPolygon(myLinearGradientBrush, ptsArray);
            }
            //填充
            g.FillPath(myLinearGradientBrush, DrawRoundRect(rectangle.X, rectangle.Y, rectangle.Width - span, rectangle.Height - 1, _radius));
        }
        public static GraphicsPath DrawRoundRect(int x, int y, int width, int height, int radius)
        {
            //四边圆角
            GraphicsPath gp = new GraphicsPath();
            gp.AddArc(x, y, radius, radius, 180, 90);
            gp.AddArc(width - radius, y, radius, radius, 270, 90);
            gp.AddArc(width - radius, height - radius, radius, radius, 0, 90);
            gp.AddArc(x, height - radius, radius, radius, 90, 90);
            gp.CloseAllFigures();
            return gp;
        }

        #endregion


        #region 显示和隐藏版本信息

        //点击窗体左上角图标,显示程序信息
        private void pictureBox_titlelogo_Click(object sender, EventArgs e)
        {
            if (!panel_exeinfo.Visible)
            {
                panel_exeinfo.Visible = true;
                statusStrip1.Visible = false;
            }
            else
            {
                panel_exeinfo.Visible = false;
                statusStrip1.Visible = true;
            }

        }

        //点击按钮隐藏程序说明信息
        private void pictureBox_infoclose_Click(object sender, EventArgs e)
        {
            panel_exeinfo.Visible = false;
            statusStrip1.Visible = true;
        }

        #endregion


        #region 界面小工具

        //打开我的电脑
        private void button_computer_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start("explorer");
            System.Diagnostics.Process.Start("explorer.exe", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
        }
        //打开设备管理器
        private void button_devmgmt_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("devmgmt.msc");
        }

        //双击复制文本到剪贴板
        private void tssLabel_pc_value_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(tssLabel_pc_value.Text);
        }

        private void tssLabel_pc_value_MouseDown(object sender, MouseEventArgs e)
        {
            tssLabel_pc_value.BackColor = System.Drawing.Color.DarkSeaGreen;
        }

        private void tssLabel_pc_value_MouseUp(object sender, MouseEventArgs e)
        {
            //tssLabel_pc_value.BackColor = System.Drawing.SystemColors.Control;
            tssLabel_pc_value.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
        }

        //打开cmd
        private void button_cmd_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("cmd.exe");
        }
        //打开小画家
        private void button_mspaint_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("mspaint.exe");

        }

        //打开记事本
        private void button_notepad_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe");
        }

        #endregion



        
    }
}
