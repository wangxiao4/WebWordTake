using CefSharp;
using CefSharp.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OcrLiteLib;
using System.Runtime.InteropServices;
using System.Threading;

namespace WebWordTake
{
    public partial class Form : System.Windows.Forms.Form
    {
        private ChromiumWebBrowser browser;
        private OcrLite ocrEngin;
        private List<string> selectWordList;
        private bool CanRun = false;
        private List<string> WordValue;

        public Form()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 屏幕截图
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        public string screenCut(int x, int y, int width, int height)
        {
            try
            {
                Image img = new Bitmap(width, height);
                //从一个继承自Image类的对象中创建Graphics对象
                Graphics gc = Graphics.FromImage(img);
                //抓屏并拷贝到myimage里
                gc.CopyFromScreen(new Point(x, y), new Point(0, 0), new Size(width, height));
                //this.BackgroundImage = img;
                //保存位图
                string filePath = @"这里是缓存的图片/" + Guid.NewGuid().ToString() + ".jpg";
                img.Save(filePath);
                Console.WriteLine("d", "截图完成！" + this.Location.X + "," + this.Location.Y + "," + width + "," + height);
                return filePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return "";
        }

        /// <summary>
        /// 删除图片
        /// </summary>
        /// <param name="filePath"></param>
        private void DelFileByName(string filePath)
        {
            if (!Properties.Settings.Default.IsCatch)
            {
                File.Delete(filePath);
            }
        }
        /// <summary>
        /// 检查窗体状态
        /// </summary>
        private void CheckState()
        {
            IntPtr mainHandle = FindWindow(null, "页面内容提取");
            if (mainHandle != IntPtr.Zero)
            {
                if (mainHandle != GetForegroundWindow())
                {
                    button.Invoke(new Action(() =>
                    {
                        button.Visible = true;
                        CanRun = false;
                    }));
                }
            }
            else
            {
                MessageBox.Show("获取窗体句柄失败,将无法继续抓取，并且即将退出");
            }
        }

        private void Worker()
        {
            CanRun = true;
            Thread t = new Thread((ThreadStart)(() =>
            {
                for (int i = 0; i < selectWordList.Count;)
                {
                    Thread.Sleep(100);
                    CheckState();
                    if (CanRun)
                    {
                        SelectWord(selectWordList[i]);
                        i++;
                    }
                }
            }));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            File.WriteAllText("这里存放的是结果/" + DateTime.Now.ToString("MMddHHmmss") +".txt", string.Join("\r\n", WordValue));
        }

        private void SelectWord(string word)
        {
            System.Windows.Forms.Clipboard.SetText(word);
            System.Windows.Forms.SendKeys.SendWait("^{V}");
            Thread.Sleep(Properties.Settings.Default.SelectSpeed);

            string filePath = screenCut(this.Location.X, this.Location.Y + 100, this.Width, this.Height - 120);
            if (!string.IsNullOrEmpty(filePath))
            {
                OcrResult ocrResult = ocrEngin.Detect(filePath, 50, 1024, 0.618F, 0.3F, 2, true, true);

                StringBuilder sb = new StringBuilder();
                foreach (var block in ocrResult.TextBlocks)
                {
                    sb.Append(block.Text + ",");
                }

                WordValue.Add(word + "->" + sb.ToString());
                DelFileByName(filePath);
            }

            System.Windows.Forms.SendKeys.SendWait("^{A}");
            System.Windows.Forms.SendKeys.SendWait("{DEL}");
            for (int i=0; i< word.Length;i++)
            {
                System.Windows.Forms.SendKeys.SendWait("{BACKSPACE}");
            }

        }

        private void Form_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("这里是缓存的图片"))
            {
                Directory.CreateDirectory("这里是缓存的图片");
            }
            if (!Directory.Exists("这里存放的是结果"))
            {
                Directory.CreateDirectory("这里存放的是结果");
            }
            if (!Directory.Exists("需要查询的文件放这里"))
            {
                Directory.CreateDirectory("需要查询的文件放这里");
            }

            MessageBox.Show("初始化前请尽量  不要  挪动  鼠标");

            InitSelectData();
            OcrLiteInit();
            browser = new ChromiumWebBrowser(Properties.Settings.Default.BaseURL);
            browser.LoadingStateChanged += Browser_LoadingStateChanged;
            panel.Controls.Add(browser);
            browser.Dock = DockStyle.Fill;
        }

        private void InitSelectData()
        {
            WordValue = new List<string>();
            selectWordList = new List<string>();
            string[] files = Directory.GetFiles("需要查询的文件放这里", "*.txt", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                selectWordList.AddRange(File.ReadAllLines(file));
            }
        }

        /// <summary>
        /// 浏览器状态刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Browser_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (e.CanReload)
            {
                bool step1 = false;
                //保存图片
                string filePath = screenCut(this.Location.X, this.Location.Y, this.Width, 300);
                if (!string.IsNullOrEmpty(filePath))
                {
                    OcrResult ocrResult = ocrEngin.Detect(filePath, 50, 1024, 0.618F, 0.3F, 2, true, true);
                    foreach (var block in ocrResult.TextBlocks)
                    {
                        if (block.Text.Equals("百度一下"))
                        {
                            int x = this.Location.X + block.BoxPoints[0].X + Properties.Settings.Default.FloatValueX;
                            int y = this.Location.Y + block.BoxPoints[0].Y + Properties.Settings.Default.FloatValueY;
                            MoveMouseToPoint(new Point(x, y));
                            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                            step1 = true;
                        }
                    }
                    DelFileByName(filePath);
                }
                //检测是否以准备好
                if (step1)
                {
                    Task.Run(() =>
                    {
                        Thread.Sleep(1000);
                        filePath = screenCut(this.Location.X, this.Location.Y, this.Width, 300);
                        if (!string.IsNullOrEmpty(filePath))
                        {
                            OcrResult ocrResult = ocrEngin.Detect(filePath, 50, 1024, 0.618F, 0.3F, 2, true, true);
                            bool step2 = false;
                            foreach (var block in ocrResult.TextBlocks)
                            {
                                if (block.Text.Equals("取消")|| block.Text.Equals("百度一下") || block.Text.Equals("登陆查看历史"))
                                {
                                    step2 = true;
                                    MessageBoxButtons messButton = MessageBoxButtons.OKCancel;
                                    DialogResult dr = MessageBox.Show("初始化完成 是否开始抓取?", "信息确认", messButton);
                                    if (dr == DialogResult.OK)
                                    {
                                        Worker();
                                    }
                                }
                            }
                            if (!step2)
                            {
                                MessageBox.Show("初始化步骤2-》未能寻找到指定位置，请修改配置参数后重新启动");
                            }
                            DelFileByName(filePath);
                        }
                    });
                }
                else
                {
                    MessageBox.Show("初始化步骤1-》未能寻找到指定位置，请修改配置参数后重新启动");
                }
            }
        }

        /// <summary>
        /// 光学扫描初始化
        /// </summary>
        private void OcrLiteInit()
        {
            string appPath = AppDomain.CurrentDomain.BaseDirectory;
            //string appDir = Directory.GetParent(appPath).FullName;
            string modelsDir = appPath + "models";
            string detPath = modelsDir + "\\dbnet.onnx";
            string clsPath = modelsDir + "\\angle_net.onnx";
            string recPath = modelsDir + "\\crnn_lite_lstm.onnx";
            string keysPath = modelsDir + "\\keys.txt";
            bool isDetExists = File.Exists(detPath);
            if (!isDetExists)
            {
                MessageBox.Show("模型文件不存在:" + detPath);
            }
            bool isClsExists = File.Exists(clsPath);
            if (!isClsExists)
            {
                MessageBox.Show("模型文件不存在:" + clsPath);
            }
            bool isRecExists = File.Exists(recPath);
            if (!isRecExists)
            {
                MessageBox.Show("模型文件不存在:" + recPath);
            }
            bool isKeysExists = File.Exists(recPath);
            if (!isKeysExists)
            {
                MessageBox.Show("Keys文件不存在:" + keysPath);
            }
            if (isDetExists && isClsExists && isRecExists && isKeysExists)
            {
                ocrEngin = new OcrLite();
                ocrEngin.InitModels(detPath, clsPath, recPath, keysPath, Properties.Settings.Default.OCRThreadCount);
            }
            else
            {
                MessageBox.Show("初始化失败，请确认模型文件夹和文件后，重新初始化！");
            }
        }

        /// <summary>
        /// 引用user32.dll动态链接库（windows api），
        /// 使用库中定义 API：SetCursorPos 
        /// </summary>
        [DllImport("user32.dll")]
        private static extern int SetCursorPos(int x, int y);
        /// <summary>
        /// 移动鼠标到指定的坐标点
        /// </summary>
        public void MoveMouseToPoint(Point p)
        {
            SetCursorPos(p.X, p.Y);
        }
        /// <summary>
        /// 设置鼠标的移动范围
        /// </summary>
        public void SetMouseRectangle(Rectangle rectangle)
        {
            System.Windows.Forms.Cursor.Clip = rectangle;
        }
        /// <summary>
        /// 设置鼠标位于屏幕中心
        /// </summary>
        public void SetMouseAtCenterScreen()
        {
            //当前屏幕的宽高
            int winHeight = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height;
            int winWidth = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            //设置鼠标的x，y位置
            int loginx = winWidth / 2;
            int loginy = winHeight / 2;
            Point centerP = new Point(loginx, loginy);
            //移动鼠标
            MoveMouseToPoint(centerP);
        }
        //点击事件
        [DllImport("User32")]
        //下面这一行对应着下面的点击事件
        //    public extern static void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);
        public extern static void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        public const int MOUSEEVENTF_LEFTDOWN = 0x2;
        public const int MOUSEEVENTF_LEFTUP = 0x4;
        public enum MouseEventFlags
        {
            Move = 0x0001, //移动鼠标
            LeftDown = 0x0002,//模拟鼠标左键按下
            LeftUp = 0x0004,//模拟鼠标左键抬起
            RightDown = 0x0008,//鼠标右键按下
            RightUp = 0x0010,//鼠标右键抬起
            MiddleDown = 0x0020,//鼠标中键按下 
            MiddleUp = 0x0040,//中键抬起
            Wheel = 0x0800,
            Absolute = 0x8000//标示是否采用绝对坐标
        }

        private void button_Click(object sender, EventArgs e)
        {
            button.Visible = false;
            CanRun = true;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        /// <summary>
        /// 获取窗体的句柄函数
        /// </summary>
        /// <param name="lpClassName">窗口类名</param>
        /// <param name="lpWindowName">窗口标题名</param>
        /// <returns>返回句柄</returns>
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    }
}
