using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AutomateKey
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        private const int MOUSEEVENTF_Rotate = 0x0800;

        public static void DoMouseClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);


        public static void MoveCursorToPoint(int x, int y)
        {
            SetCursorPos(x, y);
        }

        // <summary>
        /// Struct representing a point.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public static implicit operator System.Drawing.Point(POINT point)
            {
                return new System.Drawing.Point(point.X, point.Y);
            }
        }

        /// <summary>
        /// Retrieves the cursor's position, in screen coordinates.
        /// </summary>
        /// <see>See MSDN documentation for further information.</see>
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        public static System.Drawing.Point GetCursorPosition()
        {
            POINT lpPoint;
            GetCursorPos(out lpPoint);
            //bool success = User32.GetCursorPos(out lpPoint);
            // if (!success)

            return lpPoint;
        }
        int count = 0;
        private void Runbutton_Click(object sender, RoutedEventArgs e)
        {
            try
            {


                int numberofRun = 1;
                int.TryParse(numberofruntxt.Text, out numberofRun);
                count = 0;
                string[] inputtext = Maintext.Text.Split(new string[] { ";", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < numberofRun; i++)
                {
                    for (int j = 0; j < inputtext.Length; j++)
                    {
                        string[] BlocksinLine = inputtext[j].Split(new string[] { ",", "\r" }, StringSplitOptions.RemoveEmptyEntries);
                        if (BlocksinLine.Length > 0)
                        {
                            int numberofrepeat = 1;
                            int.TryParse(BlocksinLine[2], out numberofrepeat);
                            for (int b = 0; b < numberofrepeat; b++)
                                try
                                {
                                    Thread.Sleep(int.Parse(BlocksinLine[1]));
                                    switch (BlocksinLine[0].Trim().ToLower())
                                    {
                                        case "move":
                                            {

                                                MoveCursorToPoint(int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()));
                                                break;
                                            }
                                        case "scroll":
                                            {
                                                mouse_event(MOUSEEVENTF_Rotate, int.Parse(BlocksinLine[3]), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                break;
                                            }
                                        case "leftclick":
                                            {
                                                MoveCursorToPoint(int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()));
                                                mouse_event(MOUSEEVENTF_LEFTDOWN, int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                mouse_event(MOUSEEVENTF_LEFTUP, int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                break;
                                            }
                                        case "leftdown":
                                            {
                                                MoveCursorToPoint(int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()));
                                                mouse_event(MOUSEEVENTF_LEFTDOWN, int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                break;
                                            }
                                        case "leftup":
                                            {
                                                MoveCursorToPoint(int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()));
                                                mouse_event(MOUSEEVENTF_LEFTUP, int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                break;
                                            }
                                        case "rightclick":
                                            {
                                                MoveCursorToPoint(int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()));
                                                mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, int.Parse(BlocksinLine[3].Trim()), int.Parse(BlocksinLine[4].Trim()), 0, 0);
                                                break;
                                            }

                                        case "type":
                                            {
                                                SendKeys.Send(BlocksinLine[3].Trim());
                                                break;
                                            }
                                        case "snapshot":
                                            {
                                                //CopyScreen();
                                                PrintScreen();



                                                break;
                                            }
                                        case "enter":
                                            {
                                                SendKeys.Send("{ENTER}");
                                                break;
                                            }
                                        case "minimize":
                                            {
                                                mainwin.WindowState = WindowState.Minimized;
                                                break;
                                            }
                                        default:
                                            {

                                                break;
                                            }
                                    }
                                }
#if DEBUG
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);

                            }

#else
                                catch
                                {


                                }

#endif
                        }


                    }

                }
            }
#if DEBUG
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);

            }

#else   
                    catch
                        {
                            

                        }
                                        
#endif

        }

        private void CopyScreen()
        {
            count++;
            string nume = "img" + String.Format("{0}", count);
            string format = "Jpeg";
            string director = Directory.GetCurrentDirectory();
            string filePath = director + "\\" + nume + "." + format;

            while (File.Exists(director + "\\" + nume + "." + format))
            {
                count++;
                nume = "img" + String.Format("{0}", count);
                format = "Jpeg";
                director = Directory.GetCurrentDirectory();
                filePath = director + "\\" + nume + "." + format;
            }


            var image = System.Windows.Clipboard.GetImage();
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(image));
                encoder.Save(fileStream);
            }

        }

        private void PrintScreen()

        {
            count++;
            string nume = "img" + String.Format("{0}", count);
            string format = "Jpeg";
            string director = Directory.GetCurrentDirectory();
            string filePath = director + "\\" + nume + "." + format;

            while (File.Exists(director + "\\" + nume + "." + format))
            {
                count++;
                nume = "img" + String.Format("{0}", count);
                format = "Jpeg";
                director = Directory.GetCurrentDirectory();
                filePath = director + "\\" + nume + "." + format;
            }

            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            Graphics graphics = Graphics.FromImage(printscreen as System.Drawing.Image);

            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            printscreen.Save(filePath, ImageFormat.Jpeg);

        }
        private void Browsebutton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            System.Windows.Forms.DialogResult result = opg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                Maintext.Text = File.ReadAllText(opg.FileName);
            }

        }
        private void Savebutton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog opg = new SaveFileDialog();
            System.Windows.Forms.DialogResult result = opg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                File.WriteAllText(opg.FileName, Maintext.Text);

            }
        }
        DispatcherTimer dt = new DispatcherTimer();
        bool isstop = false;
        private void getposbutton_Click(object sender, RoutedEventArgs e)
        {
            isstop = false;
            dt.Interval = new TimeSpan(0, 0, 0, 0, 100);
            dt.Tick += Dt_Tick;
            dt.Start();


        }

        private void Dt_Tick(object sender, EventArgs e)
        {

            POINT po = new POINT();
            GetCursorPos(out po);

            postextBlock.Text = String.Format("position = {0},{1}", po.X, po.Y);
            if (isstop) dt.Stop();
        }



        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Title = "Command,Delay,Number of repeat, the rest, , ;";
            Maintext.Text =
            "leftclick,1000,1,20,40;" + "\n" +
            "scroll,1000,1,0,10;" + "\n" +
            "minimize,1000,1;" + "\n" +
            "move,1000,1,20,20;" + "\n" +
            "leftclick,1000,1,20,20;" + "\n" +
            "leftdown,1000,1,20,20;" + "\n" +
              "leftup,1000,1,20,20;" + "\n" +
            "rightclick,1000,1,20,20;" + "\n" +
            "type,1000,1,hello;" + "\n" +
            "snapshot,1000,1;" + "\n";
        }

        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                isstop = true;
            }
        }
    }
}
