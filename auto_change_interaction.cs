using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace InteractionStatus
{
    class Program
    {
        public const uint BM_CLICK = 0x00F5;
        public const uint BM_SETSTATE = 0x00F3;
        public const uint CB_SELCHANGE = 0x0001;
        public const uint CB_SETCURSEL = 0x014E;
        public const uint CB_GETLBTEXT = 0x0148;
        public const uint CB_GETCOUNT = 0x0146;
        public const uint WM_SETFOCUS = 0x0007;
        public const uint WM_CHAR = 0x0102;
        public const uint WM_KILLFOCUS = 0x0008;
        public const uint WM_KEYDOWN = 0x0100;
        public const uint WM_KEYUP = 0x0101;
        public const uint WM_MOUSEMOVE = 0x0200;
        public const uint WM_LBUTTONDOWN = 0x0201;
        public const uint WM_LBUTTONUP = 0x0202;

        public const int VK_RETURN = 13;
        public const int VK_ENTER = 0x0D;
        public const int VK_UP = 0x26;  //UP ARROW key
        public const int VK_DOWN = 0x28; //DOWN ARROW key
        public const int VK_RIGHT = 0x27;
        public const int BN_CLICKED = 245;  

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hwnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDlgItem(IntPtr hWnd, int nIDDlgItem);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false, EntryPoint = "SendMessage")]
        public static extern IntPtr SendRefMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static public extern IntPtr GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        
        [DllImport("user32.Dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr parentHandle, Win32Callback callback, IntPtr lParam);

        private const string appName = @"C:\Program Files (x86)\Interactive Intelligence\ICUserApps\InteractionClient.exe";

        public delegate bool Win32Callback(IntPtr hwnd, IntPtr lParam);

        [STAThread]
        static void Main(string[] args)
        {

            FileStream filestream = null;
            StreamWriter streamwriter;
            TextWriter oldOut = Console.Out;
            string filePath = AppDomain.CurrentDomain.BaseDirectory + @"InteractionStatus.txt";
            try
            {
                if (File.Exists(filePath)) filestream = new FileStream(filePath, FileMode.Append);
                else filestream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
                streamwriter = new StreamWriter(filestream);
                streamwriter.AutoFlush = true;
                Console.SetOut(streamwriter);
                Console.WriteLine("-------------------------------------------------------------");
                Console.WriteLine("Interaction Status Change Start : " + DateTime.Now);
                Console.WriteLine("-------------------------------------------------------------");

                if (args.Length == 0)
                {
                    Console.WriteLine(@"Please enter a InteractionClient Status argument. Ex : InteractionStatus.exe ""Abailable""");
                    return;
                }
            }
            catch (Exception e)
            {
                Console.SetOut(oldOut);
                Console.WriteLine("Cannot writin");
                Console.WriteLine(e.Message);
                filestream.Close();
                return;
            }


            string sInteractionStatus = args[0];
            Process[] pname = Process.GetProcessesByName("InteractionClient");
            if (pname.Length == 0) {
                clickLogin(Process.Start(appName));
                System.Threading.Thread.Sleep(5000);
                Console.WriteLine("InteractionClient is not running : Starting again");
            }

            pname = Process.GetProcessesByName("InteractionClient");
            foreach (Process p in pname)
            {
                if (p.MainWindowHandle == IntPtr.Zero)
                {
                    p.Kill();
                }
                else
                {
                    var hwndChild = EnumAllWindows(p.MainWindowHandle, @"WindowsForms10.COMBOBOX.app.0.33c0d9d").ToList<IntPtr>();
                    IntPtr controlId = hwndChild[7];
                    ChangeStatus(controlId, sInteractionStatus);
                    Console.WriteLine("Changed Status : "+GetComboBoxList(controlId).FirstOrDefault().ToString());
                }
                if (sInteractionStatus == "Gone Home" || sInteractionStatus == "At Lunch")
                {
                    p.CloseMainWindow();
                }
            }

            Console.WriteLine("-------------------------------------------------------------");
            streamwriter.Close();
            filestream.Close();
        }

        //Available
        //In a Meeting
        //Cancellation
        //At Lunch
        //Gone Home
        //Programming Room
        //Restroom
        //On Break
        //On Vacation
        //Out of the Office
        //Programming Room
        //Restroom
        //Working At Home
        //Working on Project  
        private static void ChangeStatus(IntPtr controlId, string InteractionStatus)
        {      
            System.Threading.Thread.Sleep(2000);
            SendMessage(controlId, WM_SETFOCUS, 0, 0);
            SendDownKey(controlId, GetComboBoxList(controlId).IndexOf(InteractionStatus));
            SendMessage(controlId, WM_KILLFOCUS, 0, 0);
            System.Threading.Thread.Sleep(4000);
        }

        private static List<string> GetComboBoxList(IntPtr controlId)
        {
            List<string> listCombo = new List<string>();
            IntPtr iCount = SendMessage(controlId, CB_GETCOUNT, 0, 0);
            StringBuilder ssb = new StringBuilder(256, 256);
            for (int i = 0; i < iCount.ToInt32(); i++)
            {
                ssb.Clear();
                SendRefMessage(controlId, CB_GETLBTEXT, i, ssb);
                listCombo.Add(ssb.ToString());
            }
            return listCombo.Distinct().ToList();
        }

        public static void SendDownKey(IntPtr index, int count)
        {
            for (int i = 0; i < count; i++)
            {
                PostMessage(index, WM_KEYDOWN, (IntPtr)VK_DOWN, IntPtr.Zero);
            }
        }

        public static IEnumerable<IntPtr> EnumAllWindows(IntPtr hwnd, string childClassName)
        {
            List<IntPtr> children = GetChildWindows(hwnd);
            if (children == null)
                yield break;
            foreach (IntPtr child in children)
            {
                if (GetWinClass(child) == childClassName)
                    yield return child;
                foreach (var childchild in EnumAllWindows(child, childClassName))
                    yield return childchild;
            }
        }

        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
                throw new InvalidCastException("GCHandle Target could not be cast as List");
            list.Add(handle);
            return true;
        }

        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                Win32Callback childProc = new Win32Callback(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        public static string GetWinClass(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                return null;
            StringBuilder classname = new StringBuilder(100);
            IntPtr result = GetClassName(hwnd, classname, classname.Capacity);
            if (result != IntPtr.Zero)
                return classname.ToString();
            return null;
        }

 
        private static void clickLogin(Process p) 
        {
            IntPtr handle = IntPtr.Zero;
            while (p.MainWindowHandle == IntPtr.Zero)
            {
                System.Threading.Thread.Sleep(100);
                p.Refresh();
            }
            if (p.MainModule.FileName == appName)
            {
                handle = p.MainWindowHandle;
                IntPtr ButtonHandle = FindWindowEx(handle, IntPtr.Zero, null, "C&onnect");
                SendMessage(ButtonHandle, BM_CLICK, 0, 0);
            }
            if (handle == IntPtr.Zero)
            {
                Console.WriteLine("Interaction Client Not found");
            }
        }
    }
}
