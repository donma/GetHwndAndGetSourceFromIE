using mshtml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetHwndAndGetSourceFromIE
{
    public partial class Form1 : Form
    {
        #region API CALLS

        [DllImport("user32.dll", EntryPoint = "GetClassNameA")]
        public static extern int GetClassName(IntPtr hwnd, StringBuilder lpClassName, int nMaxCount);

        /*delegate to handle EnumChildWindows*/
        public delegate int EnumProc(IntPtr hWnd, ref IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern int EnumChildWindows(IntPtr hWndParent, EnumProc lpEnumFunc, ref IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "RegisterWindowMessageA")]
        public static extern int RegisterWindowMessage(string lpString);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutA")]
        public static extern int SendMessageTimeout(IntPtr hwnd, int msg, int wParam, int lParam, int fuFlags, int uTimeout, out int lpdwResult);
        [DllImport("OLEACC.dll")]
        public static extern int ObjectFromLresult(int lResult, ref Guid riid, int wParam, ref IHTMLDocument2 ppvObject);
        public const int SMTO_ABORTIFHUNG = 0x2;
        public Guid IID_IHTMLDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");


        [DllImport("user32.dll")]
        static extern int EnumWindows(CallbackDef callback, int lParam);

        [DllImport("user32.dll")]
        static extern int GetWindowText(int hWnd, StringBuilder text, int count);

        delegate bool CallbackDef(int hWnd, int lParam);

        #endregion

        public IHTMLDocument2 document;


        public Form1()
        {
            InitializeComponent();
        }

        private bool ShowIEWindowHandler(int hWnd, int lParam)
        {
            string mystring;

            StringBuilder text = new StringBuilder(255);
            GetWindowText(hWnd, text, 255);

            mystring = text.ToString();
            if (mystring.Contains("Internet Explorer"))
            {
                listBox1.Items.Insert(0, text + "," + hWnd);
                
            }
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();


            listBox1.Items.Clear();
            CallbackDef callback = new CallbackDef(ShowIEWindowHandler);
            EnumWindows(callback, 0);

        }

        private IHTMLDocument2 documentFromDOM(IntPtr hWnd)
        {


            int lngMsg = 0;
            int lRes;

            EnumProc proc = new EnumProc(EnumWindows);
            EnumChildWindows(hWnd, proc, ref hWnd);
            if (!hWnd.Equals(IntPtr.Zero))
            {
                lngMsg = RegisterWindowMessage("WM_HTML_GETOBJECT");
                if (lngMsg != 0)
                {
                    SendMessageTimeout(hWnd, lngMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, out lRes);
                    if (!(bool)(lRes == 0))
                    {
                        int hr = ObjectFromLresult(lRes, ref IID_IHTMLDocument, 0, ref document);
                        if ((bool)(document == null))
                        {
                            MessageBox.Show("CANNOT FOUND IHTMLDocument");

                        }
                    }

                }

            }

            return document;
        }

        private int EnumWindows(IntPtr hWnd, ref IntPtr lParam)
        {
            int retVal = 1;
            StringBuilder classname = new StringBuilder(128);
            GetClassName(hWnd, classname, classname.Capacity);
            if ((bool)(string.Compare(classname.ToString(), "Internet Explorer_Server") == 0))
            {
                lParam = hWnd;
                retVal = 0;
            }
            return retVal;
        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtHwnd.Text = listBox1.Text.Split(',')[1];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("Choose hWnd first.");
                return;
            }
            txtSource.Text = "";
            txtTitle.Text = "";

            var document = documentFromDOM(new IntPtr(int.Parse(txtHwnd.Text)));

            if (document != null)
            {
                txtTitle.Text = document.title;
                txtSource.Text = document.body.innerHTML;
            }
            else {
                MessageBox.Show("Sorry I cant fetch it.");
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
