using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SonMakro
{
    public partial class Form2 : Form
    {
        private static IntPtr _hookID = IntPtr.Zero;
        private static bool _firstClickCaptured = false;
        Boolean keyLastState = false;
        public Form2()
        {
            InitializeComponent();
            this.button2.Click += new EventHandler(button2_Click);
        }

        public string MacroName
        {
            get { return textBox1.Text; }
        }

        public string MacroKey
        {
            get { return button1.Text.ToString(); } // Varsayalım ki Makro tuşu butonun Tag'inde saklanıyor
        }

        public string Coordinate
        {
            get { return button2.Text.ToString(); } // Koordinatlar butonun Tag'inde saklanıyor
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.TextLength > 0 && button1.Text != "Set" && button2.Text != "Set") {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please fill the areas");
            }
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (this.ActiveControl == button1)
            {
                button1.Text = e.KeyCode.ToString();
                e.Handled = true;
                if (textBox1.TextLength <= 0 && keyLastState)
                {
                    textBox1.Text = button1.Text;
                }
                button2.PerformClick();
            }
        }

        public void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            _firstClickCaptured = false;
            _hookID = SetHook(HookCallback);
            this.Show();
        }
        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (System.Diagnostics.Process curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (System.Diagnostics.ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        public static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && MouseMessages.WM_LBUTTONDOWN == (MouseMessages)wParam && !_firstClickCaptured)
            {
                _firstClickCaptured = true;
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                string coordinates = $"{hookStruct.pt.x},{hookStruct.pt.y}";

                // Form'u doğru bir şekilde referansla
                Form mainForm = Application.OpenForms["Form2"];
                if (mainForm != null)
                {
                    mainForm.Invoke(new MethodInvoker(() => {
                        Button button2 = mainForm.Controls["button2"] as Button;
                        Button button4 = mainForm.Controls["button4"] as Button;
                        if (button2 != null)
                        {
                            button2.Text = coordinates;  // Button1 metnini güncelle
                            button4.PerformClick();
                        }
                        mainForm.Show();  // Formu tekrar göster

                    }));
                }

                UnhookWindowsHookEx(_hookID); // Hook'u kaldır
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private const int WH_MOUSE_LL = 14;

        private enum MouseMessages
        {
            WM_LBUTTONDOWN = 0x0201
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private void button1_Click(object sender, EventArgs e)
        {
            keyLastState = textBox1.TextLength <= 0;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            button1.PerformClick();
        }
    }
}
