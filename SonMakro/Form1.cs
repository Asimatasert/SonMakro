using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;
using Newtonsoft.Json;
using System.IO;
using System.Threading;

namespace SonMakro
{
    public partial class Form1 : Form
    {
        int screenWidth = Screen.PrimaryScreen.Bounds.Width;
        int screenHeight = Screen.PrimaryScreen.Bounds.Height;
        private IKeyboardMouseEvents m_GlobalHook;
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        public Boolean macroIsActive = false;
        public Boolean macroLastState = false;
        public Form1()
        {
            InitializeComponent();
            InitializeDataGridView();
            Subscribe();
        }
        private void SetupFunctionColumn()
        {
            DataGridViewComboBoxColumn funcColumn = new DataGridViewComboBoxColumn();
            funcColumn.HeaderText = "Fonksiyon";
            funcColumn.Name = "Fonksiyon";
            funcColumn.DataPropertyName = "Fonksiyon";  // DataTable'daki ilgili sütun adı
            funcColumn.FlatStyle = FlatStyle.Flat;  // ComboBox stilini ayarla

            // ComboBox'a eklenecek değerler
            funcColumn.Items.AddRange(new string[] { "Click", "Hover" });

            // Eski Fonksiyon sütununu kaldır (varsayılan olarak oluşturulduysa)
            if (dataGridView1.Columns["Fonksiyon"] != null)
            {
                dataGridView1.Columns.Remove("Fonksiyon");
            }

            // Yeni ComboBox sütununu ekle
            dataGridView1.Columns.Add(funcColumn);
        }

        public bool IsKeyValid(string keyString, out Keys key)
        {
            return Enum.TryParse(keyString, out key) && Enum.IsDefined(typeof(Keys), key);
        }
        DataTable dataTable;
        private void InitializeDataGridView()
        {
            dataTable = new DataTable();
            //dataTable.Columns.Add("Makro Adı", typeof(string));
            dataTable.Columns.Add("Makro Tuşu", typeof(string));
            dataTable.Columns.Add("Koordinatlar", typeof(string));
            dataTable.Columns.Add("Fonksiyon", typeof(string));
            //dataTable.Columns.Add("Ekran Çözünürlüğü", typeof(string));
            dataGridView1.DataSource = dataTable;
            SetupFunctionColumn();
        }

        public void Subscribe()
        {
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.KeyPress += GlobalHookKeyPress;
            m_GlobalHook.KeyDown += GlobalKeyDown;
            m_GlobalHook.KeyUp += GlobalKeyUp;
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            char keyChar = char.ToUpper(e.KeyChar);

            if (keyChar == '|')
            {
                makroEkleToolStripMenuItem.PerformClick();
            }

            if (keyChar == 'Æ')
            {
                if (macroIsActive)
                {
                    toolStripMenuItem2.BackColor = Color.Red;
                    this.Text = "Son Macro";
                    int formWidth = this.Width;
                    int formHeight = this.Height;
                    this.Location = new Point((screenWidth - formWidth) / 2, (screenHeight - formHeight) / 2);
                    this.TopMost = false;
                    this.ControlBox = true;
                    this.MinimizeBox = true;
                    macroIsActive = false;
                    macroLastState = false;
                    toolStripMenuItem1.BackColor = Color.Red;
                    this.Size = new Size(321, 600);
                }
                else
                {
                    toolStripMenuItem2.BackColor = Color.Green;
                    this.Text = "";
                    this.TopMost = true;
                    this.Location = new Point(500, 0);
                    this.ControlBox = false;
                    this.MinimizeBox = false;
                    macroIsActive = true;
                    macroLastState = true;
                    toolStripMenuItem1.BackColor = Color.Green;
                    this.Size = new Size(160, 10);
                }
            }
        }


        int asdasd = 0;
        Point oldPozition = Cursor.Position;
        private void GlobalKeyDown(object sender, KeyEventArgs e)
        {
            char keyChar = char.ToUpper((char)e.KeyCode);

            if (macroIsActive)
            {
                foreach (DataRow row in ((DataTable)dataGridView1.DataSource).Rows)
                {
                    string keyString = row["Makro Tuşu"].ToString();
                    string keyStringa = row["Fonksiyon"].ToString();
                    if (keyStringa == "Hold")
                    {
                        if (IsKeyValid(keyString, out Keys key) && (key == e.KeyCode))
                        {
                            Point currentPosition = Cursor.Position;
                            PerformClickAction(currentPosition, keyChar, true);
                        }
                    }
                }
            }
        }
        public static class ScreenHelper
        {
            public static Point GetScreenCenter()
            {
                // Ana ekranı al
                var primaryScreen = Screen.PrimaryScreen;

                // Ekranın genişlik ve yükseklik değerlerini al
                int width = primaryScreen.Bounds.Width;
                int height = primaryScreen.Bounds.Height;

                // Ekranın orta noktasını hesapla
                int centerX = width / 2;
                int centerY = height / 2;

                // Orta noktayı Point olarak döndür
                return new Point(centerX, centerY);
            }
        }

        private void GlobalKeyUp(object sender, KeyEventArgs e)
        {
            char keyChar = char.ToUpper((char)e.KeyCode);

            if (macroIsActive)
            {
                foreach (DataRow row in ((DataTable)dataGridView1.DataSource).Rows)
                {
                    string keyString = row["Makro Tuşu"].ToString();
                    string keyStringa = row["Fonksiyon"].ToString();
                    if (IsKeyValid(keyString, out Keys key) && (key == e.KeyCode))
                    {
                        Point currentPosition = Cursor.Position;
                        if (oldPozition != currentPosition)
                        {
                            oldPozition = Cursor.Position;
                        }
                        if (keyStringa == "Hold")
                        {
                            Point center = ScreenHelper.GetScreenCenter();
                            PerformClickAction(center, keyChar, false);
                        }
                        if (keyStringa == "Click")
                        {
                            PerformClickAction(oldPozition, keyChar, true);
                            PerformClickAction(oldPozition, keyChar, false);
                        }
                    }
                }
            }
        }



        public static void MouseLeftButtonUp(int x, int y)
        {
            SetCursorPos(x, y);  // Fare imlecini istenilen konuma getir (opsiyonel)
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);  // Sol mouse tuşunu serbest bırak
        }
        public static void MouseLeftButtonDown(int x, int y)
        {
            SetCursorPos(x, y);  // Fare imlecini istenilen konuma getir
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);  // Sol mouse tuşunu basılı tut
        }

        private void PerformClickAction(Point d, char mykey, Boolean z)
        {
            if (macroIsActive)
            {
                foreach (DataRow row in ((DataTable)dataGridView1.DataSource).Rows)
                {
                    string keyString = row["Makro Tuşu"].ToString();
                    if (row["Koordinatlar"].ToString() != "Null" && IsKeyValid(keyString, out Keys key) && (key == (Keys)mykey))
                    {
                        string[] coords = row["Koordinatlar"].ToString().Split(',');
                        int x = int.Parse(coords[0]);
                        int y = int.Parse(coords[1]);
                        SetCursorPos(x, y);
                        if (z)
                        {
                            MouseLeftButtonDown(x, y);
                        }
                        else
                        {
                            MouseLeftButtonUp(x, y);
                            SetCursorPos(d.X, d.Y);
                        }
                    }
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Unsubscribe();
        }

        public void Unsubscribe()
        {
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;
            m_GlobalHook.Dispose();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (macroIsActive)
            {
                toolStripMenuItem2.BackColor = Color.Red;
                this.Text = "Son Macro";
                int formWidth = this.Width;
                int formHeight = this.Height;
                this.Location = new Point((screenWidth - formWidth) / 2, (screenHeight - formHeight) / 2);
                this.TopMost = false;
                this.ControlBox = true;
                this.MinimizeBox = true;
                macroIsActive = false;
                macroLastState = false;
                toolStripMenuItem1.BackColor = Color.Red;
                this.Size = new Size(321, 600);
            }
            else
            {
                toolStripMenuItem2.BackColor = Color.Green;
                this.Text = "";
                this.TopMost = true;
                this.Location = new Point(500, 0);
                this.ControlBox = false;
                this.MinimizeBox = false;
                macroIsActive = true;
                macroLastState = true;
                toolStripMenuItem1.BackColor = Color.Green;
                this.Size = new Size(160, 10);
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
                toolStripMenuItem2.BackColor = Color.Red;
            }
            else
            {
                this.TopMost = true;
                toolStripMenuItem2.BackColor = Color.Green;
            }
        }

        private void makroEkleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //macroLastState = macroIsActive;
            macroIsActive = false;
            this.TopMost = true;
            toolStripMenuItem2.BackColor = Color.Green;
            toolStripMenuItem1.BackColor = Color.Red;
            using (Form2 form2 = new Form2())
            {
                if (form2.ShowDialog() == DialogResult.OK)
                {
                    int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                    int screenHeight = Screen.PrimaryScreen.Bounds.Height;
                    string resolution = $"{screenWidth} x {screenHeight}";
                    //dataTable.Rows.Add(form2.MacroName, form2.MacroKey, form2.Coordinate, "-", resolution);
                    dataTable.Rows.Add(form2.MacroKey, form2.Coordinate, "Click");
                    this.TopMost = false;
                    //if (macroLastState)
                    //{
                    //    macroIsActive = macroLastState;
                    //    toolStripMenuItem1.BackColor = Color.Red;
                    //}
                    //else
                    //{
                    //    macroIsActive = macroLastState;
                    //    toolStripMenuItem1.BackColor = Color.Green;
                    //}
                }
            }
        }

        private void importToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = "JSON Files (*.json)|*.json|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    string jsonData = File.ReadAllText(filePath);
                    DataTable tempTable = JsonConvert.DeserializeObject<DataTable>(jsonData);
                    MergeDataTables(dataTable, tempTable);
                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = dataTable;
                }
            }
        }
        private void MergeDataTables(DataTable mainTable, DataTable newTable)
        {
            mainTable.Merge(newTable);
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                dt.Columns.Add(col.Name);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    DataRow dRow = dt.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dRow[cell.ColumnIndex] = cell.Value ?? DBNull.Value;
                    }
                    dt.Rows.Add(dRow);
                }
            }

            string json = JsonConvert.SerializeObject(dt, Formatting.Indented);

            string dosyaYolu = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), toolStripTextBox1.Text + ".json");
            File.WriteAllText(dosyaYolu, json);

            MessageBox.Show($"Veriler kaydedildi: {dosyaYolu}");
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            toolStripTextBox1.SelectAll();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (MessageBox.Show("Bu saatırı silmek istediğinden eminmisin?", "Doğrulama", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    dataGridView1.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "key.json");

            if (File.Exists(filePath))
            {
                try
                {
                    string jsonData = File.ReadAllText(filePath);
                    DataTable tempTable = JsonConvert.DeserializeObject<DataTable>(jsonData);
                    MergeDataTables(dataTable, tempTable);
                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = dataTable;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading JSON: " + ex.Message);
                }
            }
            else
            {
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                dt.Columns.Add(col.Name);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    DataRow dRow = dt.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dRow[cell.ColumnIndex] = cell.Value ?? DBNull.Value;
                    }
                    dt.Rows.Add(dRow);
                }
            }

            string json = JsonConvert.SerializeObject(dt, Formatting.Indented);

            string dosyaYolu = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "key.json");
            File.WriteAllText(dosyaYolu, json);
        }

        private void yardımToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("2-3 tane kısa yol var. \n 1.) ü harfine bas sonra farklı bir tuşa bas sonra ekrandan biyere tıkla makro hızlı kayıt edersin. \n 2.) - Tuşu ile Makrolar Aktif/Pasif olur. \n 3.) * Tuşu ile program Yukarıda/Aşşağıda olur.", ":)");


        }
    }
}
