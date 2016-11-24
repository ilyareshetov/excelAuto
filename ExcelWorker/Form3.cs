using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelWorker
{
    public partial class Form3 : Form
    {
        public Dictionary<int, String> index;
        Dictionary<int, String> ttt;
        public string excelPath;
        // Путь к папке с исходными документами
        string documentPath = Application.StartupPath + "\\Data\\";

        public Form3()
        {
            InitializeComponent();
        }

        public Form3(Dictionary<int, String> Selected)
        {
            InitializeComponent();
            index = Selected;

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = documentPath;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                excelPath = openFileDialog1.FileName;
                IWorkbook workbook;
                //string excelPath = @"C:\Users\xancolo\Desktop\Kursach\ExcelWorker\Data\master.xls";
                try
                {
                    using (FileStream file = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    ISheet sheet1 = workbook.GetSheet("КТ");
                    ttt = new Dictionary<int, string>();
                    int i = 9;
                    while (1 == 1)
                    {
                        if (sheet1.GetRow(i) == null)
                        {
                            i++;
                        }
                        if (string.IsNullOrEmpty(sheet1.GetRow(i).GetCell(0).StringCellValue) && (string.IsNullOrEmpty(sheet1.GetRow(i + 1).GetCell(0).StringCellValue)))
                            break;
                        ttt.Add(i, (sheet1.GetRow(i).GetCell(0).StringCellValue) + " " + (sheet1.GetRow(i).GetCell(12).NumericCellValue) + " семестр");
                        i++;
                    }

                    listBox1.DataSource = ttt.Values.ToList<string>();
                    listBox1.Refresh();
                }
                catch (Exception)
                {
                    MessageBox.Show("Выберите соответствующий тип файла!");
                    this.Close();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!(listBox1.SelectedItem == null))
            {
                int t = ttt.First(f => f.Value.ToString() == listBox1.SelectedItem).Key;
                if (!index.ContainsKey(t))
                {
                    listBox2.Items.Add(listBox1.SelectedItem);
                    index.Add(t, ttt[t]);
                    listBox2.Refresh();
                }
                else
                    MessageBox.Show("Выбран тот же предмет!");
            }
            else
                MessageBox.Show("Не выбран элемент для переноса!");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!(listBox2.SelectedItem == null))
            {
                int t = index.First(f => f.Value.ToString() == listBox2.SelectedItem).Key;
                index.Remove(t);
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
                listBox2.Refresh();
            }
            else
                MessageBox.Show("Не выбран элемент для удаления!");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }


}
