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
using System.Text.RegularExpressions;

namespace ExcelWorker
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            sourceFile = Application.StartupPath + "\\Shablon_Plana\\" + "Шаблон_плана.xls";
        }
        Dictionary<int, String> Selected1= new Dictionary<int, string>();
        Dictionary<int, String> Selected2 = new Dictionary<int, string>();
        Dictionary<int, String> ttt;
        string temp;
        string excelPath1;
        string excelPath2;
        // Путь к папке с исходными документами
        string documentPath = Application.StartupPath + "\\Data\\";
        // Путь к папке выходного документа
        string targetPath = Application.StartupPath + "\\Plan\\";
        // Путь к шаблону индивидуального плана
        string sourceFile = "";

        int i1 = 0, i2 = 0, i3 = 0, i4 = 0;

        // Массивы содержат колонки из основного документа где искать цифры
        int[] index_free = new int[20] { 3, 4, 26, 27, 28, 29, 31, 32, 33, 34, 35, 0, 0, 0, 38, 37, 0, 0, 0, 0 };
        int[] index_pay = new int[20] { 3, 5, 40, 41, 42, 43, 45, 46, 47, 48, 49, 0, 0, 0, 52, 51, 0, 0, 0, 0 };

        IWorkbook workbook_slave = null;
        ISheet sheet1_slave = null;
        ISheet sheet2_slave = null;

        IWorkbook workbook_master = null;
        ISheet sheet1_master = null;


        private void button1_Click(object sender, EventArgs e)
        {
            using (var FirstForm = new Form2(Selected1))
            {
                var result = FirstForm.ShowDialog();
                if (result == DialogResult.OK || result == DialogResult.Cancel)
                {
                    Dictionary<int, String> val = FirstForm.index;
                    this.Selected1 = val;
                    this.excelPath1 = FirstForm.excelPath;
                }
            } 
            
            foreach (int t1 in Selected1.Keys)
            {
                listBox1.Items.Add(Selected1[t1]);
            }
            listBox1.Refresh();         
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (var SecondForm = new Form3(Selected2))
            {
                var result = SecondForm.ShowDialog();
                if (result == DialogResult.OK || result == DialogResult.Cancel)
                {
                    Dictionary<int, String> val = SecondForm.index;
                    this.Selected2 = val;
                    this.excelPath2 = SecondForm.excelPath;
                }
            }

            foreach (int t1 in Selected2.Keys)
            {
                listBox1.Items.Add(Selected2[t1]);
            }
            listBox1.Refresh();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelPath1 != null || excelPath2 != null)
                {
                    string pattern = @"[*|\:""<>?/]";
                    string text = textBox_name.Text;
                    Regex newReg = new Regex(pattern);
                    MatchCollection matches = newReg.Matches(text);
                    if (matches.Count > 0)
                    {
                        MessageBox.Show("Недопустимый символ в имени!");
                    }
                    else
                    {
                        //Создание нового документа на основе шаблона                   
                        string destName = "Индивидуальный план работы преподавателя.xls";
                        if (!string.IsNullOrEmpty(textBox_name.Text))
                            destName = textBox_name.Text + ".xls";
                        string destFile = System.IO.Path.Combine(targetPath, destName);

                        System.IO.File.Copy(sourceFile, destFile, true);

                        //Открытие созданного пустого документа
                        using (FileStream file_slave = new FileStream(destFile, FileMode.Open, FileAccess.ReadWrite))
                        {
                            workbook_slave = new HSSFWorkbook(file_slave);
                        }
                        sheet1_slave = workbook_slave.GetSheet("I1");
                        sheet2_slave = workbook_slave.GetSheet("I1_1");


                        if (excelPath1 != null)
                        {
                            SetRows(Selected1, excelPath1);
                        }

                        if (excelPath2 != null)
                        {
                            SetRows(Selected2, excelPath2);
                        }


                        //Добавление курсовых работ и ВКР
                        //Бюджет Курсовые и ВКР
                        int i_free = 5 + i3;
                        if (numericUpDown1.Value != 0)
                        {
                            AddCourse(numericUpDown1.Value, i_free, 0, workbook_slave, sheet2_slave, 2);
                            i_free++;
                        }
                        if (numericUpDown2.Value != 0)
                            AddCourse(numericUpDown2.Value, 5 + i1, 0, workbook_slave, sheet1_slave, 3);
                        if (numericUpDown3.Value != 0)
                        {
                            AddCourse(numericUpDown3.Value, i_free, 0, workbook_slave, sheet2_slave, 3);
                            i_free++;
                        }
                        if (numericUpDown4.Value != 0)
                        {
                            AddCourse(numericUpDown4.Value, i_free, 1, workbook_slave, sheet2_slave, 4);
                            i_free++;
                        }

                        //Контракт Курсовые и ВКР
                        int i_pay = 15 + i4;
                        if (numericUpDown5.Value != 0)
                        {
                            AddCourse(numericUpDown5.Value, i_pay, 0, workbook_slave, sheet2_slave, 2);
                            i_pay++;
                        }
                        if (numericUpDown6.Value != 0)
                            AddCourse(numericUpDown6.Value, 15 + i2, 0, workbook_slave, sheet1_slave, 3);
                        if (numericUpDown7.Value != 0)
                        {
                            AddCourse(numericUpDown7.Value, i_pay, 0, workbook_slave, sheet2_slave, 3);
                            i_pay++;
                        }
                        if (numericUpDown8.Value != 0)
                        {
                            AddCourse(numericUpDown8.Value, i_pay, 1, workbook_slave, sheet2_slave, 4);
                            i_pay++;
                        }


                        //Добавление строк с формулами
                        //1 семестр
                        FormulaRow(workbook_slave, sheet1_slave, 6, 13, true);  //дневная бюджет
                        FormulaRow(workbook_slave, sheet1_slave, 16, 23, true); //дневная контракт
                        FormulaRow(workbook_slave, sheet1_slave, 14, 24, false);//итого по дневной форме
                        FormulaRow(workbook_slave, sheet1_slave, 27, 34, true); //итого по заочной форме
                        FormulaRow(workbook_slave, sheet1_slave, 24, 35, false);//итого за 1 семестр (контракт)
                        FormulaRow(workbook_slave, sheet1_slave, 14, 36, false);//итого за 1 семестр

                        //2 семестр
                        FormulaRow(workbook_slave, sheet2_slave, 6, 13, true);  //дневная бюджет
                        FormulaRow(workbook_slave, sheet2_slave, 16, 23, true); //дневная контракт
                        FormulaRow(workbook_slave, sheet2_slave, 14, 24, false);//итого по дневной форме
                        FormulaRow(workbook_slave, sheet2_slave, 27, 34, true); //итого по заочной форме
                        FormulaRow(workbook_slave, sheet2_slave, 24, 35, false);//итого за 2 семестр (контракт)
                        FormulaRow(workbook_slave, sheet2_slave, 14, 36, false);//итого за 2 семестр
                        FormulaRow(workbook_slave, sheet2_slave, 14, 37, false, false);//часов по плану за год (бюджет)
                        FormulaRow(workbook_slave, sheet2_slave, 36, 38, false, false);//часов по плану за год (контракт)
                        FormulaRow(workbook_slave, sheet2_slave, 38, 39, false);    //всего по плану за год

                        //Сохранение изменений в документе
                        using (FileStream file_slave = new FileStream(destFile, FileMode.Open, FileAccess.ReadWrite))
                        {
                            workbook_slave.Write(file_slave);
                            file_slave.Close();
                        }

                        //Сообщение об успешном создании документа
                        label5.Visible = true;
                    }
                }
                else
                {
                    MessageBox.Show("Не выбраны документы!");
                }
                
            }
            catch (Exception a)
            {
                MessageBox.Show(a.Message.ToString());
            }
        }

        public void SetRows(Dictionary<int, String> Selected, string excelPath)
        {
            //Открытие расчета учебной нагрузки
            using (FileStream file_master = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                workbook_master = new HSSFWorkbook(file_master);
            }
            sheet1_master = workbook_master.GetSheet("КТ");

            foreach (int t1 in Selected.Keys)
            {
                //Распределяем выбранные предметы по семестрам и форме обучения
                //Если бюджет и 1 семестр
                if ((Convert.ToInt32(sheet1_master.GetRow(t1).GetCell(12).NumericCellValue) % 2 == 1) && (sheet1_master.GetRow(t1).GetCell(39).NumericCellValue != 0))
                {
                    AddRow(t1, sheet1_slave, 5 + i1, sheet1_master, index_free, workbook_slave, Selected);
                    i1++;
                }
                //Если контракт и 1 семестр
                if ((Convert.ToInt32(sheet1_master.GetRow(t1).GetCell(12).NumericCellValue) % 2 == 1) && (sheet1_master.GetRow(t1).GetCell(43).NumericCellValue != 0) || (sheet1_master.GetRow(t1).GetCell(50).NumericCellValue != 0) || (sheet1_master.GetRow(t1).GetCell(51).NumericCellValue != 0))
                {
                    AddRow(t1, sheet1_slave, 15 + i2, sheet1_master, index_pay, workbook_slave, Selected);
                    i2++;
                }
                //Если бюджет и 2 семестр
                if ((Convert.ToInt32(sheet1_master.GetRow(t1).GetCell(12).NumericCellValue) % 2 == 0) && (sheet1_master.GetRow(t1).GetCell(39).NumericCellValue != 0))
                {
                    AddRow(t1, sheet2_slave, 5 + i3, sheet1_master, index_free, workbook_slave, Selected);
                    i3++;
                }
                //Если контракт и 2 семестр
                if ((Convert.ToInt32(sheet1_master.GetRow(t1).GetCell(12).NumericCellValue) % 2 == 0) && (sheet1_master.GetRow(t1).GetCell(43).NumericCellValue != 0) || (sheet1_master.GetRow(t1).GetCell(50).NumericCellValue != 0) || (sheet1_master.GetRow(t1).GetCell(51).NumericCellValue != 0))
                {
                    AddRow(t1, sheet2_slave, 15 + i4, sheet1_master, index_pay, workbook_slave, Selected);
                    i4++;
                }

            }
        }

        public void AddRow(int key, ISheet sheet, int row_number, ISheet master, int[] index, IWorkbook wb, Dictionary<int, String> Selected) 
        {
            //Добавление строки с предметом (здесь заполняются все ячейки одной строки)
            IRow row = sheet.GetRow(row_number);
            ICell cell = row.CreateCell(0);     //Название предмета
            string name = master.GetRow(key).GetCell(0).StringCellValue;
            cell.SetCellValue(Selected[key].Remove(Selected[key].Length-10,10));
            cell.CellStyle = Style(wb, 11);     //Стиль текста
            Border(cell, 1);   //Создание границ ячейки

            //Добавление названия специальности и количества групп
            SetNameGroup(master, sheet, key, row_number, wb);

            //Перенос чисел по предмету из основного документа
            int j = 0;
            for (int i = 3; i <= 22; i++)
            {
                SetCell(master, sheet, key, row_number, i, index[j], wb);
                j++;
            }

            //Добавление крайней колонки с формулами
            SetFormulaColumn(sheet, row_number, wb);
        }

        public void SetNameGroup(ISheet master, ISheet slave, int key, int row_number, IWorkbook wb)
        {
            //Установка значения специальности
            string value = master.GetRow(key).GetCell(2).StringCellValue;
            IRow row = slave.GetRow(row_number);
            ICell cell = row.CreateCell(1);
            string sub = "Информатика";
            string name = "test";
            if (value.Contains(sub))
                name = "ИВТ";
            sub = "Безопасность";

            if (value.Contains(sub))
                name = "БИКС";
            sub = "физика";

            if (value.Contains(sub))
                name = "Физика";

            cell.SetCellValue(name);
            cell.CellStyle = Style(wb, 11);
            Border(cell, 1);

            //Установка значения кол-ва групп
            int count = Convert.ToInt32(master.GetRow(key).GetCell(10).NumericCellValue);
            if (count == 0)
                count = Convert.ToInt32(master.GetRow(key).GetCell(8).NumericCellValue);
            row = slave.GetRow(row_number);
            cell = row.CreateCell(2);
            string res = count + " групп";
            cell.SetCellValue(res);
            cell.CellStyle = Style(wb, 10);
            Border(cell, 1);
        }

        public void SetFormulaColumn(ISheet sheet, int row_number, IWorkbook wb)
        {
            //Создание крайней правой ячейки с формулой
            IRow row = sheet.GetRow(row_number);
            ICell cell = row.CreateCell(23);
            cell.SetCellType(CellType.Formula);
            int i = row_number + 1;
            string formula = "SUM(F" + i + ":W" + i + ")";
            cell.SetCellFormula(formula);
            cell.CellStyle = Style(wb, 10);
            sheet.SetColumnWidth(23, 1500);
            Border(cell, 3);
        }

        public void SetCell(ISheet master, ISheet slave, int key, int row_number, int i, int j, IWorkbook wb)
        {
            //Перенос чисел из ячеек основной таблицы, если они не равны 0
            if (j != 0)
            {
                if (master.GetRow(key).GetCell(j).NumericCellValue != 0)
                {
                    bool fl = false;
                    double value = master.GetRow(key).GetCell(j).NumericCellValue;
                    fl = ((int)value == (float)value ? false : true);
                    IRow row = slave.GetRow(row_number);
                    ICell cell = row.CreateCell(i);
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(value);
                    cell.CellStyle = Style(wb, 10);
                    cell.CellStyle.DataFormat = (fl ? HSSFDataFormat.GetBuiltinFormat("#,##0.00") : HSSFDataFormat.GetBuiltinFormat("#,##0"));
                    slave.AutoSizeColumn(i);
                    Border(cell, 1);
                }
            }           
        }

        public HSSFCellStyle Style(IWorkbook wb, short size)
        {
            //Определение стиля шрифта для содержимого ячеек
            HSSFFont hFont = (HSSFFont)wb.CreateFont();

            hFont.FontHeightInPoints = size;
            hFont.FontName = "Times New Roman";

            HSSFCellStyle hStyle = (HSSFCellStyle)wb.CreateCellStyle();
  
            hStyle.SetFont(hFont);

            return hStyle;
        }

        public void Border(ICell cell, short i)
        {
            switch (i)
            {
                case 1:
                    //Стиль границ ячейки
                    //Стандартное черное выделение
                    cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    break;
                case 2:
                    //Стиль границ ячейки
                    //Верхняя и нижняя границы - жирные
                    cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    break;
                case 3:
                    //Стиль границ ячейки
                    //Правая и левая границы - жирные
                    cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                    break;
                case 4:
                    //Стиль границ ячейки
                    //Вся ячейка жирная
                    cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                    cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                    break;
            }    
        }

        public void FormulaRow(IWorkbook wb, ISheet sheet, int start_row, int end_row, bool flag, bool second = true)
        {
            /* If flag == true then SUM range of columns
             * If flag == false then SUM start_row and end_row 
             * If second == false then SUM previos sheet and current sheet
             * where end_row - position SUM cell and start_row - number of cells witch SUM*/
            IRow row = sheet.GetRow(end_row);
            int[] column;
            column = new int[11] { 5, 7, 8, 9, 10, 11, 12, 13, 14, 17, 23 };
            char[] letter;
            letter = new char[11] { 'F', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'R', 'X' };
            for (int i = 0; i <= 10; i++)
            {
                ICell cell = row.CreateCell(column[i]);
                cell.SetCellType(CellType.Formula);
                string formula = "";
                if (second)
                {
                    if (flag)
                        formula = "SUM(" + letter[i] + start_row + ":" + letter[i] + end_row + ")";
                    else
                        formula = "SUM(" + letter[i] + start_row + "," + letter[i] + end_row + ")";
                }
                else
                    formula = "" + letter[i] + start_row + "+'I1'!" + letter[i] + start_row;    
                cell.SetCellFormula(formula);
                cell.CellStyle = Style(wb, 10);
                sheet.SetColumnWidth(5, 1100);
                sheet.SetColumnWidth(7, 1300);
                if (i == 10)
                    Border(cell, 4);
                else Border(cell, 2);
            }           
        }

        public void AddCourse(decimal count, int lastPos, int i, IWorkbook wb, ISheet sheet, int course)
        {
            string[] str;
            str = new string[2] { "Курсовая работа", "ВКР" };

            IRow row = sheet.GetRow(lastPos);
            ICell cell = row.CreateCell(0);     //Название предмета
            cell.SetCellValue(str[i]);
            cell.CellStyle = Style(wb, 11);     //Стиль текста
            Border(cell, 1);   //Создание границ ячейки

            string name = "ИВТ";
            cell = row.CreateCell(1);
            cell.SetCellValue(name);
            cell.CellStyle = Style(wb, 11);
            Border(cell, 1);

            cell = row.CreateCell(3);
            cell.SetCellValue(course);
            cell.CellStyle = Style(wb, 10);
            Border(cell, 1);

            cell = row.CreateCell(4);
            cell.SetCellValue(Convert.ToInt32(count));
            cell.CellStyle = Style(wb, 10);
            Border(cell, 1);

            
            if (course == 4)
            {
                cell = row.CreateCell(14);
                cell.SetCellValue(Convert.ToInt32(count) * 20);
                sheet.AutoSizeColumn(14);
            }
            else
            {
                cell = row.CreateCell(13);
                cell.SetCellValue(Convert.ToInt32(count) * 3);
                sheet.AutoSizeColumn(13);
            }               
            cell.CellStyle = Style(wb, 10);
            Border(cell, 1);

            SetFormulaColumn(sheet, lastPos, wb);
        }       
    }
}
