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
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ExcellParseWindowsForm
{
    public partial class Form1 : Form
    {

        String dir_parse;
        String dir_templ;
        String stop_cycle = "no";

        public Form1()
        {
            InitializeComponent();
            textBox1.Text = Directory.GetCurrentDirectory() + "\\data_excells";
            textBox2.Text = Directory.GetCurrentDirectory() + "\\template.xlsx";

            dir_parse = textBox1.Text;
            dir_templ = textBox2.Text;
            label3.Text = "";

            progressBar1.Visible = false;

            String[,] param = {
                // Лист 1
                { "Номер платежного документа", "1", "I1", "C", "0" },
                { "Расчетный период", "1", "D21", "D", "0" },
                { "Общая площадь", "1", "F26", "E", "0" },
                { "Отапливаемая площадь", "1", "I26", "G", "0" },
                { "Количество проживающих", "1", "M26", "H", "0" },
                { "Сумма к оплате за расчетный период", "1", "N14", "X", "0" },
                { "БИК банка", "1", "046015602", "P", "0" }, // "J3"
                { "Расчетный счет", "1", "40703810152090101552", "Q", "0" },   // "I2"
                
                // Лист 2
                { "Техобслуживание (Наименование)", "2", "A42", "B", "0" },
                { "     Техобслуживание (тариф)", "2", "G42", "G", "0"  },
                { "     Техобслуживание (инд. потребление)", "2", "H42", "H", "0" },
                { "     Техобслуживание (всего)", "2", "N42", "AC", "0" },
                
                { "Вывоз ТОП (Наименование)", "2", "A43", "B", "1" },
                { "     Вывоз ТОП (тариф)", "2", "G43", "G", "1" },
                { "     Вывоз ТОП (инд. потребление)", "2", "H43", "H", "1" },
                { "     Вывоз ТОП (всего)", "2", "N43", "AC", "1" },

                { "ХВС (Наименование)", "2", "A44", "B", "2" },
                { "     ХВС (объём)", "2", "F44", "F", "2" },
                { "     ХВС (тариф)", "2", "G44", "G", "2" },
                { "     ХВС (инд. потребление)", "2", "H44", "H", "2" },
                { "     ХВС (общ. потребление)", "2", "I44", "I", "2" },
                { "     ХВС (всего)", "2", "N44", "AC", "2" },

                { "ГВС (Наименование)", "2", "A45", "B", "3" },
                { "     ГВС (объём)", "2", "F45", "F", "3" },
                { "     ГВС (тариф)", "2", "G45", "G", "3" },
                { "     ГВС (инд. потребление)", "2", "H45", "H", "3" },
                { "     ГВС (общ. потребление)", "2", "I45", "I", "3" },
                { "     ГВС (всего)", "2", "N45", "AC", "3" },

                { "Водоотведение (Наименование)", "2", "A46", "B", "4" },
                { "     Водоотведение (объём)", "2", "F46", "F", "4" },
                { "     Водоотведение (тариф)", "2", "G46", "G", "4" },
                { "     Водоотведение (инд. потребление)", "2", "H46", "H", "4" },
                { "     Водоотведение (общ. потребление)", "2", "I46", "I", "4" },
                { "     Водоотведение (всего)", "2", "N46", "AC", "4" },

                { "Электроснабжение (Наименование)", "2", "A47", "B", "5" },
                { "     Электроснабжение (объём)", "2", "F47", "F", "5" },
                { "     Электроснабжение (тариф)", "2", "G47", "G", "5" },
                { "     Электроснабжение (инд. потребление)", "2", "H47", "H", "5" },
                { "     Электроснабжение (общ. потребление)", "2", "I47", "I", "5" },
                { "     Электроснабжение (всего)", "2", "N47", "AC", "5" },

                { "Отопление (Наименование)", "2", "A49", "B", "6" },
                { "     Отопление (тариф)", "2", "G49", "G", "6" },
                { "     Отопление (инд. потребление)", "2", "H49", "H", "6" },
                { "     Отопление (всего)", "2", "N49", "AC", "6" },

                { "ХВС (Наименование)", "2", "A50", "B", "7" },
                { "     ХВС (тариф)", "2", "G50", "G", "7" },
                { "     ХВС (инд. потребление)", "2", "H50", "H", "7" },
                { "     ХВС (всего)", "2", "N50", "AC", "7" },

                { "Стоки ХВС (Наименование)", "2", "A51", "B", "8" },
                { "     Стоки ХВС (тариф)", "2", "G51", "G", "8" },
                { "     Стоки ХВС (инд. потребление)", "2", "H51", "H", "8" },
                { "     Стоки ХВС (всего)", "2", "N51", "AC", "8" },

                { "ГВС (Наименование)", "2", "A52", "B", "9" },
                { "     ГВС (тариф)", "2", "G52", "G", "9" },
                { "     ГВС (инд. потребление)", "2", "H52", "H", "9" },
                { "     ГВС (всего)", "2", "N52", "AC", "9" },

                { "Стоки ГВС (Наименование)", "2", "A53", "B", "10" },
                { "     Стоки ГВС (тариф)", "2", "G53", "G", "10" },
                { "     Стоки ГВС (инд. потребление)", "2", "H53", "H", "10" },
                { "     Стоки ГВС (всего)", "2", "N53", "AC", "10" },

                { "Электроснабжение (Наименование)", "2", "A54", "B", "11" },
                { "     Электроснабжение (норматив 96)", "2", "E54", "D", "11" },
                { "     Электроснабжение (тариф)", "2", "G54", "G", "11" },
                { "     Электроснабжение (инд. потребление)", "2", "H54", "H", "11" },
                { "     Электроснабжение (всего)", "2", "N54", "AC", "11" },

                { "Электроснабжение (Наименование)", "2", "A55", "B", "12" },
                { "     Электроснабжение (тариф)", "2", "G55", "G", "12" },
                { "     Электроснабжение (инд. потребление)", "2", "H55", "H", "12" },
                { "     Электроснабжение (всего)", "2", "N55", "AC", "12" },

                { "Текущий ремонт (Наименование)", "2", "A56", "B", "13" },
                { "     Текущий ремонт (тариф)", "2", "G56", "G", "13" },
                { "     Текущий ремонт (инд. потребление)", "2", "H56", "H", "13" },
                { "     Текущий ремонт (всего)", "2", "N56", "AC", "13" },

                { "Антенна (Наименование)", "2", "A57", "B", "14" },
                { "     Антенна (тариф)", "2", "G57", "G", "14" },
                { "     Антенна (инд. потребление)",  "2", "H57", "H", "14" },
                { "     Антенна (всего)", "2", "N57", "AC", "14" },

                { "     Услуги банка (Наименование)", "2", "A58", "B", "15" },
                { "     Услуги банка (тариф)", "2", "G58", "G", "15" },
                { "     Услуги банка (инд. потребление)", "2", "H58", "H", "15" },
                { "     Услуги банка (всего)", "2", "N58", "AC", "15" },
            };

            for (int i = 0; i < param.Length / 5; i++)
            {
                dataGridView1.Rows.Add(param[i, 0], param[i, 1], param[i, 2], param[i, 3], param[i, 4]);
            }

        }

        // Обработка нажатия на кнопку "Старт"
        private void button3_Click(object sender, EventArgs e)
        {
            Form1.ActiveForm.Enabled = true;
            // Очищаем старые значения в файле. На входе название файла-шаблона в текущей 
            // директории и количество строк для очистки.
            Clear(dir_templ, 50000);

            // Получаем массив с адресами файлов:
            string[] dirs = SelectFiles(dir_parse);

            progressBar1.Minimum = 0;
            progressBar1.Maximum = dirs.Length;


            // Запускаем программу парсинга
            progressBar1.Visible = true;
            label3.Text = "Запуск программы...";
            ReadFile1(1, dirs, dir_templ, progressBar1, label3, Cancel);

            //DataGrid_test();
        }

        // Обработка нажатия кнопки "Обзор" для открытия папки с файлами данных
        private void button1_Click(object sender, EventArgs e)
        {
            var path = Directory.GetCurrentDirectory();
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(FBD.SelectedPath);
                textBox1.Text = FBD.SelectedPath;
                dir_parse = textBox1.Text;
            }
        }

        // Обработка нажатия кнопки "Обзор" для открытия папки с файла-шаблона
        private void button2_Click(object sender, EventArgs e)
        {
            var path = Directory.GetCurrentDirectory();
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(FBD.SelectedPath);
                textBox2.Text = FBD.SelectedPath;
                dir_templ = textBox2.Text;
            }
        }

        // Метод возвращает список с адресами файлов-доноров для парсинга в массиве строк
        public static string[] SelectFiles(String path = "")
        {
            if (path == "")
            {
                path = Directory.GetCurrentDirectory();
            }
            try
            {
                //path = Directory.GetCurrentDirectory();
                Console.WriteLine("Адрес пути к файлам {0}.", path);
                string[] dirs = Directory.GetFiles(path, "*.xls");
                Console.WriteLine("Количество найденых файлов {0}.", dirs.Length);
                Console.WriteLine();

                int t = 1;
                foreach (string dir in dirs)
                {
                    //Console.WriteLine(t.ToString() + " " + dir);
                    t += 1;
                }

                return dirs;


            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибка чтения файлов.", e.ToString());
                string[] dirs = null;
                return dirs;
            }


        }

        // Метод автоматического перебора файлов-доноров из указанной папки и записи значений в шаблон
        /*
        public void ReadFile(int counter, String[] dirs, String template, ProgressBar progressBar1, Label label3, Button Cancel)
        {

            //CloseProcess("start");

            var app = new Excel.Application();
            app.Visible = false;

            var outbook = app.Workbooks.Open(template);
            //Console.WriteLine(Directory.GetCurrentDirectory() + "/" + template);

            progressBar1.Visible = true;

            int t = 1;
            int n = 1;
            //int stoper = 0;

            foreach (string dir in dirs)
            {
                var inbook = app.Workbooks.Open(dir);
                Application.DoEvents();
                //button4.Focus();
                Cancel.Enabled = true;
                // Лист 1

                // 3 - Номер платежного документа
                CopyRange(
                inbook.Sheets[1].Range["I1", "I1"],
                outbook.Sheets[1].Range["C" + (t + 3).ToString(), "C" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 4 - Расчетный период
                CopyRange(
                inbook.Sheets[1].Range["D21", "D21"],
                outbook.Sheets[1].Range["D" + (t + 3).ToString(), "D" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 5 - Общая площадь
                CopyRange(
                inbook.Sheets[1].Range["F26", "F26"],
                outbook.Sheets[1].Range["E" + (t + 3).ToString(), "E" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 7 - Отапливаемая площадь
                CopyRange(
                inbook.Sheets[1].Range["I26", "I26"],
                outbook.Sheets[1].Range["G" + (t + 3).ToString(), "G" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 8 - Количество проживающих (зарегестрированных)
                CopyRange(
                inbook.Sheets[1].Range["M26", "M26"],
                outbook.Sheets[1].Range["H" + (t + 3).ToString(), "H" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 9 - Сумма к оплате за расчетный период
                CopyRange(
                inbook.Sheets[1].Range["N14", "N14"],
                outbook.Sheets[1].Range["X" + (t + 3).ToString(), "X" + (t + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 16 - БИК банка

                CopyRangeReg(
                inbook.Sheets[1].Range["J3", "J3"],
                outbook.Sheets[1].Range["P" + (t + 3).ToString(), "P" + (t + 3).ToString()],
                @"Кор\/сч: \d+ БИК: (\d+)");

                Application.DoEvents();
                Cancel.Focus();
                //button4.Focus();

                // 17 - Расчетный счет

                CopyRangeReg(
                inbook.Sheets[1].Range["I2", "I2"],
                outbook.Sheets[1].Range["Q" + (t + 3).ToString(), "Q" + (t + 3).ToString()],
                @"ИНН\/КПП: \d+ \/ \d+  Р\/сч: (\d+)");

                Application.DoEvents();

                // Лист 2

                if (t == 1) n = 1;
                else n += 16;

                // 01 - Номер платежного документа
                CopyRange(
                inbook.Sheets[1].Range["I1", "I1"],
                outbook.Sheets[2].Range["A" + (n + 4).ToString(), "A" + (n + 4).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // 02  - Заполняем поля затрат

                // Техобслуживание
                outbook.Sheets[2].Range["B" + (n + 4).ToString(), "B" + (n + 4).ToString()] = "Техобслуживание";

                CopyRange(
                inbook.Sheets[1].Range["F42", "I42"],
                outbook.Sheets[2].Range["G" + (n + 4).ToString(), "H" + (n + 4).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N42", "N42"],
                outbook.Sheets[2].Range["AC" + (n + 4).ToString(), "AC" + (n + 4).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Вывоз ТОП
                outbook.Sheets[2].Range["B" + (n + 4 + 1).ToString(), "B" + (n + 4 + 1).ToString()] = "Вывоз ТОП";

                CopyRange(
                inbook.Sheets[1].Range["F43", "I43"],
                outbook.Sheets[2].Range["G" + (n + 4 + 1).ToString(), "H" + (n + 4 + 1).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N43", "N43"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 1).ToString(), "AC" + (n + 4 + 1).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // ХВС
                outbook.Sheets[2].Range["B" + (n + 4 + 2).ToString(), "B" + (n + 4 + 2).ToString()] = "ХВС";

                CopyRange(
                inbook.Sheets[1].Range["F44", "I44"],
                outbook.Sheets[2].Range["G" + (n + 4 + 2).ToString(), "H" + (n + 4 + 2).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N44", "N44"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 2).ToString(), "AC" + (n + 4 + 2).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // ГВС
                outbook.Sheets[2].Range["B" + (n + 4 + 3).ToString(), "B" + (n + 4 + 3).ToString()] = "ГВС";

                CopyRange(
                inbook.Sheets[1].Range["F45", "I45"],
                outbook.Sheets[2].Range["G" + (n + 4 + 3).ToString(), "H" + (n + 4 + 3).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N45", "N45"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 3).ToString(), "AC" + (n + 4 + 3).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Водоотведение
                outbook.Sheets[2].Range["B" + (n + 4 + 4).ToString(), "B" + (n + 4 + 4).ToString()] = "Водоотведение";

                CopyRange(
                inbook.Sheets[1].Range["F46", "I46"],
                outbook.Sheets[2].Range["G" + (n + 4 + 4).ToString(), "H" + (n + 4 + 4).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N46", "N46"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 4).ToString(), "AC" + (n + 4 + 4).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Электроснабжение
                outbook.Sheets[2].Range["B" + (n + 4 + 5).ToString(), "B" + (n + 4 + 5).ToString()] = "Электроснабжение";

                CopyRange(
                inbook.Sheets[1].Range["F47", "I47"],
                outbook.Sheets[2].Range["G" + (n + 4 + 5).ToString(), "H" + (n + 4 + 5).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N47", "N47"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 5).ToString(), "AC" + (n + 4 + 5).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Отопление
                outbook.Sheets[2].Range["B" + (n + 4 + 6).ToString(), "B" + (n + 4 + 6).ToString()] = "Отопление";

                CopyRange(
                inbook.Sheets[1].Range["F49", "I49"],
                outbook.Sheets[2].Range["G" + (n + 4 + 6).ToString(), "H" + (n + 4 + 6).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N49", "N49"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 6).ToString(), "AC" + (n + 4 + 6).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // ХВС
                outbook.Sheets[2].Range["B" + (n + 4 + 7).ToString(), "B" + (n + 4 + 7).ToString()] = "ХВС";

                CopyRange(
                inbook.Sheets[1].Range["F50", "I50"],
                outbook.Sheets[2].Range["G" + (n + 4 + 7).ToString(), "H" + (n + 4 + 7).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N50", "N50"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 7).ToString(), "AC" + (n + 4 + 7).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Стоки ХВС
                outbook.Sheets[2].Range["B" + (n + 4 + 8).ToString(), "B" + (n + 4 + 8).ToString()] = "Стоки ХВС";

                CopyRange(
                inbook.Sheets[1].Range["F51", "I51"],
                outbook.Sheets[2].Range["G" + (n + 4 + 8).ToString(), "H" + (n + 4 + 8).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N51", "N51"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 8).ToString(), "AC" + (n + 4 + 8).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // ГВС
                outbook.Sheets[2].Range["B" + (n + 4 + 9).ToString(), "B" + (n + 4 + 9).ToString()] = "ГВС";

                CopyRange(
                inbook.Sheets[1].Range["F52", "I52"],
                outbook.Sheets[2].Range["G" + (n + 4 + 9).ToString(), "H" + (n + 4 + 9).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N52", "N52"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 9).ToString(), "AC" + (n + 4 + 9).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Стоки ГВС
                outbook.Sheets[2].Range["B" + (n + 4 + 10).ToString(), "B" + (n + 4 + 10).ToString()] = "Стоки ГВС";

                CopyRange(
                inbook.Sheets[1].Range["F53", "I53"],
                outbook.Sheets[2].Range["G" + (n + 4 + 10).ToString(), "H" + (n + 4 + 10).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N53", "N53"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 10).ToString(), "AC" + (n + 4 + 10).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Электроснабжение
                outbook.Sheets[2].Range["B" + (n + 4 + 11).ToString(), "B" + (n + 4 + 11).ToString()] = "Электроснабжение";

                CopyRange(
                inbook.Sheets[1].Range["F54", "I54"],
                outbook.Sheets[2].Range["G" + (n + 4 + 11).ToString(), "H" + (n + 4 + 11).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N54", "N54"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 11).ToString(), "AC" + (n + 4 + 11).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Электроснабжение
                outbook.Sheets[2].Range["B" + (n + 4 + 12).ToString(), "B" + (n + 4 + 12).ToString()] = "Электроснабжение";

                CopyRange(
                inbook.Sheets[1].Range["F55", "I55"],
                outbook.Sheets[2].Range["G" + (n + 4 + 12).ToString(), "H" + (n + 4 + 12).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N55", "N55"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 12).ToString(), "AC" + (n + 4 + 12).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Текущ. ремонт
                outbook.Sheets[2].Range["B" + (n + 4 + 13).ToString(), "B" + (n + 4 + 13).ToString()] = "Текущ. ремонт";

                CopyRange(
                inbook.Sheets[1].Range["F56", "I56"],
                outbook.Sheets[2].Range["G" + (n + 4 + 13).ToString(), "H" + (n + 4 + 13).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N56", "N56"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 13).ToString(), "AC" + (n + 4 + 13).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Антенна
                outbook.Sheets[2].Range["B" + (n + 4 + 14).ToString(), "B" + (n + 4 + 14).ToString()] = "Антенна";

                CopyRange(
                inbook.Sheets[1].Range["F57", "I57"],
                outbook.Sheets[2].Range["G" + (n + 4 + 14).ToString(), "H" + (n + 4 + 14).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N57", "N57"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 14).ToString(), "AC" + (n + 4 + 14).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                // Услуги банка
                outbook.Sheets[2].Range["B" + (n + 4 + 15).ToString(), "B" + (n + 4 + 15).ToString()] = "Услуги банка";

                CopyRange(
                inbook.Sheets[1].Range["F58", "I58"],
                outbook.Sheets[2].Range["G" + (n + 4 + 15).ToString(), "H" + (n + 4 + 15).ToString()]);

                CopyRange(
                inbook.Sheets[1].Range["N58", "N58"],
                outbook.Sheets[2].Range["AC" + (n + 4 + 15).ToString(), "AC" + (n + 4 + 15).ToString()]);

                Application.DoEvents();
                Cancel.Focus();

                Console.WriteLine("Обработано файлов {0} из {1}.", t, dirs.Length);

                inbook.Close(false);
                progressBar1.Value += 1;
                //MessageBox.Show("Обработан файл " + t);
                label3.Text = "Обработано " + t + " из " + dirs.Length;
                t += 1;

                if (stop_cycle == "stop")
                {
                    label3.Text = "Обработка остановлена.";
                    stop_cycle = "no";
                    progressBar1.Value = 0;
                    break;
                }
            }

            outbook.Close(true);
            app.Visible = false;

            app.Quit();

            Console.WriteLine();
            Console.WriteLine("Шаблон \"{0}\" обновлён.", template);
            Console.WriteLine("Нажмите клавишу Enter для выхода...");
            Console.Read();
        }
        */

        // Метод автоматического перебора файлов-доноров из указанной папки и записи значений в шаблон
        public void ReadFile1(int counter, String[] dirs, String template, ProgressBar progressBar1, Label label3, Button Cancel)
        {
            Regex regex = new Regex(@"[A-Z]{1,2}[1-9]{0,2}");

            //CloseProcess("start");

            var app = new Excel.Application();
            app.Visible = false;

            var outbook = app.Workbooks.Open(template);

            progressBar1.Visible = true;

            int t = 1;
            int n = 1;
            //int stoper = 0;

            foreach (var dir in dirs)
            {
                var inbook = app.Workbooks.Open(dir);
                Application.DoEvents();;
                Cancel.Enabled = true;
                button3.Enabled = false;

                String account = dataGridView1.Rows[0].Cells[2].Value.ToString();

                if (t == 1) n = 1;
                else n += 16;
              

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {                   
                    String name = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    int list_num = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    String kvitancia_link = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    String template_link = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    int count_shift = Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value.ToString());

                        if (i < 8) // Заполняем первый лист
                    {
                        /*
                        CopyRange(
                            inbook.Sheets[1].Range[kvitancia_link, kvitancia_link],
                            outbook.Sheets[(int)list_num].Range[template_link + (t + 3).ToString(), template_link + (t + 3).ToString()]);
                        */
                        if (regex.IsMatch(kvitancia_link))
                        {
                            CopyRange(
                            inbook.Sheets[1].Range[kvitancia_link, kvitancia_link],
                            outbook.Sheets[(int)list_num].Range[template_link + (t + 3).ToString(), template_link + (t + 3).ToString()]);                        
                        }
                        else
                        {
                            outbook.Sheets[(int)list_num].Range[template_link + (t + 3).ToString(), template_link + (t + 3).ToString()] = kvitancia_link;
                        }

                        Application.DoEvents();
                        Cancel.Focus();
                    }
                    else      // Заполняем второй лист
                    {
                        if (regex.IsMatch(account))
                            CopyRange(
                            inbook.Sheets[1].Range[account, account],
                            outbook.Sheets[(int)list_num].Range["A" + (n + 4 + count_shift).ToString(), "A" + (n + 4 + count_shift).ToString()]);
                        else
                        {
                            outbook.Sheets[(int)list_num].Range["A" + (n + 4 + count_shift).ToString(), "A" + (n + 4 + count_shift).ToString()] = account;
                        }
                        /*
                        CopyRange(
                            inbook.Sheets[1].Range[kvitancia_link, kvitancia_link],
                            outbook.Sheets[(int)list_num].Range[template_link + (n + 4 + count_shift).ToString(), template_link + (n + 4 + count_shift).ToString()]);
                        */
                        if (regex.IsMatch(kvitancia_link))
                        {
                            CopyRange(
                            inbook.Sheets[1].Range[kvitancia_link, kvitancia_link],
                            outbook.Sheets[(int)list_num].Range[template_link + (n + 4 + count_shift).ToString(), template_link + (n + 4 + count_shift).ToString()]);
                        }
                        else
                        {
                            outbook.Sheets[(int)list_num].Range[template_link + (n + 4 + count_shift).ToString(), template_link + (n + 4 + count_shift).ToString()] = kvitancia_link;
                        }
                        Application.DoEvents();
                        Cancel.Focus();

                    }
                }

                Console.WriteLine("Обработано файлов {0} из {1}.", t, dirs.Length);

                inbook.Close(false);
                progressBar1.Value += 1;
                
                //MessageBox.Show("Обработан файл " + t);
                label3.Text = "Обработано " + t + " из " + dirs.Length;
                t += 1;

                if (stop_cycle == "stop")
                {
                    label3.Text = "Обработка остановлена.";
                    stop_cycle = "no";
                    progressBar1.Value = 0;
                    button3.Enabled = true;
                    break;
                }
            }

            outbook.Close(true);
            app.Visible = false;

            app.Quit();

            Console.WriteLine();
            Console.WriteLine("Шаблон \"{0}\" обновлён.", template);
            Console.WriteLine("Нажмите клавишу Enter для выхода...");
            Console.Read();
            if (label3.Text != "Обработка остановлена.")
            {
                label3.Text = "Обработка завершена.";
            }
        }

        // Метод копирует диапазон из файла-донора в шаблон
        static void CopyRange(Excel.Range Source, Excel.Range Destination)
        {
            Destination.Cells[1, 1]
                .Resize(Source.Rows.Count, Source.Columns.Count)
                .Value = Source.Value;
        }

        // Метод копирует диапазон из файла-донора в шаблон с фильтрацией по регулярному выражению
        static void CopyRangeReg(Excel.Range Source, Excel.Range Destination, String link, String reg = @"[A-Z]{1,2}[1-9]{0,2}")
        {
            Regex regex = new Regex(reg);
            //Match match = regex.Match(Source.Value);
            
            if (regex.IsMatch(link))
            {
                Destination.Cells[1, 1]
                .Resize(Source.Rows.Count, Source.Columns.Count)
                .Value = Source.Value;
            }
            else
            {
                Destination.Cells[1, 1]
                .Resize(1, 1)
                .Value = link;
            }
        }

        // Метод очищает шаблон от старых значений
        public void Clear(String template, int length)
        {
            CloseProcess("start");

            var app1 = new Excel.Application();
            var clear = app1.Workbooks.Open(template);
            clear.Sheets[1].Range["A4", "AD" + length.ToString()].Value = null;
            clear.Sheets[2].Range["A4", "AF" + length.ToString()].Value = null;
            clear.Close(true);
            Console.WriteLine("Шаблон \"{0}\" очищен от старых значений.", template);
        }

        // Обработка нажатия кнопки "Стоп" для остановки работы программы после запуска
        private void cancel_Click(object sender, EventArgs e)
        {
            label3.Text = "Обработка останавливается...";
            stop_cycle = "stop";
            //CloseProcess("cancel");
            //CloseProcess("start");
        }

        // Метод закрытия и очистки потока и запущенных файлов Excel 
        public void CloseProcess(String but = "stop")
        {
            if (but == "start")
            {
                var app = new Excel.Application();
                app.Quit();
            }

            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }

            if (but == "stop")
            {
                this.Dispose(true);
                GC.SuppressFinalize(this);
            }

        }

        // Метод обрабатывает очистку открытых потоков и файлов при закрытии окна формы
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseProcess();
        }

        private void DataGrid_test()
        {
            var size = dataGridView1.Rows[0].Cells[0].Value;
            MessageBox.Show(size.ToString());

            //for (var i = 0; i < dataGridView1.RowCount)
            {
                
            }
        }
    }
}
