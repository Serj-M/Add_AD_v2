using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.DirectoryServices;

namespace EPS_31._10
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        //excel.visible = true;
          private string fileName;
        // ключь кнопки "проверка файла". запрещает проверять дважды 
          bool key = true;
        // ключь кнопки "создать пользователей в AD". запрещает проверять дважды
          bool key2 = true;
        string directoryFileDodelat="";
        string directoryFile = "";
        // Имя администратора и пароль для соединения с сервером
          string Username = null;
          string Password = null;
          string PathLDAP = null;
        // Акаунт, полное имя и пароль для поиска в AD
          string newAccount = null;
          string newFullname = null;
          string newPassword = null;
        //формирование учетной записи
          string[] arrFIO = {"", "", ""};
          string formulaUZ = null;
          //нельзя нажимать дважды на кнопки
          bool keyF = true;
          bool keyI = true;
          bool keyO = true;
          bool keyTochka = true;
          bool keyTire = true;
          bool keyPodcher = true;
          bool keySteret = true;
        //пароль
          int longPass = 6; //длина пароля

        // генератор случайных чисел для генерации паролей. нужно мспользовать один экземпляр для всей последовательности
          Random rnd = new Random();

        //состовление словаря транслитерации
        static string funcDictionary(char letter)
        {
            Dictionary<char, string> dic = new Dictionary<char, string>();
            dic.Add('а', "a"); dic.Add('б', "b"); dic.Add('в', "v"); dic.Add('г', "g"); dic.Add('д', "d"); dic.Add('е', "e"); dic.Add('ё', "e"); dic.Add('ж', "zh");
            dic.Add('з', "z"); dic.Add('и', "i"); dic.Add('й', "y"); dic.Add('к', "k"); dic.Add('л', "l"); dic.Add('м', "m"); dic.Add('н', "n"); dic.Add('о', "o");
            dic.Add('п', "p"); dic.Add('р', "r"); dic.Add('с', "s"); dic.Add('т', "t"); dic.Add('у', "u"); dic.Add('ф', "f"); dic.Add('х', "kh"); dic.Add('ц', "ts");
            dic.Add('ч', "ch"); dic.Add('ш', "sh"); dic.Add('щ', "sch"); dic.Add('ъ', ""); dic.Add('ы', "y"); dic.Add('ь', ""); dic.Add('э', "e"); dic.Add('ю', "yu");
            dic.Add('я', "ya"); dic.Add('А', "A"); dic.Add('Б', "B"); dic.Add('В', "V"); dic.Add('Г', "G"); dic.Add('Д', "D"); dic.Add('Е', "E"); dic.Add('Ё', "E");
            dic.Add('Ж', "Zh"); dic.Add('З', "Z"); dic.Add('И', "I"); dic.Add('Й', "Y"); dic.Add('К', "K"); dic.Add('Л', "L"); dic.Add('М', "M"); dic.Add('Н', "N");
            dic.Add('О', "O"); dic.Add('П', "P"); dic.Add('Р', "R"); dic.Add('С', "S"); dic.Add('Т', "T"); dic.Add('У', "U"); dic.Add('Ф', "F"); dic.Add('Х', "Kh");
            dic.Add('Ц', "Ts"); dic.Add('Ч', "Ch"); dic.Add('Ш', "Sh"); dic.Add('Щ', "Sch"); dic.Add('Ъ', ""); dic.Add('Ы', "Y"); dic.Add('Ь', ""); dic.Add('Э', "E");
            dic.Add('Ю', "Yu"); dic.Add('Я', "Ya"); dic.Add('-', "-"); dic.Add('y', "y"); dic.Add('i', "i"); dic.Add('e', "e");
            return dic[letter];
        }
        //состовление словаря транслитерации для инициалов
        static string funcDictionaryInic(char letter)
        {
            Dictionary<char, string> dicInic = new Dictionary<char, string>();
            dicInic.Add('а', "a"); dicInic.Add('б', "b"); dicInic.Add('в', "v"); dicInic.Add('г', "g"); dicInic.Add('д', "d"); dicInic.Add('е', "e"); dicInic.Add('ё', "e"); dicInic.Add('ж', "z");
            dicInic.Add('з', "z"); dicInic.Add('и', "i"); dicInic.Add('й', "y"); dicInic.Add('к', "k"); dicInic.Add('л', "l"); dicInic.Add('м', "m"); dicInic.Add('н', "n"); dicInic.Add('о', "o");
            dicInic.Add('п', "p"); dicInic.Add('р', "r"); dicInic.Add('с', "s"); dicInic.Add('т', "t"); dicInic.Add('у', "u"); dicInic.Add('ф', "f"); dicInic.Add('х', "k"); dicInic.Add('ц', "t");
            dicInic.Add('ч', "c"); dicInic.Add('ш', "s"); dicInic.Add('щ', "s"); dicInic.Add('ъ', ""); dicInic.Add('ы', "y"); dicInic.Add('ь', ""); dicInic.Add('э', "e"); dicInic.Add('ю', "y");
            dicInic.Add('я', "y"); dicInic.Add('А', "A"); dicInic.Add('Б', "B"); dicInic.Add('В', "V"); dicInic.Add('Г', "G"); dicInic.Add('Д', "D"); dicInic.Add('Е', "E"); dicInic.Add('Ё', "E");
            dicInic.Add('Ж', "Z"); dicInic.Add('З', "Z"); dicInic.Add('И', "I"); dicInic.Add('Й', "Y"); dicInic.Add('К', "K"); dicInic.Add('Л', "L"); dicInic.Add('М', "M"); dicInic.Add('Н', "N");
            dicInic.Add('О', "O"); dicInic.Add('П', "P"); dicInic.Add('Р', "R"); dicInic.Add('С', "S"); dicInic.Add('Т', "T"); dicInic.Add('У', "U"); dicInic.Add('Ф', "F"); dicInic.Add('Х', "K");
            dicInic.Add('Ц', "T"); dicInic.Add('Ч', "C"); dicInic.Add('Ш', "S"); dicInic.Add('Щ', "S"); dicInic.Add('Ъ', ""); dicInic.Add('Ы', "Y"); dicInic.Add('Ь', ""); dicInic.Add('Э', "E");
            dicInic.Add('Ю', "Y"); dicInic.Add('Я', "Y");
            return dicInic[letter];
        }

        // функция для формирования фамилии в формуле учетной записи 
        static string F(string familya)
        {
            string loginF = null;
            for (int j = 0; j < familya.Length; j++)
            {
                //если найдена буква из Dictionary
                try
                {
                    // заменяеться в соответствии с транслитерацией и присваиваеться временному результату
                    loginF = loginF + funcDictionary(familya[j]);
                }
                catch
                {
                    continue;
                }
            }
            return loginF;
        }

        // функция для формирования имени в формуле учетной записи 
        static string I(string imya)
        {
            string loginI = null;
            try
            {
                // заменяеться в соответствии с транслитерацией и присваиваеться временному результату
                loginI = loginI + funcDictionaryInic(imya[0]);
            }
            catch
            {
                loginI = loginI + imya[0];
            }
            return loginI;
        }

        // функция для формирования отчества в формуле учетной записи 
        static string O(string otchestvo)
        {
            string loginO = null;
            try
            {
                // заменяеться в соответствии с транслитерацией и присваиваеться временному результату
                loginO = loginO + funcDictionaryInic(otchestvo[0]);
            }
            catch
            {
                loginO = loginO + otchestvo[0];
            }
            return loginO;
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        void TextBox1KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsLetterOrDigit(e.KeyChar)) return;
            else
                e.Handled = true;
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // скрывакет вводимые символы
            textBox1.PasswordChar = '*';
            comboBox2.SelectedItem = "6";
        }  

        private void buttonExport_Click_1(object sender, EventArgs e)
        {
            //int n = 0;
            string st = "";
            if (key == false) // ключь кнопки "проверка файла". запрещает проверять дважды 
            {
                fileName = directoryFileDodelat + " - Доработать.xlsx";
                try
                {
                    ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    //Книга.
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                    // экспорт обработанных пользователей
                    for (int i = 0; i < dataGridViewMain.Rows.Count; i++)
                    {
                        DataGridViewRow row = dataGridViewMain.Rows[i]; // строки

                        for (int j = 0; j < row.Cells.Count; j++) //цикл по ячейкам строки
                        {
                            ObjExcel.Cells[i + 1, j + 1] = row.Cells[j].Value;
                        }
                    }
                   
                    ObjWorkBook.SaveAs(/*directoryFile*/directoryFileDodelat + " - Успешно созданы" + st + ".xlsx");
                    ObjWorkBook.Close();
                    
                    // экспорт пользователей пользователей на доработку. В конце имени файла !!! 
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                    if (dataGridView1.Rows.Count != 1)
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[i]; // строки

                            for (int j = 0; j < row.Cells.Count; j++) //цикл по ячейкам строки
                            {
                                ObjExcel.Cells[i + 1, j + 1] = row.Cells[j].Value;
                            }
                        }
                        ObjWorkBook.SaveAs(fileName);
                    }
                    dataGridViewMain.Rows.Clear();
                    dataGridView1.Rows.Clear();



                    /*finally
                    {
                        {*/
                    ObjWorkBook.Close();
                    // Закрытие приложения Excel.
                    ObjExcel.Quit();
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    ObjExcel = null;
                    GC.Collect();
                    /*}
                }*/
                }
                catch { ObjWorkBook.Close(0); /*MessageBox.Show(ex.Message, "Error");*/ }
                textBoxFileName.Text= "";
                //this.Text = this.Text + " - " + textBoxFileName.Text + ".xlsx";
            }
            else
            {
                MessageBox.Show("Выполните Проверку файла и если требуеться Создайте пользователей в AD. Затем повторите экспорт.");
            }
        }

        private void buttonImport_Click_1(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            dataGridView1.Rows.Clear();
            key = true;
            key2 = true;
            keySteret = true;
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            openDialog.ShowDialog();

            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                //Таблица.
                ObjWorkSheet = ObjExcel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                Microsoft.Office.Interop.Excel.Range rg = null;
                //ObjExcel.Visible = true;

                // directoryFileDodelat - хранит директорию и имя файла без расширения (.xls или .xlsx). Понадобиться для экспорта в туже директорию 
                directoryFileDodelat = openDialog.FileName;
                directoryFile = openDialog.FileName;
                Regex rgxXLS = new Regex(@".xls[x]?$", RegexOptions.IgnoreCase);
                directoryFileDodelat = rgxXLS.Replace(directoryFileDodelat, string.Empty);
                //MessageBox.Show(directoryFile);

                // оформление dataGridView1
                dataGridView1.Rows.Add();
                dataGridView1[0, 0].Value = "ФИО";
                dataGridView1[1, 0].Value = "Подразделение";
                dataGridView1[2, 0].Value = "Должность";
                dataGridView1[3, 0].Value = "Комната";
                dataGridView1[4, 0].Value = "Телефон";
                dataGridView1[5, 0].Value = "Улица";
                dataGridView1[6, 0].Value = "";
                dataGridView1[7, 0].Value = "";
                dataGridView1[8, 0].Value = "";
                dataGridView1[9, 0].Value = "";
                dataGridView1[10, 0].Value = "";

                Int32 row = 1;
                Int32 more_Renge = 0; 
                dataGridViewMain.Rows.Clear();
                List<String> arr = new List<string>();
                // пока не конец файла формируем датаГрид
                while (ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null)
                {
                    // Читаем данные из ячейки
                    rg = ObjWorkSheet.get_Range("a" + row, "k" + row);
                    foreach (Microsoft.Office.Interop.Excel.Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch { arr.Add(""); }
                    }
                    dataGridViewMain.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8], arr[9], arr[10]);
                    arr.Clear();
                    //more_Renge = 0;
                    row++;
                }
                //проверка разрыва в Excel документе (поиск пустых строк в заполненом деопазоне)
                while ((ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null) || (more_Renge < 5))
                {
                    if (ObjWorkSheet.get_Range("a" + row, "a" + row).Value == null) // проверка пустых сторк (more_Renge < 5)
                    {
                        more_Renge++;
                        row++;
                        continue;
                    }
                    else
                    {
                        if (more_Renge > 0)
                            MessageBox.Show("Исходный файл офрмлен не правильно. Не все данные из файла загружены! Существуют разрывы");
                        
                        ObjWorkBook.Close(false, "", null);
                        // Закрытие приложения Excel.
                        ObjExcel.Quit();
                        ObjWorkBook = null;
                        ObjWorkSheet = null;
                        ObjExcel = null;
                        GC.Collect();

                        dataGridViewMain.Rows.Clear();
                        dataGridView1.Rows.Clear();
                        return;
                    }
                    //break;
                }
                
                MessageBox.Show("Файл успешно считан!", "Считывания excel файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch /*(Exception ex)*/ {/* MessageBox.Show("Ошибка: " + ex.Message, "Ошибка при считывании excel файла", MessageBoxButtons.OK, MessageBoxIcon.Error); */}
            finally
            {
                try
                {
                    ObjWorkBook.Close(false, "", null);
                    // Закрытие приложения Excel.
                    ObjExcel.Quit();
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    ObjExcel = null;
                    GC.Collect();
                }
                catch
                { 
                }
            }

            this.Text = this.Text + " - " + openDialog.SafeFileName;
            textBoxFileName.Text = openDialog.SafeFileName;
        }

        private void textBoxFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridViewMain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

 //----------------------------------------------проверка структуры файла--------------------------------------------------------

        private void button1_Click(object sender, EventArgs e)  // проверка шапки
        {
            if (textBox4.Text != "")
            {   
                int temp2 = 0;
                string[] arrFormulaUZ = {"F", "FIO", "F.IO", "F_IO", "F-IO", "IOF", "IO.F", "IO_F", "IO-F", "FI", "F.I", "F_I", "F-I", "IF", "I.F", "I_F", "I-F"};
                for (int i = 0; i < arrFormulaUZ.Length; i++)
                {  
                    if (textBox4.Text == arrFormulaUZ[i])
                    {
                        temp2++;
                    }
                }
                if (temp2 == 0)
                {
                    MessageBox.Show("Неправильная формула УЗ.");
                    keyF = true;
                    keyI = true;
                    keyO = true;
                    keyTochka = true;
                    keyTire = true;
                    keyPodcher = true;

                    formulaUZ = null;
                    textBox4.Text = null;
                    textBox4.Clear();
                    return;
                }
                // ----------------------------------проверка шапки------------------------------------------------------------------

                if (key) // ключь кнопки "проверка файла". запрещает проверять дважды 
                {
                    try
                    {
                        longPass = comboBox2.SelectedIndex;
                        keySteret = false;
                        dataGridViewMain.Columns[6].Visible = true;
                        dataGridViewMain[5, 0].Value = "Улица";
                        string shapka = "";
                        Int32 shapkaInt = 0;
                        for (int i = 0; i < dataGridViewMain.Rows.Count; i++)
                        {
                            if ((dataGridViewMain.Rows[i].Cells[0].Value != null) &&
                                (dataGridViewMain.Rows[i].Cells[0].FormattedValue.ToString().IndexOf("ФИО", StringComparison.CurrentCultureIgnoreCase) != -1))
                            {
                                shapka += dataGridViewMain.Rows[i].Cells[0].Value.ToString();
                                shapkaInt++;
                                //for (int j2 = 1; j2 < 7 /*размер шапки*/; j2++) // проверка оформления шапки в строке в которой найдено ФИО
                                //{
                                if ((dataGridViewMain.Rows[i].Cells[1].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[1].FormattedValue.ToString().IndexOf("Подразделение", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[1].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }
                                if ((dataGridViewMain.Rows[i].Cells[2].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[2].FormattedValue.ToString().IndexOf("Должность", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[2].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }
                                if ((dataGridViewMain.Rows[i].Cells[3].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[3].Value.ToString().IndexOf("Комната", StringComparison.CurrentCultureIgnoreCase) != -1) ||
                                    (dataGridViewMain.Rows[i].Cells[3].Value.ToString().IndexOf("Кабинет", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[3].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }
                                if ((dataGridViewMain.Rows[i].Cells[4].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[4].Value.ToString().IndexOf("Телефон", StringComparison.CurrentCultureIgnoreCase) != -1) ||
                                    (dataGridViewMain.Rows[i].Cells[4].Value.ToString().IndexOf("Номер телефона", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[4].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }
                                if ((dataGridViewMain.Rows[i].Cells[5].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[5].Value.ToString().IndexOf("Улица", StringComparison.CurrentCultureIgnoreCase) != -1) ||
                                    (dataGridViewMain.Rows[i].Cells[5].Value.ToString().IndexOf("Адрес", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[5].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }
                                /*if ((dataGridViewMain.Rows[i].Cells[6].Value != null) &&
                                    (dataGridViewMain.Rows[i].Cells[6].Value.ToString().IndexOf("Дом", StringComparison.CurrentCultureIgnoreCase) != -1))
                                {
                                    shapka += " " + dataGridViewMain.Rows[i].Cells[6].Value.ToString();
                                    shapkaInt++;
                                    //continue;
                                }*/
                                //}
                                dataGridViewMain.ClearSelection();
                                dataGridViewMain.Rows[i].Selected = true;
                                if (shapkaInt == 6)
                                {
                                    //MessageBox.Show("ОК! Файл проверен.\n" + shapka);
                                    dataGridViewMain[7, 0].Value = "Логин";
                                    dataGridViewMain[8, 0].Value = "Пароль";
                                    dataGridViewMain[9, 0].Value = "Соmpany";
                                }
                                else
                                {
                                    MessageBox.Show("Не все поля шапки найдены ! Необходимо исользовать стандатрный шаблон.");

                                    dataGridViewMain.Rows.Clear();
                                    dataGridView1.Rows.Clear();
                                    return;
                                }
                                i = dataGridViewMain.Rows.Count;
                                break;
                            }
                            else
                            {
                                MessageBox.Show("Импортируемый файл оформлен не правильно!");

                                dataGridViewMain.Rows.Clear();
                                dataGridView1.Rows.Clear();
                                return;
                            }
                            i = dataGridViewMain.Rows.Count; //для выхода из основного цикла
                            break;
                        }
                        // -------------------------------------проверка оформления файла --------------------------------------------------

                        // ********* проверка ФИО **********
                        Regex rgxFIO = new Regex(@"\.");
                        Regex rgxProbel2 = new Regex("\\s+");

                        // цикл по столбцу ФИО в датагрде
                        for (int i = 1; i < dataGridViewMain.RowCount; i++)
                        {
                            DataGridViewRow row = new DataGridViewRow();
                            row.CreateCells(dataGridViewMain);
                            // в пустых ячейках дальнейшая проверка не нужна
                            if (dataGridViewMain[0, i].Value != null)
                            {
                                //ФИО не должно содержать инициал
                                /*if (Regex.IsMatch(dataGridViewMain[0, i].Value.ToString(), rgxFIO.ToString()))
                                {
                                    MessageBox.Show("ФИО содержит инициалы");
                                    //формируеться колекция стороки 
                                    row.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                         dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                         dataGridViewMain[6, i].Value.ToString(), dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                         dataGridViewMain[9, i].Value.ToString(), "Не правильный формат ФИО" });
                                    //Добавление в dataGridView1 строки
                                    dataGridView1.Rows.Add(row);
                                    //Удаление из dataGridViewMain этой строки
                                    dataGridViewMain.Rows.RemoveAt(i);
                                    //Уменьшает цикл по датагриду 
                                    i--;
                                    //Завершает текущую итерацию
                                    continue;
                                }*/

                                //заменяет точки на пробелами
                                dataGridViewMain[0, i].Value = rgxFIO.Replace(dataGridViewMain[0, i].Value.ToString(), " ");
                                //удаляет крайние пробелы
                                dataGridViewMain[0, i].Value = dataGridViewMain[0, i].Value.ToString().Trim();
                                //заменяет сдвоеные (и более) пробелы одним
                                dataGridViewMain[0, i].Value = rgxProbel2.Replace(dataGridViewMain[0, i].Value.ToString(), " ");                               
                                //заменяет (удаляет) все символы кроме букв, тире и пробела на пустые
                                dataGridViewMain[0, i].Value = Regex.Replace(dataGridViewMain[0, i].Value.ToString(), @"[^\w\s@-]", "", RegexOptions.None);
                                //dataGridViewMain[0, i].Value = Regex.Replace(dataGridViewMain[0, i].Value.ToString(), @"^[а-яА-ЯёЁ]+$", RegexOptions.Compiled);
 

                                //проверяет количество слов в ФИО (с помощью подсчета пробелов)
                                string[] tempStr = dataGridViewMain[0, i].Value.ToString().Split(' ');
                                int kolProbelov = 0;
                                foreach (string _s in tempStr)
                                {
                                    kolProbelov++;
                                }
                                if (((kolProbelov - 1) < 1) && ((textBox4.Text == "FI") || (textBox4.Text == "F.I") || (textBox4.Text == "F-I") || (textBox4.Text == "F_I") || (textBox4.Text == "IF") || (textBox4.Text == "I.F") || (textBox4.Text == "I-F") || (textBox4.Text == "I_F") || (textBox4.Text == "FO") || (textBox4.Text == "F.O") || (textBox4.Text == "F-O") || (textBox4.Text == "F_O") || (textBox4.Text == "OF") || (textBox4.Text == "O.F") || (textBox4.Text == "O-F") || (textBox4.Text == "O_F")))
                                {
                                    MessageBox.Show("ФИО - '" + dataGridViewMain[0, i].Value.ToString() + "' Для выбранной формулы УЗ необходимо Имя пользователя");
                                    row.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                         dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                         dataGridViewMain[6, i].Value.ToString(), dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                         dataGridViewMain[9, i].Value.ToString(), "Не правильный формат ФИО" });
                                    dataGridView1.Rows.Add(row);
                                    dataGridViewMain.Rows.RemoveAt(i);
                                    i--;
                                    continue;
                                }
                                if (((kolProbelov - 1) < 2) && ((textBox4.Text == "FIO") || (textBox4.Text == "F.IO") || (textBox4.Text == "F-IO") || (textBox4.Text == "F_IO") || (textBox4.Text == "IOF") || (textBox4.Text == "IO.F") || (textBox4.Text == "IO-F") || (textBox4.Text == "IO_F") || (textBox4.Text == "FOI") || (textBox4.Text == "F.OI") || (textBox4.Text == "F-OI") || (textBox4.Text == "F_OI") || (textBox4.Text == "OIF") || (textBox4.Text == "OI.F") || (textBox4.Text == "OI-F") || (textBox4.Text == "OI_F")))
                                {
                                    MessageBox.Show("ФИО - '" + dataGridViewMain[0, i].Value.ToString() + "' Для выбранной формулы УЗ необходимо Имя Отчество пользователя");
                                    row.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                         dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                         dataGridViewMain[6, i].Value.ToString(), dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                         dataGridViewMain[9, i].Value.ToString(), "Не правильный формат ФИО" });
                                    dataGridView1.Rows.Add(row);
                                    dataGridViewMain.Rows.RemoveAt(i);
                                    i--;
                                    continue;
                                }
                                if ((kolProbelov - 1) > 2)
                                {
                                    MessageBox.Show("ФИО - '" + dataGridViewMain[0, i].Value.ToString() + "' содержит больше 3-х слов");
                                    row.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                        dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                        dataGridViewMain[6, i].Value.ToString(), dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                        dataGridViewMain[9, i].Value.ToString(), "Не правильный формат ФИО" });
                                    dataGridView1.Rows.Add(row);
                                    dataGridViewMain.Rows.RemoveAt(i);
                                    i--;
                                    continue;
                                }

                            }
                        }

                        // ******** Телефонный номер ********

                        // цикл по столбцу Телефон в датагрде
                        Regex rgxPhone495 = new Regex(@"^.*(495).*((\d{3}[\-\s+]\d{4})|(\d{3}[\-\s+]\d\d[\-\s+]\d\d)|(\d{7})|(\d{3}[\-\s+]\d[\-\s+]\d{3})|(\d{2}[\-\s+]\d{5})|(\d{2}[\-\s+]\d{3}[\-\s+]\d{2})|(\d{2}[\-\s+]\d{2}[\-\s+]\d{3})).*$");
                        Regex rgxPhoneXXX = new Regex(@"^.*(\d\d\d).*((\d{3}[\-\s+]\d{4})|(\d{3}[\-\s+]\d\d[\-\s+]\d\d)|(\d{7})|(\d{3}[\-\s+]\d[\-\s+]\d{3})|(\d{2}[\-\s+]\d{5})|(\d{2}[\-\s+]\d{3}[\-\s+]\d{2})|(\d{2}[\-\s+]\d{2}[\-\s+]\d{3})).*$");

                        for (int i = 1; i < dataGridViewMain.RowCount; i++)
                        {
                            if (dataGridViewMain[4, i].Value != null)
                            {
                                Match match495 = rgxPhone495.Match(dataGridViewMain[4, i].Value.ToString());
                                if (match495.Success)
                                {
                                    dataGridViewMain[4, i].Value = rgxPhone495.Replace(dataGridViewMain[4, i].Value.ToString(), "$2");
                                    //MessageBox.Show(dataGridViewMain[4, i].Value.ToString());
                                }
                                else
                                {
                                    Match matchXXX = rgxPhoneXXX.Match(dataGridViewMain[4, i].Value.ToString());
                                    if (matchXXX.Success)
                                    {
                                        dataGridViewMain[4, i].Value = rgxPhoneXXX.Replace(dataGridViewMain[4, i].Value.ToString(), "($1)$2");
                                        //MessageBox.Show(dataGridViewMain[4, i].Value.ToString());
                                    }
                                    /* else
                                     { 
                            
                                     }*/
                                }
                            }
                        }

                        // ******** Адрес ********
                        dataGridViewMain[5, 0].Value = "Адрес";
                        Regex rgxUlitsa = new Regex(@"^\s?ул.\s+|^\s?ул\s+|^\s?улица\s+| ул.\s?$| ул\s?$| улица\s?$| ул.,\s?$|^\s?Ул.\s+|^\s?Ул\s+|^\s?Улица\s+| Ул.\s?$| Ул\s?$| Улица\s?$| Ул.,\s?$", RegexOptions.IgnoreCase);
                        Regex rgxDom = new Regex(@"^\s?д.\s?|^\s?д\s?|^\s?дом\s?| д.\s?$| д\s?$| дом\s?$|^\s?Д.\s?|^\s?Д\s?|^\s?Дом\s?| Д.\s?$| Д\s?$| Дом\s?$", RegexOptions.IgnoreCase);
                        Regex rgxMoscow = new Regex(@"Москва", RegexOptions.IgnoreCase);

                        // цикл по столбцам улица, дом, город и company
                        for (int i = 1; i < dataGridViewMain.RowCount; i++)
                        {
                            //редактирование улици
                            if (dataGridViewMain[5, i].Value != null)
                            {
                                if (Regex.IsMatch(dataGridViewMain[5, i].Value.ToString(), rgxUlitsa.ToString()))
                                {
                                    dataGridViewMain[5, i].Value = rgxUlitsa.Replace(dataGridViewMain[5, i].Value.ToString(), string.Empty);
                                    dataGridViewMain[5, i].Value += " ул., ";
                                }
                                else
                                {
                                    //dataGridViewMain[5, i].Value += ", ";
                                }
                            }

                            //редактирование дома
                            if (dataGridViewMain[6, i].Value != null)
                            {
                                if (!Regex.IsMatch(dataGridViewMain[6, i].Value.ToString(), rgxMoscow.ToString()))
                                {
                                    dataGridViewMain[6, i].Value = rgxDom.Replace(dataGridViewMain[6, i].Value.ToString(), string.Empty);
                                    dataGridViewMain[5, i].Value += ", " + dataGridViewMain[6, i].Value.ToString();
                                }
                            }

                            //редактирование company
                            dataGridViewMain[9, i].Value = comboBox1.Text;

                            // !!! Доделать условие. Добавить выбор города : Москва, Зеленоград
                            dataGridViewMain[6, i].Value = " ";

                        }
                        dataGridViewMain[6, 0].Value = "Город";
                        //dataGridViewMain.Columns[6].Visible = false;

                        // ******** Заглавные буквы, Сдвоенные пробелы, Пробелы по краям, Проверка длины строк  ********
                        Regex rgxProbel = new Regex("\\s+");
                        //Regex rgxProbelS_E = new Regex(@"^\s+|\s+$");
                        for (int i = 1; i < dataGridViewMain.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridViewMain.Columns.Count; j++)
                            {
                                if (dataGridViewMain.Rows[i].Cells[j].Value != null)
                                {
                                    //заменяет сдвоеные (и более) пробелы на один
                                    dataGridViewMain[j, i].Value = rgxProbel.Replace(dataGridViewMain[j, i].Value.ToString(), " ");

                                    //убирает пробелы по краям
                                    dataGridViewMain[j, i].Value = dataGridViewMain[j, i].Value.ToString().Trim();
                                    /*if (Regex.IsMatch(dataGridViewMain[j, i].Value.ToString(), rgxProbelS_E.ToString()))
                                        dataGridViewMain[j, i].Value = rgxProbelS_E.Replace(dataGridViewMain[j, i].Value.ToString(), string.Empty);*/

                                    //Делает первую букву заглавной, для каждой ячейки датаГрида
                                    dataGridViewMain[j, i].Value = Regex.Replace(dataGridViewMain[j, i].Value.ToString(), @"^\s?[а-я]", m => m.Value.ToUpper());

                                    //проверка длины строк
                                    if (dataGridViewMain[j, i].Value.ToString().Length > 62)
                                    {
                                        dataGridViewMain[j, i].Value = dataGridViewMain[j, i].Value.ToString().Substring(0, 62);
                                    }

                                }
                            }
                        }

                        // ******** Логин ********
                        string tempLogin = string.Empty;
                        Regex rgxY = new Regex(@"ий", RegexOptions.IgnoreCase);
                        Regex rgxIE = new Regex(@"ье", RegexOptions.IgnoreCase);

                        // цикл столбцу ФИО
                        for (int i = 1; i < dataGridViewMain.Rows.Count; i++)
                        {
                            tempLogin = string.Empty;
                            //замена буквосочетаний 
                            string tempFIO = dataGridViewMain[0, i].Value.ToString(); // на время состовления логина запаминаем ФИО 
                            dataGridViewMain[0, i].Value = rgxY.Replace(dataGridViewMain[0, i].Value.ToString(), "y");
                            dataGridViewMain[0, i].Value = rgxIE.Replace(dataGridViewMain[0, i].Value.ToString(), "ie");

                            //формирование массива [ Фамилия, Имя, Отчество ]
                            arrFIO = dataGridViewMain[0, i].Value.ToString().Split(' '); //ФИО
                            formulaUZ = textBox4.Text;
                            //цикл из формулы УЗ делает логин
                            for (int g = 0; g < formulaUZ.Length; g++)
                            {
                                switch (formulaUZ[g].ToString())
                                {
                                    case "F": tempLogin = tempLogin + F(arrFIO[0]); break;
                                    case "I": tempLogin = tempLogin + I(arrFIO[1]); break;
                                    case "O": tempLogin = tempLogin + O(arrFIO[2]); break;
                                    case ".": tempLogin = tempLogin + "."; break;
                                    case "-": tempLogin = tempLogin + "-"; break;
                                    case "_": tempLogin = tempLogin + "_"; break;
                                    default: MessageBox.Show("Ошибка в формуле"); break;
                                }
                            }

                            dataGridViewMain[7, i].Value = tempLogin; // логин в датаГрид
                            dataGridViewMain[0, i].Value = tempFIO; // возвращаем исходное ФИО в датаГрид
                            dataGridViewMain[8, i].Value = GetPass(rnd, longPass); //пароль в датаГрид
                        }

                        key = false;
                    }
                    catch
                    {
                        MessageBox.Show("Необходимо импортировать файл. Затем повторить проверку.");
                        //Environment.Exit(0);
                    }

                }
                else
                {
                    MessageBox.Show("Нельзя проверять дважды");
                }
            }
            else
            {
                MessageBox.Show("Выберите вид учетной записи");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        // *********функция генерации пароля***********
        string GetPass(Random rnd, int longPass)
        {
            string all = "qwertyupasdfghjklzxcvbnmQWERTYUPASDFGHJKZXCVBNM123456789"; //набор символов для пароля
            string abc = "qwertyupasdfghjkzxcvbnm"; //набор символов для пароля
            string ABC = "QWERTYUPASDFGHJKZXCVBNM"; //набор символов для пароля
            string digit = "23456789"; //набор символов для пароля
            //int kol = 8; // кол-во символов
            string result = "";

            //Random rnd = new Random();          
            int lng = all.Length;
            int lng2 = abc.Length;
            int lng3 = ABC.Length;
            int lng4 = digit.Length;

            for (int j = 0; j < longPass; j++)
            {
                if (j == 1)
                    result += digit[rnd.Next(lng4)];
                else
                    if (j == 2)
                        result += abc[rnd.Next(lng2)];
                    else
                        if (j == 4)
                            result += ABC[rnd.Next(lng3)];
                        else
                            result += all[rnd.Next(lng)];
            }
            return result;
        }
        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

 //----------------------------------------------Поиск пользователя(ей) в AD--------------------------------------------------------

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.PasswordChar = '*'; // скрывакет вводимые символы
        }
           
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        // Метод проверяет существует ли SAMAccountName в AD 
        static bool UserAccountExist(string newAccount, string PathLDAP, string Username, string Password)
        {
            bool isAccountExist = false;

            DirectoryEntry entry = new DirectoryEntry();
            entry.Username = Username;
            entry.Password = Password;
            entry.Path = PathLDAP;  //"LDAP://192.168.100.1/DC=test,DC=local"; // Путь до контроллера домена

            DirectorySearcher searchaccount = new DirectorySearcher();
            searchaccount.SearchRoot = entry;
            searchaccount.Filter = string.Format("(SAMAccountName={0})", newAccount );
            //SearchResult account = searchaccount.FindOne(); // поиск до первого найденого
            try
            {
                SearchResultCollection account = searchaccount.FindAll();
                var check = (account.Count > 0) ? isAccountExist = true : isAccountExist = false; // поиск всех совпадении (true - совпадения найдены)
            }
            catch
            {
                MessageBox.Show("Не определено имя пользователя или пароль.\nВозможно подключение к AD идет через несколько VPN.\nНе верно указан LDAP-запрос.");
                Environment.Exit(0);
            }
            return isAccountExist;
        }   
        
        private void button2_Click(object sender, EventArgs e)
        {
            if ((key == false) && (key2 == true) && (textBox1.Text != "") && (textBox2.Text != "") && (textBox3.Text != ""))
            {
                // Пользователь с правами на AD
                Username = textBox2.Text;
                // Пароль для Username. Скрывает вводимый текст звездочками
                Password = textBox1.Text;
                // Путь до контролера домена в формате LDAP
                PathLDAP = textBox3.Text;
                // для userPrincipalName
                string UPN = "@";
                // счетчик совпадении
                int kolSovp = 0;

                // получение части имени userPrincipalName из LDAP-запроса 
                if (PathLDAP.Contains("DC=")) //строка Содержит подстроку
                {
                    string tempPathLDAP = PathLDAP; //LDAP будет редактироваться
                    string str2="DC=";
                    string tempUPN = "";
                    int in1 = 0;
                    int count = 1;
                    int temp = 0;
                    int tempLen = 0;

                    while (tempPathLDAP.IndexOf(str2) != -1) //пока есть вхождения подстроки "DC=" в строке PathLDAP 
                    {
                        temp = tempPathLDAP.IndexOf(str2);
                        if (count > 1)
                        {
                            tempLen = temp - in1;
                            tempUPN = tempPathLDAP.Substring(in1, tempLen - 1);
                            UPN = UPN + tempUPN + ".";
                        }
                        tempPathLDAP = tempPathLDAP.Remove(temp, str2.Length); // удаление вхождения подстроки
                        count++;
                        in1 = temp;
                    }
                    temp=tempPathLDAP.Length;
                    tempLen = temp - in1;
                    tempUPN = tempPathLDAP.Substring(in1, tempLen); // удаление вхождения подстроки
                    UPN = UPN + tempUPN; // результат
                }
                else
                {
                    MessageBox.Show("Не правильно оформлен LDAP-запрос.\nОбразец: LDAP://192.168.100.1/DC=test,DC=local");
                    return;
                }

                //цикл по всему датаГриду
                for (int i = 1; i < dataGridViewMain.Rows.Count; i++)
                {
                    // Аккаунд нового пользователя, для проверки (поиска) на существование
                    newAccount = dataGridViewMain[7, i].Value.ToString(); //textBox4.Text;
                    // Полное имя нового пользователя, для проверки (поиска) если аккаунд совпал
                    newFullname = dataGridViewMain[0, i].Value.ToString(); //textBox5.Text;
                    // Пароль для нового пользователя
                    newPassword = dataGridViewMain[8, i].Value.ToString();

                    //bool isuseraccountExist = false;
                    //bool modyficationLogin = false;
                    bool notCreate = false;
                    string notify = null;
                    string[] FIOStr = null;
                    string propertiesUser = null;
                    //int login_i = 1;

                    FIOStr = newFullname.Split(' '); //ФИО

             //*******************поиск по ФИО (новый)
                    bool isExist = false;

                    DirectoryEntry entry = new DirectoryEntry();
                    entry.Username = Username;
                    entry.Password = Password;
                    entry.Path = PathLDAP;

                    DirectorySearcher searchusername = new DirectorySearcher();
                    searchusername.SearchRoot = entry;
                    searchusername.Filter = "(&(objectCategory=user)(cn=" + newFullname + "))"; // фильтр поиска
                    // формируется список своиств, которые нас интересуют
                    //searchusername.PropertiesToLoad.Clear();
                    searchusername.PropertiesToLoad.Add("canonicalName");
                    //searchusername.PropertiesToLoad.Add("cn");
                    try
                    {
                        SearchResultCollection result = searchusername.FindAll();
                        var validate = (result.Count > 0) ? isExist = true : isExist = false; // поиск всех совпадении (true - совпадения найдены)

                        if (isExist)
                        {
                            // полное совпадения ФИО
                            notify = newFullname + " Уже существует";
                            // находит своиство пользователя - canonicalName. которое показывает место нахождение пользователя
                            foreach (SearchResult res in result)
                            {
                                propertiesUser += res.Properties["canonicalName"][0].ToString() + "  ";
                                // вставить в ДатаГрид
                                DataGridViewRow row2 = new DataGridViewRow();
                                row2.CreateCells(dataGridViewMain);
                                //формируеться колекция стороки 
                                row2.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                         dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                         dataGridViewMain[6, i].Value.ToString(), /*dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                         dataGridViewMain[9, i].Value.ToString(),*/ "Пользователь с таким ФИО уже существует: " +propertiesUser });
                                //Добавление в dataGridView1 строки
                                dataGridView1.Rows.Add(row2);
                                //Удаление из dataGridViewMain этой строки
                                dataGridViewMain.Rows.RemoveAt(i);
                                //Возвращает цикл на предыдущую итерацию (учитывая уменьшение цикла (уменьшение строк в датаГриде))
                                i--;
                            }
                            notCreate = true; //не будет создоваться УЗ
                            kolSovp += 1; //счетчик совпадений
                            continue; // досрочное завершение итерации. переход на следующую строку ДатаГрид
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Не определено имя пользователя или пароль.\nВозможно подключение к AD идет через несколько VPN.\nНе верно указан LDAP-запрос.");
                        //Environment.Exit(0);
                        return;
                    }

          //**************Поиск совпадений по логину (новый)

                    bool isAccountExist = false;

                    DirectoryEntry entryL = new DirectoryEntry();
                    entryL.Username = Username;
                    entryL.Password = Password;
                    entryL.Path = PathLDAP;  //"LDAP://192.168.100.1/DC=test,DC=local"; // Путь до контроллера домена

                    DirectorySearcher searchaccount = new DirectorySearcher();
                    searchaccount.SearchRoot = entryL;
                    searchaccount.Filter = string.Format("(SAMAccountName={0})", newAccount); // фильтр поиска
                    searchaccount.PropertiesToLoad.Add("canonicalName"); // формируется список своиств, которые нас интересуют 
                    //SearchResult account = searchaccount.FindOne(); // поиск до первого найденого
                    try
                    {
                        SearchResultCollection account = searchaccount.FindAll();
                        var check = (account.Count > 0) ? isAccountExist = true : isAccountExist = false; // поиск всех совпадении (true - совпадения найдены)

                        if (isAccountExist)
                        {
                            // полное совпадения Логина
                            notify = newFullname + " Уже существует";
                            // находит своиство пользователя - canonicalName. которое показывает место нахождение пользователя
                            foreach (SearchResult res in account)
                            {
                                propertiesUser += res.Properties["canonicalName"][0].ToString() + "  ";
                                // вставить в ДатаГрид
                                DataGridViewRow row2 = new DataGridViewRow();
                                row2.CreateCells(dataGridViewMain);
                                //формируеться колекция стороки 
                                row2.SetValues(new object[] { dataGridViewMain[0, i].Value.ToString(), dataGridViewMain[1, i].Value.ToString(), dataGridViewMain[2, i].Value.ToString(), 
                                                         dataGridViewMain[3, i].Value.ToString(), dataGridViewMain[4, i].Value.ToString(), dataGridViewMain[5, i].Value.ToString(), 
                                                         dataGridViewMain[6, i].Value.ToString(), /*dataGridViewMain[7, i].Value.ToString(), dataGridViewMain[8, i].Value.ToString(), 
                                                         dataGridViewMain[9, i].Value.ToString(),*/ "Пользователь с таким ЛОГИНОМ уже существует: " +propertiesUser });
                                //Добавление в dataGridView1 строки
                                dataGridView1.Rows.Add(row2);
                                //Удаление из dataGridViewMain этой строки
                                dataGridViewMain.Rows.RemoveAt(i);
                                //Возвращает цикл на предыдущую итерацию (учитывая уменьшение цикла (уменьшение строк в датаГриде))
                                i--;
                            }
                            notCreate = true; //не будет создоваться УЗ
                            kolSovp += 1; //счетчик совпадений
                            continue; //  досрочное завершение итерации. переход на следующую строку ДатаГрид
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Не определено имя пользователя или пароль.\nВозможно подключение к AD идет через несколько VPN.\nНе верно указан LDAP-запрос.");
                        //Environment.Exit(0);
                        return;
                    }

                    // **************создание УЗ**************
                    if (!notCreate)
                    {
                        try
                        {
                            DirectoryEntry obDirEntry = new DirectoryEntry();
                            obDirEntry.Username = Username;
                            obDirEntry.Password = Password;
                            obDirEntry.Path = PathLDAP; //"LDAP://192.168.100.1/OU=TempProgrammKraftway,DC=test,DC=local";
                            // создаем пользователя 
                            DirectoryEntry obUser = obDirEntry.Children.Add("CN=" + newFullname, "user");
                            obUser.CommitChanges();
                            //Устанавливаем пароль
                            obUser.Invoke("SetPassword", newPassword);
                            //Устанавливаем логин (Logon Name)
                            obUser.Properties["userPrincipalName"].Value = newAccount + UPN; /* + "@test.local";*/
                            obUser.Properties["samAccountName"].Value = newAccount;

                            //Активируем созданную учетную запись
                            const int UF_PASSWD_NOTREQD = 0x0020;
                            const int UF_NORMAL_ACCOUNT = 0x0200;
                            const int UF_PASSWORD_EXPIRED = 0x800000;
                            obUser.Properties["userAccountControl"][0] = UF_NORMAL_ACCOUNT + UF_PASSWD_NOTREQD + UF_PASSWORD_EXPIRED;
                            obUser.CommitChanges();

                            //заполнение атрибутов УЗ
                            obUser.Properties["givenName"].Value = (FIOStr[1]);
                            obUser.Properties["sn"].Value = FIOStr[0];
                            obUser.Properties["displayName"].Value = newFullname;
                            if (dataGridViewMain[3, i].Value.ToString() != "")
                                obUser.Properties["physicalDeliveryOfficeName"].Value = dataGridViewMain[3, i].Value.ToString();
                            if (dataGridViewMain[4, i].Value.ToString() != "")
                                obUser.Properties["telephoneNumber"].Value = dataGridViewMain[4, i].Value.ToString();

                            if (dataGridViewMain[5, i].Value.ToString() != "")
                                obUser.Properties["streetAddress"].Value = dataGridViewMain[5, i].Value.ToString();
                            if (dataGridViewMain[6, i].Value.ToString() != "")
                                obUser.Properties["l"].Value = dataGridViewMain[6, i].Value.ToString();

                            if (dataGridViewMain[2, i].Value.ToString() != "")
                                obUser.Properties["Title"].Value = dataGridViewMain[2, i].Value.ToString();
                            if (dataGridViewMain[1, i].Value.ToString() != "")
                                obUser.Properties["Department"].Value = dataGridViewMain[1, i].Value.ToString();
                            if (dataGridViewMain[9, i].Value.ToString() != "")
                                obUser.Properties["Company"].Value = dataGridViewMain[9, i].Value.ToString();
                            obUser.CommitChanges();
                        }
                        catch
                        {
                            MessageBox.Show("Стоп. Что-то пошло не так!\nРекомендую удалить всех только что созданых пользователей. По списку до " + newFullname);
                            return;
                        }
                    }
                }
                MessageBox.Show("Создание УЗ завершено\nКолличество совпадении = " + kolSovp);
                key2 = false;
            }
            else 
            {
                MessageBox.Show("1) Необходидо запустить Проверку файла и повторить действие.\n2)Не указан LDAP-запрос для подключения или Имя пользователя или Пароль");
            }
        }

        /*private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }*/

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }


//фамилия
        private void button3_Click(object sender, EventArgs e)
        {
            if (keyF)
            {
                textBox4.Text = textBox4.Text+"F";
                keyF = false;
            }
        }
//Имя
        private void button4_Click(object sender, EventArgs e)
        {
            if (keyI)
            {
                textBox4.Text = textBox4.Text + "I";
                keyI = false;
            }
        }
//Отчество
        private void button5_Click(object sender, EventArgs e)
        {
            if (keyO)
            {
                textBox4.Text = textBox4.Text+"O";
                keyO = false;
            }
        }
//Точка
        private void button6_Click(object sender, EventArgs e)
        {
            if (keyTochka)
            {
                textBox4.Text = textBox4.Text + ".";
                keyTochka = false;
            }
        }
//Тире
        private void button8_Click(object sender, EventArgs e)
        {
            if (keyTire)
            {
                textBox4.Text = textBox4.Text + "-";
                keyTire = false;
            }
        }
//Подчеркивание
        private void button7_Click(object sender, EventArgs e)
        {
            if (keyPodcher)
            {
                textBox4.Text = textBox4.Text + "_";
                keyPodcher = false;
            }
        }
//Стереть
        private void button9_Click(object sender, EventArgs e)
        {
            if (keySteret)
            {
                keyF = true;
                keyI = true;
                keyO = true;
                keyTochka = true;
                keyTire = true;
                keyPodcher = true;

                formulaUZ = null;
                textBox4.Text = null;
            }
            else
                MessageBox.Show("Нельзя вносить изменения в процессе создания учетных записей.\nПосле иморта нового файла, функция будет доступна.");
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }     
    }
}