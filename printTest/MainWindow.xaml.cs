using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management;
using System.Data;
using System.Windows.Forms;
using System.IO;
using printTest.modul_s;
using System.Threading.Tasks;
using Spire.Pdf.Exporting.XPS.Schema;
using System.Timers;
using System.Windows.Threading;
using static System.Windows.Forms.AxHost;
using System.Collections;
using Spire.Pdf;
using System.Drawing.Printing;
using static System.Windows.Forms.LinkLabel;
using System.Net;
using Path = System.IO.Path;
using System.Reflection;
using System.ComponentModel;
using DataGrid = System.Windows.Controls.DataGrid;

namespace printTest
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        // переменные путей
        string directoryName = null; string dbName = null;
        // переменные состояний
        bool clo = false; int statusPrinter = 0; bool pause = false;
        // переменные
        string print_name = null; double envelopeC5 = 0.00; double envelopeDL = 0.00; double onePage = 0.04;
        string delimetr = ",";
        DataTable dt = new DataTable();
        #region Загрузка и Закрытие формы Иные компоненты
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            version.Content=Assembly.GetExecutingAssembly().GetName().Version.ToString();
            string delimetr = (System.Convert.ToDouble(1.0 / 2)).ToString().Substring(1, 1);
            dt.Columns.Add("Статус");
            dt.Columns.Add("Имя файла");
            dt.Columns.Add("Адрес");
            if (!File.Exists($@"{Environment.CurrentDirectory}\ini\settings.ini"))
            {
                OneRowPanel.IsEnabled = false;
                System.Windows.MessageBox.Show("Критическая ошибка\nОтсутствует файл настроек settings.ini\nОбратитесь к файлу справка",
                    "Ошибка", MessageBoxButton.OK);
                return;
            }
            var SettingsIni = new IniFile($@"{Environment.CurrentDirectory}\ini\settings.ini");
            if (!SettingsIni.KeyExists("envelopeC5", "main"))
            {
                System.Windows.MessageBox.Show("Отсутствует значение 'envelopeC5' (вес конверта С5) файл настроек settings.ini\n Будет использованно значение по умолчанию\nОбратитесь к файлу справка",
                   "Ошибка", MessageBoxButton.OK);
            }
            else
            {
                envelopeC5 = System.Convert.ToDouble(SettingsIni.Read("envelopeC5", "main").Replace(".", delimetr).Replace(",", delimetr));
            }
            if (!SettingsIni.KeyExists("envelopeDL", "main"))
            {
                System.Windows.MessageBox.Show("Отсутствует значение 'envelopeDL' (вес конверта DL) файл настроек settings.ini\n Будет использованно значение по умолчанию\nОбратитесь к файлу справка",
                   "Ошибка", MessageBoxButton.OK);
            }
            else
            {
                envelopeDL = System.Convert.ToDouble(SettingsIni.Read("envelopeDL", "main").Replace(".", delimetr).Replace(",", delimetr));
            }
            if (!SettingsIni.KeyExists("onePage", "main"))
            {
                System.Windows.MessageBox.Show("Отсутствует значение 'onePage' (вес конверта Одного листа) файл настроек settings.ini\n Будет использованно значение по умолчанию\nОбратитесь к файлу справка",
                   "Ошибка", MessageBoxButton.OK);
            }
            else
            {
                envelopeC5 = System.Convert.ToDouble(SettingsIni.Read("onePage", "main").Replace(".", delimetr).Replace(",", delimetr));
            }

            Thread invaitToStepTh = new Thread(InvaitToStep);
            invaitToStepTh.Start();
            Thread LoadPrinterTh = new Thread(LoadPrinter);
            LoadPrinterTh.Start();

        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            clo = true;
        }
        #endregion
        #region Поиск принтеров в ситеме вывод найденых принтеров в combobox choosePrinter
        private void LoadPrinter()
        {
            Dispatcher.Invoke(() => {
                choosePrinter.IsEnabled = false;
                choosePrinter.Items.Add("выбрать принтер");});
            
            ManagementObjectSearcher PrinterSet = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");

            foreach (ManagementObject Printer in PrinterSet.Get())
            {
                Dispatcher.Invoke(() => choosePrinter.Items.Add(Printer.Properties["Name"].Value.ToString()));
            }
            Dispatcher.Invoke(() => { choosePrinter.IsEnabled = true;
                choosePrinter.SelectedValue= "выбрать принтер";
            });

        }
        #endregion
        private void docCount_PreviewTextInput(object sender, TextCompositionEventArgs e)//фильтр ввода
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c))
                {
                    e.Handled = true;
                    break;
                }
            }
        }
        private void printEnd_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c))
                {
                    e.Handled = true;
                    break;
                }
            }
        }
        private void printStart_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c))
                {
                    e.Handled = true;
                    break;
                }
            }
        }
        private void InvaitToStep() // определенине шагов для последующей работы
        {
            while ((directoryName is null || dbName is null || print_name is null) && !clo)
            {
                if (directoryName is null)
                {
                    Dispatcher.Invoke(() =>
                                    statusBarLabel.Content = "Начните работу с выбора каталога ");
                    Dispatcher.Invoke(() =>
                                         buttonOpenDBFile.IsEnabled = false);
                    Dispatcher.Invoke(() =>
                                         Print.IsEnabled = false);
                }
                else if (dbName is null)
                {
                    Dispatcher.Invoke(() =>
                                       statusBarLabel.Content = "Продолжите работу с выбора файла БД содержащую адреса НП ");
                    Dispatcher.Invoke(() =>
                                              Print.IsEnabled = false);
                }
                else if (print_name is null)
                {
                    Dispatcher.Invoke(() =>
                                      statusBarLabel.Content = "Выберите принтер");
                    Dispatcher.Invoke(() =>
                                         Print.IsEnabled = false);
                }

                if (!clo)
                {
                    Dispatcher.Invoke(() =>
                                       statusBarLabel.Background = Brushes.Coral);
                }
                Thread.Sleep(1000);
                if (!clo)
                {
                    Dispatcher.Invoke(() =>
                                           statusBarLabel.Background = Brushes.Transparent);
                }
                Thread.Sleep(1000);
                Console.WriteLine($"{directoryName}-{dbName}-{print_name}");
            }
            if (print_name != null && directoryName != null && dbName != null)
            {

                Dispatcher.Invoke(() => Print.IsEnabled = true);
                Dispatcher.Invoke(() => statusBarLabel.Content = "Отсортируйте список по имени файла или столбцу 'Адрес'");
            }

        }
        #region Выбор путей к каталогу файлов и БД приведение списка
        private void openDirectory_Click(object sender, RoutedEventArgs e) //выбор каталога с файлами
        {
            directoryName = null;
            Dispatcher.Invoke(() => buttonOpenDBFile.IsEnabled = false);
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                while (dt.Rows.Count - 1 >= 0)
                {
                    dt.Rows.RemoveAt(0);
                }
                string path = dialog.SelectedPath;
                DataTable dtTemp = new DataTable();
                dtTemp.Columns.Add("Статус");
                dtTemp.Columns.Add("Имя файла");
                dtTemp.Columns.Add("Адрес");
                dt.Merge(get_directories(path, dtTemp));
                if (dt.Rows.Count - 1 >= 0) { directoryName = path; Dispatcher.Invoke(() => buttonOpenDBFile.IsEnabled = true); }
                DataGridView1.ItemsSource = dt.DefaultView;
                DataView view = new DataView();
                view = dt.DefaultView;
                view.Sort = "Имя файла";
                DataGridView1.ItemsSource = (System.Collections.IEnumerable)view;
                ProgressBar1.Maximum = DataGridView1.Items.Count;
                statusBarAllLabel.Content = $"Документов в каталоге: {dt.Rows.Count}";
            }
        }
        private DataTable get_directories(string _path, DataTable dtTemp) //поиск файлов в папке
        {

            foreach (string dir_name in Directory.GetDirectories(_path))
            {
                get_directories(dir_name, dtTemp);
            }

            foreach (string file_name in Directory.GetFiles(_path, "*.pdf"))
            {
                dt.Rows.Add("не определено", file_name.Replace($@"{_path}\", string.Empty), "не загружался");
            }
            return dtTemp;
        }
        private void buttonOpenDBFile_Click(object sender, RoutedEventArgs e) //выбор файла БД с адресами
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "db files (*.db)|*.db";
            openFileDialog.FilterIndex = 1;
            openFileDialog.FileName = "return";
            openFileDialog.Multiselect = false;
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = openFileDialog.FileName; // Путь к файлу
                SQLite SQLlite = new SQLite();
                DataTable dtAddress = SQLlite.FillGrid(filePath, "select * from quest");
                object lockObject = new object();
                int searchCount = 0;
                Parallel.ForEach(dt.AsEnumerable(), row =>
                {
                    string searchRow = row["Имя файла"].ToString().ToLower().Replace(".pdf", string.Empty);
                    foreach (DataRow rowslave in dtAddress.Rows)
                    {
                        string searchRowSlave = $"{rowslave["fio"].ToString().ToLower()}_{rowslave["inn"].ToString().ToLower()}";

                        if (searchRow == searchRowSlave)
                        {
                            lock (lockObject)
                            {
                                searchCount++;
                                row["Адрес"] = rowslave["address"];
                                break;
                            }
                        }
                    }

                });
                if (searchCount != dt.Rows.Count)
                {
                    if (!myPopup.IsOpen)
                    {
                        labelCurrentData.Content = $"Найдено адресов: {searchCount} из {dt.Rows.Count}";
                        myPopup.IsOpen = true;
                    }
                }
                dbName = filePath;
                DataGridView1.ItemsSource = dt.DefaultView;
            }
        }
        private void removeRowsNoAddress_Click(object sender, RoutedEventArgs e)//вызов функции удаления строк не содержащих адрес
        {
            removeRowsDTDatatable();
            if (dt.Rows.Count - 1 < 0) { dbName = null; directoryName = null; }
            myPopup.IsOpen = false;
        }
        void removeRowsDTDatatable() //удаление строк не содержащих адрес
        {
            try
            {
                int icount = 0;
                foreach (DataRow row in dt.Rows)
                {
                    if (row["Адрес"].ToString() == "не загружался")
                    {
                        dt.Rows.RemoveAt(icount);
                        removeRowsDTDatatable();
                    }
                    icount++;
                }
                DataGridView1.ItemsSource = dt.DefaultView;
            }
            catch (Exception e)
            {

            }
        }
        private void printNoAddress_Click(object sender, RoutedEventArgs e) //игнорирование строк не содержащих адрес
        {
            myPopup.IsOpen = false;
        }
        #endregion
        #region Функции печати и статуса принтера
        private void Pause_Click(object sender, RoutedEventArgs e) // события кнопки Пауза
        {
            if (pause)
            {
                pause = false;
                Pause.Content = "Пауза";
                Dispatcher.Invoke(() => Pause.Content = "Пауза");
                Pause.Background = Brushes.Red;
            }
            else
            {
                pause = true;
                Dispatcher.Invoke(() => Pause.Content = "Продолжить");
                Pause.Content = "Продолжить";
                Pause.Background = Brushes.Green;
            }
        }
        private string Status_print()//функция возращения строкового параметра состояния принтера
        {
            string stat = "";

            if (print_name != null) {
                ManagementObjectSearcher PrinterSet = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
                foreach (ManagementObject Printer in PrinterSet.Get())
                {
                    if (Printer.Properties["Name"].Value.ToString().ToLower() == print_name.ToLower())
                    {
                        stat = Printer.Properties["PrinterStatus"].Value.ToString();
                    }
                }
                switch (stat)
                {
                    case "3": { stat = "Готов"; break; };
                    case "4": { stat = "Печать"; break; };
                    case "5": { stat = "Предупреждение"; break; };
                    default: { stat = "Ошибка"; break; };
                } }
            return stat;
        }
        private void choosePrinter_SelectionChanged(object sender, SelectionChangedEventArgs e) // смена принтера
        {
            if (choosePrinter.SelectedValue != null)               
                    {
                        print_name = choosePrinter.SelectedValue.ToString();
                        labelStatusPrinter.Content = Status_print();
                        switch (labelStatusPrinter.Content)
                        {
                            case "Готов": { labelStatusPrinter.Foreground = Brushes.Green; break; }
                            case "Печать": { labelStatusPrinter.Foreground = Brushes.Blue; break; }
                            default: { labelStatusPrinter.Foreground = Brushes.Red; break; }
                        }                        
                    }                                      
                              
        }
        private void Print_Click(object sender, RoutedEventArgs e) //события кнопки Печать
        {
                       printDocMain();
        }
        private void printDocMain() //основная функция печати документов
        {
            if (!int.TryParse(docCount.Text, out int printCountVale) || printCountVale <= 0)
            {
                statusBarLabel.Content = "Ошибка при определении количества документов";
                docCount.Background = Brushes.Coral;
                return;
            }


            docCount.Background = Brushes.Green;
            PrinterStatusTracking();
            double massa = 0.00;
            massa = typeLetter.Text == "C5" ? envelopeC5
                    : typeLetter.Text == "DL" ? envelopeDL
                    : 0;
            int printTwoSideValue = printTwoSide.Text == "односторонняя" ? 1
                                    : printTwoSide.Text == "двухсторонняя" ? 2
                                    : 1;
            int printTwoListValue = printTwoList.Text == "1 на листе" ? 1
                                    : printTwoList.Text == "2 на листе" ? 2
                                    : 1;
            bool skeepRow = printNaRow.IsChecked.Value ? true
                                    : printNaRow.IsChecked.Value == false ? false
                                    : false;
            DataTable dtReport = new DataTable();
            dtReport.Columns.Add("Номер заказа"); dtReport.Columns.Add("ШПИ"); dtReport.Columns.Add("Масса");
            dtReport.Columns.Add("Стоимость пересылки (с НДС)"); dtReport.Columns.Add("Индекс"); dtReport.Columns.Add("Адрес");
            dtReport.Columns.Add("Телефон"); dtReport.Columns.Add("ФИО"); dtReport.Columns.Add("ID_PO");
            dtReport.Columns.Add("Комментарий"); dtReport.Columns.Add("Вид РПО");
            TwoRowPanel.IsEnabled = false;
            DataGridView1.CanUserSortColumns = false;
            Thread printTh = new Thread(() =>
            {
                Dispatcher.Invoke(() => Print.IsEnabled = false);               
                
                DataTable dtTemp = new DataTable();
                Dispatcher.Invoke(() => {
                    Pause.IsEnabled = true;
                    ProgressBar1.Maximum = dt.Rows.Count;
                    ProgressBar1.Value = 0;
                });
                string nameFolder = createFolder();
                int jCount = 0;
                int iCount = 0;
                foreach (DataRowView row in DataGridView1.Items)
                {
                    if (row["Статус"].ToString() == "не определено")
                    {
                        Console.WriteLine($"{skeepRow} {row["Адрес"].ToString()}");
                        if (!skeepRow && row["Адрес"].ToString() == "не загружался")
                        {
                            Dispatcher.Invoke(() => statusBarLabel.Content = "Документ пропущен");
                            row["Статус"] = $"документ пропущен";
                            continue;
                        }
                        while (pause)
                        {
                            if (clo) { return; }
                            Dispatcher.Invoke(() => statusBarLabel.Content = "Выполнение задания приостановленно");
                            Thread.Sleep(100);
                        }
                        while (statusPrinter != 3 )
                        {
                            if (clo) { return; }
                            Dispatcher.Invoke(() => statusBarLabel.Content = "Ожидание готовности принтера");
                            Thread.Sleep(100);
                        }
                        Dispatcher.Invoke(() => statusBarLabel.Content = "Документ отправлен на печать");
                        Dictionary<string, int> resultPrint = PrintDocSlave($@"{directoryName}\{row["Имя файла"].ToString()}", printTwoSideValue, printTwoListValue);
                        Dispatcher.Invoke(() => statusBarLabel.Content = "Документ распечатан");
                        row["Статус"] = $"распечатан, количество листов:{resultPrint["page"]}";
                        Dictionary<string, string> resultGetAddresstoRow = getAddresstoRow(row["Адрес"].ToString());
                        File.Move($@"{directoryName}\{row["Имя файла"]}", $@"{nameFolder}\{row["Имя файла"]}");
                        string[] parts = row["Имя файла"].ToString().Split('_')[0].Split(' ');
                        string fullName = $"{parts[0]} {parts[1]} {parts[2]}";
                        dtReport.Rows.Add();
                        dtReport.Rows[jCount]["Масса"] = resultPrint["page"] * onePage + massa;
                        dtReport.Rows[jCount]["Индекс"] = resultGetAddresstoRow["index"];
                        dtReport.Rows[jCount]["Адрес"] = resultGetAddresstoRow["address"];
                        dtReport.Rows[jCount]["ФИО"] = fullName;
                        dtReport.Rows[jCount]["Вид РПО"] = "Письмо";                        
                        jCount++;   
                        if (jCount == printCountVale)
                        {
                            datatable2excel datatable2Excel = new datatable2excel(dtReport, nameFolder);
                            nameFolder = createFolder();
                            while (dtReport.Rows.Count - 1 >= 0) { dtReport.Rows.RemoveAt(0); }
                            jCount = 0;
                        }
                        
                        iCount++;
                        Dispatcher.Invoke(() => ProgressBar1.Value = iCount);
                    }
                }
                if (dtReport.Rows.Count - 1 >= 0) { datatable2excel datatable2Excel = new datatable2excel(dtReport, nameFolder); }
                Dispatcher.Invoke(() => {
                    DataGridView1.CanUserSortColumns = true;
                    Pause.IsEnabled = false; TwoRowPanel.IsEnabled = true;
                    Print.IsEnabled = false;
                    directoryName = null; dbName = null; choosePrinter.SelectedValue="выбрать принтер"; print_name =null; 
                    
                });
                Thread invaitToStepTh = new Thread(InvaitToStep);
                invaitToStepTh.Start();
            });
            printTh.Start();
        }
        private string createFolder() 
        {
            string currentFolder = "";
            string rootFolder = $@"{Environment.CurrentDirectory}\outFolder";
            if (!Directory.Exists(rootFolder)) { Directory.CreateDirectory(rootFolder); }
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            string userName = Environment.UserName;
            Random random = new Random();
            string randomNumbers = "";
            for (int i = 0; i < 5; i++)
            {
                randomNumbers += random.Next(0, 10).ToString();
            }
            string directoryPath = Path.Combine(rootFolder, currentDate + "_" + userName + "_" + randomNumbers);
         
            if (!Directory.Exists(directoryPath))
            {
                
                Directory.CreateDirectory(directoryPath);
                currentFolder= directoryPath;
            }
            else
            {
                createFolder();
            }
            return currentFolder;
        }
        private Dictionary<string, string> getAddresstoRow(string row)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("index", ""); dict.Add("address", "");
            string[] parts = row.Split(',');
            if (int.TryParse(parts[0], out int index) && parts.Length>= 2)
            {
                string address = string.Join(",", parts, 1, parts.Length - 1);
                dict["index"]=index.ToString();
                dict["address"] = address.ToString();               
            }
            else
            {
                dict["index"] = "error_find";
                dict["address"] = "error_find";              
            }
            return dict;    
        }
        private Dictionary<string,int>  PrintDocSlave(string _path, int printTwoSideValue,int printTwoListValue)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            dict.Add("status", -1); dict.Add("page", -1);
            Dispatcher.Invoke(() =>
            {
                webview2panel.Source = new Uri(_path);
            });       
            PdfDocument pdfDoc = new PdfDocument();
            pdfDoc.LoadFromFile(_path);
            int pageCount = pdfDoc.Pages.Count;
            dict["page"] = pageCount;
            if (pageCount > 0) {                
                pdfDoc.PrintSettings.PrinterName = print_name;
                if (printTwoSideValue == 1) { pdfDoc.PrintSettings.Duplex = Duplex.Simplex; } else { pdfDoc.PrintSettings.Duplex = Duplex.Horizontal; }
                if (printTwoListValue == 2) { pdfDoc.PrintSettings.SelectMultiPageLayout(1, 2); } else { pdfDoc.PrintSettings.SelectMultiPageLayout(1, 1); }
                pdfDoc.Print();
                dict["status"] = 1;
            }
            pdfDoc.Close();

            return dict;
        }
        private void PrinterStatusTracking()
        {       
            Thread th = new Thread(() =>
            {
                while (!clo)
                {
                    ManagementObjectSearcher PrinterSet = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
                    foreach (ManagementObject Printer in PrinterSet.Get())
                    {
                        if (print_name != null)
                        {
                            if (Printer.Properties["Name"].Value.ToString().ToLower() == print_name.ToLower())
                            {
                                statusPrinter = System.Convert.ToInt32(Printer.Properties["PrinterStatus"].Value);
                                switch (statusPrinter)
                                {
                                    case 3: { Dispatcher.Invoke(() => labelStatusPrinter.Content = "Готов"); break; };
                                    case 4: { Dispatcher.Invoke(() => labelStatusPrinter.Content = "Печать"); break; };
                                    case 5: { Dispatcher.Invoke(() => labelStatusPrinter.Content = "Предупреждение"); break; };
                                    default: { Dispatcher.Invoke(() => labelStatusPrinter.Content = "Ошибка"); break; };
                                }
                            }
                        }
                    }
                    Thread.Sleep(100);
                }
            });
            th.Start();            
        }
        #endregion

        

        
    }

}
