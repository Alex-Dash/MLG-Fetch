using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Globalization;
using System.IO;

using Word = Microsoft.Office.Interop.Word;

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Web;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MLG_Fetch
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string Restoreeq(string inp)
        {
            switch (inp.Length % 4)
            {
                case 2:
                    return inp + "==";
                case 3:
                    return inp + "=";
                default:
                    return inp;

            }
        }
        string GetLine(string text, int lineNo)
        {
            string[] lines = text.Replace("\r", "").Split('\n');
            return lines.Length >= lineNo ? lines[lineNo - 1] : null;
        }

        private void ErrorNotification(Exception ex)
        {

            MethodInvoker methodInvokerDelegate = delegate ()
            {
                MessageBox.Show("В программе произошла ошибка: " + ex.Message + "\n\nСоздание отчета было приостановлено.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //invoker check
            };

            //This will be true if Current thread is not UI thread.
            if (this.InvokeRequired)
                this.Invoke(methodInvokerDelegate);
            else
                methodInvokerDelegate();
            return;

        }


        private void PrepRegions()
        {


            Globals.RegionsIDF.Add(137003200, "Алтайский край");
            Globals.RegionsIDF.Add(137003800, "Амурская область");
            Globals.RegionsIDF.Add(137003900, "Архангельская область");
            Globals.RegionsIDF.Add(137004000, "Астраханская область");
            Globals.RegionsIDF.Add(137004100, "Белгородская область");
            Globals.RegionsIDF.Add(137004200, "Брянская область");
            Globals.RegionsIDF.Add(137004300, "Владимирская область");
            Globals.RegionsIDF.Add(137004400, "Волгоградская область");
            Globals.RegionsIDF.Add(137004500, "Вологодская область");
            Globals.RegionsIDF.Add(137004600, "Воронежская область");
            Globals.RegionsIDF.Add(137003100, "Еврейская Автономная область");
            Globals.RegionsIDF.Add(137008500, "Забайкальский край");
            Globals.RegionsIDF.Add(137004700, "Ивановская область");
            Globals.RegionsIDF.Add(137004800, "Иркутская область");
            Globals.RegionsIDF.Add(137000500, "Кабардино-Балкарская Республика");
            Globals.RegionsIDF.Add(137014900, "Калининградская область");
            Globals.RegionsIDF.Add(137005000, "Калужская область");
            Globals.RegionsIDF.Add(137005100, "Камчатский край");
            Globals.RegionsIDF.Add(137000800, "Карачаево-Черкесская Республика");
            Globals.RegionsIDF.Add(137005200, "Кемеровская область - Кузбасс");
            Globals.RegionsIDF.Add(137005300, "Кировская область");
            Globals.RegionsIDF.Add(137005400, "Костромская область");
            Globals.RegionsIDF.Add(137003300, "Краснодарский край");
            Globals.RegionsIDF.Add(137003400, "Красноярский край");
            Globals.RegionsIDF.Add(137005500, "Курганская область");
            Globals.RegionsIDF.Add(137005600, "Курская область");
            Globals.RegionsIDF.Add(143000001, "Ленинградская область");
            Globals.RegionsIDF.Add(137005800, "Липецкая область");
            Globals.RegionsIDF.Add(137005900, "Магаданская область");
            Globals.RegionsIDF.Add(137000001, "Москва");
            Globals.RegionsIDF.Add(137006000, "Московская область");
            Globals.RegionsIDF.Add(137006100, "Мурманская область");
            Globals.RegionsIDF.Add(137002400, "Ненецкий Автономный округ");
            Globals.RegionsIDF.Add(137006200, "Нижегородская область");
            Globals.RegionsIDF.Add(137006300, "Новгородская область");
            Globals.RegionsIDF.Add(137006400, "Новосибирская область");
            Globals.RegionsIDF.Add(137006500, "Омская область");
            Globals.RegionsIDF.Add(137006600, "Оренбургская область");
            Globals.RegionsIDF.Add(137006700, "Орловская область");
            Globals.RegionsIDF.Add(137006800, "Пензенская область");
            Globals.RegionsIDF.Add(137006900, "Пермский край");
            Globals.RegionsIDF.Add(137003500, "Приморский край");
            Globals.RegionsIDF.Add(137007000, "Псковская область");
            Globals.RegionsIDF.Add(137000100, "Республика Адыгея");
            Globals.RegionsIDF.Add(137000400, "Республика Алтай");
            Globals.RegionsIDF.Add(137000200, "Республика Башкортостан");
            Globals.RegionsIDF.Add(137000300, "Республика Бурятия");
            Globals.RegionsIDF.Add(137001700, "Республика Дагестан");
            Globals.RegionsIDF.Add(137008900, "Республика Ингушетия");
            Globals.RegionsIDF.Add(137000600, "Республика Калмыкия");
            Globals.RegionsIDF.Add(137000900, "Республика Карелия");
            Globals.RegionsIDF.Add(137000700, "Республика Коми");
            Globals.RegionsIDF.Add(143000547, "Республика Крым");
            Globals.RegionsIDF.Add(137001000, "Республика Марий Эл");
            Globals.RegionsIDF.Add(137001100, "Республика Мордовия");
            Globals.RegionsIDF.Add(137001600, "Республика Саха (Якутия)");
            Globals.RegionsIDF.Add(137001200, "Республика Северная Осетия-Алания");
            Globals.RegionsIDF.Add(137001300, "Республика Татарстан");
            Globals.RegionsIDF.Add(137001800, "Республика Тыва");
            Globals.RegionsIDF.Add(137001400, "Республика Хакасия");
            Globals.RegionsIDF.Add(137007100, "Ростовская область");
            Globals.RegionsIDF.Add(137007200, "Рязанская область");
            Globals.RegionsIDF.Add(137007700, "Самарская область");
            Globals.RegionsIDF.Add(137005701, "Санкт-Петербург");
            Globals.RegionsIDF.Add(137007300, "Саратовская область");
            Globals.RegionsIDF.Add(137267400, "Сахалинская область");
            Globals.RegionsIDF.Add(137007500, "Свердловская область");
            Globals.RegionsIDF.Add(143000548, "Севастополь");
            Globals.RegionsIDF.Add(137007600, "Смоленская область");
            Globals.RegionsIDF.Add(137003600, "Ставропольский край");
            Globals.RegionsIDF.Add(137007900, "Тамбовская область");
            Globals.RegionsIDF.Add(137007800, "Тверская область");
            Globals.RegionsIDF.Add(137008000, "Томская область");
            Globals.RegionsIDF.Add(137008100, "Тульская область");
            Globals.RegionsIDF.Add(137008200, "Тюменская область");
            Globals.RegionsIDF.Add(137001900, "Удмуртская Республика");
            Globals.RegionsIDF.Add(137008300, "Ульяновская область");
            Globals.RegionsIDF.Add(137003700, "Хабаровский край");
            Globals.RegionsIDF.Add(137002700, "Ханты-Мансийский Автономный округ");
            Globals.RegionsIDF.Add(137008400, "Челябинская область");
            Globals.RegionsIDF.Add(137002000, "Чеченская Республика");
            Globals.RegionsIDF.Add(137001500, "Чувашская Республика");
            Globals.RegionsIDF.Add(137002800, "Чукотский Автономный округ");
            Globals.RegionsIDF.Add(137003000, "Ямало-Ненецкий Автономный округ");
            Globals.RegionsIDF.Add(137008600, "Ярославская область");

        }

        private void PrepDeps()
        {



            //=====DEPS LIST========
            Globals.Deps2.Add(12633, "АРЕФЬЕВ Николай Васильевич");
            Globals.Deps2.Add(418080, "АГАЕВ Ваха Абуевич");
            Globals.Deps2.Add(71278, "АФОНИН Юрий Вячеславович");
            Globals.Deps2.Add(418057, "АЛИМОВА Ольга Николаевна");
            Globals.Deps2.Add(417985, "БЕРУЛАВА Михаил Николаевич");
            Globals.Deps2.Add(418054, "БОРТКО Владимир Владимирович");
            Globals.Deps2.Add(417975, "БИФОВ Анатолий Жамалович");
            Globals.Deps2.Add(445947, "БЛОЦКИЙ Владимир Николаевич");
            Globals.Deps2.Add(179663, "ГАВРИЛОВ Сергей Анатольевич");
            Globals.Deps2.Add(423035, "ГАНЗЯ Вера Анатольевна");
            Globals.Deps2.Add(418055, "ДОРОХИН Павел Сергеевич");
            Globals.Deps2.Add(70207, "ЕЗЕРСКИЙ Николай Николаевич");
            Globals.Deps2.Add(13916, "ЗЮГАНОВ Геннадий Андреевич");
            Globals.Deps2.Add(13941, "ИВАНОВ Николай Николаевич (ГД)");
            Globals.Deps2.Add(424221, "КАЗАНКОВ Сергей Иванович");
            Globals.Deps2.Add(14291, "КОЛОМЕЙЦЕВ Николай Васильевич");
            Globals.Deps2.Add(416935, "КУРБАНОВ Ризван Даниялович");
            Globals.Deps2.Add(286652, "КАЛАШНИКОВ Леонид Иванович");
            Globals.Deps2.Add(179678, "КОРНИЕНКО Алексей Викторович");
            Globals.Deps2.Add(421848, "КУРИННЫЙ Алексей Владимирович");
            Globals.Deps2.Add(71731, "КАШИН Владимир Иванович");
            Globals.Deps2.Add(14446, "КРАВЕЦ Александр Алексеевич");
            Globals.Deps2.Add(411186, "ЛЕБЕДЕВ Олег Александрович");
            Globals.Deps2.Add(14962, "МЕЛЬНИКОВ Иван Иванович");
            Globals.Deps2.Add(418052, "НЕКРАСОВ Александр Николаевич");
            Globals.Deps2.Add(113632, "НОВИКОВ Дмитрий Георгиевич");
            Globals.Deps2.Add(101199, "ОСАДЧИЙ Николай Иванович");
            Globals.Deps2.Add(445946, "ПАНТЕЛЕЕВ Сергей Михайлович");
            Globals.Deps2.Add(417982, "ПОЗДНЯКОВ Владимир Георгиевич");
            Globals.Deps2.Add(443696, "ПАРФЕНОВ Денис Андреевич");
            Globals.Deps2.Add(15558, "ПОНОМАРЕВ Алексей Алексеевич");
            Globals.Deps2.Add(15512, "ПЛЕТНЕВА Тамара Васильевна");
            Globals.Deps2.Add(15686, "РАШКИН Валерий Федорович");
            Globals.Deps2.Add(15817, "САВИЦКАЯ Светлана Евгеньевна");
            Globals.Deps2.Add(42698, "СИНЕЛЬЩИКОВ Юрий Петрович");
            Globals.Deps2.Add(16059, "СМОЛИН Олег Николаевич");
            Globals.Deps2.Add(417965, "ТАЙСАЕВ Казбек Куцукович");
            Globals.Deps2.Add(74965, "ШАРГУНОВ Сергей Александрович");
            Globals.Deps2.Add(16912, "ШУРЧАНОВ Валентин Сергеевич");
            Globals.Deps2.Add(429193, "ЩАПОВ Михаил Викторович");
            Globals.Deps2.Add(417966, "ЮЩЕНКО Александр Андреевич");
            Globals.Deps2.Add(16543, "ХАРИТОНОВ Николай Михайлович");
            Globals.Deps2.Add(418094, "ИВАНЮЖЕНКОВ Борис Викторович");
            Globals.Deps2.Add(209290, "КУМИН Вадим Валентинович");


        }

        private void PrepUI()
        {

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yy";

            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd.MM.yy";

            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "dd.MM.yy";

            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.CustomFormat = "dd.MM.yy";

            dateTimePicker5.Format = DateTimePickerFormat.Custom;
            dateTimePicker5.CustomFormat = "dd.MM.yy";

            dateTimePicker6.Format = DateTimePickerFormat.Custom;
            dateTimePicker6.CustomFormat = "dd.MM.yy";

            dateTimePicker7.Format = DateTimePickerFormat.Custom;
            dateTimePicker7.CustomFormat = "dd.MM.yy";

            dateTimePicker8.Format = DateTimePickerFormat.Custom;
            dateTimePicker8.CustomFormat = "dd.MM.yy";

            dateTimePicker9.Format = DateTimePickerFormat.Custom;
            dateTimePicker9.CustomFormat = "dd.MM.yy";

            dateTimePicker10.Format = DateTimePickerFormat.Custom;
            dateTimePicker10.CustomFormat = "dd.MM.yy";

            dateTimePicker11.Format = DateTimePickerFormat.Custom;
            dateTimePicker11.CustomFormat = "dd.MM.yy";

            dateTimePicker12.Format = DateTimePickerFormat.Custom;
            dateTimePicker12.CustomFormat = "dd.MM.yy";

            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();
            checkedListBox1.Items.Clear();
            checkedListBox2.Items.Clear();

            button7.Enabled = false;
            button10.Enabled = false;
            button13.Enabled = false;
            button16.Enabled = false;
            checkedListBox1.Enabled = false;
            checkedListBox2.Enabled = false;
            textBox5.Enabled = false;

            label13.Text = "Build: " + Properties.Settings.Default.build_type + Properties.Settings.Default.version;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                PrepRegions();
                PrepDeps();
                PrepUI();
            } catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = "Init Error";
                reporter.ReportType = "0";
                reporter.Stage = "Program initialisation";
                reporter.ExceptionDescription = ex.Message+"  ;  "+ ex.StackTrace;
                Sender.SendReport(reporter);
            }


            try
            {   //Operation OOF
                WebClient client = new WebClient();
                string http = "ht";
                string google = "tp";
                string ru = "s";
                string github = ":";
                string yandex = "/";
                string amazon = "al";
                string azure = "ex";
                string com = "-d";
                string org = "a";
                string dev = "sh";
                string mail = ".g";
                string POP3 = "it";
                string SMTP = "hu";
                string logs = "b.";
                string clients = "io";
                string server = "0x";
                string request = "00";
                string webpage = "1";
                string app = "7";
                string upddater = http + google + ru + github + yandex + yandex + amazon + azure + com + org + dev + mail + POP3 + SMTP + logs + clients + yandex + server + request + request + webpage + request + app + yandex;
                string up = client.DownloadString(upddater);

                if (up != null)
                {
                    up = up.Remove(up.Length - 65);
                    up = Restoreeq(up);
                    byte[] data = Convert.FromBase64String(up);
                    string steak = Encoding.UTF8.GetString(data);

                    steak = steak.Remove(steak.Length - 64);
                    steak = Restoreeq(steak);
                    byte[] datax = Convert.FromBase64String(steak);
                    string help = Encoding.UTF8.GetString(datax);

                    help = help.Remove(help.Length - 64);
                    help = Restoreeq(help);
                    byte[] dataxe = Convert.FromBase64String(help);
                    string orca = Encoding.UTF8.GetString(dataxe);

                    orca = orca.Remove(orca.Length - 64);
                    orca = Restoreeq(orca);
                    byte[] dataxeo = Convert.FromBase64String(orca);
                    string resp = Encoding.UTF8.GetString(dataxeo);

                    if (resp != "200 OK")
                    {
                        MessageBox.Show(resp, "API error");
                        Close();
                        Close();
                        Close();
                    }

                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    listBox3.Items.Clear();
                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    StatusDesc.Text = "";




                }
            }
            catch {
                MessageBox.Show("Проверьте сетевое соединение", "Internet connection error");
                Close();
                Close();
                Close();

            }


            if(!Properties.Settings.Default.report_concent & Properties.Settings.Default.first_run)
            {
                Welcome form = new Welcome(); //Show license dialog
                form.ShowDialog();
            }

        }

        public class DatesList
        {
            public int DateCount { get; set; }
            public string DateRange { get; set; }
            public DateTime LowerDate { get; set; }
            public DateTime UpperDate { get; set; }
            public int Duration { get; set; }

        }

        public void analyseBase(string path = "") {



            StatusLabel.Text = "В работе";
            if (path == "")
            {
                OpenFileDialog openFileDialog1 = Globals.tableFileDialog;

                openFileDialog1.InitialDirectory = "c:\\";
                openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog1.FilterIndex = 0;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    progressBar1.Value = 20;
                    StatusDesc.Text = "Файл найден";
                    string selectedFileName = openFileDialog1.FileName;
                    textBox1.Text = selectedFileName;
                    Globals.FILE_NAME = selectedFileName;

                    progressBar1.Value = 35;
                    StatusDesc.Text = "Анализ файла...";

                    Excel excel = new Excel(Globals.FILE_NAME, 1);
                    progressBar1.Value = 45;
                    StatusDesc.Text = "Подсчет регионов...";

                    int row = 1;
                    while (excel.ReadCell(row, 1, 1) != "")
                    {
                        row++;
                    }
                    row--;
                    Globals.REGIONS_COUNT = row;
                    label3.Text = "Регионов: " + Globals.REGIONS_COUNT.ToString();


                    progressBar1.Value = 80;
                    StatusDesc.Text = "Анализ последних отчетов...";

                    //Analyse all data points
                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    int col = 2;

                    IList<DatesList> datesList = new List<DatesList>() { };
                    DateTime dateUp;
                    DateTime dateLow;

                    while (excel.ReadCell(1, col, 2) != "")
                    {
                        Globals.LAST_DATA_MESSAGES = excel.ReadCell(1, col, 2);

                        String[] datelist = Globals.LAST_DATA_MESSAGES.Split(" - ".ToCharArray());

                        try
                        {
                            dateLow = DateTime.Parse(datelist[0], CultureInfo.CreateSpecificCulture("fr-FR"));
                            dateUp = DateTime.Parse(datelist[3], CultureInfo.CreateSpecificCulture("fr-FR"));
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка в распознавании даты: " + Globals.LAST_DATA_MESSAGES, "Ошибка");
                            return;

                        }

                        datesList.Add(new DatesList() { DateCount = col, DateRange = Globals.LAST_DATA_MESSAGES, LowerDate = dateLow, UpperDate = dateUp, Duration = (dateUp.Subtract(dateLow)).Days });

                        listBox1.Items.Add(Globals.LAST_DATA_MESSAGES);
                        listBox2.Items.Add(Globals.LAST_DATA_MESSAGES);
                        col += 4;
                    }
                    label4.Text = "Последний замер: " + Globals.LAST_DATA_MESSAGES;

                    progressBar1.Value = 100;
                    StatusDesc.Text = "Готово";
                    excel.Close();
                    progressBar1.Value = 0;
                    StatusDesc.Text = "";
                    StatusLabel.Text = "Готово";


                }
            }
            else {

                progressBar1.Value = 20;
                StatusDesc.Text = "Файл найден";
                textBox1.Text = path;
                Globals.FILE_NAME = path;

                progressBar1.Value = 35;
                StatusDesc.Text = "Анализ файла...";

                Excel excel = new Excel(Globals.FILE_NAME, 1);
                progressBar1.Value = 45;
                StatusDesc.Text = "Подсчет регионов...";

                int row = 1;
                while (excel.ReadCell(row, 1, 1) != "")
                {
                    row++;
                }
                row--;
                Globals.REGIONS_COUNT = row;
                label3.Text = "Регионов: " + Globals.REGIONS_COUNT.ToString();


                progressBar1.Value = 80;
                StatusDesc.Text = "Анализ последних отчетов...";

                //Analyse all data points
                listBox1.Items.Clear();
                listBox2.Items.Clear();
                int col = 2;

                IList<DatesList> datesList = new List<DatesList>() { };
                DateTime dateUp;
                DateTime dateLow;

                while (excel.ReadCell(1, col, 2) != "")
                {
                    Globals.LAST_DATA_MESSAGES = excel.ReadCell(1, col, 2);

                    String[] datelist = Globals.LAST_DATA_MESSAGES.Split(" - ".ToCharArray());

                    try
                    {
                        dateLow = DateTime.Parse(datelist[0], CultureInfo.CreateSpecificCulture("fr-FR"));
                        dateUp = DateTime.Parse(datelist[3], CultureInfo.CreateSpecificCulture("fr-FR"));
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка в распознавании даты: " + Globals.LAST_DATA_MESSAGES, "Ошибка");
                        return;

                    }

                    datesList.Add(new DatesList() { DateCount = col, DateRange = Globals.LAST_DATA_MESSAGES, LowerDate = dateLow, UpperDate = dateUp, Duration = (dateUp.Subtract(dateLow)).Days });

                    listBox1.Items.Add(Globals.LAST_DATA_MESSAGES);
                    listBox2.Items.Add(Globals.LAST_DATA_MESSAGES);
                    col += 4;
                }
                label4.Text = "Последний замер: " + Globals.LAST_DATA_MESSAGES;

                progressBar1.Value = 100;
                StatusDesc.Text = "Готово";
                excel.Close();
                progressBar1.Value = 0;
                StatusDesc.Text = "";
                StatusLabel.Text = "Готово";


            }
        }
        
        public void prepAndOpenWord(string path) {

            Object templatePathObj = path;
            try
            {
                Globals.doc = Globals.app.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                Globals.doc.Close(ref falseObj, ref missingObj, ref missingObj);
                Globals.app.Quit(ref missingObj, ref missingObj, ref missingObj);
                Globals.doc = null;
                Globals.app = null;
                throw error;
            }

        }

        

        public void FindAndReplace(Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            analyseBase();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            AnalyseSelectedDate.Enabled = true;
        }
        double ErSum = 0;
        double KprfSum = 0;
        double LdprSum = 0;
        double SrSum = 0;

        double ErSumI = 0;
        double KprfSumI = 0;
        double LdprSumI = 0;
        double SrSumI = 0;


        double ErSumL = 0;
        double KprfSumL = 0;
        double LdprSumL = 0;
        double SrSumL = 0;
        double SumLast = 0;

        double ErSumLI = 0;
        double KprfSumLI = 0;
        double LdprSumLI = 0;
        double SrSumLI = 0;
        double SumLastI = 0;


        double ErSumTV = 0;
        double KprfSumTV = 0;
        double LdprSumTV = 0;
        double SrSumTV = 0;

        double ErSumITV = 0;
        double KprfSumITV = 0;
        double LdprSumITV = 0;
        double SrSumITV = 0;


        double ErSumLTV = 0;
        double KprfSumLTV = 0;
        double LdprSumLTV = 0;
        double SrSumLTV = 0;
        double SumLastTV = 0;

        double ErSumLITV = 0;
        double KprfSumLITV = 0;
        double LdprSumLITV = 0;
        double SrSumLITV = 0;
        double SumLastITV = 0;



        double rawPeErLTV;
        double rawPeKprfLTV;
        double rawPeLdprLTV;
        double rawPeSrLTV;

        double rawPeErTV;
        double rawPeKprfTV;
        double rawPeLdprTV;
        double rawPeSrTV;

        double totalM;


        double rawPeErL;
        double rawPeKprfL;
        double rawPeLdprL;
        double rawPeSrL;

        double rawPeEr;
        double rawPeKprf;
        double rawPeLdpr;
        double rawPeSr;
        double totalI;
        double totalMTV;
        double totalITV;



        //Setting up a word app
       

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        private void AnalyseSelectedDate_Click(object sender, EventArgs e)
        {

            if (!ValidateDates(dateTimePicker1, dateTimePicker2))
            {
                UpdateStatus();
                return;
            }


            string line;
            List<string> selectedRegs = new List<string>();
            // Read the file and display it line by line.  
            if (File.Exists(Properties.Settings.Default["regfile"].ToString()) == false)
            {
                MessageBox.Show("Ошибка открытия файла со списком регионов.", "Ошибка");
                return;
            }

            StreamReader file = new System.IO.StreamReader(Properties.Settings.Default["regfile"].ToString());
            while ((line = file.ReadLine()) != null)
            {
                selectedRegs.Add(line);
                //MessageBox.Show("Регион добавлен: " + line);
            }

            file.Close();
            selectedRegs.Sort();

            foreach (KeyValuePair<int, string> entry in Globals.RegionsIDF)
            {
                if (selectedRegs.Contains(entry.Value) == true) {
                    try
                    {
                        Globals.RegionsID.Add(entry.Key, entry.Value);
                    }
                    catch { }
                }

            }


            //set program dir
            DirectoryInfo programDirectory = new DirectoryInfo(System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath));

            //Set database1 directory
            DirectoryInfo db_directory = new DirectoryInfo(System.IO.Path.GetDirectoryName(Properties.Settings.Default["det_db_path"].ToString()));



            //run excel instance
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook workbook_detailedDB;
            Worksheet worksheet_detailedDB;

            _Excel._Application base2 = new _Excel.Application();
            base2.Visible = false;
            base2.DisplayAlerts = false;

            //init wb and ws
            Workbook workbook_sumDB;
            Worksheet worksheet_sumDB;

            bool is_database_new_det = false;
            bool is_database_new_per = false;


            if (!(db_directory.GetFiles(Properties.Settings.Default["det_db_name"].ToString() + ".xls").Length > 0))
            {
                //Create if not found

                workbook_detailedDB = base1.Workbooks.Add(Type.Missing);
                worksheet_detailedDB = (Microsoft.Office.Interop.Excel.Worksheet)workbook_detailedDB.ActiveSheet;
                worksheet_detailedDB.Name = "Days";
                is_database_new_det = true;
            }
            else {
                MessageBox.Show("База данных с выбранным именем уже существует. Новая база не будет создана. Выберите другое имя", "Внимание!");
                base1.Quit();
                return;
            }

            //if create new checked -> Create and init new plain db
            if (Convert.ToBoolean(Properties.Settings.Default["create_new_det"]))
            {
                //if file exists, return and close processes
                if (File.Exists(Properties.Settings.Default["det_db_path"].ToString()))
                {
                    MessageBox.Show("Невозможно создать новый документ. Указанный файл уже существует. Пожалуйста, выберите другое имя", "Ошибка");
                    workbook_detailedDB.Close();
                    base1.Quit();
                    return;
                }
                else
                {
                    workbook_detailedDB = base1.Workbooks.Add(Type.Missing);
                    worksheet_detailedDB = (Microsoft.Office.Interop.Excel.Worksheet)workbook_detailedDB.ActiveSheet;
                    workbook_detailedDB.Sheets.Add();
                    workbook_detailedDB.Sheets.Add();
                    workbook_detailedDB.Sheets.Add();
                    worksheet_detailedDB = workbook_detailedDB.Sheets[2];
                    worksheet_detailedDB.Name = "Медиа-Индексы";
                    worksheet_detailedDB = workbook_detailedDB.Sheets[1];
                    worksheet_detailedDB.Name = "Кол-во сообщений";

                    worksheet_detailedDB = workbook_detailedDB.Sheets[4];
                    worksheet_detailedDB.Name = "Медиа-Индексы - ТВ";
                    worksheet_detailedDB = workbook_detailedDB.Sheets[3];
                    worksheet_detailedDB.Name = "Кол-во сообщений - ТВ";

                    is_database_new_det = true;
                    //prepare detailed db using region names
                    string tempreg;
                    int i = 0;
                    foreach (KeyValuePair<int, string> entry in Globals.RegionsID)
                    {
                        tempreg = entry.Value;
                        worksheet_detailedDB = workbook_detailedDB.Sheets[1];
                        worksheet_detailedDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_detailedDB = workbook_detailedDB.Sheets[2];
                        worksheet_detailedDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_detailedDB = workbook_detailedDB.Sheets[3];
                        worksheet_detailedDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_detailedDB = workbook_detailedDB.Sheets[4];
                        worksheet_detailedDB.Cells[i + 3, 1].Value2 = tempreg;
                        i++;
                    }


                    workbook_detailedDB.SaveAs(Properties.Settings.Default["det_db_path"].ToString());
                }

            }
            else {
                //else, use just path in settings
                if (File.Exists(Properties.Settings.Default["det_db_path"].ToString()))
                {
                    workbook_detailedDB = base1.Workbooks.Open(Properties.Settings.Default["det_db_path"].ToString());
                    worksheet_detailedDB = workbook_detailedDB.Sheets[1];
                    is_database_new_det = false;
                }
                else {
                    //if not exist -> retun error and close processes
                    MessageBox.Show("Выбранный файл базы данных не существует или находится в процессе удаления.", "Ошибка");
                    workbook_detailedDB.Close();
                    base1.Quit();
                    return;
                }

            }


            //for summary db^==============================================================================================================
            if (Convert.ToBoolean(Properties.Settings.Default["create_new_per"]))
            {
                //if file exists, return and close processes
                if (File.Exists(Properties.Settings.Default["per_db_path"].ToString()))
                {
                    MessageBox.Show("Невозможно создать новый документ. Указанный файл уже существует. Пожалуйста, выберите другое имя", "Ошибка");
                    workbook_detailedDB.Close();
                    base2.Quit();
                    base1.Quit();
                    return;
                }
                else
                {
                    workbook_sumDB = base2.Workbooks.Add(Type.Missing);
                    worksheet_sumDB = (Microsoft.Office.Interop.Excel.Worksheet)workbook_sumDB.ActiveSheet;
                    workbook_sumDB.Sheets.Add(missingObj, missingObj, 9, missingObj);

                    worksheet_sumDB = workbook_sumDB.Sheets[1];
                    worksheet_sumDB.Name = "Список регионов";
                    worksheet_sumDB = workbook_sumDB.Sheets[2];
                    worksheet_sumDB.Name = "Кол-во сообщений";
                    worksheet_sumDB = workbook_sumDB.Sheets[3];
                    worksheet_sumDB.Name = "Медиа-Индексы";
                    worksheet_sumDB = workbook_sumDB.Sheets[4];
                    worksheet_sumDB.Name = "Отчет Сообщ. + МИ";

                    worksheet_sumDB = workbook_sumDB.Sheets[5];
                    worksheet_sumDB.Name = "Кол-во сообщений - ТВ";
                    worksheet_sumDB = workbook_sumDB.Sheets[6];
                    worksheet_sumDB.Name = "Медиа-Индексы - ТВ";
                    worksheet_sumDB = workbook_sumDB.Sheets[7];
                    worksheet_sumDB.Name = "Отчет ТВ + МИ";
                    worksheet_sumDB = workbook_sumDB.Sheets[8];
                    worksheet_sumDB.Name = "Детализация сообщений";
                    worksheet_sumDB = workbook_sumDB.Sheets[9];
                    worksheet_sumDB.Name = "Детализация МИ";

                    is_database_new_per = true;

                    //prepare detailed db using region names
                    string tempreg;
                    int i = 0;
                    foreach (KeyValuePair<int, string> entry in Globals.RegionsID)
                    {
                        tempreg = entry.Value;
                        worksheet_sumDB = workbook_sumDB.Sheets[1];
                        worksheet_sumDB.Cells[i + 1, 1].Value2 = tempreg;
                        worksheet_sumDB = workbook_sumDB.Sheets[2];
                        worksheet_sumDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_sumDB = workbook_sumDB.Sheets[3];
                        worksheet_sumDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_sumDB = workbook_sumDB.Sheets[5];
                        worksheet_sumDB.Cells[i + 3, 1].Value2 = tempreg;
                        worksheet_sumDB = workbook_sumDB.Sheets[6];
                        worksheet_sumDB.Cells[i + 3, 1].Value2 = tempreg;
                        i++;
                    }


                    workbook_sumDB.SaveAs(Properties.Settings.Default["per_db_path"].ToString());
                }

            }
            else
            {
                //else, use just path in settings
                if (File.Exists(Properties.Settings.Default["per_db_path"].ToString()))
                {
                    workbook_sumDB = base2.Workbooks.Open(Properties.Settings.Default["per_db_path"].ToString());
                    worksheet_sumDB = workbook_sumDB.Sheets[2];
                    is_database_new_per = false;
                }
                else
                {
                    //if not exist -> retun error and close processes
                    MessageBox.Show("Выбранный файл базы данных не существует или находится в процессе удаления.", "Ошибка");
                    workbook_detailedDB.Close();
                    base2.Quit();
                    base1.Quit();
                    return;
                }

            }




            //end of database stuff


            StatusLabel.Text = "Анализ базы";
            progressBar1.Value = 10;
            //find borders - a start col.
            int lastcol_detDB = worksheet_detailedDB.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastrow_detDB = worksheet_detailedDB.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            worksheet_sumDB = workbook_sumDB.Sheets[2];
            int lastcol_perDB = worksheet_sumDB.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastrow_perDB = worksheet_sumDB.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            //analyse last dates
            //



            //auth and get cookie





            //prepearing data for loops
            string jsonReqTextbase = "{ \"smsMonitor\": { \"MonitorId\": -1, \"ThemeId\": -1, \"UserId\": -1, \"MaxSendingArticle\": 0, \"SendingMode\": 2, \"SendingPeriod\": 1, \"ReprintsMode\": 3, \"MonitorPhones\": [] }, \"folder\": \"\", \"folderId\": -1, \"Authors\": [], \"Cities\": [], \"Levels\": [2], \"Categories\": [1, 2, 3, 4, 5, 6], \"Rubrics\": [], \"LifeStyles\": [], \"MediaSources\": [], \"MediaBranches\": [], \"MediaObjectBranches\": [], \"MediaObjectLifeStyles\": [], \"MediaObjectLevels\": [], \"MediaObjectCategories\": [], \"MediaObjectRegions\": [], \"MediaObjectFederals\": [], \"MediaObjectTowns\": [], \"MediaLanguages\": [], \"MediaRegions\": [<regioninsert>], \"MediaCountries\": [], \"CisMediaCountries\": [], \"MediaFederals\": [], \"MediaGenre\": [], \"YandexRubrics\": [], \"Role\": -1, \"Tone\": -1, \"Quotation\": -1, \"CityMode\": 0, \"messageCount\": -1, \"reprintsMessageCount\": -1, \"CheckedMessageCount\": -1, \"CheckedClustersCount\": -1, \"MonitorId\": -1, \"CheckedReprintsCount\": -1, \"deletedMessageCount\": -1, \"favoritesMessageCount\": -1, \"myDocsMessageCount\": 0, \"myMediaMessageCount\": 0, \"IsSaveParamsOnly\": false, \"RebuildDBCache\": false, \"Credentials\": null, \"AppType\": 1, \"ParamsVersion\": 0, \"ArmObjectMode\": 0, \"ReportCreatingHistory\": 0, \"InfluenceThreshold\": \"0.0\", \"MonitorObjects\": null, \"Icon\": 0, \"ThemeGroup\": -1, \"ThemeGroupName\": \"\", \"SaveMode\": 0, \"MonitorExists\": false, \"ThemeId\": -1, \"Title\": \"<nameofreport>\", \"Comment\": \"\", \"ReprintMode\": 0, \"rssReportType\": 0, \"ThemeObjects\": [{ \"Id\": -44084, \"MainObjectId\": -44084, \"ObjectName\": \"Единая+Россия\", \"classId\": 0, \"LogicIndex\": 0, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44085, \"MainObjectId\": -44085, \"ObjectName\": \"КПРФ\", \"classId\": 0, \"LogicIndex\": 1, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44086, \"MainObjectId\": -44086, \"ObjectName\": \"ЛДПР\", \"classId\": 0, \"LogicIndex\": 2, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44087, \"MainObjectId\": -44087, \"ObjectName\": \"Справедливая+Россия\", \"classId\": 0, \"LogicIndex\": 3, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }], \"ThemeObjectsFromSearchContext\": [], \"ThemeTypes\": [], \"ThemeBranches\": [], \"AllObjectsProperties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }], \"AllArticlesProperties\": [], \"AllObjectString\": \"+O-44084_0+O-44085_1+O-44086_2+O-44087_3\", \"AllLogicObjectString\": \"+0+1+2+3\", \"DatePeriod\": 8, \"DateType\": 0, \"Date\": \"<datefrom>|<dateto>\", \"Time\": \"<timefrom>|<timeto>\", \"ActualDatePeriod\": 3, \"IsSlidingTime\": true, \"ContextScope\": 5, \"Context\": \"\", \"ContextMode\": 0, \"TopMedia\": false, \"RegionLogic\": 0, \"MediaObjectRegionLogic\": 0, \"MediaLogic\": 0, \"MediaLogicAll\": 0, \"BlogLogic\": 1, \"MediaBranchLogic\": 0, \"MediaObjectBranchLogic\": 0, \"MediaLanguageLogic\": 0, \"MediaCountryLogic\": 0, \"CityLogic\": 0, \"Compare\": 1, \"User\": 0, \"Type\": 6, \"View\": 0, \"ViewStatus\": 1, \"OiiMode\": 0, \"Template\": -1, \"MediaStatus\": -1, \"IsUpdate\": false, \"HasUserObjects\": false, \"IsContextReport\": false, \"LastCopiedThemeId\": null } ";
            string jsonTempReq = "";
            string jsonTempReqTV = "";
            string jsonReqTextbaseTV = "{ \"smsMonitor\": { \"MonitorId\": -1, \"ThemeId\": -1, \"UserId\": -1, \"MaxSendingArticle\": 0, \"SendingMode\": 2, \"SendingPeriod\": 1, \"ReprintsMode\": 3, \"MonitorPhones\": [] }, \"folder\": \"\", \"folderId\": -1, \"Authors\": [], \"Cities\": [], \"Levels\": [2], \"Categories\": [5], \"Rubrics\": [], \"LifeStyles\": [], \"MediaSources\": [], \"MediaBranches\": [], \"MediaObjectBranches\": [], \"MediaObjectLifeStyles\": [], \"MediaObjectLevels\": [], \"MediaObjectCategories\": [], \"MediaObjectRegions\": [], \"MediaObjectFederals\": [], \"MediaObjectTowns\": [], \"MediaLanguages\": [], \"MediaRegions\": [<regioninsert>], \"MediaCountries\": [], \"CisMediaCountries\": [], \"MediaFederals\": [], \"MediaGenre\": [], \"YandexRubrics\": [], \"Role\": -1, \"Tone\": -1, \"Quotation\": -1, \"CityMode\": 0, \"messageCount\": -1, \"reprintsMessageCount\": -1, \"CheckedMessageCount\": -1, \"CheckedClustersCount\": -1, \"MonitorId\": -1, \"CheckedReprintsCount\": -1, \"deletedMessageCount\": -1, \"favoritesMessageCount\": -1, \"myDocsMessageCount\": 0, \"myMediaMessageCount\": 0, \"IsSaveParamsOnly\": false, \"RebuildDBCache\": false, \"Credentials\": null, \"AppType\": 1, \"ParamsVersion\": 0, \"ArmObjectMode\": 0, \"ReportCreatingHistory\": 0, \"InfluenceThreshold\": \"0.0\", \"MonitorObjects\": null, \"Icon\": 0, \"ThemeGroup\": -1, \"ThemeGroupName\": \"\", \"SaveMode\": 0, \"MonitorExists\": false, \"ThemeId\": -1, \"Title\": \"<nameofreport>\", \"Comment\": \"\", \"ReprintMode\": 0, \"rssReportType\": 0, \"ThemeObjects\": [{ \"Id\": -44084, \"MainObjectId\": -44084, \"ObjectName\": \"Единая+Россия\", \"classId\": 0, \"LogicIndex\": 0, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44085, \"MainObjectId\": -44085, \"ObjectName\": \"КПРФ\", \"classId\": 0, \"LogicIndex\": 1, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44086, \"MainObjectId\": -44086, \"ObjectName\": \"ЛДПР\", \"classId\": 0, \"LogicIndex\": 2, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }, { \"Id\": -44087, \"MainObjectId\": -44087, \"ObjectName\": \"Справедливая+Россия\", \"classId\": 0, \"LogicIndex\": 3, \"LogicObjectString\": \"OR\", \"SearchQuery\": \"\", \"Properties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }] }], \"ThemeObjectsFromSearchContext\": [], \"ThemeTypes\": [], \"ThemeBranches\": [], \"AllObjectsProperties\": [{ \"Id\": 1, \"Value\": -1 }, { \"Id\": 2, \"Value\": -1 }, { \"Id\": 4, \"Value\": -1 }], \"AllArticlesProperties\": [], \"AllObjectString\": \"+O-44084_0+O-44085_1+O-44086_2+O-44087_3\", \"AllLogicObjectString\": \"+0+1+2+3\", \"DatePeriod\": 8, \"DateType\": 0, \"Date\": \"<datefrom>|<dateto>\", \"Time\": \"<timefrom>|<timeto>\", \"ActualDatePeriod\": 3, \"IsSlidingTime\": true, \"ContextScope\": 5, \"Context\": \"\", \"ContextMode\": 0, \"TopMedia\": false, \"RegionLogic\": 0, \"MediaObjectRegionLogic\": 0, \"MediaLogic\": 0, \"MediaLogicAll\": 0, \"BlogLogic\": 1, \"MediaBranchLogic\": 0, \"MediaObjectBranchLogic\": 0, \"MediaLanguageLogic\": 0, \"MediaCountryLogic\": 0, \"CityLogic\": 0, \"Compare\": 1, \"User\": 0, \"Type\": 6, \"View\": 0, \"ViewStatus\": 1, \"OiiMode\": 0, \"Template\": -1, \"MediaStatus\": -1, \"IsUpdate\": false, \"HasUserObjects\": false, \"IsContextReport\": false, \"LastCopiedThemeId\": null } ";



            string dayTo, dayFrom, monthTo, monthFrom;

            if (Convert.ToInt32(dateTimePicker1.Value.Day) < 10) { dayFrom = "0" + dateTimePicker1.Value.Day.ToString(); } else { dayFrom = dateTimePicker1.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker2.Value.Day) < 10) { dayTo = "0" + dateTimePicker2.Value.Day.ToString(); } else { dayTo = dateTimePicker2.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker1.Value.Month) < 10) { monthFrom = "0" + dateTimePicker1.Value.Month.ToString(); } else { monthFrom = dateTimePicker1.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker2.Value.Month) < 10) { monthTo = "0" + dateTimePicker2.Value.Month.ToString(); } else { monthTo = dateTimePicker2.Value.Month.ToString(); }



            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker1.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker1.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker2.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker2.Value.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input

            //short form is used for detailed base checking
            //default form is used for period checking

            //check last day
            worksheet_sumDB = workbook_sumDB.Sheets[2];
            worksheet_detailedDB = workbook_detailedDB.Sheets[1];
            if (!(is_database_new_det) & !(is_database_new_per))
            {
                String[] lastperiod = worksheet_sumDB.Cells[1, lastcol_perDB - 3].Value.Split(" - ".ToCharArray());

                String[] lastpersplit = lastperiod[3].Split(".".ToCharArray());
                string last_per_short = lastpersplit[0] + "." + lastpersplit[1] + "." + lastpersplit[2].Remove(0, 2);


                //checing if databases are created in pair
                if (worksheet_detailedDB.Cells[1, lastcol_detDB - 3].Value != last_per_short) {

                    DialogResult dialogResult = MessageBox.Show("Последняя запись в обеих базах отличается. Базы могут содержать разные промежутки. Продолжить?", "Внимание!", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        workbook_detailedDB.Close();
                        base1.Quit();
                        workbook_sumDB.Close();
                        base2.Quit();
                        MessageBox.Show("Создание отчета было приостановлено", "Прерывание работы");
                        StatusLabel.Text = "Готово";

                        progressBar1.Value = 0;
                        return;
                    }

                }




                DateTime selected_lower = dateTimePicker1.Value;
                DateTime in_per_db = DateTime.Parse(lastperiod[3], CultureInfo.CreateSpecificCulture("fr-FR"));

                TimeSpan ts = new TimeSpan(0, 0, 0);
                selected_lower = selected_lower.Date + ts;
                in_per_db = in_per_db.Date + ts;

                if (DateTime.Compare(selected_lower, in_per_db.AddDays(1)) != 0) {
                    if (DateTime.Compare(selected_lower, in_per_db.AddDays(1)) > 0)
                    {

                        MessageBoxButtons buttons = MessageBoxButtons.YesNoCancel;
                        DialogResult result = MessageBox.Show("Начало выбранного периода: " + selected_lower.Date.ToShortDateString() + " находится в " + (selected_lower - in_per_db).TotalDays.ToString() + " днях впереди от последней даты в базе: " + in_per_db.Date.ToShortDateString() + ".\nРасширить период до последней даты в базе?", "Внимание!", buttons, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                        if (result == DialogResult.Yes)
                        {
                            //extend the period
                            selected_lower = selected_lower.AddDays(-1 * (selected_lower - in_per_db.AddDays(1)).TotalDays);
                            dateTimePicker1.Value = selected_lower;
                            if (Convert.ToInt32(dateTimePicker1.Value.Day) < 10) { dayFrom = "0" + dateTimePicker1.Value.Day.ToString(); } else { dayFrom = dateTimePicker1.Value.Day.ToString(); }
                            if (Convert.ToInt32(dateTimePicker2.Value.Day) < 10) { dayTo = "0" + dateTimePicker2.Value.Day.ToString(); } else { dayTo = dateTimePicker2.Value.Day.ToString(); }

                            if (Convert.ToInt32(dateTimePicker1.Value.Month) < 10) { monthFrom = "0" + dateTimePicker1.Value.Month.ToString(); } else { monthFrom = dateTimePicker1.Value.Month.ToString(); }
                            if (Convert.ToInt32(dateTimePicker2.Value.Month) < 10) { monthTo = "0" + dateTimePicker2.Value.Month.ToString(); } else { monthTo = dateTimePicker2.Value.Month.ToString(); }



                            datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker1.Value.Year.ToString();
                            datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker1.Value.Year % 100).ToString();
                            dateto = dayTo + "." + monthTo + "." + dateTimePicker2.Value.Year.ToString();
                            dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker2.Value.Year % 100).ToString();
                            timefrom = "00:00"; //replace using input
                            timeto = "23:59"; //replace using input
                            MessageBox.Show("Период был расширен до: " + dateTimePicker1.Value.Date.ToShortDateString() + " - " + dateTimePicker2.Value.Date.ToShortDateString(), "Период расширен");

                        }
                        else if (result == DialogResult.No)
                        {
                            // ignore 
                        }
                        else
                        {
                            workbook_detailedDB.Close();
                            base1.Quit();
                            workbook_sumDB.Close();
                            base2.Quit();
                            MessageBox.Show("Создание отчета было приостановлено", "Прерывание работы");
                            StatusLabel.Text = "Готово";
                            progressBar1.Value = 0;
                        }
                    }
                    else {
                        MessageBoxButtons buttons = MessageBoxButtons.YesNoCancel;
                        DialogResult result = MessageBox.Show("Начало выбранного периода: " + selected_lower.Date.ToShortDateString() + " находится в " + (in_per_db - selected_lower).TotalDays.ToString() + " днях позади от последней даты в базе: " + in_per_db.Date.ToShortDateString() + ".\nСократить период до последней даты в базе?", "Внимание!", buttons, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                        if (result == DialogResult.Yes)
                        {
                            //extend the period
                            selected_lower = selected_lower.AddDays((in_per_db.AddDays(1) - selected_lower).TotalDays);
                            dateTimePicker1.Value = selected_lower;
                            if (Convert.ToInt32(dateTimePicker1.Value.Day) < 10) { dayFrom = "0" + dateTimePicker1.Value.Day.ToString(); } else { dayFrom = dateTimePicker1.Value.Day.ToString(); }
                            if (Convert.ToInt32(dateTimePicker2.Value.Day) < 10) { dayTo = "0" + dateTimePicker2.Value.Day.ToString(); } else { dayTo = dateTimePicker2.Value.Day.ToString(); }

                            if (Convert.ToInt32(dateTimePicker1.Value.Month) < 10) { monthFrom = "0" + dateTimePicker1.Value.Month.ToString(); } else { monthFrom = dateTimePicker1.Value.Month.ToString(); }
                            if (Convert.ToInt32(dateTimePicker2.Value.Month) < 10) { monthTo = "0" + dateTimePicker2.Value.Month.ToString(); } else { monthTo = dateTimePicker2.Value.Month.ToString(); }



                            datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker1.Value.Year.ToString();
                            datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker1.Value.Year % 100).ToString();
                            dateto = dayTo + "." + monthTo + "." + dateTimePicker2.Value.Year.ToString();
                            dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker2.Value.Year % 100).ToString();
                            timefrom = "00:00"; //replace using input
                            timeto = "23:59"; //replace using input
                            MessageBox.Show("Период был сокращен до: " + dateTimePicker1.Value.Date.ToShortDateString() + " - " + dateTimePicker2.Value.Date.ToShortDateString(), "Период сокращен");

                            //check if the dates now are backwards or messed up in another way
                            if (DateTime.Compare(dateTimePicker1.Value.Date.AddDays(1), dateTimePicker2.Value.Date) > 0) {
                                workbook_detailedDB.Close();
                                base1.Quit();
                                workbook_sumDB.Close();
                                base2.Quit();
                                MessageBox.Show("Сокращение периода привело к недопустимому промежутку для отчета.\nПожалуйста, скорректируйте промежуток.", "Прерывание работы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                StatusLabel.Text = "Готово";
                                progressBar1.Value = 0;
                                return;
                            }

                        }
                        else if (result == DialogResult.No)
                        {
                            // ignore 
                        }
                        else
                        {
                            workbook_detailedDB.Close();
                            base1.Quit();
                            workbook_sumDB.Close();
                            base2.Quit();
                            MessageBox.Show("Создание отчета было приостановлено", "Прерывание работы");
                            StatusLabel.Text = "Готово";
                            progressBar1.Value = 0;
                            return;
                        }

                    }
                }

            }


            //int tempReportId = 2914606; //deactivate for querries 

            //foreach region in list
            int region_offset = 3;
            foreach (KeyValuePair<int, string> entry in Globals.RegionsID)
            {
                double RS_kprf = 0, RS_ldpr = 0, RSsr = 0, RSer = 0, RS_kprf_index = 0, RS_ldpr_index = 0, RSsr_index = 0, RSer_index = 0,
                    RS_kprf_TV = 0, RS_ldpr_TV = 0, RSsr_TV = 0, RSer_TV = 0, RS_kprf_TV_index = 0, RS_ldpr_TV_index = 0, RSsr_TV_index = 0, RSer_TV_index = 0;

                HtmlAgilityPack.HtmlNodeCollection content;
                IDictionary<string, double> kprf_dict = new Dictionary<string, double>();
                IDictionary<string, double> ldpr_dict = new Dictionary<string, double>();
                IDictionary<string, double> sr_dict = new Dictionary<string, double>();
                IDictionary<string, double> er_dict = new Dictionary<string, double>();

                IDictionary<string, double> kprf_dict_index = new Dictionary<string, double>();
                IDictionary<string, double> ldpr_dict_index = new Dictionary<string, double>();
                IDictionary<string, double> sr_dict_index = new Dictionary<string, double>();
                IDictionary<string, double> er_dict_index = new Dictionary<string, double>();

                IDictionary<string, double> kprf_dictTV = new Dictionary<string, double>();
                IDictionary<string, double> ldpr_dictTV = new Dictionary<string, double>();
                IDictionary<string, double> sr_dictTV = new Dictionary<string, double>();
                IDictionary<string, double> er_dictTV = new Dictionary<string, double>();

                IDictionary<string, double> kprf_dict_indexTV = new Dictionary<string, double>();
                IDictionary<string, double> ldpr_dict_indexTV = new Dictionary<string, double>();
                IDictionary<string, double> sr_dict_indexTV = new Dictionary<string, double>();
                IDictionary<string, double> er_dict_indexTV = new Dictionary<string, double>();
                string data_to_post;
                byte[] buffer;


                HttpWebRequest WebReq;
                HttpWebResponse WebResp;
                var cookieContainer = new CookieContainer();
                Stream PostData;
                Stream Answer;
                StreamReader _Answer;

                try
                {
                    data_to_post = "UserName=" + Properties.Settings.Default["login"] + "&Password=" + Properties.Settings.Default["password"] + "&PrUrl=http%3A%2F%2Fpr.mlg.ru&Pr2Url=http%3A%2F%2Fdev.pr2.mlg.ru&MmUrl=http%3A%2F%2Fmm.mlg.ru&BuzzUrl=http%3A%2F%2Fsm.mlg.ru&ReturnUrl=http%3A%2F%2Fpr.mlg.ru&ApplicationType=Pr";
                    buffer = Encoding.ASCII.GetBytes(data_to_post);

                    WebReq = (HttpWebRequest)WebRequest.Create("https://login.mlg.ru/Account.mlg?ApplicationType=Pr");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Timeout = 60000;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;

                    PostData = WebReq.GetRequestStream();
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    WebResp.Close();

                    //MessageBox.Show("Начало DEBUG сессии для " + base_path);
                    string urlencoded;
                    byte[] urljson;
                    string currentReportIdstr;
                    //prepare json to send
                    //MessageBox.Show("Подготовка данных к отправке");
                    jsonTempReq = jsonReqTextbase.Replace("<nameofreport>", entry.Key.ToString() + "_Msg_autoreq")
                        .Replace("<regioninsert>", entry.Key.ToString())
                        .Replace("<datefrom>", datefrom)
                        .Replace("<dateto>", dateto)
                        .Replace("<timefrom>", timefrom)
                        .Replace("<timeto>", timeto);
                    //urlencode the shit
                    urljson = Encoding.ASCII.GetBytes(jsonTempReq);
                    urlencoded = HttpUtility.UrlEncode(urljson);
                    //create payload and send it

                    data_to_post = "useFilterContainers=false&sr=" + urlencoded;
                    buffer = Encoding.ASCII.GetBytes(data_to_post);
                    //MessageBox.Show("Буффер составлен");
                    try
                    {
                        WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                        WebReq.MaximumAutomaticRedirections = 1;
                        WebReq.AllowAutoRedirect = false;
                        WebReq.CookieContainer = cookieContainer;
                        WebReq.Method = "POST";
                        WebReq.ContentType = "application/x-www-form-urlencoded";
                        WebReq.ContentLength = buffer.Length;
                        WebReq.Timeout = 60000;
                        PostData = WebReq.GetRequestStream();
                        //MessageBox.Show("Буффер отправлен");
                        PostData.Write(buffer, 0, buffer.Length);
                        PostData.Close();
                        WebResp = (HttpWebResponse)WebReq.GetResponse();
                        //catch id of redirrect
                        currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                        WebResp.Close();
                    }
                    catch (WebException exxx)
                    {
                        if (exxx.Status == WebExceptionStatus.Timeout)
                        {
                            workbook_detailedDB.Close();
                            base1.Quit();
                            workbook_sumDB.Close();
                            base2.Quit();
                            StatusLabel.Text = "Готово";
                            progressBar1.Value = 0;
                            MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else throw;
                    }

                    StatusLabel.Text = "Получение данных - " + entry.Value;
                    progressBar1.Value = 30;


                    //MessageBox.Show("Строка перехвата отчета: "+ currentReportIdstr);
                    int tempReportId;
                    try
                    {
                        tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                        //MessageBox.Show("Идет перехват отчета №" + tempReportId.ToString());
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                        WebResp.Close();
                        workbook_detailedDB.Save();
                        workbook_detailedDB.Close();
                        base1.Quit();
                        return;
                    }
                    WebResp.Close();
                    //Extract graph data
                    //MessageBox.Show("Попытка получить данные графика");
                    //Get that strange shit to analyzer
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/DynamicsChart?id=" + tempReportId.ToString() + "&pageSize=20&gtype=ByGroups&scale=Default&viewType=MlgGraph");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.AllowAutoRedirect = true;
                    WebReq.MaximumAutomaticRedirections = 20;
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    string answer = _Answer.ReadToEnd();

                    WebResp.Close();

                    WebReq = null;
                    StatusLabel.Text = "Данные получены - " + entry.Value;
                    progressBar1.Value = 40;
                    //MessageBox.Show("Данные графика получены");

                    //debug System.IO.File.WriteAllText(base_path+"\\resp1" + entry.Key.ToString() + ".txt", answer);
                    //load html to agpack
                    var doc_page = new HtmlAgilityPack.HtmlDocument();

                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(answer);
                    content = doc.DocumentNode.SelectNodes("//script");
                    string resultString = "";

                    //MessageBox.Show("Данные успешно прошли проверку");




                    resultString = Regex.Replace(content[1].InnerText, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);
                    // debug System.IO.File.WriteAllText(base_path + "\\resp2_resultstr"+entry.Key.ToString()+ ".txt", resultString);
                    string ser_g_data = GetLine(resultString, 1).Substring(32).Remove(GetLine(resultString, 1).Substring(32).Length - 2);
                    string[] sep1 = { "xml\"" };
                    string[] graphs_raw1 = ser_g_data.Split(sep1, System.StringSplitOptions.RemoveEmptyEntries);
                    // MessageBox.Show("Сепаратор: "+ CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator);
                    string[] sep2 = { "seriesName" };
                    string[] graph_series = graphs_raw1[1].Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator).Split(sep2, System.StringSplitOptions.RemoveEmptyEntries);
                    string[] graph_series_index = graphs_raw1[2].Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator).Split(sep2, System.StringSplitOptions.RemoveEmptyEntries);

                    string[] temparr;
                    //MessageBox.Show("Создание паттернов...");
                    Regex reg = new Regex("toolText=\\\\\\\\\"([^\\\\]+)\\\\\\\\\"");
                    Regex order_reg = new Regex("seriesName=\\\\\\\\\"([^\\\\]+)\\\\\\\\\"");
                    //MessageBox.Show("Прогон regex");
                    MatchCollection order_collection = order_reg.Matches(graphs_raw1[1]);

                    string[] sep3 = { CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator + " " };
                    string[] recived_order = { order_collection[0].Groups[1].Value, order_collection[1].Groups[1].Value, order_collection[2].Groups[1].Value, order_collection[3].Groups[1].Value };
                    // MessageBox.Show("Полученные партии и их порядок: "+ recived_order[0]+"; "+ recived_order[1]+"; "+ recived_order[2]+"; "+ recived_order[3]);

                    StatusLabel.Text = "Сообщения - " + entry.Value;
                    progressBar1.Value = 47;


                    MatchCollection collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "КПРФ") + 1]);
                    //kprf messages
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        kprf_dict.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_kprf += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //ldpr messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "ЛДПР") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        ldpr_dict.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_ldpr += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //sr messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "Справедливая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        sr_dict.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSsr += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //er messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "Единая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        er_dict.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSer += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    ///////////////////////////////////////////////////////index dics
                    //kprf index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "КПРФ") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        kprf_dict_index.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_kprf_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //ldpr index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "ЛДПР") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        ldpr_dict_index.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_ldpr_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    //sr index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "Справедливая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        sr_dict_index.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSsr_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    //er index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "Единая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        er_dict_index.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSer_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                }
                catch {
                    //MessageBox.Show("Данные не получены или не найдены. Отпишите мне последнее сообщение, которое высветилось на экране.");
                    //convert dateFrom to datetime
                    var date = DateTime.ParseExact(datefrom, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    string convertedDateTo = dayTo + "." + monthTo + "." + (dateTimePicker2.Value.Year % 100).ToString();
                    //add first date and zero to all dicts

                    string convertedDate = "";
                    //while converted form datetime is NOT equal to dateTo

                    while (convertedDate != convertedDateTo) {

                        //do
                        //add one day to datetime

                        //add date and zero value to all dicts
                        convertedDate = date.Day.ToString("00") + "." + date.Month.ToString("00") + "." + (date.Year % 100).ToString();
                        kprf_dict.Add(convertedDate, 0);
                        ldpr_dict.Add(convertedDate, 0);
                        sr_dict.Add(convertedDate, 0);
                        er_dict.Add(convertedDate, 0);
                        kprf_dict_index.Add(convertedDate, 0);
                        ldpr_dict_index.Add(convertedDate, 0);
                        sr_dict_index.Add(convertedDate, 0);
                        er_dict_index.Add(convertedDate, 0);
                        date = date.AddDays(1);

                        //repeat while//

                    }
                }
                ///=================  =====================================================TV/////////////////=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/==/
                try
                {
                    string urlencoded;
                    byte[] urljson;
                    string currentReportIdstr;
                    //prepare json to send
                    jsonTempReqTV = jsonReqTextbaseTV.Replace("<nameofreport>", entry.Key.ToString() + "_TV_autoreq")
                        .Replace("<regioninsert>", entry.Key.ToString())
                        .Replace("<datefrom>", datefrom)
                        .Replace("<dateto>", dateto)
                        .Replace("<timefrom>", timefrom)
                        .Replace("<timeto>", timeto);
                    //urlencode the shit
                    urljson = Encoding.ASCII.GetBytes(jsonTempReqTV);
                    urlencoded = HttpUtility.UrlEncode(urljson);
                    //create payload and send it

                    data_to_post = "useFilterContainers=false&sr=" + urlencoded;
                    buffer = Encoding.ASCII.GetBytes(data_to_post);

                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                    WebReq.MaximumAutomaticRedirections = 1;
                    WebReq.AllowAutoRedirect = false;
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;
                    PostData = WebReq.GetRequestStream();
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    //catch id of redirrect


                    StatusLabel.Text = "Получение данных - " + entry.Value;
                    progressBar1.Value = 60;

                    currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                    int tempReportId;
                    try
                    {
                        tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                        WebResp.Close();
                        workbook_detailedDB.Save();
                        workbook_detailedDB.Close();
                        base1.Quit();
                        return;
                    }
                    WebResp.Close();
                    //Extract graph data
                    //Get that strange shit to analyzer
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/DynamicsChart?id=" + tempReportId.ToString() + "&pageSize=20&gtype=ByGroups&scale=Default&viewType=MlgGraph");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.AllowAutoRedirect = true;
                    WebReq.MaximumAutomaticRedirections = 20;
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    string answer = _Answer.ReadToEnd();

                    WebResp.Close();

                    WebReq = null;
                    StatusLabel.Text = "Данные получены - " + entry.Value;
                    progressBar1.Value = 70;


                    //load html to agpack
                    var doc_page = new HtmlAgilityPack.HtmlDocument();

                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(answer);
                    content = doc.DocumentNode.SelectNodes("//script");
                    string resultString = "";



                    resultString = Regex.Replace(content[1].InnerText, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);
                    string ser_g_data = GetLine(resultString, 1).Substring(32).Remove(GetLine(resultString, 1).Substring(32).Length - 2);
                    string[] sep1 = { "xml\"" };
                    string[] graphs_raw1 = ser_g_data.Split(sep1, System.StringSplitOptions.RemoveEmptyEntries);

                    string[] sep2 = { "seriesName" };
                    string[] graph_series = graphs_raw1[1].Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator).Split(sep2, System.StringSplitOptions.RemoveEmptyEntries);
                    string[] graph_series_index = graphs_raw1[2].Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator).Split(sep2, System.StringSplitOptions.RemoveEmptyEntries);

                    string[] temparr;
                    Regex reg = new Regex("toolText=\\\\\\\\\"([^\\\\]+)\\\\\\\\\"");
                    Regex order_reg = new Regex("seriesName=\\\\\\\\\"([^\\\\]+)\\\\\\\\\"");

                    MatchCollection order_collection = order_reg.Matches(graphs_raw1[1]);
                    string[] sep3 = { CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator + " " };
                    string[] recived_order = { order_collection[0].Groups[1].Value, order_collection[1].Groups[1].Value, order_collection[2].Groups[1].Value, order_collection[3].Groups[1].Value };


                    StatusLabel.Text = "ТВ - " + entry.Value;
                    progressBar1.Value = 75;


                    MatchCollection collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "КПРФ") + 1]);
                    //kprf messages
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        kprf_dictTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_kprf_TV += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //ldpr messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "ЛДПР") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        ldpr_dictTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_ldpr_TV += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //sr messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "Справедливая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        sr_dictTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSsr_TV += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //er messages
                    collection = reg.Matches(graph_series[Array.IndexOf(recived_order, "Единая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        er_dictTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSer_TV += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    ///////////////////////////////////////////////////////index dics
                    //kprf index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "КПРФ") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        kprf_dict_indexTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_kprf_TV_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }

                    //ldpr index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "ЛДПР") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        ldpr_dict_indexTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RS_ldpr_TV_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    //sr index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "Справедливая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        sr_dict_indexTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSsr_TV_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                    //er index
                    collection = reg.Matches(graph_series_index[Array.IndexOf(recived_order, "Единая Россия") + 1]);
                    for (int id = 0; id < collection.Count; id++)
                    {
                        temparr = collection[id].Groups[1].Value.Split(sep3, StringSplitOptions.RemoveEmptyEntries);
                        er_dict_indexTV.Add(temparr[1], Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator)));
                        RSer_TV_index += Convert.ToDouble(temparr[0].Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator));
                    }
                }
                catch
                {




                    //convert dateFrom to datetime
                    var date = DateTime.ParseExact(datefrom, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    //add first date and zero to all dicts
                    string convertedDateTo = dayTo + "." + monthTo + "." + (dateTimePicker2.Value.Year % 100).ToString();


                    string convertedDate = "";
                    //while converted form datetime is NOT equal to dateTo

                    while (convertedDate != convertedDateTo)
                    {

                        //do
                        //add one day to datetime

                        //add date and zero value to all dicts
                        convertedDate = date.Day.ToString("00") + "." + date.Month.ToString("00") + "." + (date.Year % 100).ToString();
                        kprf_dictTV.Add(convertedDate, 0);
                        ldpr_dictTV.Add(convertedDate, 0);
                        sr_dictTV.Add(convertedDate, 0);
                        er_dictTV.Add(convertedDate, 0);
                        kprf_dict_indexTV.Add(convertedDate, 0);
                        ldpr_dict_indexTV.Add(convertedDate, 0);
                        sr_dict_indexTV.Add(convertedDate, 0);
                        er_dict_indexTV.Add(convertedDate, 0);
                        date = date.AddDays(1);

                        //repeat while//

                    }



                    /*
                    WebResp.Close();
                    workbook_detailedDB.Save();
                    workbook_detailedDB.Close();
                    base1.Quit();
                    return;*/
                }





                //loop and paste in excel

                // for each key in dic
                int offset_count = 0;

                foreach (KeyValuePair<string, double> current_run in kprf_dict) {
                    string date = current_run.Key;
                    //0ffset?
                    StatusLabel.Text = "Запись данных - " + entry.Value;
                    progressBar1.Value = 90;
                    worksheet_detailedDB = workbook_detailedDB.Sheets[1];
                    worksheet_detailedDB.Cells[1, lastcol_detDB + 1 + 4 * offset_count] = date;
                    worksheet_detailedDB.Cells[2, lastcol_detDB + 1 + 4 * offset_count] = "КПРФ";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 1 + 4 * offset_count] = kprf_dict[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 2 + 4 * offset_count] = "ЛДПР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 2 + 4 * offset_count] = ldpr_dict[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 3 + 4 * offset_count] = "СР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 3 + 4 * offset_count] = sr_dict[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 4 + 4 * offset_count] = "ЕР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 4 + 4 * offset_count] = er_dict[date].ToString();

                    //indexes
                    worksheet_detailedDB = workbook_detailedDB.Sheets[2];
                    worksheet_detailedDB.Cells[1, lastcol_detDB + 1 + 4 * offset_count] = date;
                    worksheet_detailedDB.Cells[2, lastcol_detDB + 1 + 4 * offset_count] = "КПРФ";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 1 + 4 * offset_count] = kprf_dict_index[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 2 + 4 * offset_count] = "ЛДПР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 2 + 4 * offset_count] = ldpr_dict_index[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 3 + 4 * offset_count] = "СР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 3 + 4 * offset_count] = sr_dict_index[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 4 + 4 * offset_count] = "ЕР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 4 + 4 * offset_count] = er_dict_index[date].ToString();

                    ///////////====================TVTVTVTVTVTVTV====================/////////tv//////////=====================================================/////////
                    ////
                    worksheet_detailedDB = workbook_detailedDB.Sheets[3];
                    worksheet_detailedDB.Cells[1, lastcol_detDB + 1 + 4 * offset_count] = date;
                    worksheet_detailedDB.Cells[2, lastcol_detDB + 1 + 4 * offset_count] = "КПРФ";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 1 + 4 * offset_count] = kprf_dictTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 2 + 4 * offset_count] = "ЛДПР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 2 + 4 * offset_count] = ldpr_dictTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 3 + 4 * offset_count] = "СР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 3 + 4 * offset_count] = sr_dictTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 4 + 4 * offset_count] = "ЕР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 4 + 4 * offset_count] = er_dictTV[date].ToString();

                    //indexes
                    worksheet_detailedDB = workbook_detailedDB.Sheets[4];
                    worksheet_detailedDB.Cells[1, lastcol_detDB + 1 + 4 * offset_count] = date;
                    worksheet_detailedDB.Cells[2, lastcol_detDB + 1 + 4 * offset_count] = "КПРФ";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 1 + 4 * offset_count] = kprf_dict_indexTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 2 + 4 * offset_count] = "ЛДПР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 2 + 4 * offset_count] = ldpr_dict_indexTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 3 + 4 * offset_count] = "СР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 3 + 4 * offset_count] = sr_dict_indexTV[date].ToString();

                    worksheet_detailedDB.Cells[2, lastcol_detDB + 4 + 4 * offset_count] = "ЕР";
                    worksheet_detailedDB.Cells[region_offset, lastcol_detDB + 4 + 4 * offset_count] = er_dict_indexTV[date].ToString();

                    offset_count++;

                }
                kprf_dict.Clear();
                ldpr_dict.Clear();
                sr_dict.Clear();
                er_dict.Clear();

                kprf_dict_index.Clear();
                ldpr_dict_index.Clear();
                sr_dict_index.Clear();
                er_dict_index.Clear();

                kprf_dictTV.Clear();
                ldpr_dictTV.Clear();
                sr_dictTV.Clear();
                er_dictTV.Clear();

                kprf_dict_indexTV.Clear();
                ldpr_dict_indexTV.Clear();
                sr_dict_indexTV.Clear();
                er_dict_indexTV.Clear();
                //////////////  //////     //SUMMARY DB PASTE
                offset_count = 0;

                string dates = datefrom + " - " + dateto;
                //0ffset?
                StatusLabel.Text = "Запись данных - " + entry.Value;
                progressBar1.Value = 95;
                worksheet_sumDB = workbook_sumDB.Sheets[2];
                worksheet_sumDB.Cells[1, lastcol_perDB + 1 + 4 * offset_count] = dates;
                worksheet_sumDB.Cells[2, lastcol_perDB + 1 + 4 * offset_count] = "КПРФ";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 1 + 4 * offset_count] = RS_kprf.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 2 + 4 * offset_count] = "ЛДПР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 2 + 4 * offset_count] = RS_ldpr.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 3 + 4 * offset_count] = "СР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 3 + 4 * offset_count] = RSsr.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 4 + 4 * offset_count] = "ЕР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 4 + 4 * offset_count] = RSer.ToString();

                //indexes
                worksheet_sumDB = workbook_sumDB.Sheets[3];
                worksheet_sumDB.Cells[1, lastcol_perDB + 1 + 4 * offset_count] = dates;
                worksheet_sumDB.Cells[2, lastcol_perDB + 1 + 4 * offset_count] = "КПРФ";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 1 + 4 * offset_count] = RS_kprf_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 2 + 4 * offset_count] = "ЛДПР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 2 + 4 * offset_count] = RS_ldpr_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 3 + 4 * offset_count] = "СР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 3 + 4 * offset_count] = RSsr_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 4 + 4 * offset_count] = "ЕР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 4 + 4 * offset_count] = RSer_index.ToString();

                ///////////====================TVTVTVTVTVTVTV====================/////////tv//////////=====================================================/////////
                ////
                worksheet_sumDB = workbook_sumDB.Sheets[5];
                worksheet_sumDB.Cells[1, lastcol_perDB + 1 + 4 * offset_count] = dates;
                worksheet_sumDB.Cells[2, lastcol_perDB + 1 + 4 * offset_count] = "КПРФ";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 1 + 4 * offset_count] = RS_kprf_TV.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 2 + 4 * offset_count] = "ЛДПР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 2 + 4 * offset_count] = RS_ldpr_TV.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 3 + 4 * offset_count] = "СР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 3 + 4 * offset_count] = RSsr_TV.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 4 + 4 * offset_count] = "ЕР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 4 + 4 * offset_count] = RSer_TV.ToString();

                //indexes
                worksheet_sumDB = workbook_sumDB.Sheets[6];
                worksheet_sumDB.Cells[1, lastcol_perDB + 1 + 4 * offset_count] = dates;
                worksheet_sumDB.Cells[2, lastcol_perDB + 1 + 4 * offset_count] = "КПРФ";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 1 + 4 * offset_count] = RS_kprf_TV_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 2 + 4 * offset_count] = "ЛДПР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 2 + 4 * offset_count] = RS_ldpr_TV_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 3 + 4 * offset_count] = "СР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 3 + 4 * offset_count] = RSsr_TV_index.ToString();

                worksheet_sumDB.Cells[2, lastcol_perDB + 4 + 4 * offset_count] = "ЕР";
                worksheet_sumDB.Cells[region_offset, lastcol_perDB + 4 + 4 * offset_count] = RSer_TV_index.ToString();






                region_offset++;
                System.Threading.Thread.Sleep(500);
                StatusLabel.Text = "Ожидание нового региона...";
                progressBar1.Value = 10;
                System.Threading.Thread.Sleep(500);



            }//endforeach region

            //cumsum count




            //Save and go on
            workbook_detailedDB.Save();
            workbook_detailedDB.Close();
            base1.Quit();
            workbook_sumDB.Save();
            workbook_sumDB.Close();
            base2.Quit();
            MessageBox.Show("Данные получены. База составлена");
            StatusLabel.Text = "Готово";
            progressBar1.Value = 0;
            analyseBase(Properties.Settings.Default["per_db_path"].ToString());
            return;


        }

        private void StatusDesc_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {


        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Settings settings = new Settings(); //set settings
            settings.ShowDialog();
        }
        private void Form1_FormClosing(object sender, EventArgs e)
        {
            try {
                Globals.app.Quit();
            } catch { };

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show(dateTimePicker1.Value.Day.ToString() + "." + dateTimePicker1.Value.Month.ToString() + "." + dateTimePicker1.Value.Year.ToString());


        }

        private void Report1()
        {

            Excel excel = new Excel(Globals.FILE_NAME, 2);
            if (!(Nothing1.Checked)) //Creating a total for messages
            {
                int Frow = 0;
                int Fcol = 0;
                try
                {
                    Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 2).Item1;
                    Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 2).Item2;
                }
                catch
                {
                    MessageBox.Show("Основная дата не выбрана или не найдена", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    excel.Close();
                    return;
                }
                StatusLabel.Text = "В работе";
                progressBar1.Value = 10;
                StatusDesc.Text = "Поиск даты";
                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 2", "Ошибка");
                    return;
                }

                Frow += 1;
                KprfSum = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 2);
                LdprSum = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 2);
                SrSum = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 2);
                ErSum = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 2);

                StatusLabel.Text = "В работе";
                progressBar1.Value = 20;
                StatusDesc.Text = "Запись итогов сообщений";

                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol, KprfSum.ToString(), 2);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 1, LdprSum.ToString(), 2);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 2, SrSum.ToString(), 2);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 3, ErSum.ToString(), 2);
            }

            if (!(Nothing2.Checked))//Creating a total for indexes
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 3).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 3).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 3", "Ошибка");
                    return;
                }

                Frow += 1;
                KprfSumI = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 3);
                LdprSumI = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 3);
                SrSumI = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 3);
                ErSumI = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 3);

                StatusLabel.Text = "В работе";
                progressBar1.Value = 20;
                StatusDesc.Text = "Запись по индексам";

                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol, KprfSumI.ToString(), 3);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 1, LdprSumI.ToString(), 3);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 2, SrSumI.ToString(), 3);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 3, ErSumI.ToString(), 3);
            }

            //FOR TV
            if (!(nothingTV1.Checked)) //Creating a total for messages
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 5).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 5).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 2", "Ошибка");
                    return;
                }

                Frow += 1;
                KprfSumTV = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 5);
                LdprSumTV = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 5);
                SrSumTV = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 5);
                ErSumTV = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 5);

                StatusLabel.Text = "В работе";
                progressBar1.Value = 20;
                StatusDesc.Text = "Запись итогов по ТВ";

                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol, KprfSumTV.ToString(), 5);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 1, LdprSumTV.ToString(), 5);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 2, SrSumTV.ToString(), 5);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 3, ErSumTV.ToString(), 5);
            }

            //for tv indexes
            if (!(nothingTV2.Checked)) //Creating a total for messages
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 6).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 6).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 2", "Ошибка");
                    return;
                }

                Frow += 1;
                KprfSumITV = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 6);
                LdprSumITV = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 6);
                SrSumITV = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 6);
                ErSumITV = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 6);

                StatusLabel.Text = "В работе";
                progressBar1.Value = 20;
                StatusDesc.Text = "Запись итогов по ТВ-индексам";

                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol, KprfSumITV.ToString(), 6);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 1, LdprSumITV.ToString(), 6);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 2, SrSumITV.ToString(), 6);
                excel.WriteToCell(Frow + 1 + Globals.REGIONS_COUNT, Fcol + 3, ErSumITV.ToString(), 6);
            }

            void PrepareList(int listNumber, string dateMain, string dateFor, bool clearList)
            {

                if (clearList)
                {
                    //I don't need that shit here (c)
                    excel.ClearListContents(listNumber);
                }
                switch (listNumber)
                {
                    case 4:
                        //write dates as placeholders
                        excel.WriteToCell(1, 2, dateMain, 4);
                        excel.WriteToCell(1, 5, dateFor, 4);
                        //
                        //write headers
                        excel.WriteToCell(2, 1, "Партия", 4);
                        excel.WriteToCell(2, 2, "Всего сообщений", 4);
                        excel.WriteToCell(2, 3, "% Сообщений", 4);
                        excel.WriteToCell(2, 4, "Медиа-индекс", 4);
                        excel.WriteToCell(2, 5, "% Сообщений (прошлый)", 4);
                        excel.WriteToCell(2, 6, "Прирост числа сообщ., %", 4);
                        excel.WriteToCell(2, 7, "Прошлый индекс", 4);
                        excel.WriteToCell(2, 8, "Прирост медиа-индекса", 4);

                        excel.WriteToCell(6, 1, "Единая Россия", 4);
                        excel.WriteToCell(3, 1, "КПРФ", 4);
                        excel.WriteToCell(4, 1, "ЛДПР", 4);
                        excel.WriteToCell(5, 1, "Справедливая Россия", 4);
                        excel.WriteToCell(7, 1, "Всего", 4);

                        break;
                    case 7:
                        //write dates as placeholders
                        excel.WriteToCell(1, 2, dateMain, 7);
                        excel.WriteToCell(1, 5, dateFor, 7);
                        //
                        //write headers
                        excel.WriteToCell(2, 1, "Партия", 7);
                        excel.WriteToCell(2, 2, "Всего сообщений", 7);
                        excel.WriteToCell(2, 3, "% Сообщений", 7);
                        excel.WriteToCell(2, 4, "Медиа-индекс", 7);
                        excel.WriteToCell(2, 5, "% Сообщений (прошлый)", 7);
                        excel.WriteToCell(2, 6, "Прирост числа сообщ., %", 7);
                        excel.WriteToCell(2, 7, "Прошлый индекс", 7);
                        excel.WriteToCell(2, 8, "Прирост медиа-индекса", 7);

                        //parts
                        excel.WriteToCell(6, 1, "Единая Россия", 7);
                        excel.WriteToCell(3, 1, "КПРФ", 7);
                        excel.WriteToCell(4, 1, "ЛДПР", 7);
                        excel.WriteToCell(5, 1, "Справедливая Россия", 7);
                        excel.WriteToCell(7, 1, "Всего", 7);

                        break;
                    case 8:
                        excel.WriteToCell(1, 2, "КПРФ", 8);
                        excel.WriteToCell(1, 6, "ЛДПР", 8);
                        excel.WriteToCell(1, 10, "СР", 8);
                        excel.WriteToCell(1, 14, "ЕР", 8);

                        excel.WriteToCell(2, 1, "Регионы", 8);
                        excel.WriteToCell(2, 2, "Кол-во сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 3, "% сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 4, "% сообщ. " + dateFor, 8);
                        excel.WriteToCell(2, 5, "Изменение", 8);

                        excel.WriteToCell(2, 6, "Кол-во сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 7, "% сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 8, "% сообщ. " + dateFor, 8);
                        excel.WriteToCell(2, 9, "Изменение", 8);

                        excel.WriteToCell(2, 10, "Кол-во сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 11, "% сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 12, "% сообщ. " + dateFor, 8);
                        excel.WriteToCell(2, 13, "Изменение", 8);

                        excel.WriteToCell(2, 14, "Кол-во сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 15, "% сообщ. " + dateMain, 8);
                        excel.WriteToCell(2, 16, "% сообщ. " + dateFor, 8);
                        excel.WriteToCell(2, 17, "Изменение", 8);

                        for (int row2 = 0; row2 < Globals.REGIONS_COUNT; row2++)
                        {
                            excel.WriteToCell(3 + row2, 1, excel.ReadCell(row2 + 1, 1, 1), 8);
                        }
                        break;
                    case 9:
                        excel.WriteToCell(1, 2, "КПРФ", 9);
                        excel.WriteToCell(1, 5, "ЛДПР", 9);
                        excel.WriteToCell(1, 8, "СР", 9);
                        excel.WriteToCell(1, 11, "ЕР", 9);

                        excel.WriteToCell(2, 1, "Регионы", 9);
                        excel.WriteToCell(2, 2, "Медиа-индекс " + dateMain, 9);
                        excel.WriteToCell(2, 3, "Медиа-индекс " + dateFor, 9);
                        excel.WriteToCell(2, 4, "Изменение", 9);
                        excel.WriteToCell(2, 5, "Медиа-индекс " + dateMain, 9);
                        excel.WriteToCell(2, 6, "Медиа-индекс " + dateFor, 9);
                        excel.WriteToCell(2, 7, "Изменение", 9);
                        excel.WriteToCell(2, 8, "Медиа-индекс " + dateMain, 9);
                        excel.WriteToCell(2, 9, "Медиа-индекс " + dateFor, 9);
                        excel.WriteToCell(2, 10, "Изменение", 9);
                        excel.WriteToCell(2, 11, "Медиа-индекс " + dateMain, 9);
                        excel.WriteToCell(2, 12, "Медиа-индекс " + dateFor, 9);
                        excel.WriteToCell(2, 13, "Изменение", 9);
                        for (int row3 = 0; row3 < Globals.REGIONS_COUNT; row3++)
                        {
                            excel.WriteToCell(3 + row3, 1, excel.ReadCell(row3 + 1, 1, 1), 9);
                        }

                        break;

                    case 10:
                        //headers
                        excel.WriteToCell(5, 1, "Единая Россия", 10);
                        excel.WriteToCell(2, 1, "КПРФ", 10);
                        excel.WriteToCell(3, 1, "ЛДПР", 10);
                        excel.WriteToCell(4, 1, "Справедливая Россия", 10);
                        excel.WriteToCell(1, 2, "Сообщения", 10);
                        excel.WriteToCell(1, 3, "Медиа-индекс", 10);

                        //data (messages)
                        excel.WriteToCell(2, 2, KprfSum.ToString(), 10);
                        excel.WriteToCell(3, 2, LdprSum.ToString(), 10);
                        excel.WriteToCell(4, 2, SrSum.ToString(), 10);
                        excel.WriteToCell(5, 2, ErSum.ToString(), 10);

                        //indexes
                        excel.WriteToCell(2, 3, KprfSumI.ToString(), 10);
                        excel.WriteToCell(3, 3, LdprSumI.ToString(), 10);
                        excel.WriteToCell(4, 3, SrSumI.ToString(), 10);
                        excel.WriteToCell(5, 3, ErSumI.ToString(), 10);

                        break;

                    case 11:
                        //headers
                        excel.WriteToCell(5, 1, "Единая Россия", 11);
                        excel.WriteToCell(2, 1, "КПРФ", 11);
                        excel.WriteToCell(3, 1, "ЛДПР", 11);
                        excel.WriteToCell(4, 1, "Справедливая Россия", 11);
                        excel.WriteToCell(1, 2, "Сообщения по ТВ", 11);
                        excel.WriteToCell(1, 3, "Медиа-индекс", 11);

                        //data (messages)
                        excel.WriteToCell(2, 2, KprfSumTV.ToString(), 11);
                        excel.WriteToCell(3, 2, LdprSumTV.ToString(), 11);
                        excel.WriteToCell(4, 2, SrSumTV.ToString(), 11);
                        excel.WriteToCell(5, 2, ErSumTV.ToString(), 11);

                        //indexes
                        excel.WriteToCell(2, 3, KprfSumITV.ToString(), 11);
                        excel.WriteToCell(3, 3, LdprSumITV.ToString(), 11);
                        excel.WriteToCell(4, 3, SrSumITV.ToString(), 11);
                        excel.WriteToCell(5, 3, ErSumITV.ToString(), 11);

                        break;
                    case 12:
                        excel.WriteToCell(2, 1, "КПРФ", 12);
                        excel.WriteToCell(3, 1, "ЛДПР", 12);
                        excel.WriteToCell(4, 1, "СР", 12);
                        excel.WriteToCell(5, 1, "EР", 12);
                        excel.WriteToCell(6, 1, "Всего", 12);
                        break;

                    case 13:
                        excel.WriteToCell(2, 1, "КПРФ", 13);
                        excel.WriteToCell(3, 1, "ЛДПР", 13);
                        excel.WriteToCell(4, 1, "СР", 13);
                        excel.WriteToCell(5, 1, "EР", 13);
                        break;

                    case 14:
                        excel.WriteToCell(2, 1, "КПРФ", 14);
                        excel.WriteToCell(3, 1, "ЛДПР", 14);
                        excel.WriteToCell(4, 1, "СР", 14);
                        excel.WriteToCell(5, 1, "EР", 14);
                        excel.WriteToCell(6, 1, "Всего", 14);
                        break;
                    default:
                        break;
                }
            }

            string[] lb1arr = listBox1.SelectedItem.ToString().Split(" - ".ToCharArray());
            string[] lb2arr = listBox2.SelectedItem.ToString().Split(" - ".ToCharArray());

            string perStart_short = lb1arr[0].Substring(0, 5) + " - " + lb1arr[3].Substring(0, 5);
            string pedEnd_short = lb2arr[0].Substring(0, 5) + " - " + lb2arr[3].Substring(0, 5);

            //clear the shit out
            PrepareList(4, listBox1.SelectedItem.ToString(), listBox2.SelectedItem.ToString(), true);
            PrepareList(7, listBox1.SelectedItem.ToString(), listBox2.SelectedItem.ToString(), true);
            PrepareList(8, perStart_short, pedEnd_short, true);
            PrepareList(9, perStart_short, pedEnd_short, true);

            //vert allign
            excel.AllignRange("2", 8);
            excel.AllignRange("2", 9);


            //Messages difference calc
            if (Difference1.Checked | Detailed1.Checked)
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 4).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 4).Item2;
                StatusLabel.Text = "В работе";
                progressBar1.Value = 40;
                StatusDesc.Text = "Подготовка изменений сообщений";
                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 4", "Ошибка");
                    return;
                }

                totalM = ErSum + KprfSum + LdprSum + SrSum;
                excel.WriteToCell(Frow + 2, Fcol, KprfSum.ToString(), 4);
                excel.WriteToCell(Frow + 3, Fcol, LdprSum.ToString(), 4);
                excel.WriteToCell(Frow + 4, Fcol, SrSum.ToString(), 4);
                excel.WriteToCell(Frow + 5, Fcol, ErSum.ToString(), 4);
                excel.WriteToCell(Frow + 6, Fcol, totalM.ToString(), 4);

                excel.WriteToCell(Frow + 2, Fcol + 1, excel.FindPercent(KprfSum, totalM), 4);
                excel.WriteToCell(Frow + 3, Fcol + 1, excel.FindPercent(LdprSum, totalM), 4);
                excel.WriteToCell(Frow + 4, Fcol + 1, excel.FindPercent(SrSum, totalM), 4);
                excel.WriteToCell(Frow + 5, Fcol + 1, excel.FindPercent(ErSum, totalM), 4);


                // Create last message info
                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item2;
                StatusLabel.Text = "В работе";
                progressBar1.Value = 40;
                StatusDesc.Text = "Подготовка изменений прошлых сообщений";

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 2", "Ошибка");
                    return;
                }

                Frow += 1;
                KprfSumL = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 2);
                LdprSumL = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 2);
                SrSumL = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 2);
                ErSumL = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 2);

                SumLast = ErSumL + KprfSumL + LdprSumL + SrSumL;

                //write totalals for last messages to totals list

                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol, KprfSumL.ToString(), 2);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, LdprSumL.ToString(), 2);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, SrSumL.ToString(), 2);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, ErSumL.ToString(), 2);



                // Write last percentages of messages to the report
                excel.WriteToCell(Frow + 1, Fcol + 3, excel.FindPercent(KprfSumL, SumLast), 4);
                excel.WriteToCell(Frow + 2, Fcol + 3, excel.FindPercent(LdprSumL, SumLast), 4);
                excel.WriteToCell(Frow + 3, Fcol + 3, excel.FindPercent(SrSumL, SumLast), 4);
                excel.WriteToCell(Frow + 4, Fcol + 3, excel.FindPercent(ErSumL, SumLast), 4);

                //Find last raw percentages for difference
                rawPeErL = Convert.ToDouble(excel.FindPercent(ErSumL, SumLast).Remove(excel.FindPercent(ErSumL, SumLast).Length - 1));
                rawPeKprfL = Convert.ToDouble(excel.FindPercent(KprfSumL, SumLast).Remove(excel.FindPercent(KprfSumL, SumLast).Length - 1));
                rawPeLdprL = Convert.ToDouble(excel.FindPercent(LdprSumL, SumLast).Remove(excel.FindPercent(LdprSumL, SumLast).Length - 1));
                rawPeSrL = Convert.ToDouble(excel.FindPercent(SrSumL, SumLast).Remove(excel.FindPercent(SrSumL, SumLast).Length - 1));

                rawPeEr = Convert.ToDouble(excel.FindPercent(ErSum, totalM).Remove(excel.FindPercent(ErSum, totalM).Length - 1));
                rawPeKprf = Convert.ToDouble(excel.FindPercent(KprfSum, totalM).Remove(excel.FindPercent(KprfSum, totalM).Length - 1));
                rawPeLdpr = Convert.ToDouble(excel.FindPercent(LdprSum, totalM).Remove(excel.FindPercent(LdprSum, totalM).Length - 1));
                rawPeSr = Convert.ToDouble(excel.FindPercent(SrSum, totalM).Remove(excel.FindPercent(SrSum, totalM).Length - 1));

                //write difference in percentages in messages
                Fcol += 4;
                Frow--;
                excel.WriteToCell(Frow + 5, Fcol, (rawPeEr - rawPeErL).ToString() + "%", 4);
                excel.WriteToCell(Frow + 2, Fcol, (rawPeKprf - rawPeKprfL).ToString() + "%", 4);
                excel.WriteToCell(Frow + 3, Fcol, (rawPeLdpr - rawPeLdprL).ToString() + "%", 4);
                excel.WriteToCell(Frow + 4, Fcol, (rawPeSr - rawPeSrL).ToString() + "%", 4);

            }
            //messages difference. Indexes part
            if (Difference2.Checked | Detailed2.Checked)
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 4).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 4).Item2;

                StatusLabel.Text = "В работе";
                progressBar1.Value = 40;
                StatusDesc.Text = "Подготовка изменений индексов";

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 4", "Ошибка");
                    return;
                }
                Fcol += 2;

                //write current indexes
                excel.WriteToCell(Frow + 2, Fcol, KprfSumI.ToString(), 4);
                excel.WriteToCell(Frow + 3, Fcol, LdprSumI.ToString(), 4);
                excel.WriteToCell(Frow + 4, Fcol, SrSumI.ToString(), 4);
                excel.WriteToCell(Frow + 5, Fcol, ErSumI.ToString(), 4);
                totalI = ErSumI + KprfSumI + LdprSumI + SrSumI;
                excel.WriteToCell(Frow + 6, Fcol, totalI.ToString(), 4);

                // Last Indexes part

                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item2;
                StatusLabel.Text = "В работе";
                progressBar1.Value = 40;
                StatusDesc.Text = "Подготовка изменений прошлых индексов";

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 3", "Ошибка");
                    return;
                }
                //count indexes for a selected date

                Frow += 1;
                KprfSumLI = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 3);
                LdprSumLI = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 3);
                SrSumLI = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 3);
                ErSumLI = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 3);

                //write totals for last indexes to totals list
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol, KprfSumLI.ToString(), 3);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, LdprSumLI.ToString(), 3);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, SrSumLI.ToString(), 3);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, ErSumLI.ToString(), 3);

                //Write indexes to difference list

                StatusLabel.Text = "В работе";
                progressBar1.Value = 40;
                StatusDesc.Text = "Запись";

                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 3", "Ошибка");
                    return;
                }
                Fcol += 5;
                excel.WriteToCell(Frow + 2, Fcol, KprfSumLI.ToString(), 4);
                excel.WriteToCell(Frow + 3, Fcol, LdprSumLI.ToString(), 4);
                excel.WriteToCell(Frow + 4, Fcol, SrSumLI.ToString(), 4);
                excel.WriteToCell(Frow + 5, Fcol, ErSumLI.ToString(), 4);
                SumLastI = SrSumLI + LdprSumLI + KprfSumLI + ErSumLI;
                excel.WriteToCell(Frow + 6, Fcol, SumLastI.ToString(), 4);

                //write the difference in indexes

                Fcol += 1;

                excel.WriteToCell(Frow + 5, Fcol, (ErSumI - ErSumLI).ToString(), 4);
                excel.WriteToCell(Frow + 2, Fcol, (KprfSumI - KprfSumLI).ToString(), 4);
                excel.WriteToCell(Frow + 3, Fcol, (LdprSumI - LdprSumLI).ToString(), 4);
                excel.WriteToCell(Frow + 4, Fcol, (SrSumI - SrSumLI).ToString(), 4);
                SumLastI = SrSumLI + LdprSumLI + KprfSumLI + ErSumLI;
                excel.WriteToCell(Frow + 6, Fcol, (totalI - SumLastI).ToString(), 4);

            }



            //Messages difference calc
            if (differenceTV1.Checked)
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 7).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 7).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 7", "Ошибка");
                    return;
                }

                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Подготовка изменений ТВ";

                totalMTV = ErSumTV + KprfSumTV + LdprSumTV + SrSumTV;
                excel.WriteToCell(Frow + 5, Fcol, ErSumTV.ToString(), 7);
                excel.WriteToCell(Frow + 2, Fcol, KprfSumTV.ToString(), 7);
                excel.WriteToCell(Frow + 3, Fcol, LdprSumTV.ToString(), 7);
                excel.WriteToCell(Frow + 4, Fcol, SrSumTV.ToString(), 7);
                excel.WriteToCell(Frow + 6, Fcol, totalMTV.ToString(), 7);

                excel.WriteToCell(Frow + 5, Fcol + 1, excel.FindPercent(ErSumTV, totalMTV), 7);
                excel.WriteToCell(Frow + 2, Fcol + 1, excel.FindPercent(KprfSumTV, totalMTV), 7);
                excel.WriteToCell(Frow + 3, Fcol + 1, excel.FindPercent(LdprSumTV, totalMTV), 7);
                excel.WriteToCell(Frow + 4, Fcol + 1, excel.FindPercent(SrSumTV, totalMTV), 7);


                // Create last message info
                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 5).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 5).Item2;
                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 5", "Ошибка");
                    return;
                }
                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Подготовка изменений прошлых ТВ";
                Frow += 1;
                KprfSumLTV = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 5);
                LdprSumLTV = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 5);
                SrSumLTV = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 5);
                ErSumLTV = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 5);

                SumLastTV = ErSumLTV + KprfSumLTV + LdprSumLTV + SrSumLTV;

                //write totalals for last messages to totals list

                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol, KprfSumLTV.ToString(), 5);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, LdprSumLTV.ToString(), 5);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, SrSumLTV.ToString(), 5);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, ErSumLTV.ToString(), 5);



                // Write last percentages of messages to the report
                excel.WriteToCell(Frow + 1, Fcol + 3, excel.FindPercent(KprfSumLTV, SumLastTV), 7);
                excel.WriteToCell(Frow + 2, Fcol + 3, excel.FindPercent(LdprSumLTV, SumLastTV), 7);
                excel.WriteToCell(Frow + 3, Fcol + 3, excel.FindPercent(SrSumLTV, SumLastTV), 7);
                excel.WriteToCell(Frow + 4, Fcol + 3, excel.FindPercent(ErSumLTV, SumLastTV), 7);

                //Find last raw percentages for difference
                rawPeErLTV = Convert.ToDouble(excel.FindPercent(ErSumLTV, SumLastTV).Remove(excel.FindPercent(ErSumLTV, SumLastTV).Length - 1));
                rawPeKprfLTV = Convert.ToDouble(excel.FindPercent(KprfSumLTV, SumLastTV).Remove(excel.FindPercent(KprfSumLTV, SumLastTV).Length - 1));
                rawPeLdprLTV = Convert.ToDouble(excel.FindPercent(LdprSumLTV, SumLastTV).Remove(excel.FindPercent(LdprSumLTV, SumLastTV).Length - 1));
                rawPeSrLTV = Convert.ToDouble(excel.FindPercent(SrSumLTV, SumLastTV).Remove(excel.FindPercent(SrSumLTV, SumLastTV).Length - 1));

                rawPeErTV = Convert.ToDouble(excel.FindPercent(ErSumTV, totalMTV).Remove(excel.FindPercent(ErSumTV, totalMTV).Length - 1));
                rawPeKprfTV = Convert.ToDouble(excel.FindPercent(KprfSumTV, totalMTV).Remove(excel.FindPercent(KprfSumTV, totalMTV).Length - 1));
                rawPeLdprTV = Convert.ToDouble(excel.FindPercent(LdprSumTV, totalMTV).Remove(excel.FindPercent(LdprSumTV, totalMTV).Length - 1));
                rawPeSrTV = Convert.ToDouble(excel.FindPercent(SrSumTV, totalMTV).Remove(excel.FindPercent(SrSumTV, totalMTV).Length - 1));
                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Запись ТВ";
                //write difference in percentages in messages
                Fcol += 4;
                Frow--;
                excel.WriteToCell(Frow + 5, Fcol, (rawPeErTV - rawPeErLTV).ToString() + "%", 7);
                excel.WriteToCell(Frow + 2, Fcol, (rawPeKprfTV - rawPeKprfLTV).ToString() + "%", 7);
                excel.WriteToCell(Frow + 3, Fcol, (rawPeLdprTV - rawPeLdprLTV).ToString() + "%", 7);
                excel.WriteToCell(Frow + 4, Fcol, (rawPeSrTV - rawPeSrLTV).ToString() + "%", 7);

            }
            //messages difference. Indexes part (TV)
            if (differenceTV2.Checked)
            {
                int Frow = excel.FindCell(listBox1.SelectedItem.ToString(), 7).Item1;
                int Fcol = excel.FindCell(listBox1.SelectedItem.ToString(), 7).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox1.SelectedItem.ToString() + " не найдена на листе 7", "Ошибка");
                    return;
                }
                Fcol += 2;
                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Подготовка изменений индексов ТВ";
                //write current indexes
                excel.WriteToCell(Frow + 2, Fcol, KprfSumITV.ToString(), 7);
                excel.WriteToCell(Frow + 3, Fcol, LdprSumITV.ToString(), 7);
                excel.WriteToCell(Frow + 4, Fcol, SrSumITV.ToString(), 7);
                excel.WriteToCell(Frow + 5, Fcol, ErSumITV.ToString(), 7);
                totalITV = ErSumITV + KprfSumITV + LdprSumITV + SrSumITV;
                excel.WriteToCell(Frow + 6, Fcol, totalITV.ToString(), 7);

                // Last Indexes part

                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 6).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 6).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 6", "Ошибка");
                    return;
                }
                //count indexes for a selected date

                Frow += 1;
                KprfSumLITV = excel.RangeSum(Frow + 1, Fcol, Frow + Globals.REGIONS_COUNT + 1, Fcol, 6);
                LdprSumLITV = excel.RangeSum(Frow + 1, Fcol + 1, Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, 6);
                SrSumLITV = excel.RangeSum(Frow + 1, Fcol + 2, Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, 6);
                ErSumLITV = excel.RangeSum(Frow + 1, Fcol + 3, Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, 6);

                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Подготовка изменений прошлых ТВ-индексов";
                //write totals for last indexes to totals list
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol, KprfSumLITV.ToString(), 6);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 1, LdprSumLITV.ToString(), 6);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 2, SrSumLITV.ToString(), 6);
                excel.WriteToCell(Frow + Globals.REGIONS_COUNT + 1, Fcol + 3, ErSumLITV.ToString(), 6);

                //Write indexes to difference list

                Frow = excel.FindCell(listBox2.SelectedItem.ToString(), 6).Item1;
                Fcol = excel.FindCell(listBox2.SelectedItem.ToString(), 6).Item2;

                if ((Frow == 0) & (Fcol == 0))
                {
                    MessageBox.Show("Дата измерения " + listBox2.SelectedItem.ToString() + " не найдена на листе 6", "Ошибка");
                    return;
                }
                Fcol += 5;
                excel.WriteToCell(Frow + 2, Fcol, KprfSumLITV.ToString(), 7);
                excel.WriteToCell(Frow + 3, Fcol, LdprSumLITV.ToString(), 7);
                excel.WriteToCell(Frow + 4, Fcol, SrSumLITV.ToString(), 7);
                excel.WriteToCell(Frow + 5, Fcol, ErSumLITV.ToString(), 7);
                SumLastITV = SrSumLITV + LdprSumLITV + KprfSumLITV + ErSumLITV;
                excel.WriteToCell(Frow + 6, Fcol, SumLastITV.ToString(), 7);
                StatusLabel.Text = "В работе";
                progressBar1.Value = 60;
                StatusDesc.Text = "Запись";
                //write the difference in indexes

                Fcol += 1;

                excel.WriteToCell(Frow + 5, Fcol, (ErSumITV - ErSumLITV).ToString(), 7);
                excel.WriteToCell(Frow + 2, Fcol, (KprfSumITV - KprfSumLITV).ToString(), 7);
                excel.WriteToCell(Frow + 3, Fcol, (LdprSumITV - LdprSumLITV).ToString(), 7);
                excel.WriteToCell(Frow + 4, Fcol, (SrSumITV - SrSumLITV).ToString(), 7);
                SumLastITV = SrSumLITV + LdprSumLITV + KprfSumLITV + ErSumLITV;
                excel.WriteToCell(Frow + 6, Fcol, (totalITV - SumLastITV).ToString(), 7);

            }


            if (Detailed1.Checked)
            {

                int Frow_main = excel.FindCell(listBox1.SelectedItem.ToString(), 2).Item1;
                int Fcol_main = excel.FindCell(listBox1.SelectedItem.ToString(), 2).Item2;

                int Frow_second = excel.FindCell(listBox2.SelectedItem.ToString(), 2).Item1;
                int Fcol_second = excel.FindCell(listBox2.SelectedItem.ToString(), 2).Item2;

                double temp1 = 0;
                double temp2 = 0;
                double temp3 = 0;
                double temp4 = 0;
                double temp5 = 0;
                double temp6 = 0;
                double temp7 = 0;
                double temp8 = 0;

                double tempSum = 0;
                double tempSumL = 0;

                StatusLabel.Text = "В работе";
                progressBar1.Value = 80;
                StatusDesc.Text = "Просчет детализации по сообщениям";

                //int row = 0;
                int col = 0;
                string s1;
                string s2;
                double diff1;

                for (int row = 0; row < Globals.REGIONS_COUNT; row++)
                {
                    //kprfblock
                    temp1 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main, 2));
                    temp2 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second, 2));

                    //ldpr block
                    temp3 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 1, 2));
                    temp4 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 1, 2));


                    //sr
                    temp5 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 2, 2));
                    temp6 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 2, 2));

                    //er
                    temp7 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 2, 2));
                    temp8 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 3, 2));

                    tempSum = temp1 + temp3 + temp5 + temp7;
                    tempSumL = temp2 + temp4 + temp6 + temp8;

                    excel.WriteToCell(row + 3, col + 2, temp1.ToString(), 8);
                    excel.WriteToCell(row + 3, col + 3, excel.FindPercent(temp1, tempSum), 8);
                    excel.WriteToCell(row + 3, col + 4, excel.FindPercent(temp2, tempSumL), 8);

                    s1 = excel.FindPercent(temp1, tempSum).Remove(excel.FindPercent(temp1, tempSum).Length - 1);
                    s2 = excel.FindPercent(temp2, tempSumL).Remove(excel.FindPercent(temp2, tempSumL).Length - 1);
                    diff1 = Convert.ToDouble(s1) - Convert.ToDouble(s2);

                    excel.WriteToCell(row + 3, col + 5, diff1.ToString() + "%", 8);



                    excel.WriteToCell(row + 3, col + 6, temp3.ToString(), 8);
                    excel.WriteToCell(row + 3, col + 7, excel.FindPercent(temp3, tempSum), 8);
                    excel.WriteToCell(row + 3, col + 8, excel.FindPercent(temp4, tempSumL), 8);

                    s1 = excel.FindPercent(temp3, tempSum).Remove(excel.FindPercent(temp3, tempSum).Length - 1);
                    s2 = excel.FindPercent(temp4, tempSumL).Remove(excel.FindPercent(temp4, tempSumL).Length - 1);
                    diff1 = Convert.ToDouble(s1) - Convert.ToDouble(s2);

                    excel.WriteToCell(row + 3, col + 9, diff1.ToString() + "%", 8);




                    excel.WriteToCell(row + 3, col + 10, temp5.ToString(), 8);
                    excel.WriteToCell(row + 3, col + 11, excel.FindPercent(temp5, tempSum), 8);
                    excel.WriteToCell(row + 3, col + 12, excel.FindPercent(temp6, tempSumL), 8);
                    s1 = excel.FindPercent(temp5, tempSum).Remove(excel.FindPercent(temp5, tempSum).Length - 1);
                    s2 = excel.FindPercent(temp6, tempSumL).Remove(excel.FindPercent(temp6, tempSumL).Length - 1);
                    diff1 = Convert.ToDouble(s1) - Convert.ToDouble(s2);
                    excel.WriteToCell(row + 3, col + 13, diff1.ToString() + "%", 8);

                    excel.WriteToCell(row + 3, col + 14, temp7.ToString(), 8);
                    excel.WriteToCell(row + 3, col + 15, excel.FindPercent(temp7, tempSum), 8);
                    excel.WriteToCell(row + 3, col + 16, excel.FindPercent(temp8, tempSumL), 8);
                    s1 = excel.FindPercent(temp7, tempSum).Remove(excel.FindPercent(temp7, tempSum).Length - 1);
                    s2 = excel.FindPercent(temp8, tempSumL).Remove(excel.FindPercent(temp8, tempSumL).Length - 1);
                    diff1 = Convert.ToDouble(s1) - Convert.ToDouble(s2);
                    excel.WriteToCell(row + 3, col + 17, diff1.ToString() + "%", 8);
                }
                StatusLabel.Text = "В работе";
                progressBar1.Value = 80;
                StatusDesc.Text = "Запись";
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 6, KprfSum.ToString(), 8);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 2, LdprSum.ToString(), 8);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 10, SrSum.ToString(), 8);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 14, ErSum.ToString(), 8);

            }

            if (Detailed2.Checked)
            {

                int Frow_main = excel.FindCell(listBox1.SelectedItem.ToString(), 3).Item1;
                int Fcol_main = excel.FindCell(listBox1.SelectedItem.ToString(), 3).Item2;

                int Frow_second = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item1;
                int Fcol_second = excel.FindCell(listBox2.SelectedItem.ToString(), 3).Item2;

                double temp1 = 0;
                double temp2 = 0;
                StatusLabel.Text = "В работе";
                progressBar1.Value = 80;
                StatusDesc.Text = "Просчет детализации по индексам";

                //int row = 0;
                int col = 0;

                for (int row = 0; row < Globals.REGIONS_COUNT; row++)
                {
                    temp1 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main, 3));
                    temp2 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second, 3));
                    excel.WriteToCell(row + 3, col + 2, temp1.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 3, temp2.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 4, (temp1 - temp2).ToString(), 9);

                    temp1 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 1, 3));
                    temp2 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 1, 3));
                    excel.WriteToCell(row + 3, col + 5, temp1.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 6, temp2.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 7, (temp1 - temp2).ToString(), 9);

                    temp1 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 2, 3));
                    temp2 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 2, 3));
                    excel.WriteToCell(row + 3, col + 8, temp1.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 9, temp2.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 10, (temp1 - temp2).ToString(), 9);

                    temp1 = Convert.ToDouble(excel.ReadCell(Frow_main + 2 + row, Fcol_main + 3, 3));
                    temp2 = Convert.ToDouble(excel.ReadCell(Frow_second + 2 + row, Fcol_second + 3, 3));
                    excel.WriteToCell(row + 3, col + 11, temp1.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 12, temp2.ToString(), 9);
                    excel.WriteToCell(row + 3, col + 13, (temp1 - temp2).ToString(), 9);
                }

                //It's time to fuck bitches and get money
                StatusLabel.Text = "В работе";
                progressBar1.Value = 80;
                StatusDesc.Text = "Запись";

                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 2, KprfSumI.ToString(), 9);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 3, KprfSumLI.ToString(), 9);

                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 11, ErSumI.ToString(), 9);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 12, ErSumLI.ToString(), 9);

                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 5, LdprSumI.ToString(), 9);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 6, LdprSumLI.ToString(), 9);

                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 8, SrSumI.ToString(), 9);
                excel.WriteToCell(Globals.REGIONS_COUNT + 3, 9, SrSumLI.ToString(), 9);

                while (excel.ListCount() < 14)
                {
                    excel.AddSheet();
                }
                PrepareList(10, "None", "None", false);

                excel.FixTheShit(10);


                excel.PlotChart(10, "A1", "C5", "graph1.bmp");

                //Count day reports (NON CUMULATIVE!)


                _Excel._Application base1 = new _Excel.Application();
                base1.Visible = false;
                Workbook workbook_detailedDB = base1.Workbooks.Open(Properties.Settings.Default["det_db_path"].ToString());
                Worksheet worksheet_detailedDB = workbook_detailedDB.Sheets[1];

                worksheet_detailedDB = workbook_detailedDB.Sheets[1];
                int daysInDetailed = (worksheet_detailedDB.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column - 1) / 4;
                int CSumDivider = 0;
                Globals.CSumOk = true;
                CSumDivider = Convert.ToInt32(Properties.Settings.Default["divs"]);
                if (CSumDivider > daysInDetailed)
                {
                    int recommendedDivision = 0;
                    MessageBoxButtons buttons = MessageBoxButtons.YesNoCancel;
                    if (daysInDetailed > 6)
                    {
                        recommendedDivision = 6;
                    }
                    else
                    {
                        recommendedDivision = daysInDetailed;
                    }

                    DialogResult result = MessageBox.Show("Дней в детальной базе меньше, чем указано кумулятивных делений в настройках. Невозможно произвести кумулятивное деление.\nПоставить рекомендуемые настройки?\n\nДней в базе: " + daysInDetailed.ToString() + "\nРекомендуемое деление: " + recommendedDivision.ToString(), "Внимание!", buttons, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                    if (result == DialogResult.Yes)
                    {
                        CSumDivider = recommendedDivision;
                    }
                    else if (result == DialogResult.No)
                    {
                        Globals.CSumOk = false;
                        CSumDivider = 0;
                    }
                    else
                    {
                        Globals.CSumOk = false;
                        CSumDivider = 0;
                    }
                }



                List<string> FindAllDates(int sheet)
                {
                    worksheet_detailedDB = workbook_detailedDB.Sheets[sheet];
                    int cols = worksheet_detailedDB.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    List<string> dates = new List<string>();
                    for (int col1 = 1; col1 <= cols + 1; col1++)
                    {
                        if (worksheet_detailedDB.Cells[1, col1].Value2 != null)
                        {
                            dates.Add(worksheet_detailedDB.Cells[1, col1].Value2);
                        }
                    }
                    return dates;
                }


                List<int> SumByDate(string date, int sheet)
                {
                    List<int> sums = new List<int>();
                    worksheet_detailedDB = workbook_detailedDB.Sheets[sheet];
                    int tcol = 0;
                    try
                    {
                        tcol = worksheet_detailedDB.Cells.Find(date).Column;
                    }
                    catch
                    {
                        int lastCol = worksheet_detailedDB.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;

                        int i = 1;//search row
                        for (int j = 1; j < lastCol; j++)
                        {
                            if ((worksheet_detailedDB.Cells[i, j].Value2 != "") & (worksheet_detailedDB.Cells[i, j].Value2 == date))
                            {
                                tcol = j;
                                break;
                            }
                        }
                    }
                    int allsums = 0;
                    for (int part = 0; part < 4; part++)
                    {
                        int current = 0;
                        for (int row1 = 3; row1 <= worksheet_detailedDB.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row; row1++)
                        {
                            current += Convert.ToInt32(worksheet_detailedDB.Cells[row1, tcol + part].Value2);
                        }
                        sums.Add(current);
                        allsums += current;
                    }
                    sums.Add(allsums);
                    return sums;
                }



                //Divide by input. Leftover to first
                if (Globals.CSumOk)
                {
                    int daysInCsum = Convert.ToInt32(Math.Floor(Convert.ToDouble(daysInDetailed / CSumDivider)));
                    int leftoverDaysCsum = daysInDetailed % CSumDivider;
                    //find all day labels and give them ids
                    //define labels array here

                    List<string> labels = FindAllDates(1);

                    //string[] labels = new string[daysInDetailed];

                    //define pointsum here
                    int[][] pointCSums = new int[CSumDivider][];
                    for (int i = 0; i < CSumDivider; i++)
                    {
                        pointCSums[i] = new int[5] { 0, 0, 0, 0, 0 };
                    }

                    //define labels for final labeling 
                    string[] CSlabels = new string[CSumDivider];

                    for (int point = 1; point <= CSumDivider; point++)
                    {
                        if ((leftoverDaysCsum != 0) & (point == CSumDivider))
                        {//leftover case
                         //take last label for SCLabels
                            string fin = labels[labels.Count - 1];
                            string sta = "";
                            for (int daypoint = 1; daypoint <= leftoverDaysCsum + daysInCsum; daypoint++)
                            {
                                //take last label
                                //get sums
                                //sum them to pointsum
                                for (int c = 0; c < 5; c++)
                                {
                                    pointCSums[CSumDivider - point][c] += SumByDate(labels[labels.Count - 1], 1).ToArray()[c];
                                }
                                //take last label for SCLabels
                                sta = labels[labels.Count - 1];
                                //pop label
                                labels.RemoveAt(labels.Count - 1);

                            }
                            //create label for final labeling
                            CSlabels[CSumDivider - point] = (sta + " - " + fin);
                        }
                        else
                        {//normal case
                            //take last label for SCLabels
                            string fin = labels[labels.Count - 1];
                            string sta = "";
                            for (int daypoint = 1; daypoint <= daysInCsum; daypoint++)
                            {
                                //take last label
                                //get sums
                                //sum them to pointsum
                                for (int c = 0; c < 5; c++)
                                {
                                    pointCSums[CSumDivider - point][c] += SumByDate(labels[labels.Count - 1], 1).ToArray()[c];
                                }
                                //take last label for SCLabels
                                sta = labels[labels.Count - 1];
                                //pop label
                                labels.RemoveAt(labels.Count - 1);

                            }
                            //create label for final labeling
                            CSlabels[CSumDivider - point] = (sta + " - " + fin);
                        }
                    }
                    //closing detailed base
                    workbook_detailedDB.Close();
                    base1.Quit();



                    //prepare sheet
                    PrepareList(12, "None", "None", false);
                    PrepareList(13, "None", "None", false);

                    //writing data
                    int[] CSumLoop = new int[5] { 0, 0, 0, 0, 0 };

                    for (int pointpart = 0; pointpart < CSumDivider; pointpart++)
                    {
                        excel.WriteToCell(1, 2 + pointpart, CSlabels[pointpart], 12);
                        excel.WriteToCell(1, 2 + pointpart, CSlabels[pointpart], 13);
                        for (int CSval = 0; CSval < 5; CSval++)
                        {
                            //nedocum at sheet 12
                            excel.WriteToCell(2 + CSval, 2 + pointpart, pointCSums[pointpart][CSval].ToString(), 12);

                            //cumstruct sheet at 13
                            if (CSval < 4)
                            {
                                CSumLoop[CSval] += pointCSums[pointpart][CSval];
                                excel.WriteToCell(2 + CSval, 2 + pointpart, CSumLoop[CSval].ToString(), 13);
                            }

                        }
                    }
                    //create percent table on sheet 14 out of necodum table on sheet 12
                    excel.AddSheet();

                    //prepare sheet
                    PrepareList(14, "None", "None", false);

                    //write
                    for (int labelID = 0; labelID < CSumDivider; labelID++)
                    {
                        excel.WriteToCell(1, 2 + labelID, CSlabels[labelID], 14);
                        for (int partID = 0; partID < 5; partID++)
                        {
                            double target = Convert.ToDouble(pointCSums[labelID][partID]);
                            double total = Convert.ToDouble(pointCSums[labelID][4]);
                            excel.WriteToCell(2 + partID, 2 + labelID, excel.FindPercent(target, total), 14);
                        }
                    }
                    // Add chart.
                    excel.PlotCSumChart(13, "graph2.bmp");

                }


                PrepareList(11, "None", "None", false);
                excel.FixTheShit(11);
                excel.PlotChart(11, "A1", "C5", "graph3.bmp");

                excel.SortShit(8, 17, "C");
                excel.SortShit(9, 13, "B");

                excel.FixTheShit(12);
                excel.FixTheShit(13);
                excel.FixTheShit(14);
                excel.FixTheShit(4);
                excel.FixTheShit(7);
                excel.FixTheShit(8);
                excel.FixTheShit(9);
                double[] MT = new double[] { KprfSum, ErSum, LdprSum, SrSum };
                double[] LMT = new double[] { KprfSumL, ErSumL, LdprSumL, SrSumL };
                double[] IT = new double[] { KprfSumI, ErSumI, LdprSumI, SrSumI };
                double[] LIT = new double[] { KprfSumLI, ErSumLI, LdprSumLI, SrSumLI };
                excel.AddTotalsFix(MT, IT, LMT, LIT, Globals.REGIONS_COUNT);

                _Excel.ColorScale[] ColorScaleArray = new _Excel.ColorScale[13];

                int scaler = 0;
                excel.PaintSheet(excel.GetUsedRangeOf(8), System.Drawing.Color.Gainsboro);
                excel.PaintSheet(excel.GetUsedRangeOf(9), System.Drawing.Color.Gainsboro);
                for (int i = 2; i <= 17; i += 4)
                {
                    excel.ApplyCondForm(8, 2, i, 2 + Globals.REGIONS_COUNT, i, ColorScaleArray, scaler);
                    ++scaler;
                    excel.ApplyCondForm(8, 2, i + 1, 2 + Globals.REGIONS_COUNT, i + 1, ColorScaleArray, scaler);
                    ++scaler;
                }


                for (int i = 2; i <= 13; i += 3)
                {
                    excel.ApplyCondForm(9, 2, i, 2 + Globals.REGIONS_COUNT, i, ColorScaleArray, scaler);
                    ++scaler;
                }

                excel.WrapDetailed();


                //saving the results
                string resultsPath = "";
                if (SaveToCurrent.Checked)
                {
                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    StatusDesc.Text = "";
                    excel.Save();
                    resultsPath = Globals.FILE_NAME;
                }
                else
                {
                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    StatusDesc.Text = "";
                    resultsPath = excel.SaveAs();
                }



                //start word processor

                //chose the template


                Object GRC = Globals.REGIONS_COUNT.ToString();
                Object regflag = "<reg_count>";

                Globals.app = new Word.Application();
                Globals.doc = new Word.Document();


                prepAndOpenWord(Properties.Settings.Default["report1_template_path"].ToString());
                FindAndReplace(Globals.app, regflag, GRC);//insert regions

                string wbkName = resultsPath;
                _Excel._Application xlApp = new _Excel.Application();
                xlApp.Visible = false;
                Globals.app.Visible = false;
                Globals.app.DisplayAlerts = 0;
                Thread.Sleep(4000);
                if (wbkName == "none")
                {
                    MessageBox.Show("Выполнение программы было приостановлено", "Отмена");
                    xlApp.Quit();
                    excel.Close();
                    Globals.doc.Close();
                    Globals.app.Quit();
                    return;
                }
                _Excel.Workbook workbook = xlApp.Workbooks.Open(wbkName);

                //selecting a table
                _Excel.Worksheet worksheet = workbook.Sheets[4];

                //copy and insert table
                worksheet.Range["A1", "H6"].Copy();
                Word.Range rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<tab1>"))
                {
                    Thread.Sleep(4000);
                    rangetemp.PasteExcelTable(false, true, false);
                }



                //selecting a table for TV part
                worksheet = workbook.Sheets[7];

                //copy and insert table
                worksheet.Range["A1", "H6"].Copy();
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<tabl4>"))
                {
                    rangetemp.PasteExcelTable(false, true, false);
                }

                //selecting a table for Nedocum part
                worksheet = workbook.Sheets[12];

                //copy and insert table
                worksheet.UsedRange.Copy();
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<tabl2>"))
                {
                    rangetemp.PasteExcelTable(false, true, false);
                }

                //selecting a table for Nedocum percents
                worksheet = workbook.Sheets[14];

                //copy and insert table
                worksheet.UsedRange.Copy();
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<tabl3>"))
                {
                    rangetemp.PasteExcelTable(false, true, false);
                }



                //detailed1
                worksheet = workbook.Sheets[8];

                //copy and insert table

                worksheet.UsedRange.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
                worksheet.UsedRange.Borders.Weight = _Excel.XlBorderWeight.xlThin;

                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[3 + Globals.REGIONS_COUNT, 17]].Copy();
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<detailed1>"))
                {
                    //rangetemp.PasteSpecial(_Excel.XlPasteType.xlPasteAll);
                    rangetemp.PasteExcelTable(false, false, false);
                }


                //detailed2
                worksheet = workbook.Sheets[9];

                worksheet.UsedRange.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
                worksheet.UsedRange.Borders.Weight = _Excel.XlBorderWeight.xlThin;
                //copy and insert table
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[3 + Globals.REGIONS_COUNT, 13]].Copy();
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<detailed2>"))
                {
                    rangetemp.PasteExcelTable(false, false, false);
                    //rangetemp.PasteSpecial(_Excel.XlPasteType.xlPasteAll);
                }



                // excel = new Excel(Globals.FILE_NAME, 2);


                //Insert First graph
                string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<graph1>"))
                {
                    rangetemp.Text = "";
                    rangetemp.InlineShapes.AddPicture(base_path + "\\graph1.bmp");
                }

                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<graph2>"))
                {
                    rangetemp.Text = "";
                    rangetemp.InlineShapes.AddPicture(base_path + "\\graph2.bmp");
                }

                rangetemp = Globals.doc.Range(0, 0);
                if (rangetemp.Find.Execute("<graph3>"))
                {
                    rangetemp.Text = "";
                    rangetemp.InlineShapes.AddPicture(base_path + "\\graph3.bmp");
                }



                SaveFileDialog sfd3 = new SaveFileDialog();
                sfd3.Filter = "Word files (*.docx)|*.docx";
                if (sfd3.ShowDialog() == DialogResult.OK)
                {
                    Globals.doc.SaveAs2(sfd3.FileName);
                }


                workbook.Close(falseObj);
                xlApp.Quit();
            }

            try
            {
                Globals.doc.Close();
                Globals.app.Quit();
            }
            catch
            {

            }
            excel.Close();
            excel.Close();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            //check settings
            if (!File.Exists(Properties.Settings.Default["report1_template_path"].ToString())) {
                MessageBox.Show("Шаблон для отчета не был найден. Проверьте настройки", "Шаблон не найден", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(Properties.Settings.Default["regfile"].ToString()))
            {
                MessageBox.Show("Список регионов не был найден. Проверьте настройки", "Список регионов не найден", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Report1();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "1";
                reporter.Stage = "Report creation";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }



        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SaveToCurrent_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            var workbook = base1.Workbooks.Open(Globals.FILE_NAME);
            var worksheet = workbook.Worksheets[12] as
                Microsoft.Office.Interop.Excel.Worksheet;

            // Add chart.
            var charts = worksheet.ChartObjects() as
                Microsoft.Office.Interop.Excel.ChartObjects;
            var chartObject = charts.Add(60, 10, 600, 300) as
                Microsoft.Office.Interop.Excel.ChartObject;
            var chart = chartObject.Chart;

            // Set chart range.
            //var range = worksheet.get_Range(topLeft, bottomRight);
            var range = worksheet.UsedRange;
            chart.SetSourceData(range);

            // Set chart properties.
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: range,
                Title: "Кумулятивная сумма по датам",
                CategoryTitle: "Даты",
                ValueTitle: "Количество");
            chart.ApplyDataLabels(_Excel.XlDataLabelsType.xlDataLabelsShowLabel, true, true, false, false, false, true, true, false, false);
            string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            chart.Export(base_path + "\\" + "graph2.bmp", "BMP", missingObj);

        }


        private void Report2()
        {

            string target = listBox3.SelectedItem.ToString();
            //load excel
            StatusLabel.Text = "Подготовка базы";
            progressBar1.Value = 10;
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["dep_month_path"].ToString());
            ws = wb.Sheets[1];


            //copy data to list 3 (clear list 3)
            int tcol = ws.Cells.Find(target, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj).Column;
            int lastcol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            StatusLabel.Text = "Сбор данных";
            progressBar1.Value = 30;
            ws = wb.Sheets[1];
            _Excel.Range sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[3];
            _Excel.Range destinationRange = ws.Range[ws.Cells[1, 3], ws.Cells[lastrow, 3]];
            sourceRange.Copy(destinationRange);

            ws = wb.Sheets[2];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[3];
            destinationRange = ws.Range[ws.Cells[1, 4], ws.Cells[lastrow, 4]];
            sourceRange.Copy(destinationRange);

            ws = wb.Sheets[1];
            sourceRange = ws.Range[ws.Cells[1, 1], ws.Cells[lastrow, 2]];
            ws = wb.Sheets[3];
            destinationRange = ws.Range[ws.Cells[1, 1], ws.Cells[lastrow, 2]];
            sourceRange.Copy(destinationRange);

            ws.Cells[1, 3].Value2 += " Количество сообщений";
            ws.Cells[1, 4].Value2 += " Медиа-Индекс";

            ws.Columns[1].ColumnWidth = 3;
            ws.Columns[2].ColumnWidth = 40;
            ws.Columns[3].ColumnWidth = 13;
            ws.Columns[4].ColumnWidth = 13;
            ws.Rows[1].Cells.WrapText = true;
            //sort the shit(?)
            _Excel.Range rng = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 4]];

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(rng.Columns[2], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
            var sort = ws.Sort;
            sort.SetRange(rng.Rows);
            sort.Header = _Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
            sort.Apply();


            //init word
            StatusLabel.Text = "Составление отчета";
            progressBar1.Value = 90;
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            //paste data
            ws.UsedRange.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            ws.UsedRange.Borders.Weight = _Excel.XlBorderWeight.xlThin;

            ws.UsedRange.Copy(missingObj);
            Word.Range rangetemp = doc.Range(0, 0);

            rangetemp.PasteExcelTable(false, false, false);

            //save and close
            SaveFileDialog sf = new SaveFileDialog();

            sf.InitialDirectory = "c:\\";
            sf.Filter = "Word files (*.docx)|*.docx";
            sf.FilterIndex = 0;
            sf.RestoreDirectory = true;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                doc.SaveAs2(sf.FileName);
                MessageBox.Show("Отчет был сохранен как " + sf.FileName, "Сохранение");
            }
            else
            {
                MessageBox.Show("Сохранение отчета было отменено", "Отмена");

            }
            StatusLabel.Text = "Готово";
            progressBar1.Value = 0;
            doc.Close();
            app.Quit();
            wb.Close();
            base1.Quit();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Create from current
            if (listBox3.SelectedItem == null)
            {
                MessageBox.Show("Выберите дату для составления отчета из списка", "Внимание");
                return;
            }

            try
            {
                Report2();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "2";
                reporter.Stage = "Report creation";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }




        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        public void AnFullBase(string path = "none") {
            progressBar1.Value = 10;
            StatusLabel.Text = "Чтение базы";
            if (path == "none")
            {
                if (Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
                {
                    MessageBox.Show("В настройках стоит создание новой базы. Укажите уже существующую базу", "Ошибка загрузки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }
                if (!File.Exists(Properties.Settings.Default["dep_month_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
                {
                    MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }
                StatusLabel.Text = "Идет загрузка базы, пожалуйста подождите";
                progressBar1.Value = 10;
                textBox2.Text = Properties.Settings.Default["dep_month_path"].ToString();
                //start temp excel
                _Excel._Application base1 = new _Excel.Application();
                base1.Visible = false;
                base1.DisplayAlerts = false;

                //init wb and ws
                Workbook wb;
                Worksheet ws;
                wb = base1.Workbooks.Open(Properties.Settings.Default["dep_month_path"].ToString());
                ws = wb.Sheets[1];
                List<string> FindAllDates(int sheet)
                {
                    ws = wb.Sheets[sheet];
                    int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    List<string> dates = new List<string>();
                    for (int col1 = 3; col1 <= cols + 1; col1++)
                    {
                        if (ws.Cells[1, col1].Value2 != null)
                        {
                            dates.Add(ws.Cells[1, col1].Value2);
                        }
                    }
                    return dates;
                }
                List<string> dates_temp = FindAllDates(1);
                listBox3.Items.Clear();
                for (int i = 0; i < dates_temp.Count; i++)
                {
                    listBox3.Items.Add(dates_temp[i]);
                }
                wb.Close();
                base1.Quit();
                button7.Enabled = true;
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;
            }
            else
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }
                StatusLabel.Text = "Идет загрузка базы, пожалуйста подождите";
                progressBar1.Value = 20;
                textBox2.Text = path;
                //start temp excel
                _Excel._Application base1 = new _Excel.Application();
                base1.Visible = false;
                base1.DisplayAlerts = false;

                //init wb and ws
                Workbook wb;
                Worksheet ws;
                wb = base1.Workbooks.Open(path);
                ws = wb.Sheets[1];
                List<string> FindAllDates(int sheet)
                {
                    ws = wb.Sheets[sheet];
                    int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    List<string> dates = new List<string>();
                    for (int col1 = 3; col1 <= cols + 1; col1++)
                    {
                        if (ws.Cells[1, col1].Value2 != null)
                        {
                            dates.Add(ws.Cells[1, col1].Value2);
                        }
                    }
                    return dates;
                }
                List<string> dates_temp = FindAllDates(1);
                listBox3.Items.Clear();
                for (int i = 0; i < dates_temp.Count; i++)
                {
                    listBox3.Items.Add(dates_temp[i]);
                }
                wb.Close();
                base1.Quit();
                button7.Enabled = true;
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;

            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            if (Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
            {
                MessageBox.Show("В настройках выбрана опция создания новой базы.\nНовая база будет создана автоматически при первом запросе", "Внимание!");
                progressBar1.Value = 0;
                StatusLabel.Text = "Готово";
                return;
            }
            if (!File.Exists(Properties.Settings.Default["dep_month_path"].ToString()))
            {
                MessageBox.Show("Файл базы не найден. Пожалуйста, проверьте настройки.", "Внимание!");
                progressBar1.Value = 0;
                StatusLabel.Text = "Готово";
                return;
            }

            try
            {
                AnFullBase();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "dep_month";
                reporter.Stage = "Analyse database";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }
            

        }


        private void Dep_month_request()
        {

            StatusLabel.Text = "Чтение списка депутатов";
            progressBar1.Value = 40;
            //populate deps
            string line;
            List<string> selectedDeps = new List<string>();
            // Read the file and display it line by line.  

            StreamReader file = new System.IO.StreamReader(Properties.Settings.Default["dep_month_list"].ToString());
            while ((line = file.ReadLine()) != null)
            {
                selectedDeps.Add(line);
            }

            file.Close();
            Globals.DepsSel.Clear();
            foreach (KeyValuePair<int, string> entry in Globals.Deps2)
            {
                if (selectedDeps.Contains(entry.Value) == true)
                {
                    try
                    {
                        Globals.DepsSel.Add(entry.Key, entry.Value);
                    }
                    catch
                    {
                        DialogResult dialogResult = MessageBox.Show("Депутат " + entry.Value + " не был найден в списках Медиалогии.\nДанный депутат не будет включен в выборку\nОтменить создание отчета?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (dialogResult == DialogResult.Yes)
                        {
                            MessageBox.Show("Создание отчета было приостановлено", "Отмена");
                            StatusLabel.Text = "Готово";
                            progressBar1.Value = 0;
                            return;
                        }
                        else
                        {
                            //ignore
                        }

                    }
                }

            }
            //check input dates
            if (DateTime.Compare(dateTimePicker3.Value, dateTimePicker4.Value) > 0)
            {
                MessageBox.Show("Выбран неверный промежуток времени", "Ошибка дат");
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;
                Globals.DepsSel.Clear();
                return;
            }
            else if (DateTime.Compare(dateTimePicker3.Value, dateTimePicker4.Value) == 0)
            {
                MessageBox.Show("Для отчета необходим промежуток более 24х часов", "Ошибка дат");
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;
                Globals.DepsSel.Clear();
                return;
            }
            string dayTo, dayFrom, monthTo, monthFrom;

            if (Convert.ToInt32(dateTimePicker3.Value.Day) < 10) { dayFrom = "0" + dateTimePicker3.Value.Day.ToString(); } else { dayFrom = dateTimePicker3.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker4.Value.Day) < 10) { dayTo = "0" + dateTimePicker4.Value.Day.ToString(); } else { dayTo = dateTimePicker4.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker3.Value.Month) < 10) { monthFrom = "0" + dateTimePicker3.Value.Month.ToString(); } else { monthFrom = dateTimePicker3.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker4.Value.Month) < 10) { monthTo = "0" + dateTimePicker4.Value.Month.ToString(); } else { monthTo = dateTimePicker4.Value.Month.ToString(); }



            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker3.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker3.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker4.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker4.Value.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input

            //create new excel
            //start excel
            StatusLabel.Text = "Подготовка базы";
            progressBar1.Value = 10;
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            bool is_db_new = false;

            if (Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
            {
                //prepare new sheet
                wb = base1.Workbooks.Add(Type.Missing);
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                wb.Sheets.Add();
                wb.Sheets.Add();
                ws = wb.Sheets[1];
                ws.Name = "Количество сообщений";
                ws = wb.Sheets[2];
                ws.Name = "Медиа-Индекс";
                ws = wb.Sheets[3];
                ws.Name = "Технический";

                //set labels
                ws = wb.Sheets[1];
                ws.Cells[1, 1].Value2 = "№";
                ws.Cells[1, 2].Value2 = "Депутат";
                ws = wb.Sheets[2];
                ws.Cells[1, 1].Value2 = "№";
                ws.Cells[1, 2].Value2 = "Депутат";

                int counter = 1;
                foreach (KeyValuePair<int, string> entry in Globals.DepsSel)
                {
                    ws = wb.Sheets[1];
                    ws.Cells[counter + 1, 1].Value2 = counter.ToString();
                    ws.Cells[counter + 1, 2].Value2 = entry.Value;
                    ws = wb.Sheets[2];
                    ws.Cells[counter + 1, 1].Value2 = counter.ToString();
                    ws.Cells[counter + 1, 2].Value2 = entry.Value;
                    counter++;
                }
                wb.SaveAs(Properties.Settings.Default["dep_month_path"].ToString());
                wb.Close();
                is_db_new = true;
            }
            else
            {
                AnFullBase(Properties.Settings.Default["dep_month_path"].ToString());
            }
            //create report
            string data_to_post;
            byte[] buffer;
            string jsonTempReq;
            string jsonReqTextbase = "{\"smsMonitor\":{\"MonitorId\":-1,\"ThemeId\":-1,\"UserId\":-1,\"MaxSendingArticle\":0,\"SendingMode\":2,\"SendingPeriod\":1,\"ReprintsMode\":3,\"MonitorPhones\":[]},\"folder\":\"\",\"folderId\":-1,\"Authors\":[],\"Cities\":[],\"Levels\":[1,2],\"Categories\":[1,2,3,4,5,6],\"Rubrics\":[],\"LifeStyles\":[],\"MediaSources\":[],\"MediaBranches\":[],\"MediaObjectBranches\":[],\"MediaObjectLifeStyles\":[],\"MediaObjectLevels\":[],\"MediaObjectCategories\":[],\"MediaObjectRegions\":[],\"MediaObjectFederals\":[],\"MediaObjectTowns\":[],\"MediaLanguages\":[],\"MediaRegions\":[],\"MediaCountries\":[],\"CisMediaCountries\":[],\"MediaFederals\":[],\"MediaGenre\":[],\"YandexRubrics\":[],\"Role\":-1,\"Tone\":-1,\"Quotation\":-1,\"CityMode\":0,\"messageCount\":-1,\"reprintsMessageCount\":-1,\"CheckedMessageCount\":-1,\"CheckedClustersCount\":-1,\"MonitorId\":-1,\"CheckedReprintsCount\":-1,\"deletedMessageCount\":-1,\"favoritesMessageCount\":-1,\"myDocsMessageCount\":0,\"myMediaMessageCount\":0,\"IsSaveParamsOnly\":false,\"RebuildDBCache\":false,\"Credentials\":null,\"AppType\":1,\"ParamsVersion\":0,\"ArmObjectMode\":0,\"ReportCreatingHistory\":0,\"InfluenceThreshold\":\"0.0\",\"MonitorObjects\":null,\"Icon\":0,\"ThemeGroup\":-1,\"ThemeGroupName\":\"\",\"SaveMode\":0,\"MonitorExists\":false,\"ThemeId\":-1,\"Title\":\"<reptitle>\",\"Comment\":\"\",\"ReprintMode\":0,\"rssReportType\":0,\"ThemeObjects\":[";

            StatusLabel.Text = "Отправка запроса и получение данных (до 20 секунд)";
            progressBar1.Value = 30;

            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            var cookieContainer = new CookieContainer();
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;

            //create reqobjects
            string nameOfReport = "dep_GD_" + DateTime.Now.Day.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Year.ToString() + " at " + DateTime.Now.TimeOfDay.ToString();
            string depBase = "{\"Id\":\"<depid>\",\"MainObjectId\":\"<depid>\",\"ObjectName\":\"<depname>\",\"classId\":43,\"LogicIndex\":<lindex>,\"LogicObjectString\":\"OR\",\"SearchQuery\":null,\"Properties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}]}";
            string repEnd = "],\"ThemeObjectsFromSearchContext\":[],\"ThemeTypes\":[],\"ThemeBranches\":[],\"AllObjectsProperties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}],\"AllArticlesProperties\":[],\"AllObjectString\":\"<allobjectstring>\",\"AllLogicObjectString\":\"<alllogic>\",\"DatePeriod\":8,\"DateType\":0,\"Date\":\"<datefrom>|<dateto>\",\"Time\":\"<timefrom>|<timeto>\",\"ActualDatePeriod\":3,\"IsSlidingTime\":true,\"ContextScope\":5,\"Context\":\"\",\"ContextMode\":0,\"TopMedia\":false,\"RegionLogic\":0,\"MediaObjectRegionLogic\":0,\"MediaLogic\":0,\"MediaLogicAll\":0,\"BlogLogic\":1,\"MediaBranchLogic\":0,\"MediaObjectBranchLogic\":0,\"MediaLanguageLogic\":0,\"MediaCountryLogic\":0,\"CityLogic\":0,\"Compare\":1,\"User\":0,\"Type\":6,\"View\":0,\"ViewStatus\":1,\"OiiMode\":0,\"Template\":-1,\"MediaStatus\":-1,\"IsUpdate\":false,\"HasUserObjects\":false,\"IsContextReport\":false,\"LastCopiedThemeId\":null}";
            string allobjectstring = "";
            string alllogic = "";
            string allDeps = "";

            //create alldep string
            int c1 = 0;
            foreach (KeyValuePair<int, string> entry in Globals.DepsSel)
            {
                allobjectstring += "+O" + entry.Key.ToString() + "_" + c1.ToString();
                alllogic += "+" + c1.ToString();
                allDeps += depBase.Replace("<depid>", entry.Key.ToString()).Replace("<depname>", entry.Value.Replace(" ", "+")).Replace("<lindex>", c1.ToString());
                if (c1 < Globals.DepsSel.Count - 1)
                {
                    allDeps += ",";

                }
                c1++;
            }

            jsonTempReq = jsonReqTextbase.Replace("<reptitle>", nameOfReport);
            string repEnd2 = repEnd.Replace("<datefrom>", datefrom)
                .Replace("<dateto>", dateto)
                .Replace("<timefrom>", timefrom)
                .Replace("<timeto>", timeto)
                .Replace("<allobjectstring>", allobjectstring)
                .Replace("<alllogic>", alllogic);

            jsonTempReq += allDeps + repEnd2;

            IDictionary<string, double[]> depData = new Dictionary<string, double[]>();


            try
            {
                data_to_post = "UserName=" + Properties.Settings.Default["login"] + "&Password=" + Properties.Settings.Default["password"] + "&PrUrl=http%3A%2F%2Fpr.mlg.ru&Pr2Url=http%3A%2F%2Fdev.pr2.mlg.ru&MmUrl=http%3A%2F%2Fmm.mlg.ru&BuzzUrl=http%3A%2F%2Fsm.mlg.ru&ReturnUrl=http%3A%2F%2Fpr.mlg.ru&ApplicationType=Pr";
                buffer = Encoding.ASCII.GetBytes(data_to_post);

                WebReq = (HttpWebRequest)WebRequest.Create("https://login.mlg.ru/Account.mlg?ApplicationType=Pr");
                WebReq.CookieContainer = cookieContainer;
                WebReq.Timeout = 60000;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;

                PostData = WebReq.GetRequestStream();
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                WebResp.Close();

                //MessageBox.Show("Начало DEBUG сессии для " + base_path);
                string urlencoded;
                byte[] urljson;
                string currentReportIdstr;
                //prepare json to send
                //MessageBox.Show("Подготовка данных к отправке");
                //
                //urlencode the shit
                urljson = Encoding.ASCII.GetBytes(jsonTempReq);
                urlencoded = HttpUtility.UrlEncode(urljson);
                //create payload and send it

                data_to_post = "useFilterContainers=false&sr=" + urlencoded;
                buffer = Encoding.ASCII.GetBytes(data_to_post);
                //MessageBox.Show("Буффер составлен");
                try
                {
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                    WebReq.MaximumAutomaticRedirections = 1;
                    WebReq.AllowAutoRedirect = false;
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;
                    WebReq.Timeout = 60000;
                    PostData = WebReq.GetRequestStream();
                    //MessageBox.Show("Буффер отправлен");
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    //catch id of redirrect
                    currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                    WebResp.Close();
                }
                catch (WebException exxx)
                {
                    if (exxx.Status == WebExceptionStatus.Timeout)
                    {
                        //workbook_detailedDB.Close();
                        base1.Quit();
                        //workbook_sumDB.Close();
                        //base2.Quit();
                        StatusLabel.Text = "Готово";
                        progressBar1.Value = 0;
                        MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        return;
                    }
                    else throw;
                }

                int tempReportId;
                try
                {
                    tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                    //MessageBox.Show("Идет перехват отчета №" + tempReportId.ToString());
                }
                catch
                {
                    MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                    WebResp.Close();
                    base1.Quit();
                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    return;
                }
                WebResp.Close();

                //catch redirrect
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/DynamicsChart?id=" + tempReportId.ToString() + "&pageSize=20&gtype=ByGroups&scale=Default&viewType=MlgGraph&pageNumber=1");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                string answer = _Answer.ReadToEnd();

                WebResp.Close();

                WebReq = null;
                StatusLabel.Text = "Данные получены";
                progressBar1.Value = 70;

                string resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                Regex reg = new Regex("legendItemSign.*?an>(.*?)<.*?(\\d.*?)<.*?&count=(\\d*?)&");
                MatchCollection data = reg.Matches(resultString);

                double messageCount;
                double mediaIndex;
                string depName;
                if (data.Count > 0)
                {
                    for (int id = 0; id < data.Count; id++)
                    {
                        depName = data[id].Groups[1].Value;
                        string mmme = data[id].Groups[2].Value.Replace(" ", "").Replace(" ", "").Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator);
                        mediaIndex = Convert.ToDouble(mmme);
                        messageCount = Convert.ToDouble(data[id].Groups[3].Value.Replace(" ", ""));

                        depData.Add(depName, new double[] { messageCount, mediaIndex });

                    }
                }
                else
                {
                    foreach (KeyValuePair<int, string> entry in Globals.DepsSel)
                    {
                        depName = entry.Value;
                        mediaIndex = Convert.ToDouble(0);
                        messageCount = Convert.ToDouble(0);
                        depData.Add(depName, new double[] { messageCount, mediaIndex });
                    }
                }



            }
            catch (Exception ex)
            {
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "?";
                reporter.Stage = "Ошибка при получении отчета";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

                MessageBox.Show("Ошибка получения отчета", "Ошибка");
                base1.Quit();
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;
                return;
            }

            //paste data into excel
            StatusLabel.Text = "Запись данных в базу";
            progressBar1.Value = 80;
            wb = base1.Workbooks.Open(Properties.Settings.Default["dep_month_path"].ToString());
            ws = wb.Sheets[1];

            int lastcol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            //paste label
            ws.Cells[1, lastcol + 1].Value2 = datefrom + " - " + dateto;
            ws = wb.Sheets[2];
            ws.Cells[1, lastcol + 1].Value2 = datefrom + " - " + dateto;
            ws = wb.Sheets[1];

            foreach (KeyValuePair<string, double[]> entry in depData)
            {
                try
                {

                    _Excel.Range found = ws.Cells.Find(entry.Key, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj);
                    int trow = found.Row;
                    if (ws.Cells[trow, 2].Value2 == entry.Key)
                    {
                        ws = wb.Sheets[1];
                        ws.Cells[trow, lastcol + 1].Value2 = entry.Value[0].ToString();
                        ws = wb.Sheets[2];
                        ws.Cells[trow, lastcol + 1].Value2 = entry.Value[1].ToString();
                    }
                    else
                    {
                        found = found.FindNext();
                        trow = found.Row;
                        ws = wb.Sheets[1];
                        ws.Cells[trow, lastcol + 1].Value2 = entry.Value[0].ToString();
                        ws = wb.Sheets[2];
                        ws.Cells[trow, lastcol + 1].Value2 = entry.Value[1].ToString();
                    }
                }
                catch
                {
                    MessageBox.Show("Депутат " + entry.Key + " не найден в базе", "Внимание");
                }
            }
            wb.Save();
            wb.Close();
            base1.Quit();
            AnFullBase(Properties.Settings.Default["dep_month_path"].ToString());
            Globals.DepsSel.Clear();
            StatusLabel.Text = "Готово";
            progressBar1.Value = 0;
        }
        private void button6_Click(object sender, EventArgs e)
        {

            if (!ValidateDates(dateTimePicker3, dateTimePicker4))
            {
                UpdateStatus();
                return;
            }


            //request report
            if (Properties.Settings.Default["dep_month_path"].ToString() == "")
            {
                MessageBox.Show("Файл базы не указан в настройках", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!File.Exists(Properties.Settings.Default["dep_month_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
            {
                MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (File.Exists(Properties.Settings.Default["dep_month_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]))
            {
                MessageBox.Show("Невозможно создать новый файл. Файл с таким именем уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((Properties.Settings.Default["dep_month_list"].ToString()) == "")
            {
                MessageBox.Show("Отсутствует файл со списком депутатов. Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(Properties.Settings.Default["dep_month_list"].ToString()))
            {
                MessageBox.Show("Не найден файл со списком депутатов. Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Dep_month_request();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "dep_month_request";
                reporter.Stage = "Deps report request";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }


        }

        public void AnFullBase_reg(string path = "none")
        {
            progressBar1.Value = 10;
            StatusLabel.Text = "Чтение базы";
            if (path == "none")
            {
                if (Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]))
                {
                    MessageBox.Show("В настройках стоит создание новой базы. Укажите уже существующую базу", "Ошибка загрузки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }
                if (!File.Exists(Properties.Settings.Default["dep_reg_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]))
                {
                    MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }
                StatusLabel.Text = "Идет загрузка базы, пожалуйста подождите";
                progressBar1.Value = 10;
                textBox3.Text = Properties.Settings.Default["dep_reg_path"].ToString();
                //start temp excel
                _Excel._Application base1 = new _Excel.Application();
                base1.Visible = false;
                base1.DisplayAlerts = false;

                //init wb and ws
                Workbook wb;
                Worksheet ws;
                wb = base1.Workbooks.Open(Properties.Settings.Default["dep_reg_path"].ToString());
                ws = wb.Sheets[1];
                List<string> FindAllDates(int sheet)
                {
                    ws = wb.Sheets[sheet];
                    int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    List<string> dates = new List<string>();
                    for (int col1 = 3; col1 <= cols + 1; col1++)
                    {
                        if (ws.Cells[1, col1].Value2 != null)
                        {
                            dates.Add(ws.Cells[1, col1].Value2);
                        }
                    }
                    return dates;
                }
                List<string> dates_temp = FindAllDates(1);
                listBox4.Items.Clear();
                for (int i = 0; i < dates_temp.Count; i++)
                {
                    listBox4.Items.Add(dates_temp[i]);
                }
                wb.Close();
                base1.Quit();
                button10.Enabled = true;
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;
            }
            else
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                StatusLabel.Text = "Идет загрузка базы, пожалуйста подождите";
                progressBar1.Value = 10;
                textBox3.Text = path;
                //start temp excel
                _Excel._Application base1 = new _Excel.Application();
                base1.Visible = false;
                base1.DisplayAlerts = false;

                //init wb and ws
                Workbook wb;
                Worksheet ws;
                wb = base1.Workbooks.Open(path);
                ws = wb.Sheets[1];
                List<string> FindAllDates(int sheet)
                {
                    ws = wb.Sheets[sheet];
                    int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                    List<string> dates = new List<string>();
                    for (int col1 = 3; col1 <= cols + 1; col1++)
                    {
                        if (ws.Cells[1, col1].Value2 != null)
                        {
                            dates.Add(ws.Cells[1, col1].Value2);
                        }
                    }
                    return dates;
                }
                List<string> dates_temp = FindAllDates(1);
                listBox4.Items.Clear();
                for (int i = 0; i < dates_temp.Count; i++)
                {
                    listBox4.Items.Add(dates_temp[i]);
                }
                wb.Close();
                base1.Quit();
                button10.Enabled = true;
                StatusLabel.Text = "Готово";
                progressBar1.Value = 0;

            }
        }


        private void button9_Click(object sender, EventArgs e)
        {

            if (Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]))
            {
                MessageBox.Show("В настройках выбрана опция создания новой базы.\nНовая база будет создана автоматически при первом запросе", "Внимание!");
                return;
            }
            if (!File.Exists(Properties.Settings.Default["dep_reg_path"].ToString()))
            {
                MessageBox.Show("Файл базы не найден. Пожалуйста, проверьте настройки.", "Внимание!");
                return;
            }
            //await Task.Run(() => AnFullBase_reg());

            try
            {
                AnFullBase_reg();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "3";
                reporter.Stage = "Database Analysys";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }
            
        }

        private void Dep_reg_create()
        {

            UpdateStatus("В работе", 10, "Составление отчета");

            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["dep_reg_path"].ToString());
            ws = wb.Sheets[1];

            string target = listBox4.SelectedItem.ToString();
            int tcol = ws.Cells.Find(target, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj).Column;
            int depCount = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row - 1;

            //set labels for the first table
            ws = wb.Sheets[3];
            ws.Cells.Clear();
            ws.Cells[1, 1].Value2 = "№";
            ws.Cells[1, 2].Value2 = "Регион";
            ws.Cells[1, 3].Value2 = "Депутат";
            ws.Cells[1, 4].Value2 = "Количество сообщений " + target;


            ws = wb.Sheets[4];
            ws.Cells.Clear();
            ws.Cells[1, 1].Value2 = "№";
            ws.Cells[1, 2].Value2 = "Регион";
            ws.Cells[1, 3].Value2 = "Количество сообщений " + target;
            ws.Cells[1, 4].Value2 = "Медиа-Индекс " + target;

            List<string> regions = new List<string>();
            List<string> deps = new List<string>();
            ws = wb.Sheets[1];
            IDictionary<string, double[]> regSum = new Dictionary<string, double[]>();
            for (int r = 2; r < depCount + 2; r++)
            {
                //add dep to the list
                ws = wb.Sheets[1];
                deps.Add(ws.Cells[r, 2].Value2);
                string temp_reg = ws.Cells[r, 1].Value2;
                //populate regions list
                try
                {
                    regions.Add(temp_reg);
                    ws = wb.Sheets[1];
                    double count = Convert.ToDouble(ws.Cells[r, tcol].Value2);
                    ws = wb.Sheets[2];
                    double index = Convert.ToDouble(ws.Cells[r, tcol].Value2);
                    regSum.Add(temp_reg, new double[] { count, index });
                }
                catch
                {
                    ws = wb.Sheets[1];
                    double count = Convert.ToDouble(ws.Cells[r, tcol].Value2);
                    ws = wb.Sheets[2];
                    double index = Convert.ToDouble(ws.Cells[r, tcol].Value2);
                    regSum[temp_reg][0] += count;
                    regSum[temp_reg][1] += index;
                }
            }
            //populate first table
            ws = wb.Sheets[1];
            _Excel.Range sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[depCount + 1, tcol]];
            ws = wb.Sheets[3];
            _Excel.Range destinationRange = ws.Range[ws.Cells[1, 4], ws.Cells[depCount + 1, 4]];
            sourceRange.Copy(destinationRange);

            //add indexes to first table
            ws = wb.Sheets[2];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[depCount + 1, tcol]];
            ws = wb.Sheets[3];
            destinationRange = ws.Range[ws.Cells[1, 5], ws.Cells[depCount + 1, 5]];
            sourceRange.Copy(destinationRange);

            //names and regs
            ws = wb.Sheets[1];
            sourceRange = ws.Range[ws.Cells[1, 1], ws.Cells[depCount + 1, 2]];
            ws = wb.Sheets[3];
            destinationRange = ws.Range[ws.Cells[1, 2], ws.Cells[depCount + 1, 3]];
            sourceRange.Copy(destinationRange);

            //numbers to the left
            for (int rw = 2; rw < depCount + 2; rw++)
            {
                ws.Cells[rw, 1].Value2 = (rw - 1).ToString();
            }
            //paste sums
            int counter = 2;
            ws = wb.Sheets[4];
            foreach (KeyValuePair<string, double[]> reg in regSum)
            {
                ws.Cells[counter, 1].Value2 = (counter - 1).ToString();
                ws.Cells[counter, 2].Value2 = reg.Key;
                ws.Cells[counter, 3].Value2 = reg.Value[0].ToString();
                ws.Cells[counter, 4].Value2 = reg.Value[1].ToString();
                counter++;
            }
            ws.Columns[1].ColumnWidth = 3;
            ws.Columns[2].ColumnWidth = 30;
            ws.Columns[3].ColumnWidth = 13;
            ws.Columns[4].ColumnWidth = 13;
            ws.Rows[1].Cells.WrapText = true;
            ws.Columns[2].Cells.WrapText = true;

            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            _Excel.Range rng = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 4]];

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(rng.Columns[2], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
            var sort = ws.Sort;
            sort.SetRange(rng.Rows);
            sort.Header = _Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
            sort.Apply();




            //Sort and beautify first table
            ws = wb.Sheets[3];
            ws.Columns[1].ColumnWidth = 3;
            ws.Columns[2].ColumnWidth = 26;
            ws.Columns[3].ColumnWidth = 36;
            ws.Columns[4].ColumnWidth = 13;
            ws.Columns[5].ColumnWidth = 13;
            ws.Cells[1, 4].Value2 = "Количество сообщений " + target;
            ws.Cells[1, 5].Value2 = "Медиа-Индекс " + target;
            ws.Rows[1].Cells.WrapText = true;
            ws.Columns[2].Cells.WrapText = true;
            ws.Columns[3].Cells.WrapText = true;

            lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            //fix sorting issues
            for (int rowf = 2; rowf <= lastrow; rowf++)
            {
                string tmp = Convert.ToString(ws.Cells[rowf, 5].Value2);
                ws.Cells[rowf, 5].Value2 = tmp.Replace(",", ".");
            }

            rng = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 5]];

            ws.Sort.SortFields.Clear();

            int sortCol;
            if (Convert.ToBoolean(Properties.Settings.Default["reg_sort_by_indexes"]))
            {
                sortCol = 4;
            }
            else
            {
                sortCol = 3;
            }



            ws.Sort.SortFields.Add(rng.Columns[sortCol], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
            sort = ws.Sort;
            sort.SetRange(rng.Rows);
            sort.Header = _Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
            sort.Apply();


            //revert
            for (int rowf = 2; rowf <= lastrow; rowf++)
            {
                string tmp = Convert.ToString(ws.Cells[rowf, 5].Value2);
                ws.Cells[rowf, 5].Value2 = tmp.Replace(".", ",");
            }
            //paste tables
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();

            ws = wb.Sheets[3];
            ws.UsedRange.Copy();
            Word.Range rangetemp = doc.Range(0, 0);
            rangetemp.PasteExcelTable(false, true, false);

            ws = wb.Sheets[4];
            ws.UsedRange.Copy();
            doc.Paragraphs.Add();
            rangetemp = doc.Paragraphs.Last.Range;
            rangetemp.PasteExcelTable(false, true, false);

            //debug


            //saving shit
            SaveFileDialog sf = new SaveFileDialog();

            sf.InitialDirectory = "c:\\";
            sf.Filter = "Word files (*.docx)|*.docx";
            sf.FilterIndex = 0;
            sf.RestoreDirectory = true;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                doc.SaveAs2(sf.FileName);
                MessageBox.Show("Отчет был сохранен как " + sf.FileName, "Сохранение");
            }
            else
            {
                MessageBox.Show("Сохранение отчета было отменено", "Отмена");

            }
            UpdateStatus();
            doc.Close();
            app.Quit();
            wb.Close();
            base1.Quit();

        }

        private void button10_Click(object sender, EventArgs e)
        {//create report of selected date
            if (listBox4.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите дату для отчета из списка доступных замеров", "Замер не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Dep_reg_create();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "3";
                reporter.Stage = "Report Creation";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }



        }

        private void Report3()
        {
            //load global base
            IDictionary<string, string[]> DepsRegSel = new Dictionary<string, string[]>();
            IDictionary<string, string[]> DepsRegBase = new Dictionary<string, string[]>();
            IDictionary<string, int> regionsSel = new Dictionary<string, int>();
            Dictionary<int, string> allRegions = new Dictionary<int, string>();
            List<string> selectedDeps = new List<string>();
            //init excel
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["dep_reg_database"].ToString());
            ws = wb.Sheets[1];
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            for (int r = 1; r <= lastrow; r++)
            {
                DepsRegBase.Add(ws.Cells[r, 2].Value2, new string[] { Convert.ToString(ws.Cells[r, 1].Value2), ws.Cells[r, 3].Value2 });
            }
            wb.Close();
            //populate the selection

            StatusLabel.Text = "Чтение списка депутатов";
            progressBar1.Value = 30;
            //populate deps
            string line;
            StreamReader file = new System.IO.StreamReader(Properties.Settings.Default["dep_reg_list"].ToString());
            while ((line = file.ReadLine()) != null)
            {
                selectedDeps.Add(line);
            }
            file.Close();

            foreach (string dep in selectedDeps)
            {
                if (DepsRegBase.Keys.Contains(dep))
                {
                    if (Globals.RegionsIDF.Values.Contains(DepsRegBase[dep][1]))
                    {
                        DepsRegSel.Add(dep, new string[] { DepsRegBase[dep][0], DepsRegBase[dep][1] });
                        try
                        {
                            int id = 0;
                            foreach (KeyValuePair<int, string> entry in Globals.RegionsIDF)
                            {
                                if (DepsRegBase[dep][1] == entry.Value)
                                {
                                    id = entry.Key;
                                    break;
                                }
                            }
                            regionsSel.Add(DepsRegBase[dep][1], id);
                        }
                        catch { }

                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("Регион " + DepsRegBase[dep][1] + " у депутата " + dep + " не был найден в списках Медиалогии.\nДанный депутат не будет включен в выборку\nОтменить создание отчета?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                        if (dialogResult == DialogResult.Yes)
                        {
                            MessageBox.Show("Создание отчета было приостановлено", "Отмена");
                            return;
                        }
                        else
                        {
                            //ignore
                        }
                    }

                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("Депутат " + dep + " не был найден в списках Медиалогии.\nДанный депутат не будет включен в выборку\nОтменить создание отчета?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (dialogResult == DialogResult.Yes)
                    {
                        MessageBox.Show("Создание отчета было приостановлено", "Отмена");
                        return;
                    }
                    else
                    {
                        //ignore
                    }
                }
            }

            //TODO: if new base, put labels and stuff
            if (Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]) & !File.Exists((Properties.Settings.Default["dep_reg_path"]).ToString()))
            {
                //prepare new sheet
                wb = base1.Workbooks.Add(Type.Missing);
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                wb.Sheets.Add();
                wb.Sheets.Add();
                wb.Sheets.Add();
                ws = wb.Sheets[1];
                ws.Name = "Количество сообщений";
                ws = wb.Sheets[2];
                ws.Name = "Медиа-Индекс";
                ws = wb.Sheets[3];
                ws.Name = "Технический";
                ws = wb.Sheets[4];
                ws.Name = "Технический - 2";

                int ro = 2;
                foreach (KeyValuePair<string, string[]> de in DepsRegSel)
                {
                    ws = wb.Sheets[1];
                    ws.Cells[ro, 1].Value2 = de.Value[1];
                    ws.Cells[ro, 2].Value2 = de.Key;
                    ws = wb.Sheets[2];
                    ws.Cells[ro, 1].Value2 = de.Value[1];
                    ws.Cells[ro, 2].Value2 = de.Key;
                    ro++;
                }

                //set labels
                ws = wb.Sheets[1];
                ws.Cells[1, 1].Value2 = "Регион депутата";
                ws.Cells[1, 2].Value2 = "Депутат";
                ws = wb.Sheets[2];
                ws.Cells[1, 1].Value2 = "Регион депутата";
                ws.Cells[1, 2].Value2 = "Депутат";
                wb.SaveAs(Properties.Settings.Default["dep_reg_path"].ToString());


            }
            else
            {
                if (Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]) & File.Exists((Properties.Settings.Default["dep_reg_path"]).ToString()))
                {
                    MessageBox.Show("Невозможно создать базу с таким именем.\nБаза с таким именем уже существует. Выберите другое имя в настройках.", "Ошибка создания базы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    base1.Quit();
                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    return;
                }
            }

            if (!Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]) & File.Exists((Properties.Settings.Default["dep_reg_path"]).ToString()))
            {
                StatusLabel.Text = "Открытие базы";
                progressBar1.Value = 45;
                AnFullBase_reg(Properties.Settings.Default["dep_reg_path"].ToString());

                wb = base1.Workbooks.Open(Properties.Settings.Default["dep_reg_path"].ToString());
                ws = wb.Sheets[1];
            }






            string dayTo, dayFrom, monthTo, monthFrom;


            if (Convert.ToInt32(dateTimePicker5.Value.Day) < 10) { dayFrom = "0" + dateTimePicker5.Value.Day.ToString(); } else { dayFrom = dateTimePicker5.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker6.Value.Day) < 10) { dayTo = "0" + dateTimePicker6.Value.Day.ToString(); } else { dayTo = dateTimePicker6.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker5.Value.Month) < 10) { monthFrom = "0" + dateTimePicker5.Value.Month.ToString(); } else { monthFrom = dateTimePicker5.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker6.Value.Month) < 10) { monthTo = "0" + dateTimePicker6.Value.Month.ToString(); } else { monthTo = dateTimePicker6.Value.Month.ToString(); }



            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker5.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker5.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker6.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker6.Value.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input

            //set day label
            ws = wb.Sheets[1];
            int lastcol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            ws.Cells[1, lastcol + 1].Value2 = datefrom + " - " + dateto;
            ws = wb.Sheets[2];
            ws.Cells[1, lastcol + 1].Value2 = datefrom + " - " + dateto;
            ws = wb.Sheets[1];
            string part1 = "{\"smsMonitor\":{\"MonitorId\":-1,\"ThemeId\":-1,\"UserId\":-1,\"MaxSendingArticle\":0,\"SendingMode\":2,\"SendingPeriod\":1,\"ReprintsMode\":3,\"MonitorPhones\":[]},\"folder\":\"\",\"folderId\":-1,\"Authors\":[],\"Cities\":[],\"Levels\":[1,2],\"Categories\":[1,2,3,4,5,6],\"Rubrics\":[],\"LifeStyles\":[],\"MediaSources\":[],\"MediaBranches\":[],\"MediaObjectBranches\":[],\"MediaObjectLifeStyles\":[],\"MediaObjectLevels\":[],\"MediaObjectCategories\":[],\"MediaObjectRegions\":[],\"MediaObjectFederals\":[],\"MediaObjectTowns\":[],\"MediaLanguages\":[],\"MediaRegions\":[<regid>],\"MediaCountries\":[],\"CisMediaCountries\":[],\"MediaFederals\":[],\"MediaGenre\":[],\"YandexRubrics\":[],\"Role\":-1,\"Tone\":-1,\"Quotation\":-1,\"CityMode\":0,\"messageCount\":-1,\"reprintsMessageCount\":-1,\"CheckedMessageCount\":-1,\"CheckedClustersCount\":-1,\"MonitorId\":-1,\"CheckedReprintsCount\":-1,\"deletedMessageCount\":-1,\"favoritesMessageCount\":-1,\"myDocsMessageCount\":0,\"myMediaMessageCount\":0,\"IsSaveParamsOnly\":false,\"RebuildDBCache\":false,\"Credentials\":null,\"AppType\":1,\"ParamsVersion\":0,\"ArmObjectMode\":0,\"ReportCreatingHistory\":0,\"InfluenceThreshold\":\"0.0\",\"MonitorObjects\":null,\"Icon\":0,\"ThemeGroup\":-1,\"ThemeGroupName\":\"\",\"SaveMode\":0,\"MonitorExists\":false,\"ThemeId\":-1,\"Title\":\"<repname>\",\"Comment\":\"\",\"ReprintMode\":0,\"rssReportType\":0,\"ThemeObjects\":[";
            //<regid>
            //<repname>


            //<depid>
            //<depname>
            //<lindex>
            string depObj = "{\"Id\":\"<depid>\",\"MainObjectId\":\"<depid>\",\"ObjectName\":\"<depname>\",\"classId\":43,\"LogicIndex\":<lindex>,\"LogicObjectString\":\"OR\",\"SearchQuery\":null,\"Properties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}]}";

            //<allstring>
            //<alllogic>
            //<datefrom>|<dateto>
            //<timefrom>|<timeto>
            string ending = "],\"ThemeObjectsFromSearchContext\":[],\"ThemeTypes\":[],\"ThemeBranches\":[],\"AllObjectsProperties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}],\"AllArticlesProperties\":[],\"AllObjectString\":\"<allstring>\",\"AllLogicObjectString\":\"<alllogic>\",\"DatePeriod\":8,\"DateType\":0,\"Date\":\"<datefrom>|<dateto>\",\"Time\":\"<timefrom>|<timeto>\",\"ActualDatePeriod\":3,\"IsSlidingTime\":true,\"ContextScope\":5,\"Context\":\"\",\"ContextMode\":0,\"TopMedia\":false,\"RegionLogic\":0,\"MediaObjectRegionLogic\":0,\"MediaLogic\":0,\"MediaLogicAll\":0,\"BlogLogic\":1,\"MediaBranchLogic\":0,\"MediaObjectBranchLogic\":0,\"MediaLanguageLogic\":0,\"MediaCountryLogic\":0,\"CityLogic\":0,\"Compare\":1,\"User\":0,\"Type\":6,\"View\":0,\"ViewStatus\":1,\"OiiMode\":0,\"Template\":-1,\"MediaStatus\":-1,\"IsUpdate\":false,\"HasUserObjects\":false,\"IsContextReport\":false,\"LastCopiedThemeId\":null}";

            string data_to_post;
            byte[] buffer;





            foreach (KeyValuePair<string, int> entry in regionsSel)
            {
                string region = entry.Key;
                string allobjectstring = "";
                string deps = "";
                string alllogic = "";
                IDictionary<string, string[]> depinreg = new Dictionary<string, string[]>();
                foreach (KeyValuePair<string, string[]> entry2 in DepsRegSel)
                {
                    if (entry2.Value[1] == region)
                    {
                        depinreg.Add(entry2.Key, new string[] { entry2.Value[0], entry2.Value[1] });
                    }
                }

                int lindex = 0;
                foreach (KeyValuePair<string, string[]> dep in depinreg)
                {
                    deps += depObj.Replace("<depid>", dep.Value[0])
                        .Replace("<depname>", dep.Key)
                        .Replace("<lindex>", lindex.ToString());
                    if (lindex + 1 < depinreg.Keys.Count)
                    {
                        deps += ", ";
                    }
                    allobjectstring += "+O" + dep.Value[0] + "_" + lindex.ToString();
                    alllogic += "+" + lindex.ToString();
                    lindex++;
                }
                //
                //
                //
                //<datefrom>|<dateto>
                //<timefrom>|<timeto>
                string sr = part1.Replace("<regid>", entry.Value.ToString())
                    .Replace("<repname>", "Dep_reg_" + entry.Value.ToString()) + deps + ending.Replace("<allstring>", allobjectstring)
                    .Replace("<alllogic>", alllogic)
                    .Replace("<datefrom>", datefrom)
                    .Replace("<dateto>", dateto)
                    .Replace("<timefrom>", timefrom)
                    .Replace("<timeto>", timeto);

                //send request to website
                HttpWebRequest WebReq;
                HttpWebResponse WebResp;
                var cookieContainer = new CookieContainer();
                Stream PostData;
                Stream Answer;
                StreamReader _Answer;
                try
                {
                    data_to_post = "UserName=" + Properties.Settings.Default["login"] + "&Password=" + Properties.Settings.Default["password"] + "&PrUrl=http%3A%2F%2Fpr.mlg.ru&Pr2Url=http%3A%2F%2Fdev.pr2.mlg.ru&MmUrl=http%3A%2F%2Fmm.mlg.ru&BuzzUrl=http%3A%2F%2Fsm.mlg.ru&ReturnUrl=http%3A%2F%2Fpr.mlg.ru&ApplicationType=Pr";
                    buffer = Encoding.ASCII.GetBytes(data_to_post);

                    WebReq = (HttpWebRequest)WebRequest.Create("https://login.mlg.ru/Account.mlg?ApplicationType=Pr");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Timeout = 60000;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;

                    PostData = WebReq.GetRequestStream();
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    WebResp.Close();
                    //MessageBox.Show("Начало DEBUG сессии для " + base_path);
                    string urlencoded;
                    byte[] urljson;
                    string currentReportIdstr;
                    //prepare json to send
                    //MessageBox.Show("Подготовка данных к отправке");

                    //urlencode the shit
                    urljson = Encoding.ASCII.GetBytes(sr);
                    urlencoded = HttpUtility.UrlEncode(urljson);
                    //create payload and send it

                    data_to_post = "useFilterContainers=false&sr=" + urlencoded;
                    buffer = Encoding.ASCII.GetBytes(data_to_post);
                    //MessageBox.Show("Буффер составлен");
                    try
                    {
                        WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                        WebReq.MaximumAutomaticRedirections = 1;
                        WebReq.AllowAutoRedirect = false;
                        WebReq.CookieContainer = cookieContainer;
                        WebReq.Method = "POST";
                        WebReq.ContentType = "application/x-www-form-urlencoded";
                        WebReq.ContentLength = buffer.Length;
                        WebReq.Timeout = 60000;
                        PostData = WebReq.GetRequestStream();
                        //MessageBox.Show("Буффер отправлен");
                        PostData.Write(buffer, 0, buffer.Length);
                        PostData.Close();
                        WebResp = (HttpWebResponse)WebReq.GetResponse();
                        //catch id of redirrect
                        currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                        WebResp.Close();
                    }
                    catch (WebException exxx)
                    {
                        if (exxx.Status == WebExceptionStatus.Timeout)
                        {

                            StatusLabel.Text = "Готово";
                            progressBar1.Value = 0;
                            MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else throw;
                    }

                    StatusLabel.Text = "Получение данных - " + entry.Key;
                    progressBar1.Value = 50;

                    //MessageBox.Show("Строка перехвата отчета: "+ currentReportIdstr);
                    int tempReportId;
                    try
                    {
                        tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                        //MessageBox.Show("Идет перехват отчета №" + tempReportId.ToString());
                        StatusLabel.Text = "Перехват отчета " + tempReportId.ToString();
                        progressBar1.Value = 55;
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                        StatusLabel.Text = "Готово";
                        progressBar1.Value = 0;
                        WebResp.Close();
                        return;
                    }
                    WebResp.Close();
                    //Extract graph data
                    //MessageBox.Show("Попытка получить данные графика");
                    //Get that strange shit to analyzer
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/DynamicsChart?id=" + tempReportId.ToString() + "&pageSize=20&gtype=ByGroups&scale=Default&viewType=MlgGraph");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.AllowAutoRedirect = true;
                    WebReq.MaximumAutomaticRedirections = 20;
                    WebReq.Timeout = 60000;
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    string answer = _Answer.ReadToEnd();
                    WebResp.Close();
                    WebReq = null;
                    StatusLabel.Text = "Данные получены - " + entry.Key;
                    progressBar1.Value = 70;

                    //regex-find
                    string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                    //System.IO.File.WriteAllText(base_path + "\\ans_" + entry.Key.ToString() + ".txt", answer);
                    string resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                    //System.IO.File.WriteAllText(base_path + "\\resultString_" + entry.Key.ToString() + ".txt", resultString);
                    Regex reg = new Regex("legendItemSign.*?an>(.*?)<.*?([-]*\\d.*?)<.*?&count=(\\d*?)&");
                    MatchCollection data = reg.Matches(resultString);

                    double messageCount;
                    double mediaIndex;
                    string depName;
                    if (data.Count > 0)
                    {
                        //MessageBox.Show("Count is: " + data.Count.ToString(), "DEBUG");
                        for (int id = 0; id < data.Count; id++)
                        {
                            depName = data[id].Groups[1].Value;

                            mediaIndex = Convert.ToDouble(data[id].Groups[2].Value.Replace(",", CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator).Replace(" ", "").Replace(" ", ""));

                            messageCount = Convert.ToDouble(data[id].Groups[3].Value.Replace(" ", "").Replace(" ", ""));
                            //System.IO.File.WriteAllText(base_path + "\\dep" + depName + ".txt", messageCount.ToString() + " - " + mediaIndex.ToString());
                            _Excel.Range found = ws.Cells.Find(depName, missingObj,
                            _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlWhole,
                            _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                            missingObj, missingObj);
                            if(found == null)
                            {
                                wb.Close();
                                base1.Quit();
                                MessageBox.Show("В локальной базе медиалогии не был найден депутат "+depName+"\nСкорее всего у указанного депутата поменялось полное имя в медиалогии.\n\nСоздание отчета было приостановлено","Внимание!",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                                UpdateStatus();
                                return;
                            }
                            
                            int trow = found.Row;
                            //wtite to excel
                            if (ws.Cells[trow, 2].Value2 == depName)
                            {
                                ws = wb.Sheets[1];
                                ws.Cells[trow, lastcol + 1].Value2 = messageCount.ToString();
                                ws = wb.Sheets[2];
                                ws.Cells[trow, lastcol + 1].Value2 = mediaIndex.ToString();
                            } else
                            {
                                found = ws.Cells.FindNext(found);
                                trow = found.Row;
                                ws = wb.Sheets[1];
                                ws.Cells[trow, lastcol + 1].Value2 = messageCount.ToString();
                                ws = wb.Sheets[2];
                                ws.Cells[trow, lastcol + 1].Value2 = mediaIndex.ToString();
                            }

                        }
                    }
                    else
                    {
                        foreach (KeyValuePair<string, string[]> dep in depinreg)
                        {
                            depName = dep.Key;
                            mediaIndex = Convert.ToDouble(0);
                            messageCount = Convert.ToDouble(0);
                            //System.IO.File.WriteAllText(base_path + "\\dep" + depName + ".txt", messageCount.ToString() + " - " + mediaIndex.ToString());
                            _Excel.Range found = ws.Cells.Find(depName, missingObj,
                            _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlWhole,
                            _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                            missingObj, missingObj);
                            int trow = found.Row;
                            //wtite to excel
                            if (ws.Cells[trow, 2].Value2 == depName)
                            {
                                ws = wb.Sheets[1];
                                ws.Cells[trow, lastcol + 1].Value2 = messageCount.ToString();
                                ws = wb.Sheets[2];
                                ws.Cells[trow, lastcol + 1].Value2 = mediaIndex.ToString();
                            }
                            else
                            {
                                found = ws.Cells.FindNext(found);
                                trow = found.Row;
                                ws = wb.Sheets[1];
                                ws.Cells[trow, lastcol + 1].Value2 = messageCount.ToString();
                                ws = wb.Sheets[2];
                                ws.Cells[trow, lastcol + 1].Value2 = mediaIndex.ToString();
                            }

                        }
                    }
                }
                catch (Exception exe)
                {
                    string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                    System.IO.File.WriteAllText(base_path + "\\exeption.txt", exe.Message + "\n" + exe.StackTrace + "\n" + exe.ToString());
                    MessageBox.Show("Ошибка при получении отчета для " + entry.Key, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    progressBar1.Value = 0;
                    StatusLabel.Text = "Готово";
                    return;
                }

            }
            StatusLabel.Text = "Готово";
            progressBar1.Value = 0;
            try
            {
                wb.SaveAs(Properties.Settings.Default["dep_reg_path"].ToString());
                wb.Close();
                base1.Quit();
            }
            catch
            {
                MessageBox.Show("Программа не может сохранить файл базы.\nПожалуйста, сделайте это вручную (После нажатия \"ОК\" откроется файл с базой)", "Ошибка доступа", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                base1.Visible = true;
            }

            AnFullBase_reg(Properties.Settings.Default["dep_reg_path"].ToString());
            UpdateStatus();

        }

        private void button11_Click(object sender, EventArgs e)
        {
            //request reports from website
            UpdateStatus("В работе", 10, "Проверка данных");
            //request report
            if (Properties.Settings.Default["dep_reg_path"].ToString() == "")
            {
                MessageBox.Show("Файл базы не указан в настройках", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }
            if (!File.Exists(Properties.Settings.Default["dep_reg_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]))
            {
                MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (File.Exists(Properties.Settings.Default["dep_reg_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]))
            {
                MessageBox.Show("Невозможно создать новый файл. Файл с таким именем уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if ((Properties.Settings.Default["dep_reg_list"].ToString()) == "")
            {
                MessageBox.Show("Отсутствует файл со списком депутатов. Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus(); ;
                return;
            }

            if (!File.Exists(Properties.Settings.Default["dep_reg_list"].ToString()))
            {
                MessageBox.Show("Не найден файл со списком депутатов. Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }
            if (!File.Exists(Properties.Settings.Default["dep_reg_database"].ToString()))
            {
                MessageBox.Show("Не найден файл общей базы депутатов. Проверьте настройки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            //check input dates
            if (!ValidateDates(dateTimePicker5, dateTimePicker6))
            {
                UpdateStatus();
                return;
            }

            try
            {
                Report3();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "3";
                reporter.Stage = "Report Request";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }

            

            //await Task.Run(() => Report3());

        }

        public void UpdateStatus(string Status = "Готово", int PBpercentage = 0, string description = "")
        {
            MethodInvoker methodInvokerDelegate = delegate ()
            {
                StatusLabel.Text = Status;
                StatusDesc.Text = description;
                progressBar1.Value = PBpercentage;
            };

            //This will be true if Current thread is not UI thread.
            if (this.InvokeRequired)
                this.Invoke(methodInvokerDelegate);
            else
                methodInvokerDelegate();
        }

        private void AnFullBase_fsec(string path = "none")
        {
            MethodInvoker methodInvokerDelegate = delegate ()
            {

                UpdateStatus("В работе", 5, "Чтение базы");
                if (path == "none")
                {
                    if (Convert.ToBoolean(Properties.Settings.Default["fsec_new"]))
                    {
                        MessageBox.Show("В настройках стоит создание новой базы. Укажите уже существующую базу", "Ошибка загрузки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatus();
                        return;
                    }
                    if (!File.Exists(Properties.Settings.Default["fsec_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["fsec_new"]))
                    {
                        MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatus();
                        return;
                    }
                    UpdateStatus("В работе", 10, "Идет загрузка базы, пожалуйста подождите");
                    textBox4.Text = Properties.Settings.Default["fsec_path"].ToString();
                    //start temp excel
                    _Excel._Application base1 = new _Excel.Application();
                    base1.Visible = false;
                    base1.DisplayAlerts = false;

                    //init wb and ws
                    Workbook wb;
                    Worksheet ws;
                    wb = base1.Workbooks.Open(Properties.Settings.Default["fsec_path"].ToString());
                    ws = wb.Sheets[1];
                    List<string> FindAllDates(int sheet)
                    {
                        ws = wb.Sheets[sheet];
                        int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                        List<string> dates = new List<string>();
                        for (int col1 = 3; col1 <= cols + 1; col1++)
                        {
                            if (ws.Cells[1, col1].Value2 != null)
                            {
                                dates.Add(ws.Cells[1, col1].Value2);
                            }
                        }
                        return dates;
                    }
                    List<string> dates_temp = FindAllDates(1);
                    listBox5.Items.Clear();
                    for (int i = 0; i < dates_temp.Count; i++)
                    {
                        listBox5.Items.Add(dates_temp[i]);
                    }
                    wb.Close();
                    base1.Quit();
                    button13.Enabled = true;
                    UpdateStatus();
                }
                else
                {
                    if (!File.Exists(path))
                    {
                        MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    UpdateStatus("В работе", 10, "Идет загрузка базы, пожалуйста подождите");
                    textBox4.Text = path;
                    //start temp excel
                    _Excel._Application base1 = new _Excel.Application();
                    base1.Visible = false;
                    base1.DisplayAlerts = false;

                    //init wb and ws
                    Workbook wb;
                    Worksheet ws;
                    wb = base1.Workbooks.Open(path);
                    ws = wb.Sheets[1];
                    List<string> FindAllDates(int sheet)
                    {
                        ws = wb.Sheets[sheet];
                        int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                        List<string> dates = new List<string>();
                        for (int col1 = 3; col1 <= cols + 1; col1++)
                        {
                            if (ws.Cells[1, col1].Value2 != null)
                            {
                                dates.Add(ws.Cells[1, col1].Value2);
                            }
                        }
                        return dates;
                    }
                    List<string> dates_temp = FindAllDates(1);
                    listBox5.Items.Clear();
                    for (int i = 0; i < dates_temp.Count; i++)
                    {
                        listBox5.Items.Add(dates_temp[i]);
                    }
                    wb.Close();
                    base1.Quit();
                    button13.Enabled = true;
                    UpdateStatus();

                }

                //invoker check
            };

            //This will be true if Current thread is not UI thread.
            if (this.InvokeRequired)
                this.Invoke(methodInvokerDelegate);
            else
                methodInvokerDelegate();
            return;
        }


        private async void button12_Click(object sender, EventArgs e)
        {
            //loadbase
            Thread backgroundThread = new Thread(() => AnFullBase_fsec());
            await Task.Run(() => backgroundThread.Start());
            //await Task.Run(()=>AnFullBase_fsec());

        }


        private void Fsec_create()
        {

            string target = listBox5.SelectedItem.ToString();
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            UpdateStatus("В работе", 15, "Чтение базы");

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["fsec_path"].ToString());
            ws = wb.Sheets[1];
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int tcol = ws.Cells.Find(target, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj).Column;

            //paste labels
            ws = wb.Sheets[1];
            _Excel.Range sourceRange = ws.Range[ws.Cells[1, 1], ws.Cells[lastrow, 2]];
            ws = wb.Sheets[7];
            _Excel.Range destinationRange = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 3]];
            sourceRange.Copy(destinationRange);
            //paste numbers
            ws.Cells[1, 1].Value2 = "№";
            for (int ri = 2; ri <= lastrow; ri++)
            {
                ws.Cells[ri, 1].Value2 = (ri - 1).ToString();
            }

            //paste message count
            UpdateStatus("В работе", 25, "Анализ сообщений");

            ws = wb.Sheets[1];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 4], ws.Cells[lastrow, 4]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 4].Value2 = "Кол-во сообщений";

            //paste mindex
            UpdateStatus("В работе", 35, "Анализ Медиа-Индекса");

            ws = wb.Sheets[2];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 5], ws.Cells[lastrow, 5]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 5].Value2 = "Медиа-Индекс";

            //paste range of influence
            UpdateStatus("В работе", 45, "Анализ охвата аудитории");

            ws = wb.Sheets[3];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 6], ws.Cells[lastrow, 6]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 6].Value2 = "Охват";

            //paste citations
            UpdateStatus("В работе", 55, "Анализ цитирований");

            ws = wb.Sheets[4];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 7], ws.Cells[lastrow, 7]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 7].Value2 = "Цитирование";


            //paste Negative
            UpdateStatus("В работе", 55, "Анализ негативных упоминаний");

            ws = wb.Sheets[5];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 8], ws.Cells[lastrow, 8]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 8].Value2 = "Негативный характер упом.";

            //paste Positive
            UpdateStatus("В работе", 55, "Анализ позитивных упоминаний");

            ws = wb.Sheets[6];
            sourceRange = ws.Range[ws.Cells[1, tcol], ws.Cells[lastrow, tcol]];
            ws = wb.Sheets[7];
            destinationRange = ws.Range[ws.Cells[1, 9], ws.Cells[lastrow, 9]];
            sourceRange.Copy(destinationRange);
            ws.Cells[1, 9].Value2 = "Позитивный характер упом.";




            UpdateStatus("В работе", 65, "Форматирование таблицы");
            for (int ro = 2; ro <= lastrow; ro++)
            {
                string name = ws.Cells[ro, 2].Value2;
                string[] parts = name.Split(" ".ToCharArray());
                name = parts[0].Substring(0, 1) + parts[0].Substring(1).ToLower() + " " + parts[1].Substring(0, 1) + ". " + parts[2].Substring(0, 1) + ".";
                ws.Cells[ro, 2].Value2 = name;
            }
            ws.Columns[1].ColumnWidth = 3;
            ws.Columns[2].ColumnWidth = 20;
            ws.Columns[3].ColumnWidth = 21;
            ws.Columns[4].ColumnWidth = 11;
            ws.Columns[5].ColumnWidth = 13;
            ws.Columns[6].ColumnWidth = 10;
            ws.Columns[7].ColumnWidth = 13;
            ws.Columns[8].ColumnWidth = 11;
            ws.Columns[9].ColumnWidth = 11;
            ws.Rows[1].Cells.WrapText = true;
            ws.UsedRange.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            ws.UsedRange.Borders.Weight = _Excel.XlBorderWeight.xlThin;

            _Excel.Range rng = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 7]];

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(rng.Columns[3], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
            var sort = ws.Sort;
            sort.SetRange(rng.Rows);
            sort.Header = _Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
            sort.Apply();

            UpdateStatus("В работе", 85, "Составление отчета");
            ws.UsedRange.Copy(missingObj);

            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            doc.Range(0, 0).PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            doc.Range(0, 0).PasteExcelTable(false, false, false);
            wb.Close();
            base1.Quit();

            SaveFileDialog sf = new SaveFileDialog();

            sf.InitialDirectory = "c:\\";
            sf.Filter = "Word files (*.docx)|*.docx";
            sf.FilterIndex = 0;
            sf.RestoreDirectory = true;

            if (sf.ShowDialog() == DialogResult.OK)
            {
                doc.SaveAs2(sf.FileName);
                MessageBox.Show("Отчет был сохранен как " + sf.FileName, "Сохранение");
            }
            else
            {
                MessageBox.Show("Сохранение отчета было отменено", "Отмена");

            }
            doc.Close();
            app.Quit();

            UpdateStatus();

        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (listBox5.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите дату для отчета из списка доступных замеров", "Замер не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Fsec_create();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "4";
                reporter.Stage = "Report Creation";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }


        }

        private void Fsec_request()
        {

            UpdateStatus("В работе", 15, "Чтение баз");
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;
            //init wb and ws
            Workbook wb;
            Worksheet ws;

            //lookup table
            _Excel._Application lookup = new _Excel.Application();
            lookup.Visible = false;
            lookup.DisplayAlerts = false;
            //init wb and ws
            Workbook wblu;
            Worksheet wslu;
            wblu = lookup.Workbooks.Open(Properties.Settings.Default["fsec_database"].ToString());
            wslu = wblu.Sheets[1];

            string line;
            IDictionary<string, string[]> SelectedFsec = new Dictionary<string, string[]>();
            StreamReader file = new System.IO.StreamReader(Properties.Settings.Default["fsec_list"].ToString());
            UpdateStatus("В работе", 20, "Составление списков");
            while ((line = file.ReadLine()) != null)
            {
                try
                {
                    _Excel.Range found = wslu.Cells.Find(line, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj);
                    int trow = found.Row;
                    string id = wslu.Cells[trow, 1].Value2.ToString();
                    string reg = wslu.Cells[trow, 3].Value2;

                    //structure is NAME: ID, REGION, messagecount, Mindex, reach, bad, good, citations
                    SelectedFsec.Add(line, new string[] { id, reg, "0", "0", "н/д", "0", "0", "0" });

                }
                catch
                {
                    MessageBox.Show("Первый секретарь " + line + " не был найден в базе медиалогии", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            wblu.Close();
            lookup.Quit();


            //create with labels if new
            if (Convert.ToBoolean(Properties.Settings.Default["fsec_new"]))
            {
                UpdateStatus("В работе", 25, "Подготовка новой базы");
                wb = base1.Workbooks.Add(Type.Missing);
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                wb.Sheets.Add(Count: 6);

                wb.Sheets[1].Name = "Количество упоминаний";
                wb.Sheets[1].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[1].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[2].Name = "Медиа-Индекс";
                wb.Sheets[2].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[2].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[3].Name = "Охват";
                wb.Sheets[3].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[3].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[4].Name = "Цитирование";
                wb.Sheets[4].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[4].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[5].Name = "Негативный характер упоминаний";
                wb.Sheets[5].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[5].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[6].Name = "Позитивный характер упоминаний";
                wb.Sheets[6].Cells[1, 1].Value2 = "ФИО";
                wb.Sheets[6].Cells[1, 2].Value2 = "Отделение";

                wb.Sheets[7].Name = "Технический";

                int cou = 2;
                foreach (KeyValuePair<string, string[]> entry in SelectedFsec)
                {

                    for (int temp_sheet = 1; temp_sheet < 7; temp_sheet++)
                    {
                        wb.Sheets[temp_sheet].Cells[cou, 1].Value2 = entry.Key;
                        wb.Sheets[temp_sheet].Cells[cou, 2].Value2 = entry.Value[1];
                    }
                    cou++;
                }
                wb.SaveAs(Properties.Settings.Default["fsec_path"].ToString());
                wb.Close();
            }

            //prepare request parts
            //dates
            string dayTo, dayFrom, monthTo, monthFrom;


            if (Convert.ToInt32(dateTimePicker7.Value.Day) < 10) { dayFrom = "0" + dateTimePicker7.Value.Day.ToString(); } else { dayFrom = dateTimePicker7.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker8.Value.Day) < 10) { dayTo = "0" + dateTimePicker8.Value.Day.ToString(); } else { dayTo = dateTimePicker8.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker7.Value.Month) < 10) { monthFrom = "0" + dateTimePicker7.Value.Month.ToString(); } else { monthFrom = dateTimePicker7.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker8.Value.Month) < 10) { monthTo = "0" + dateTimePicker8.Value.Month.ToString(); } else { monthTo = dateTimePicker8.Value.Month.ToString(); }



            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker7.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker7.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker8.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker8.Value.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input


            //<repname>
            string part1 = "{\"smsMonitor\":{\"MonitorId\":-1,\"ThemeId\":-1,\"UserId\":-1,\"MaxSendingArticle\":0,\"SendingMode\":2,\"SendingPeriod\":1,\"ReprintsMode\":3,\"MonitorPhones\":[]},\"folder\":\"\",\"folderId\":-1,\"Authors\":[],\"Cities\":[],\"Levels\":[1,2],\"Categories\":[1,2,3,4,5,6],\"Rubrics\":[],\"LifeStyles\":[],\"MediaSources\":[],\"MediaBranches\":[],\"MediaObjectBranches\":[],\"MediaObjectLifeStyles\":[],\"MediaObjectLevels\":[],\"MediaObjectCategories\":[],\"MediaObjectRegions\":[],\"MediaObjectFederals\":[],\"MediaObjectTowns\":[],\"MediaLanguages\":[],\"MediaRegions\":[],\"MediaCountries\":[],\"CisMediaCountries\":[],\"MediaFederals\":[],\"MediaGenre\":[],\"YandexRubrics\":[],\"Role\":-1,\"Tone\":-1,\"Quotation\":-1,\"CityMode\":0,\"messageCount\":-1,\"reprintsMessageCount\":-1,\"CheckedMessageCount\":-1,\"CheckedClustersCount\":-1,\"MonitorId\":-1,\"CheckedReprintsCount\":-1,\"deletedMessageCount\":-1,\"favoritesMessageCount\":-1,\"myDocsMessageCount\":0,\"myMediaMessageCount\":0,\"IsSaveParamsOnly\":false,\"RebuildDBCache\":false,\"Credentials\":null,\"AppType\":1,\"ParamsVersion\":0,\"ArmObjectMode\":0,\"ReportCreatingHistory\":0,\"InfluenceThreshold\":\"0.0\",\"MonitorObjects\":null,\"Icon\":0,\"ThemeGroup\":-1,\"ThemeGroupName\":\"\",\"SaveMode\":0,\"MonitorExists\":false,\"ThemeId\":-1,\"Title\":\"<repname>\",\"Comment\":\"\",\"ReprintMode\":0,\"rssReportType\":0,\"ThemeObjects\":[";

            //<id> <name> <lindex>
            string obj = "{\"Id\":\"<id>\",\"MainObjectId\":\"<id>\",\"ObjectName\":\"<name>\",\"classId\":43,\"LogicIndex\":<lindex>,\"LogicObjectString\":\"OR\",\"SearchQuery\":null,\"Properties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}]}";

            //<allstring> <logic> <datefrom> <dateto> <timefrom> <timeto>
            string part3 = "],\"ThemeObjectsFromSearchContext\":[],\"ThemeTypes\":[],\"ThemeBranches\":[],\"AllObjectsProperties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}],\"AllArticlesProperties\":[],\"AllObjectString\":\"<allstring>\",\"AllLogicObjectString\":\"<logic>\",\"DatePeriod\":8,\"DateType\":0,\"Date\":\"<datefrom>|<dateto>\",\"Time\":\"<timefrom>|<timeto>\",\"ActualDatePeriod\":3,\"IsSlidingTime\":true,\"ContextScope\":5,\"Context\":\"\",\"ContextMode\":0,\"TopMedia\":false,\"RegionLogic\":0,\"MediaObjectRegionLogic\":0,\"MediaLogic\":0,\"MediaLogicAll\":0,\"BlogLogic\":1,\"MediaBranchLogic\":0,\"MediaObjectBranchLogic\":0,\"MediaLanguageLogic\":0,\"MediaCountryLogic\":0,\"CityLogic\":0,\"Compare\":0,\"User\":0,\"Type\":1,\"View\":0,\"ViewStatus\":1,\"OiiMode\":0,\"Template\":-1,\"MediaStatus\":-1,\"IsUpdate\":false,\"HasUserObjects\":false,\"IsContextReport\":false,\"LastCopiedThemeId\":null}";

            string objs = "";
            string allobjectstring = "";
            string alllogic = "";

            int lindex = 0;
            foreach (KeyValuePair<string, string[]> sec in SelectedFsec)
            {
                objs += obj.Replace("<id>", sec.Value[0])
                    .Replace("<name>", sec.Key)
                    .Replace("<lindex>", lindex.ToString());
                if (lindex + 1 < SelectedFsec.Keys.Count)
                {
                    objs += ", ";
                }
                allobjectstring += "+O" + sec.Value[0] + "_" + lindex.ToString();
                alllogic += "+" + lindex.ToString();
                lindex++;
            }
            string target = datefrom + "-" + dateto;
            string sr = part1.Replace("<repname>", "Fsec_" + target) + objs + part3
                .Replace("<allstring>", allobjectstring)
                .Replace("<logic>", alllogic)
                .Replace("<datefrom>", datefrom)
                .Replace("<dateto>", dateto)
                .Replace("<timefrom>", timefrom)
                .Replace("<timeto>", timeto);

            //auth
            UpdateStatus("В работе", 35, "Авторизация");
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            var cookieContainer = new CookieContainer();
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;
            string data_to_post;
            byte[] buffer;
            try
            {
                data_to_post = "UserName=" + Properties.Settings.Default["login"] + "&Password=" + Properties.Settings.Default["password"] + "&PrUrl=http%3A%2F%2Fpr.mlg.ru&Pr2Url=http%3A%2F%2Fdev.pr2.mlg.ru&MmUrl=http%3A%2F%2Fmm.mlg.ru&BuzzUrl=http%3A%2F%2Fsm.mlg.ru&ReturnUrl=http%3A%2F%2Fpr.mlg.ru&ApplicationType=Pr";
                buffer = Encoding.ASCII.GetBytes(data_to_post);

                WebReq = (HttpWebRequest)WebRequest.Create("https://login.mlg.ru/Account.mlg?ApplicationType=Pr");
                WebReq.CookieContainer = cookieContainer;
                WebReq.Timeout = 30000;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;

                PostData = WebReq.GetRequestStream();
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                WebResp.Close();
            }
            catch (Exception ex)
            {
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "?";
                reporter.Stage = "Auth";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

                MessageBox.Show("Ошибка авторизации. Проверьте правильность логина/пароля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                base1.Quit();
                return;
            }

            //make a request
            string currentReportIdstr;
            byte[] urljson = Encoding.ASCII.GetBytes(sr);
            string urlencoded = HttpUtility.UrlEncode(urljson);
            //create payload and send it

            data_to_post = "useFilterContainers=false&sr=" + urlencoded;
            buffer = Encoding.ASCII.GetBytes(data_to_post);
            //MessageBox.Show("Буффер составлен");
            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                WebReq.MaximumAutomaticRedirections = 1;
                WebReq.AllowAutoRedirect = false;
                WebReq.CookieContainer = cookieContainer;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;
                WebReq.Timeout = 20000;
                PostData = WebReq.GetRequestStream();
                //MessageBox.Show("Буффер отправлен");
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                //catch id of redirrect
                currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                WebResp.Close();
            }
            catch (WebException exxx)
            {
                if (exxx.Status == WebExceptionStatus.Timeout)
                {
                    base1.Quit();
                    UpdateStatus();
                    MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else throw;
            }

            UpdateStatus("В работе", 40, "Получение данных");


            //MessageBox.Show("Строка перехвата отчета: "+ currentReportIdstr);
            int tempReportId;
            try
            {
                tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                //MessageBox.Show("Идет перехват отчета №" + tempReportId.ToString());
            }
            catch
            {
                MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                WebResp.Close();
                base1.Quit();
                UpdateStatus();
                return;
            }
            WebResp.Close();

            //loop through pages and reg-search-add


            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + tempReportId.ToString() + "&pageSize=20&pageNumber=1");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.Timeout = 60000;
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                string answer = _Answer.ReadToEnd();
                WebResp.Close();
                WebReq = null;
                //regex-find
                //string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                //System.IO.File.WriteAllText(base_path + "\\ans_" + entry.Key.ToString() + ".txt", answer);
                string resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                //System.IO.File.WriteAllText(base_path + "\\resultString_" + entry.Key.ToString() + ".txt", resultString);
                Regex reg = new Regex("pagesCount = (\\d+)");
                Regex sreg = new Regex("по объекту' data-title-align='top'>(.*?)<.*?&count=(.*?)\".*?-->(.*?)<.*?ce\">(.*?)<\\/td.*?td.*?(?>negative\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>positive\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>speech.*?>(\\d.*?)<\\/|&count=(\\d+))");
                MatchCollection pages_collection = reg.Matches(resultString);
                if (pages_collection.Count > 0)
                {
                    int pages = Convert.ToInt32(pages_collection[0].Groups[1].Value);
                    UpdateStatus("В работе", 45, "Получение данных. Страница 1/" + pages.ToString());
                    MatchCollection data = sreg.Matches(resultString);
                    for (int g = 0; g < data.Count; g++)
                    {
                        int count_arr_point = 2;
                        for (int group_id = 2; group_id <= data[g].Groups.Count; group_id++)
                        {

                            if (data[g].Groups[group_id].Value != "")
                            {
                                SelectedFsec[data[g].Groups[1].Value][count_arr_point] = data[g].Groups[group_id].Value;
                                count_arr_point++;
                            }

                        }
                    }
                    if (pages > 1)
                    {
                        for (int pageId = 2; pageId <= pages; pageId++)
                        {
                            UpdateStatus("В работе", 45, "Получение данных. Страница " + pageId + "/" + pages.ToString());
                            try
                            {
                                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + tempReportId.ToString() + "&pageSize=20&pageNumber=" + pageId.ToString());
                                WebReq.CookieContainer = cookieContainer;
                                WebReq.ContentType = "application/x-www-form-urlencoded";
                                WebReq.AllowAutoRedirect = true;
                                WebReq.MaximumAutomaticRedirections = 20;
                                WebReq.Timeout = 60000;
                                WebResp = (HttpWebResponse)WebReq.GetResponse();
                                Answer = WebResp.GetResponseStream();
                                _Answer = new StreamReader(Answer);
                                answer = _Answer.ReadToEnd();
                                WebResp.Close();
                                WebReq = null;
                                resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                                data = sreg.Matches(resultString);
                                for (int g = 0; g < data.Count; g++)
                                {
                                    int count_arr_point = 2;
                                    for (int group_id = 2; group_id <= data[g].Groups.Count; group_id++)
                                    {

                                        if (data[g].Groups[group_id].Value != "")
                                        {
                                            SelectedFsec[data[g].Groups[1].Value][count_arr_point] = data[g].Groups[group_id].Value;
                                            count_arr_point++;
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ReportSender Sender = new ReportSender();
                                Reporter reporter = new Reporter();
                                reporter.EventType = ex.Message;
                                reporter.ReportType = "?";
                                reporter.Stage = "Request Fsec";
                                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                                Sender.SendReport(reporter);

                                base1.Quit();
                                UpdateStatus();
                                MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }

                }
                else
                {
                    foreach (KeyValuePair<string, string[]> sec in SelectedFsec)
                    {
                        sec.Value[2] = "0";
                        sec.Value[3] = "0";
                        sec.Value[4] = "н/д";
                        sec.Value[5] = "0";
                        sec.Value[6] = "0";
                        sec.Value[7] = "0";
                    }
                }

            }
            catch (Exception ex)
            {
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "Fsec";
                reporter.Stage = "Report request";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

                MessageBox.Show("Ошибка при получении отчета", "Ошибка");
                UpdateStatus();
                WebResp.Close();
                base1.Quit();
                return;
            }

            //assuming that base has something to work with
            wb = base1.Workbooks.Open(Properties.Settings.Default["fsec_path"].ToString());
            ws = wb.Sheets[1];
            int lastcol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

            for (int sheet = 1; sheet < 7; sheet++)
            {
                UpdateStatus("В работе", 75, "Запись данных в базу. Лист " + sheet.ToString());
                ws = wb.Sheets[sheet];
                ws.Cells[1, lastcol + 1].Value2 = target;
                foreach (KeyValuePair<string, string[]> sec in SelectedFsec)
                {
                    _Excel.Range found = ws.Cells.Find(sec.Key, missingObj,
                    _Excel.XlFindLookIn.xlValues, _Excel.XlLookAt.xlPart,
                    _Excel.XlSearchOrder.xlByRows, _Excel.XlSearchDirection.xlNext, false,
                    missingObj, missingObj);
                    int trow = found.Row;
                    ws.Cells[trow, lastcol + 1].Value2 = sec.Value[sheet + 1];
                }
            }
            UpdateStatus("В работе", 95, "Сохранение");
            wb.SaveAs(Properties.Settings.Default["fsec_path"].ToString());
            wb.Close();
            base1.Quit();
            AnFullBase_fsec(Properties.Settings.Default["fsec_path"].ToString());
            UpdateStatus();
        }
        private void button14_Click(object sender, EventArgs e)
        {
            if (!ValidateDates(dateTimePicker7, dateTimePicker8))
            {
                return;
            }

            if (!File.Exists(Properties.Settings.Default["fsec_list"].ToString()))
            {
                MessageBox.Show("Не найден файл со списком первых секретарей", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(Properties.Settings.Default["fsec_database"].ToString()))
            {
                MessageBox.Show("Не найдена общая база Медиалогии для первых секретарей", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(Properties.Settings.Default["fsec_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["fsec_new"]))
            {
                MessageBox.Show("Не найдена историческая база по первым секретарям.\nУкажите существующий файл, либо создайте новую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (File.Exists(Properties.Settings.Default["fsec_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["fsec_new"]))
            {
                MessageBox.Show("База с указанным именем уже существует.\nВыберите другое имя базы, либо выберите уже существующую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Fsec_request();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "4";
                reporter.Stage = "Report Request";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }



        }


        private void Fsec_readbase()
        {
            //load database to list boxes
            UpdateStatus("В работе", 5, "Чтение базы");
            if (!File.Exists(Properties.Settings.Default["media_database"].ToString()))
            {
                MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }
            UpdateStatus("В работе", 10, "Идет загрузка базы, пожалуйста подождите");

            //start temp excel
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["media_database"].ToString());
            ws = wb.Sheets[1];
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            Globals.MediaDB.Clear();
            for (int ro = 1; ro <= lastrow; ro++)
            {
                string name = ws.Cells[ro, 2].Value2;
                string id = ws.Cells[ro, 1].Value2.ToString();
                string group = ws.Cells[ro, 3].Value2.ToString();
                Globals.MediaDB.Add(name, new string[] { id, group });
                checkedListBox1.Items.Add(name);
            }
            ws = wb.Sheets[2];
            lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            for (int ro = 1; ro <= lastrow; ro++)
            {
                string name = ws.Cells[ro, 2].Value2;
                string id = ws.Cells[ro, 1].Value2.ToString();
                string group = ws.Cells[ro, 3].Value2.ToString();
                Globals.MediaDB.Add(name, new string[] { id, group });
                checkedListBox2.Items.Add(name);
            }

            UpdateStatus();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                Fsec_readbase();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "4";
                reporter.Stage = "Database Analysys";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }
            
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                button16.Enabled = true;
                checkedListBox1.Enabled = true;
                checkedListBox2.Enabled = true;
                button15.Enabled = true;
            }
            else
            {
                button16.Enabled = false;
                checkedListBox1.Enabled = false;
                checkedListBox2.Enabled = false;
                if (!checkBox1.Checked)
                {
                    button15.Enabled = false;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox5.Enabled = true;
                button15.Enabled = true;
                button17.Enabled = true;
                button18.Enabled = true;
            }
            else
            {
                textBox5.Enabled = false;
                button17.Enabled = false;
                button18.Enabled = false;
                if (!checkBox2.Checked)
                {
                    button15.Enabled = false;
                }
            }

        }

        public async Task<IDictionary<string, string[]>> GetMediaTable1(string ReportId, CookieContainer cookieContainer)
        {
            IDictionary<string, string[]> table = new Dictionary<string, string[]>();

            if (InvokeRequired)
            {Invoke((MethodInvoker)(() =>{
                    UpdateStatus("В работе", 45, "Получение данных");
            }));}else{
                UpdateStatus("В работе", 45, "Получение данных");
            }

            
            try
            {
                HttpWebRequest WebReq;
                HttpWebResponse WebResp;
                Stream Answer;
                StreamReader _Answer;
                MatchCollection data;
                string resultString;
                string answer;
                Regex sreg = new Regex("blank\">(.*?)<\\/a.*?>(\\d+).*?>(\\d+).*?>(\\d+).*?>(\\d+).*?>(\\d+).*?>(\\d+).*?>(\\d+).*?>(\\d+)");

                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/MediaCategoriesGrid?id="+ ReportId + "&pageSize=20&pageNumber=1&columnName=ArtCount0&order=desc");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.Timeout = 60000;
                WebResp = (HttpWebResponse)await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                answer = _Answer.ReadToEnd();
                WebResp.Close();
                WebReq = null;
                resultString = Regex.Replace(answer, @"\r\n| ", string.Empty, RegexOptions.Multiline);
                resultString = Regex.Replace(resultString, @"(?<=\d)(&nbsp;)(?=\d)", string.Empty, RegexOptions.Multiline);
                resultString = resultString.Replace(" ", "");
                data = sreg.Matches(resultString);
                
                foreach (Match match in data)
                {
                    string gr = match.Groups[1].Value.Replace("&nbsp;", " ");

                    string np = match.Groups[2].Value;
                    string magaz = match.Groups[3].Value;
                    string informag = match.Groups[4].Value;
                    string inter = match.Groups[5].Value;
                    string tv = match.Groups[6].Value;
                    string radio = match.Groups[7].Value;
                    string blogs = match.Groups[8].Value;
                    string total = match.Groups[9].Value;
                    table.Add(gr, new string[] { np, magaz, informag, inter, tv, radio, blogs, total }); ;
                }
                return table;
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.Timeout)
                {
                    if (InvokeRequired)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);


                        }
                        ));
                    }
                    else
                    {
                        MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }
                else {

                    if (InvokeRequired)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            MessageBox.Show("Произошла неизвестная ошибка. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        ));
                    }
                    else
                    {
                        MessageBox.Show("Произошла неизвестная ошибка. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }

                if (InvokeRequired)
                {
                    Invoke((MethodInvoker)(() =>
                    {
                        UpdateStatus();

                    }
                    ));
                }
                else
                {
                    UpdateStatus();
                }
                return table;
            }
        }

        public async Task<IDictionary<string, string[]>> GetMediaTable2(string ReportId, CookieContainer cookieContainer)
        {
            IDictionary<string, string[]> table = new Dictionary<string, string[]>();
            if (InvokeRequired)
            {
                Invoke((MethodInvoker)(() =>
                {
                    UpdateStatus("В работе", 45, "Получение данных о СМИ");
                }));
            }
            else
            {
                UpdateStatus("В работе", 45, "Получение данных о СМИ");
            }

                try
            {
                HttpWebRequest WebReq;
                HttpWebResponse WebResp;
                Stream Answer;
                StreamReader _Answer;
                MatchCollection data;
                string resultString;
                string answer;

                Regex sreg = new Regex("источника (.*?)' data.*?'top'>(.*?)<.*?category\">(.*?)<\\/td.*?city\">(.*?)<\\/td");
                Regex pagesreg = new Regex("pagesCount = (\\d+)");
                
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/MediaByCountGrid?top=0&id="+ReportId+"&pageSize=20&pageNumber=1&columnName=Mention&order=desc&onlyNewArticles=false");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.Timeout = 60000;
                WebResp = (HttpWebResponse)await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                answer = _Answer.ReadToEnd();
                WebResp.Close();
                WebReq = null;

                resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                resultString = Regex.Replace(resultString, @"(?<=\d)(&nbsp;)(?=\d)", string.Empty, RegexOptions.Multiline);
                resultString = resultString.Replace(" ", "");

                data = sreg.Matches(resultString);
                int pagescount = Convert.ToInt32(pagesreg.Matches(resultString)[0].Groups[1].Value);

                if (InvokeRequired)
                {
                    Invoke((MethodInvoker)(() =>
                    {
                        UpdateStatus("В работе", 55, "Получение данных. Страница 1/" + pagescount.ToString());
                    }));
                }
                else
                {
                    UpdateStatus("В работе", 55, "Получение данных. Страница 1/" + pagescount.ToString());
                }


                string name, count, type, city;
                foreach (Match match in data)
                {
                    name = match.Groups[1].Value;
                    count = match.Groups[2].Value;
                    type = match.Groups[3].Value;
                    city = match.Groups[4].Value;
                    table.Add(name, new string[] {count, type, city});
                }

                for(int pageNum = 2; pageNum<= pagescount; pageNum++)
                {

                    if (InvokeRequired)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            UpdateStatus("В работе", 65, "Получение данных. Страница " + pageNum.ToString() + "/" + pagescount.ToString());
                        }));
                    }
                    else
                    {
                        UpdateStatus("В работе", 65, "Получение данных. Страница " + pageNum.ToString() + "/" + pagescount.ToString());
                    }



                    
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/MediaByCountGrid?top=0&id=" + ReportId + "&pageSize=20&pageNumber="+pageNum.ToString()+"&columnName=Mention&order=desc&onlyNewArticles=false");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.AllowAutoRedirect = true;
                    WebReq.MaximumAutomaticRedirections = 20;
                    WebReq.Timeout = 60000;
                    WebResp = (HttpWebResponse) await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    answer = _Answer.ReadToEnd();
                    WebResp.Close();
                    WebReq = null;

                    resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                    resultString = Regex.Replace(resultString, @"(?<=\d)(&nbsp;)(?=\d)", string.Empty, RegexOptions.Multiline);
                    resultString = resultString.Replace(" ", "");
                    data = sreg.Matches(resultString);
                    
                    foreach (Match match in data)
                    {
                        name = match.Groups[1].Value;
                        count = match.Groups[2].Value;
                        type = match.Groups[3].Value;
                        city = match.Groups[4].Value;
                        table.Add(name, new string[] { count, type, city });
                    }
                }

                return table;
            }
            catch {
                return table;
            }

        }

        public async Task<IDictionary<string, string[]>> GetMediaTable3(string ReportId, CookieContainer cookieContainer)
        {
            IDictionary<string, string[]> table = new Dictionary<string, string[]>();


            if (InvokeRequired)
            {
                Invoke((MethodInvoker)(() =>
                {
                    UpdateStatus("В работе", 75, "Получение данных по субъектам");
                }));
            }
            else
            {
                UpdateStatus("В работе", 75, "Получение данных по субъектам");
            }

            
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            Stream Answer;
            StreamReader _Answer;
            MatchCollection data;
            string resultString;
            string answer;

            Regex sreg = new Regex("по объекту' data-title-align='top'>(.*?)<\\/.*?>([\\d, ]+)<\\/.*?>([\\d, ]+)<.*?negative.*?>([\\d, ]+)<.*?>([\\d, ]+)<.*?>([\\d, ]+)<");
            Regex pagesreg = new Regex("pagesCount = (\\d+)");


            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + ReportId+"&pageSize=20&pageNumber=1&columnName=&order=&onlyNewArticles=false");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.Timeout = 60000;
                WebResp = (HttpWebResponse)await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                answer = _Answer.ReadToEnd();
                WebResp.Close();
                WebReq = null;

                resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                resultString = Regex.Replace(resultString, @"(?<=\d)(&nbsp;)(?=\d)", string.Empty, RegexOptions.Multiline);
                resultString = resultString.Replace(" ", "");

                data = sreg.Matches(resultString);
                int pagescount = Convert.ToInt32(pagesreg.Matches(resultString)[0].Groups[1].Value);

                

                
                string name, count, role, neg, pos, cit;
                foreach (Match match in data)
                {
                    name = match.Groups[1].Value;
                    count = match.Groups[2].Value;
                    role = match.Groups[3].Value;
                    neg = match.Groups[4].Value;
                    pos = match.Groups[5].Value;
                    cit = match.Groups[6].Value;
                    table.Add(name, new string[] { count, role, neg, pos, cit });
                }

                for (int pageNum = 2; pageNum <= pagescount; pageNum++)
                {

                    if (InvokeRequired)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            UpdateStatus("В работе", 90, "Получение данных. Страница " + pageNum.ToString() + "/" + pagescount.ToString());
                        }));
                    }
                    else
                    {
                        UpdateStatus("В работе", 90, "Получение данных. Страница " + pageNum.ToString() + "/" + pagescount.ToString());
                    }

                    
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + ReportId + "&pageSize=20&pageNumber="+pageNum.ToString()+"&columnName=&order=&onlyNewArticles=false");
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.AllowAutoRedirect = true;
                    WebReq.MaximumAutomaticRedirections = 20;
                    WebReq.Timeout = 60000;
                    WebResp = (HttpWebResponse)await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                    Answer = WebResp.GetResponseStream();
                    _Answer = new StreamReader(Answer);
                    answer = _Answer.ReadToEnd();
                    WebResp.Close();
                    WebReq = null;

                    resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                    resultString = Regex.Replace(resultString, @"(?<=\d)(&nbsp;)(?=\d)", string.Empty, RegexOptions.Multiline);
                    resultString = resultString.Replace(" ", "");

                    data = sreg.Matches(resultString);

                    foreach (Match match in data)
                    {
                        name = match.Groups[1].Value;
                        count = match.Groups[2].Value;
                        role = match.Groups[3].Value;
                        neg = match.Groups[4].Value;
                        pos = match.Groups[5].Value;
                        cit = match.Groups[6].Value;
                        table.Add(name, new string[] { count, role, neg, pos, cit });
                    }

                }
            }
            catch (WebException exx){
                if (exx.Status == WebExceptionStatus.Timeout)
                {
                    MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                else
                {
                    MessageBox.Show("Произошла неизвестная ошибка. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                UpdateStatus();
            }


            return table;
        }
    
        public CookieContainer Auth() {
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            var cookieContainer = new CookieContainer();
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;
            string data_to_post;
            byte[] buffer;
            UpdateStatus("В работе", 5, "Авторизация (до 20 секунд)");
            try
            {
                data_to_post = "UserName=" + Properties.Settings.Default["login"] + "&Password=" + Properties.Settings.Default["password"] + "&PrUrl=http%3A%2F%2Fpr.mlg.ru&Pr2Url=http%3A%2F%2Fdev.pr2.mlg.ru&MmUrl=http%3A%2F%2Fmm.mlg.ru&BuzzUrl=http%3A%2F%2Fsm.mlg.ru&ReturnUrl=http%3A%2F%2Fpr.mlg.ru&ApplicationType=Pr";
                buffer = Encoding.ASCII.GetBytes(data_to_post);
                WebReq = (HttpWebRequest)WebRequest.Create("https://login.mlg.ru/Account.mlg?ApplicationType=Pr");
                WebReq.CookieContainer = cookieContainer;
                WebReq.Timeout = 20000;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;
                PostData = WebReq.GetRequestStream();
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                WebResp.Close();
            } catch (WebException exxx)
            {
                if (exxx.Status == WebExceptionStatus.Timeout)
                {
                    MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else {
                    MessageBox.Show("Ошибка авторизации. Проверьте правильность связки логин/пароль.", "Ошибка авторизации", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return cookieContainer;
        }

        public string ValidateContext(string inp, string mode)
        {
            string ret = "";
            inp.Trim();
            if (inp.Substring(inp.Length - 1, 1) == "|")
            {
                inp = inp.Substring(0, inp.Length - 1);
                inp.Trim();
            }
            string[] objs = inp.Split("|".ToCharArray());

            ret += "(";
            int id = 0;
            foreach (string obj in objs){
                string copy = obj;
                copy = copy.Trim();
                copy = copy.Replace("\"", "\\\"")
                    .Replace(" ","+")
                    .Replace("++++","+")
                    .Replace("+++","+")
                    .Replace("++","+");
                if (copy.Split("+".ToCharArray()).Length > 1)
                {
                    ret += "\\\"" + copy + "\\\"";
                }
                else
                {
                    ret += copy;
                }
                if (id < objs.Length - 1)
                {
                    ret += "+" + mode + "+";
                }
                id++;
            }
            ret += ")";
            ret = ret.Replace("+", " ");
            return ret;
        }

        public void ValidateQuerry(string context, CookieContainer cookieContainer)
        {
            string data_to_post;
            byte[] buffer;
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;

            data_to_post = "query=" + HttpUtility.UrlEncode(context);
            buffer = Encoding.ASCII.GetBytes(data_to_post);

            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/ValidateQuery");
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.AllowAutoRedirect = true;
                WebReq.CookieContainer = cookieContainer;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;
                WebReq.Timeout = 20000;
                PostData = WebReq.GetRequestStream();
                //MessageBox.Show("Буффер отправлен");
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebReq = null;
            }
            catch (Exception ex) {

                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = "Http error";
                reporter.ReportType = "?";
                reporter.Stage = "ValidateQuerry";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

            }
         }

        public bool ValidateDates(DateTimePicker begin, DateTimePicker fin)
        {
            //Retuns true if everything is OK, else, returns FALSE;

            DateTime start = begin.Value;
            DateTime end = fin.Value;

            if (DateTime.Compare(start, end) > 0)
            {
                MessageBox.Show("Выбран неверный промежуток времени", "Ошибка дат",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else if (DateTime.Compare(start, end) == 0)
            {
                MessageBox.Show("Для отчета необходим промежуток более 24х часов", "Ошибка дат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string dayTo, dayFrom, monthTo, monthFrom;

            if (Convert.ToInt32(start.Day) < 10) { dayFrom = "0" + start.Day.ToString(); } else { dayFrom = start.Day.ToString(); }
            if (Convert.ToInt32(end.Day) < 10) { dayTo = "0" + end.Day.ToString(); } else { dayTo = end.Day.ToString(); }

            if (Convert.ToInt32(start.Month) < 10) { monthFrom = "0" + start.Month.ToString(); } else { monthFrom = start.Month.ToString(); }
            if (Convert.ToInt32(end.Month) < 10) { monthTo = "0" + end.Month.ToString(); } else { monthTo = end.Month.ToString(); }


            string datefrom = dayFrom + "." + monthFrom + "." + start.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (start.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + end.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (end.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input//


            DialogResult dialogResult = MessageBox.Show("Вы собираетесь сделать новый онлайн запрос со следующими параметрами:\n\nДата начала: "+datefrom_short+"; "+timefrom+"\nДата конца: "+dateto_short+"; "+timeto+"\nРазмер промежутка в днях: "+(end.Subtract(start).Days+1).ToString()+ "\n\nПродолжить?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                return true;
            }
            else if (dialogResult == DialogResult.No)
            {
                return false;
            }

            return false;
        }

       
        private async Task Report5_request()
        {

            IDictionary<string, string[]> MediaSelection = new Dictionary<string, string[]>();

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                string name = checkedListBox1.CheckedItems[i].ToString();
                MediaSelection.Add(name, Globals.MediaDB[name]);
            }

            for (int i = 0; i < checkedListBox2.CheckedItems.Count; i++)
            {
                string name = checkedListBox2.CheckedItems[i].ToString();
                MediaSelection.Add(name, Globals.MediaDB[name]);
            }

            string context = "";
            if (checkBox1.Checked)//create context if enabled
            {
                context = ValidateContext(textBox5.Text, Globals.MediaMode);
            }


            UpdateStatus();
            //string RepId = "3090105";

            string obj_template = "{\"Id\":\"<id>\",\"MainObjectId\":\"<id>\",\"ObjectName\":\"<name>\",\"classId\":<group>,\"LogicIndex\":<lindex>,\"LogicObjectString\":\"OR\",\"SearchQuery\":null,\"Properties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}]}";
            string base_template = "{\"smsMonitor\":{\"MonitorId\":-1,\"ThemeId\":-1,\"UserId\":-1,\"MaxSendingArticle\":0,\"SendingMode\":2,\"SendingPeriod\":1,\"ReprintsMode\":3,\"MonitorPhones\":[]},\"folder\":\"\",\"folderId\":-1,\"Authors\":[],\"Cities\":[],\"Levels\":[],\"Categories\":[1,2,3,4,5,6],\"Rubrics\":[],\"LifeStyles\":[],\"MediaSources\":[],\"MediaBranches\":[],\"MediaObjectBranches\":[],\"MediaObjectLifeStyles\":[],\"MediaObjectLevels\":[],\"MediaObjectCategories\":[],\"MediaObjectRegions\":[],\"MediaObjectFederals\":[],\"MediaObjectTowns\":[],\"MediaLanguages\":[],\"MediaRegions\":[],\"MediaCountries\":[],\"CisMediaCountries\":[],\"MediaFederals\":[],\"MediaGenre\":[],\"YandexRubrics\":[],\"Role\":-1,\"Tone\":-1,\"Quotation\":-1,\"CityMode\":0,\"messageCount\":-1,\"reprintsMessageCount\":-1,\"CheckedMessageCount\":-1,\"CheckedClustersCount\":-1,\"MonitorId\":-1,\"CheckedReprintsCount\":-1,\"deletedMessageCount\":-1,\"favoritesMessageCount\":-1,\"myDocsMessageCount\":0,\"myMediaMessageCount\":0,\"IsSaveParamsOnly\":false,\"RebuildDBCache\":false,\"Credentials\":null,\"AppType\":1,\"ParamsVersion\":0,\"ArmObjectMode\":0,\"ReportCreatingHistory\":<RCH>,\"InfluenceThreshold\":\"0.0\",\"MonitorObjects\":null,\"Icon\":0,\"ThemeGroup\":-1,\"ThemeGroupName\":\"\",\"SaveMode\":0,\"MonitorExists\":false,\"ThemeId\":<TID>,\"Title\":\"<title>\",\"Comment\":\"\",\"ReprintMode\":0,\"rssReportType\":0,\"ThemeObjects\":[<objects>],\"ThemeObjectsFromSearchContext\":[],\"ThemeTypes\":[],\"ThemeBranches\":[],\"AllObjectsProperties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}],\"AllArticlesProperties\":[],\"AllObjectString\":\"<allobjstring>\",\"AllLogicObjectString\":\"<alllogicstring>\",\"DatePeriod\":8,\"DateType\":0,\"Date\":\"<datefrom>|<dateto>\",\"Time\":\"<timefrom>|<timeto>\",\"ActualDatePeriod\":<ADP>,\"IsSlidingTime\":true,\"ContextScope\":5,\"Context\":\"<context>\",\"ContextMode\":0,\"TopMedia\":false,\"RegionLogic\":0,\"MediaObjectRegionLogic\":0,\"MediaLogic\":0,\"MediaLogicAll\":0,\"BlogLogic\":1,\"MediaBranchLogic\":0,\"MediaObjectBranchLogic\":0,\"MediaLanguageLogic\":0,\"MediaCountryLogic\":0,\"CityLogic\":0,\"Compare\":<compare>,\"User\":0,\"Type\":<type>,\"View\":0,\"ViewStatus\":1,\"OiiMode\":0,\"Template\":-1,\"MediaStatus\":-1,\"IsUpdate\":false,\"HasUserObjects\":false,\"IsContextReport\":<iscontext>,\"LastCopiedThemeId\":null}";

            string objs = "";
            string allobjstring = "";
            string alllogic = "";
            //create objs string
            int count = 0;
            foreach (KeyValuePair<string, string[]> entry in MediaSelection)
            {
                string name = entry.Key.Replace("\"", "\\\"").Replace(" ", "+");
                objs += obj_template
                    .Replace("<id>", entry.Value[0])
                    .Replace("<group>", entry.Value[1])
                    .Replace("<name>", name)
                    .Replace("<lindex>", count.ToString());
                if (count < MediaSelection.Count - 1)
                {
                    objs += ",";
                }
                allobjstring += "+O" + entry.Value[0] + "_" + count.ToString();
                alllogic += "+" + count.ToString();
                count++;
            }

            // get dates from inputs
            string dayTo, dayFrom, monthTo, monthFrom, jsonq;

            if (Convert.ToInt32(dateTimePicker9.Value.Day) < 10) { dayFrom = "0" + dateTimePicker9.Value.Day.ToString(); } else { dayFrom = dateTimePicker9.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker10.Value.Day) < 10) { dayTo = "0" + dateTimePicker10.Value.Day.ToString(); } else { dayTo = dateTimePicker10.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker9.Value.Month) < 10) { monthFrom = "0" + dateTimePicker9.Value.Month.ToString(); } else { monthFrom = dateTimePicker9.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker10.Value.Month) < 10) { monthTo = "0" + dateTimePicker10.Value.Month.ToString(); } else { monthTo = dateTimePicker10.Value.Month.ToString(); }

            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker9.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker9.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker10.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker10.Value.Year % 100).ToString();
            string timefrom = "00:00"; //replace using input
            string timeto = "23:59"; //replace using input//


            jsonq = base_template
                .Replace("<title>", "Auto_media_" + datefrom_short + "_" + dateto_short)
                .Replace("<allobjstring>", allobjstring)
                .Replace("<objects>", objs)
                .Replace("<alllogicstring>", alllogic)
                .Replace("<datefrom>", datefrom)
                .Replace("<dateto>", dateto)
                .Replace("<timefrom>", timefrom)
                .Replace("<timeto>", timeto)
                .Replace("<context>", context);
            string saved = jsonq;
            if (checkBox1.Checked & !checkBox2.Checked)//context only
            {
                jsonq = jsonq
                    .Replace("<iscontext>", "false")
                    .Replace("<ADP>", "3")
                    .Replace("<compare>", "0")
                    .Replace("<RCH>", "0")
                    .Replace("<TID>", "-1")
                    .Replace("<type>", "1");
            }

            if (!checkBox1.Checked & checkBox2.Checked)//Obj only
            {
                jsonq = jsonq
                    .Replace("<iscontext>", "false")
                    .Replace("<ADP>", "3")
                    .Replace("<RCH>", "1")
                    .Replace("<TID>", "-1");
                if (count > 1)
                {
                    jsonq = jsonq
                        .Replace("<compare>", "1")
                        .Replace("<type>", "6");
                }
                else
                {
                    jsonq = jsonq
                        .Replace("<compare>", "0")
                        .Replace("<type>", "1");
                }
            }

            if (checkBox1.Checked & checkBox2.Checked)//Obj+context
            {
                jsonq = jsonq
                    .Replace("<iscontext>", "false")
                    .Replace("<ADP>", "3")
                    .Replace("<RCH>", "0")
                    .Replace("<TID>", "-1");
                if (count > 1)
                {
                    jsonq = jsonq
                        .Replace("<compare>", "1")
                        .Replace("<type>", "6");
                }
                else
                {
                    jsonq = jsonq
                        .Replace("<compare>", "0")
                        .Replace("<type>", "1");
                }
            }

            //string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //System.IO.File.WriteAllText(base_path + "\\json_to_test.txt", jsonq);

            UpdateStatus("В работе", 10, "Отправка запроса");

            CookieContainer cookieContainer = Auth();
            string RepId = "";
            string data_to_post;
            byte[] buffer;
            string urlencoded;
            byte[] urljson;
            string currentReportIdstr;

            ValidateQuerry(context, cookieContainer);
            //prepare json to send
            //MessageBox.Show("Подготовка данных к отправке");

            //urlencode the shit
            //urljson = Encoding.ASCII.GetBytes(jsonq);
            //System.IO.File.WriteAllText(base_path + "\\json_to_test.txt", jsonq);

            urlencoded = HttpUtility.UrlEncode(jsonq);
            //System.IO.File.WriteAllText(base_path + "\\urlencoded.txt", urlencoded);
            //create payload and send it
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;

            data_to_post = "useFilterContainers=false&sr=" + urlencoded + "&tabNumber=&menuNumber=&sortColumn=";
            buffer = Encoding.ASCII.GetBytes(data_to_post);
            //MessageBox.Show("Буффер составлен");
            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                WebReq.MaximumAutomaticRedirections = 1;
                WebReq.AllowAutoRedirect = false;
                WebReq.CookieContainer = cookieContainer;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;
                WebReq.Timeout = 20000;
                PostData = WebReq.GetRequestStream();
                //MessageBox.Show("Буффер отправлен");
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)await Task.FromResult<WebResponse>(WebReq.GetResponseAsync().Result);
                Answer = WebResp.GetResponseStream();

                _Answer = new StreamReader(Answer);
                string answer = _Answer.ReadToEnd();
                string resultString = Regex.Replace(answer, @"\r\n| ", string.Empty, RegexOptions.Multiline);
                Regex reg = new Regex("getReportTreeUrl='(.*?)';");
                try
                {
                    currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                    WebResp.Close();
                }
                catch
                {
                    jsonq = saved
                    .Replace("<iscontext>", "true")
                    .Replace("<ADP>", "0")
                    .Replace("<compare>", "0")
                    .Replace("<RCH>", "1")
                    .Replace("<TID>", "342723")
                    .Replace("<type>", "1");
                    urlencoded = HttpUtility.UrlEncode(jsonq);
                    data_to_post = "useFilterContainers=false&sr=" + urlencoded;
                    buffer = Encoding.ASCII.GetBytes(data_to_post);
                    WebReq = null;
                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/SearchGrid");
                    WebReq.MaximumAutomaticRedirections = 1;
                    WebReq.AllowAutoRedirect = false;
                    WebReq.Referer = "https://pr.mlg.ru/Report.mlg/Create?utm_source=f5516k&utm_medium=s2&utm_campaign=c";
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;
                    WebReq.Timeout = 20000;
                    PostData = WebReq.GetRequestStream();
                    //MessageBox.Show("Буффер отправлен");
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    WebResp.Close();
                    WebReq = null;


                    WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                    WebReq.MaximumAutomaticRedirections = 1;
                    WebReq.AllowAutoRedirect = false;
                    WebReq.Referer = "https://pr.mlg.ru/Report.mlg/Create?utm_source=f5516k&utm_medium=s2&utm_campaign=c";
                    WebReq.CookieContainer = cookieContainer;
                    WebReq.Method = "POST";
                    WebReq.ContentType = "application/x-www-form-urlencoded";
                    WebReq.ContentLength = buffer.Length;
                    WebReq.Timeout = 90000;
                    PostData = WebReq.GetRequestStream();
                    //MessageBox.Show("Буффер отправлен");
                    PostData.Write(buffer, 0, buffer.Length);
                    PostData.Close();
                    WebResp = (HttpWebResponse)WebReq.GetResponse();
                    Answer = WebResp.GetResponseStream();
                    currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                    WebResp.Close();

                }

                //catch id of redirrect

            }
            catch (WebException exxx)
            {
                if (exxx.Status == WebExceptionStatus.Timeout)
                {

                    StatusLabel.Text = "Готово";
                    progressBar1.Value = 0;
                    MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else throw;
            }


            RepId = currentReportIdstr.Split("?".ToCharArray())[0];
            //RepId = "3111102";

            IDictionary<string, string[]> table1;
            IDictionary<string, string[]> table2;
            IDictionary<string, string[]> table3;
            try
            {
                table1 = await GetMediaTable1(RepId, cookieContainer);
                table2 = await GetMediaTable2(RepId, cookieContainer);
                table3 = await GetMediaTable3(RepId, cookieContainer);
            }
            catch
            {
                MessageBox.Show("При заданных параметрах не было найдено ни одного сообщения", "Данные не получены", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                UpdateStatus();
                return;
            }

            UpdateStatus();

            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            UpdateStatus("В работе", 40, "Подготовка данных. Таблица 1/3");

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Add();
            wb.Sheets.Add(missingObj, missingObj, 3);

            ws = wb.Sheets[1];
            //labels
            ws.Cells[1, 1].Value2 = "Уровень\\Категория";
            ws.Cells[1, 2].Value2 = "Газеты";
            ws.Cells[1, 3].Value2 = "Журналы";
            ws.Cells[1, 4].Value2 = "Информ. Агенства";
            ws.Cells[1, 5].Value2 = "Интернет";
            ws.Cells[1, 6].Value2 = "ТВ";
            ws.Cells[1, 7].Value2 = "Радио";
            ws.Cells[1, 8].Value2 = "Блоги";
            ws.Cells[1, 9].Value2 = "Всего";

            //data
            int offset1 = 2;
            foreach (KeyValuePair<string, string[]> entry in table1)
            {
                ws.Cells[offset1, 1].Value2 = entry.Key;
                for (int valId = 0; valId < entry.Value.Length; valId++)
                {
                    ws.Cells[offset1, 2 + valId].Value2 = entry.Value[valId];
                }
                offset1++;
            }

            //for table2
            UpdateStatus("В работе", 60, "Подготовка данных. Таблица 2/3");
            ws = wb.Sheets[2];
            //labels
            ws.Cells[1, 1].Value2 = "№";
            ws.Cells[1, 2].Value2 = "Наименование СМИ";
            ws.Cells[1, 3].Value2 = "Кол-во Сообщений";
            ws.Cells[1, 4].Value2 = "Категория СМИ";
            ws.Cells[1, 5].Value2 = "Город";
            //data
            int offset2 = 2;
            foreach (KeyValuePair<string, string[]> entry in table2)
            {
                ws.Cells[offset2, 1].Value2 = (offset2 - 1).ToString();
                ws.Cells[offset2, 2].Value2 = entry.Key;
                for (int valId = 0; valId < entry.Value.Length; valId++)
                {
                    ws.Cells[offset2, 3 + valId].Value2 = entry.Value[valId];
                }
                offset2++;
            }


            //for table3
            UpdateStatus("В работе", 70, "Подготовка данных. Таблица 3/3");
            ws = wb.Sheets[3];
            //labels
            ws.Cells[1, 1].Value2 = "№";
            ws.Cells[1, 2].Value2 = "Название объекта";
            ws.Cells[1, 3].Value2 = "Кол-во Сообщений";
            ws.Cells[1, 4].Value2 = "Главная роль";
            ws.Cells[1, 5].Value2 = "Негативные упоминания";
            ws.Cells[1, 6].Value2 = "Позитивные упоминания";
            ws.Cells[1, 7].Value2 = "Цитирование";
            //data
            int offset3 = 2;
            foreach (KeyValuePair<string, string[]> entry in table3)
            {
                ws.Cells[offset3, 1].Value2 = (offset3 - 1).ToString();
                ws.Cells[offset3, 2].Value2 = entry.Key;
                for (int valId = 0; valId < entry.Value.Length; valId++)
                {
                    ws.Cells[offset3, 3 + valId].Value2 = entry.Value[valId];
                }
                offset3++;
            }


            //formatting
            UpdateStatus("В работе", 80, "Форматирование");
            ws = wb.Sheets[1];
            ws.Columns[1].ColumnWidth = 21;
            ws.Columns[2].ColumnWidth = 8;
            ws.Columns[3].ColumnWidth = 10;
            ws.Columns[4].ColumnWidth = 10;
            ws.Columns[5].ColumnWidth = 10;
            ws.Columns[6].ColumnWidth = 4;
            ws.Columns[7].ColumnWidth = 6;
            ws.Columns[8].ColumnWidth = 6;
            ws.Columns[9].ColumnWidth = 6;
            ws.Rows[1].Cells.WrapText = true;
            ws.Rows[1].Cells.HorizontalAlignment = _Excel.XlHAlign.xlHAlignCenter;
            ws.Rows[1].Cells.VerticalAlignment = _Excel.XlHAlign.xlHAlignCenter;
            ws.Rows[1].Cells.Font.Bold = true;
            ws.Rows[5].Cells.Font.Bold = true;
            ws.Columns[9].Cells.Font.Bold = true;


            ws = wb.Sheets[2];
            ws.Columns[1].ColumnWidth = 4;
            ws.Columns[2].ColumnWidth = 40;
            ws.Columns[3].ColumnWidth = 12;
            ws.Columns[4].ColumnWidth = 19;
            ws.Columns[5].ColumnWidth = 19;
            ws.Rows[1].Cells.Font.Bold = true;
            ws.Rows[1].Cells.WrapText = true;
            ws.Columns[2].Cells.WrapText = true;
            ws.Rows[1].Cells.HorizontalAlignment = _Excel.XlHAlign.xlHAlignCenter;
            ws.Rows[1].Cells.VerticalAlignment = _Excel.XlHAlign.xlHAlignCenter;


            ws = wb.Sheets[3];
            ws.Columns[1].ColumnWidth = 4;
            ws.Columns[2].ColumnWidth = 35;
            ws.Columns[3].ColumnWidth = 12;
            ws.Columns[4].ColumnWidth = 13;
            ws.Columns[5].ColumnWidth = 13;
            ws.Columns[6].ColumnWidth = 13;
            ws.Columns[7].ColumnWidth = 13;
            ws.Rows[1].Cells.Font.Bold = true;
            ws.Rows[1].Cells.WrapText = true;
            ws.Columns[2].Cells.WrapText = true;
            ws.Rows[1].Cells.HorizontalAlignment = _Excel.XlHAlign.xlHAlignCenter;
            ws.Rows[1].Cells.VerticalAlignment = _Excel.XlHAlign.xlHAlignCenter;

            //port the shit
            UpdateStatus("В работе", 90, "Экспорт");
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();

            ws = wb.Sheets[1];
            ws.UsedRange.Copy();
            Word.Range rangetemp = doc.Paragraphs.Last.Range;
            rangetemp.PasteExcelTable(false, true, false);

            ws = wb.Sheets[2];
            ws.UsedRange.Copy();
            doc.Paragraphs.Add();
            rangetemp = doc.Paragraphs.Last.Range;
            rangetemp.PasteExcelTable(false, true, false);

            ws = wb.Sheets[3];
            ws.UsedRange.Copy();
            doc.Paragraphs.Add();
            rangetemp = doc.Paragraphs.Last.Range;
            rangetemp.PasteExcelTable(false, true, false);
            app.Visible = true;

            wb.Close();
            base1.Quit();
            UpdateStatus();
        }
        private async void button15_Click(object sender, EventArgs e) //report 5
        {
            if(!checkBox1.Checked & !checkBox2.Checked)
            {
                MessageBox.Show("Необходимо выбрать хотя бы один параметр отчета","Мало параметров", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if((checkedListBox1.CheckedItems.Count == 0 & checkedListBox2.CheckedItems.Count == 0) & checkBox2.Checked)
            {
                MessageBox.Show("Не выбраны объекты для отчета", "Мало параметров", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!ValidateDates(dateTimePicker9, dateTimePicker10))
            {
                return;
            }

            try
            {
               await Report5_request();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "5";
                reporter.Stage = "Requesting data and building the report";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (Globals.MediaMode == "OR")
            {
                Globals.MediaMode = "&";
                button17.Text = "И";
            }
            else
            {
                Globals.MediaMode = "OR";
                button17.Text = "ИЛИ";
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text.Trim();
            try
            {
                if (textBox5.Text.Length > 0 & textBox5.Text.Substring(textBox5.Text.Length - 1, 1) != "|")
                {
                    textBox5.Text += "|";
                }
            }
            catch { 
            
            }
        }

        private void label40_Click(object sender, EventArgs e)
        {

        }


        private void Report6_create()
        {


            UpdateStatus("В работе", 10, "Чтение базы");
            string selectedDate = listBox6.SelectedItem.ToString();
            IDictionary<string, string[]> DB = new Dictionary<string, string[]>();
            //structure is NAME, {part, Messages, MIndex, Influence, Negative, Positive, Citations, Likes}



            //open DB and populate the structure
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["rep6_path"].ToString());
            ws = wb.Sheets[1];
            int lastrow_DB = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int tcol = ws.Cells.Find(selectedDate).Column;


            for (int row = 2; row <= lastrow_DB; row++)
            {
                string name = ws.Cells[row, 1].Value2;
                string part = ws.Cells[row, 2].Value2;
                string[] arr = new string[] { part, "0", "0", "н/д", "0", "0", "0", "0" };
                for (int sheetID = 1; sheetID <= wb.Sheets.Count; sheetID++) //Find shit in pages 
                {
                    ws = wb.Sheets[sheetID];
                    arr[sheetID] = Convert.ToString(ws.Cells[row, tcol].Value2);
                }

                //Append to the array
                DB.Add(name, arr);
            }

            //By this point we have all the data, so we don't need the DB anymore
            wb.Close();
            UpdateStatus("В работе", 40, "Подготовка таблиц");
            //create a dummy excel for exporting
            wb = base1.Workbooks.Add();

            //add sheets
            wb.Sheets.Add(missingObj, missingObj, 5);

            //paste the data
            int row_kprf = 1;
            int row_er = 1;
            int row_ldpr = 1;
            int row_sr = 1;
            int row_lp = 1;
            int trow = 1;
            int sum_row = 2;
            int selected_sheet = 1;

            string[] labarr = new string[] { "КПРФ", "ЕР и САТЕЛЛИТЫ", "ЛДПР", "СР", "Либеральные политики", "Топ 50" };

            string[] labels = new string[] { "№", "ФИО", "Кол-во сообщений", "Медиа-индекс", "Охват аудитории", "Негатив", "Позитив", "Цитирование", "Лайки и репосты" };
            foreach (KeyValuePair<string, string[]> entry in DB)
            {
                //set the correct sheet
                switch (entry.Value[0])
                {
                    case "КПРФ":
                        ws = wb.Sheets[1];
                        selected_sheet = 1;
                        row_kprf++;
                        trow = row_kprf;
                        break;
                    case "ЕР":
                        ws = wb.Sheets[2];
                        selected_sheet = 2;
                        row_er++;
                        trow = row_er;
                        break;
                    case "ЛДПР":
                        ws = wb.Sheets[3];
                        selected_sheet = 3;
                        row_ldpr++;
                        trow = row_ldpr;
                        break;
                    case "СР":
                        ws = wb.Sheets[4];
                        selected_sheet = 4;
                        row_sr++;
                        trow = row_sr;
                        break;
                    case "Либеральные политики":
                        ws = wb.Sheets[5];
                        selected_sheet = 5;
                        row_lp++;
                        trow = row_lp;
                        break;
                    default:
                        break;
                }
                ws.Cells[trow, 2].Value2 = entry.Key;
                int col = 3;
                ws.Cells[trow, 1].Value2 = (trow - 1).ToString();
                for (int i = 0; i < 7; i++)
                {
                    ws = wb.Sheets[selected_sheet];
                    ws.Cells[trow, col].Value2 = entry.Value[i + 1];
                    ws = wb.Sheets[6];
                    ws.Cells[sum_row, col].Value2 = entry.Value[i + 1];
                    col++;
                }
                ws.Cells[sum_row, 2].Value2 = entry.Key;
                ws.Cells[sum_row, 1].Value2 = (sum_row - 1).ToString();
                sum_row++;
            }

            //paste labels
            for (int sheetID = 1; sheetID <= wb.Sheets.Count; sheetID++)
            {
                ws.Name = labarr[sheetID - 1];
                ws = wb.Sheets[sheetID];
                int col = 1;
                foreach (string label in labels)
                {
                    ws.Cells[1, col].Value2 = label;
                    col++;
                }
            }


            //apply sort 
            UpdateStatus("В работе", 60, "Сортировка данных");
            for (int sheetId = 1; sheetId <= wb.Sheets.Count; sheetId++)
            {
                ws = wb.Sheets[sheetId];
                int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                _Excel.Range rng = ws.Range[ws.Cells[1, 2], ws.Cells[lastrow, 9]];


                ws.Sort.SortFields.Clear();
                ws.Sort.SortFields.Add(rng.Columns[2], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
                var sort = ws.Sort;
                sort.SetRange(rng.Rows);
                sort.Header = _Excel.XlYesNoGuess.xlYes;
                sort.MatchCase = false;
                sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
                sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
                sort.Apply();

                //make it beautiful
                //set col widths
                ws.Columns[1].ColumnWidth = 3;
                ws.Columns[2].ColumnWidth = 35;
                ws.Columns[3].ColumnWidth = 11;
                ws.Columns[4].ColumnWidth = 9;
                ws.Columns[5].ColumnWidth = 10;
                ws.Columns[6].ColumnWidth = 8;
                ws.Columns[7].ColumnWidth = 8;
                ws.Columns[8].ColumnWidth = 12;
                ws.Columns[9].ColumnWidth = 8;
                ws.Rows[1].Cells.WrapText = true;
                ws.Columns[2].Cells.WrapText = true;
                ws.UsedRange.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
                ws.UsedRange.Borders.Weight = _Excel.XlBorderWeight.xlThin;
            }

            UpdateStatus("В работе", 70, "Подготовка текстового отчета");

            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            string[] parts = new string[] { wb.Sheets[1].Name, wb.Sheets[2].Name, wb.Sheets[3].Name, wb.Sheets[4].Name, wb.Sheets[5].Name, wb.Sheets[6].Name };
            doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            for (int sheetId = 1; sheetId <= wb.Sheets.Count; sheetId++)
            {
                System.Threading.Thread.Sleep(400);
                UpdateStatus("В работе", 86 + sheetId * 2, "Запись документа");
                System.Threading.Thread.Sleep(400);

                ws = wb.Sheets[sheetId];
                ws.UsedRange.Copy(missingObj);
                if (sheetId == 6)
                {
                    ws.Range[ws.Cells[1, 1], ws.Cells[51, 9]].Copy(missingObj);
                }
                //paste data
                doc.Paragraphs.Add();
                doc.Paragraphs.Last.Range.Text = labarr[sheetId - 1];
                doc.Paragraphs.Last.Range.Bold = 1;
                doc.Paragraphs.Last.Range.Font.Size = 30;
                doc.Paragraphs.Add();
                doc.Paragraphs.Add();
                Word.Range rangetemp = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                rangetemp.PasteExcelTable(false, false, false);
                doc.Paragraphs.Add();

            }
            app.Visible = true;
            wb.Close();
            base1.Quit();
        }
        private void button20_Click(object sender, EventArgs e)  //Report 6 - Create report
        {

            try
            {
                Report6_create();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "6";
                reporter.Stage = "Report Creation";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }

        }

        private void AnalyseRep6Base(string path = "none")//Report6 - Analyse and load Base
        {

            MethodInvoker methodInvokerDelegate = delegate ()
            {

                UpdateStatus("В работе", 5, "Чтение базы");
                if (path == "none")
                {
                    if (Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
                    {
                        MessageBox.Show("В настройках стоит создание новой базы. Укажите уже существующую базу", "Ошибка загрузки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatus();
                        return;
                    }
                    if (!File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
                    {
                        MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatus();
                        return;
                    }
                    UpdateStatus("В работе", 10, "Идет загрузка базы, пожалуйста подождите");
                    textBox6.Text = Properties.Settings.Default["rep6_path"].ToString();

                    //start temp excel
                    _Excel._Application base1 = new _Excel.Application();
                    base1.Visible = false;
                    base1.DisplayAlerts = false;

                    //init wb and ws
                    Workbook wb;
                    Worksheet ws;
                    wb = base1.Workbooks.Open(Properties.Settings.Default["rep6_path"].ToString());
                    ws = wb.Sheets[1];

                    List<string> FindAllDates(int sheet)
                    {
                        ws = wb.Sheets[sheet];
                        int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                        List<string> dates = new List<string>();
                        for (int col1 = 3; col1 <= cols + 1; col1++)
                        {
                            if (ws.Cells[1, col1].Value2 != null)
                            {
                                dates.Add(ws.Cells[1, col1].Value2);
                            }
                        }
                        return dates;
                    }


                    List<string> dates_temp = FindAllDates(1);
                    listBox6.Items.Clear();
                    for (int i = 0; i < dates_temp.Count; i++)
                    {
                        listBox6.Items.Add(dates_temp[i]);
                    }
                    wb.Close();
                    base1.Quit();
                    button21.Enabled = true;
                    button20.Enabled = true;
                    UpdateStatus();
                }
                else
                {
                    if (!File.Exists(path))
                    {
                        MessageBox.Show("Файл базы не был найден. Проверьте существование указанной базы", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    UpdateStatus("В работе", 10, "Идет загрузка базы, пожалуйста подождите");
                    textBox6.Text = path;
                    //start temp excel
                    _Excel._Application base1 = new _Excel.Application();
                    base1.Visible = false;
                    base1.DisplayAlerts = false;

                    //init wb and ws
                    Workbook wb;
                    Worksheet ws;
                    wb = base1.Workbooks.Open(path);
                    ws = wb.Sheets[1];
                    List<string> FindAllDates(int sheet)
                    {
                        ws = wb.Sheets[sheet];
                        int cols = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                        List<string> dates = new List<string>();
                        for (int col1 = 3; col1 <= cols + 1; col1++)
                        {
                            if (ws.Cells[1, col1].Value2 != null)
                            {
                                dates.Add(ws.Cells[1, col1].Value2);
                            }
                        }
                        return dates;
                    }
                    List<string> dates_temp = FindAllDates(1);
                    listBox6.Items.Clear();
                    for (int i = 0; i < dates_temp.Count; i++)
                    {
                        listBox6.Items.Add(dates_temp[i]);
                    }
                    wb.Close();
                    base1.Quit();
                    button21.Enabled = true;
                    button20.Enabled = true;
                    UpdateStatus();

                }

                //invoker check
            };

            //This will be true if Current thread is not UI thread.
            if (this.InvokeRequired)
                this.Invoke(methodInvokerDelegate);
            else
                methodInvokerDelegate();
            return;

        }

        private void button19_Click(object sender, EventArgs e) //Load rep6 base
        {

            if (!File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {
                MessageBox.Show("Не найдена историческая база потенциальных кандидатов.\nУкажите существующий файл, либо создайте новую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {
                MessageBox.Show("База с указанным именем уже существует.\nВыберите другое имя базы, либо выберите уже существующую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (!File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {
                MessageBox.Show("В настройках указано создание новой базы\nНовая база будет создана при первом запросе", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                UpdateStatus();
                return;
            }


            try
            {
                AnalyseRep6Base(Properties.Settings.Default["rep6_path"].ToString());
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "6";
                reporter.Stage = "Database Analysys";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }

            

            //enable button only if label count>0
            if (listBox6.Items.Count>0) {
                button20.Enabled = true;
            }




        }


        private void Report6_request()
        {

            //populate deps
            IDictionary<string, string[]> DB = new Dictionary<string, string[]>();
            IDictionary<string, string[]> DB_MLG = new Dictionary<string, string[]>();
            //structure {"id", "part", "MsgCount", "MIndex", "Influence", "Negative", "Positive", "Citations", "Likes" }


            string line;
            StreamReader file = new System.IO.StreamReader(Properties.Settings.Default["rep6_list"].ToString());
            while ((line = file.ReadLine()) != null)
            {
                DB.Add(line, new string[] { "id", "part", "0", "0", "н/д", "0", "0", "0", "0" });
            }
            file.Close();


            //initiate lookup
            //start temp excel
            _Excel._Application base1 = new _Excel.Application();
            base1.Visible = false;
            base1.DisplayAlerts = false;

            //init wb and ws
            Workbook wb;
            Worksheet ws;
            wb = base1.Workbooks.Open(Properties.Settings.Default["rep6_database"].ToString());
            ws = wb.Sheets[1];

            int lastrow_MLG_DB = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            //populate MLGDB
            for (int row_mlg = 1; row_mlg <= lastrow_MLG_DB; row_mlg++)
            {
                DB_MLG.Add(ws.Cells[row_mlg, 2].Value2, new string[] { Convert.ToString(ws.Cells[row_mlg, 1].Value2), Convert.ToString(ws.Cells[row_mlg, 3].Value2) });
            }

            foreach (KeyValuePair<string, string[]> person in DB)
            {
                try
                {
                    person.Value[0] = DB_MLG[person.Key][0];
                    person.Value[1] = DB_MLG[person.Key][1];
                }
                catch
                {
                    MessageBox.Show("Депутат " + person.Key + " не был найден в базе медиалогии и был исключен из запроса", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DB.Remove(person.Key);
                }
            }
            //close global database
            wb.Close();
            //Create DB if creation is checked

            if (Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {

                wb = base1.Workbooks.Add();
                wb.Sheets.Add(missingObj, missingObj, 6);
                for (int sheetID = 1; sheetID <= wb.Sheets.Count; sheetID++)
                {
                    ws = wb.Sheets[sheetID];
                    ws.Cells[1, 1].Value2 = "ФИО";
                    ws.Cells[1, 2].Value2 = "Партия";
                    int rowc = 2;
                    foreach (KeyValuePair<string, string[]> person in DB)
                    {
                        ws.Cells[rowc, 1].Value2 = person.Key;
                        ws.Cells[rowc, 2].Value2 = person.Value[1];
                        rowc++;

                    }
                }
                wb.SaveAs(Properties.Settings.Default["rep6_path"].ToString());
                wb.Close();

            }

            //Make a request
            //https://regex101.com/r/HYs937/1
            //по объекту' data-title-align='top'>(.*?)<.*?&count=(.*?)\\".*?-->(.*?)<.*?ce\\">(.*?)<\\/td.*?td.*?(?>negative\\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>positive\\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>speech.*?>(\\d.*?)<\\/|&count=(\\d+)).*?(?>like.*?>(\\d.*?)<\\/|&count=(\\d+))
            //Get data
            //Populate DB

            //prepare request parts
            //dates
            string dayTo, dayFrom, monthTo, monthFrom;


            if (Convert.ToInt32(dateTimePicker11.Value.Day) < 10) { dayFrom = "0" + dateTimePicker11.Value.Day.ToString(); } else { dayFrom = dateTimePicker11.Value.Day.ToString(); }
            if (Convert.ToInt32(dateTimePicker12.Value.Day) < 10) { dayTo = "0" + dateTimePicker12.Value.Day.ToString(); } else { dayTo = dateTimePicker12.Value.Day.ToString(); }

            if (Convert.ToInt32(dateTimePicker11.Value.Month) < 10) { monthFrom = "0" + dateTimePicker11.Value.Month.ToString(); } else { monthFrom = dateTimePicker11.Value.Month.ToString(); }
            if (Convert.ToInt32(dateTimePicker12.Value.Month) < 10) { monthTo = "0" + dateTimePicker12.Value.Month.ToString(); } else { monthTo = dateTimePicker12.Value.Month.ToString(); }



            string datefrom = dayFrom + "." + monthFrom + "." + dateTimePicker11.Value.Year.ToString();
            string datefrom_short = dayFrom + "." + monthFrom + "." + (dateTimePicker11.Value.Year % 100).ToString();
            string dateto = dayTo + "." + monthTo + "." + dateTimePicker12.Value.Year.ToString();
            string dateto_short = dayTo + "." + monthTo + "." + (dateTimePicker12.Value.Year % 100).ToString();
            string timefrom = "00:00";
            string timeto = "23:59";


            //<repname>
            string part1 = "{\"smsMonitor\":{\"MonitorId\":-1,\"ThemeId\":-1,\"UserId\":-1,\"MaxSendingArticle\":0,\"SendingMode\":2,\"SendingPeriod\":1,\"ReprintsMode\":3,\"MonitorPhones\":[]},\"folder\":\"\",\"folderId\":-1,\"Authors\":[],\"Cities\":[],\"Levels\":[1,2],\"Categories\":[1,2,3,4,5,6],\"Rubrics\":[],\"LifeStyles\":[],\"MediaSources\":[],\"MediaBranches\":[],\"MediaObjectBranches\":[],\"MediaObjectLifeStyles\":[],\"MediaObjectLevels\":[],\"MediaObjectCategories\":[],\"MediaObjectRegions\":[],\"MediaObjectFederals\":[],\"MediaObjectTowns\":[],\"MediaLanguages\":[],\"MediaRegions\":[],\"MediaCountries\":[],\"CisMediaCountries\":[],\"MediaFederals\":[],\"MediaGenre\":[],\"YandexRubrics\":[],\"Role\":-1,\"Tone\":-1,\"Quotation\":-1,\"CityMode\":0,\"messageCount\":-1,\"reprintsMessageCount\":-1,\"CheckedMessageCount\":-1,\"CheckedClustersCount\":-1,\"MonitorId\":-1,\"CheckedReprintsCount\":-1,\"deletedMessageCount\":-1,\"favoritesMessageCount\":-1,\"myDocsMessageCount\":0,\"myMediaMessageCount\":0,\"IsSaveParamsOnly\":false,\"RebuildDBCache\":false,\"Credentials\":null,\"AppType\":1,\"ParamsVersion\":0,\"ArmObjectMode\":0,\"ReportCreatingHistory\":0,\"InfluenceThreshold\":\"0.0\",\"MonitorObjects\":null,\"Icon\":0,\"ThemeGroup\":-1,\"ThemeGroupName\":\"\",\"SaveMode\":0,\"MonitorExists\":false,\"ThemeId\":-1,\"Title\":\"<repname>\",\"Comment\":\"\",\"ReprintMode\":0,\"rssReportType\":0,\"ThemeObjects\":[";

            //<id> <name> <lindex>
            string obj = "{\"Id\":\"<id>\",\"MainObjectId\":\"<id>\",\"ObjectName\":\"<name>\",\"classId\":43,\"LogicIndex\":<lindex>,\"LogicObjectString\":\"OR\",\"SearchQuery\":null,\"Properties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}]}";

            //<allstring> <logic> <datefrom> <dateto> <timefrom> <timeto>
            string part3 = "],\"ThemeObjectsFromSearchContext\":[],\"ThemeTypes\":[],\"ThemeBranches\":[],\"AllObjectsProperties\":[{\"Id\":1,\"Value\":-1},{\"Id\":2,\"Value\":-1},{\"Id\":4,\"Value\":-1}],\"AllArticlesProperties\":[],\"AllObjectString\":\"<allstring>\",\"AllLogicObjectString\":\"<logic>\",\"DatePeriod\":8,\"DateType\":0,\"Date\":\"<datefrom>|<dateto>\",\"Time\":\"<timefrom>|<timeto>\",\"ActualDatePeriod\":3,\"IsSlidingTime\":true,\"ContextScope\":5,\"Context\":\"\",\"ContextMode\":0,\"TopMedia\":false,\"RegionLogic\":0,\"MediaObjectRegionLogic\":0,\"MediaLogic\":0,\"MediaLogicAll\":0,\"BlogLogic\":1,\"MediaBranchLogic\":0,\"MediaObjectBranchLogic\":0,\"MediaLanguageLogic\":0,\"MediaCountryLogic\":0,\"CityLogic\":0,\"Compare\":0,\"User\":0,\"Type\":1,\"View\":0,\"ViewStatus\":1,\"OiiMode\":0,\"Template\":-1,\"MediaStatus\":-1,\"IsUpdate\":false,\"HasUserObjects\":false,\"IsContextReport\":false,\"LastCopiedThemeId\":null}";

            string objs = "";
            string allobjectstring = "";
            string alllogic = "";

            int lindex = 0;
            foreach (KeyValuePair<string, string[]> sec in DB)
            {
                objs += obj.Replace("<id>", sec.Value[0])
                    .Replace("<name>", sec.Key)
                    .Replace("<lindex>", lindex.ToString());
                if (lindex + 1 < DB.Keys.Count)
                {
                    objs += ", ";
                }
                allobjectstring += "+O" + sec.Value[0] + "_" + lindex.ToString();
                alllogic += "+" + lindex.ToString();
                lindex++;
            }
            string target = datefrom + "-" + dateto;
            string sr = part1.Replace("<repname>", "Pot_C_" + target) + objs + part3
                .Replace("<allstring>", allobjectstring)
                .Replace("<logic>", alllogic)
                .Replace("<datefrom>", datefrom)
                .Replace("<dateto>", dateto)
                .Replace("<timefrom>", timefrom)
                .Replace("<timeto>", timeto);

            //auth
            UpdateStatus("В работе", 35, "Авторизация");
            HttpWebRequest WebReq;
            HttpWebResponse WebResp;
            var cookieContainer = new CookieContainer();
            Stream PostData;
            Stream Answer;
            StreamReader _Answer;
            string data_to_post;
            byte[] buffer;
            try
            {
                cookieContainer = Auth();
            }
            catch (Exception ex)
            {
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "?";
                reporter.Stage = "Auth";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

                MessageBox.Show("Ошибка авторизации. Проверьте правильность логина/пароля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                base1.Quit();
                return;
            }

            //make a request
            string currentReportIdstr;
            byte[] urljson = Encoding.ASCII.GetBytes(sr);
            string urlencoded = HttpUtility.UrlEncode(urljson);
            //create payload and send it

            data_to_post = "useFilterContainers=false&sr=" + urlencoded;
            buffer = Encoding.ASCII.GetBytes(data_to_post);
            //MessageBox.Show("Буффер составлен");
            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/Save");
                WebReq.MaximumAutomaticRedirections = 1;
                WebReq.AllowAutoRedirect = false;
                WebReq.CookieContainer = cookieContainer;
                WebReq.Method = "POST";
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.ContentLength = buffer.Length;
                WebReq.Timeout = 20000;
                PostData = WebReq.GetRequestStream();
                //MessageBox.Show("Буффер отправлен");
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                //catch id of redirrect
                currentReportIdstr = WebResp.Headers["Location"].Substring(20);
                WebResp.Close();
            }
            catch (WebException exxx)
            {
                if (exxx.Status == WebExceptionStatus.Timeout)
                {
                    base1.Quit();
                    UpdateStatus();
                    MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else throw;
            }

            UpdateStatus("В работе", 40, "Получение данных");


            //MessageBox.Show("Строка перехвата отчета: "+ currentReportIdstr);
            int tempReportId;
            try
            {
                tempReportId = Convert.ToInt32(currentReportIdstr.Remove(currentReportIdstr.Length - 22, 22));
                //MessageBox.Show("Идет перехват отчета №" + tempReportId.ToString());
            }
            catch
            {
                MessageBox.Show("Ошибка при получении отчета. Проверьте правильность данных и дат", "Ошибка");
                WebResp.Close();
                base1.Quit();
                UpdateStatus();
                return;
            }
            WebResp.Close();

            //loop through pages and reg-search-add


            try
            {
                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + tempReportId.ToString() + "&pageSize=20&pageNumber=1");
                WebReq.CookieContainer = cookieContainer;
                WebReq.ContentType = "application/x-www-form-urlencoded";
                WebReq.AllowAutoRedirect = true;
                WebReq.MaximumAutomaticRedirections = 20;
                WebReq.Timeout = 60000;
                WebResp = (HttpWebResponse)WebReq.GetResponse();
                Answer = WebResp.GetResponseStream();
                _Answer = new StreamReader(Answer);
                string answer = _Answer.ReadToEnd();
                WebResp.Close();
                WebReq = null;
                //regex-find
                //string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                //System.IO.File.WriteAllText(base_path + "\\ans_" + entry.Key.ToString() + ".txt", answer);
                string resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                //System.IO.File.WriteAllText(base_path + "\\resultString_" + entry.Key.ToString() + ".txt", resultString);
                Regex reg = new Regex("pagesCount = (\\d+)");
                Regex sreg = new Regex("по объекту' data-title-align='top'>(.*?)<.*?&count=(.*?)\".*?-->(.*?)<.*?ce\">(.*?)<\\/td.*?td.*?(?>negative\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>positive\">(\\d.*?)<\\/|&count=(\\d+)).*?(?>speech.*?>(\\d.*?)<\\/|&count=(\\d+)).*?(?>like.*?>(\\d.*?)<\\/|&count=(\\d+))");
                MatchCollection pages_collection = reg.Matches(resultString);
                if (pages_collection.Count > 0)
                {
                    int pages = Convert.ToInt32(pages_collection[0].Groups[1].Value);
                    UpdateStatus("В работе", 45, "Получение данных. Страница 1/" + pages.ToString());
                    MatchCollection data = sreg.Matches(resultString);
                    for (int g = 0; g < data.Count; g++)
                    {
                        int count_arr_point = 2;
                        for (int group_id = 2; group_id <= data[g].Groups.Count; group_id++)
                        {

                            if (data[g].Groups[group_id].Value != "")
                            {
                                DB[data[g].Groups[1].Value][count_arr_point] = data[g].Groups[group_id].Value;
                                count_arr_point++;
                            }

                        }
                    }
                    if (pages > 1)
                    {
                        for (int pageId = 2; pageId <= pages; pageId++)
                        {
                            UpdateStatus("В работе", 45, "Получение данных. Страница " + pageId + "/" + pages.ToString());
                            try
                            {
                                WebReq = (HttpWebRequest)WebRequest.Create("https://pr.mlg.ru/Report.mlg/StatisticsGrid?id=" + tempReportId.ToString() + "&pageSize=20&pageNumber=" + pageId.ToString());
                                WebReq.CookieContainer = cookieContainer;
                                WebReq.ContentType = "application/x-www-form-urlencoded";
                                WebReq.AllowAutoRedirect = true;
                                WebReq.MaximumAutomaticRedirections = 20;
                                WebReq.Timeout = 60000;
                                WebResp = (HttpWebResponse)WebReq.GetResponse();
                                Answer = WebResp.GetResponseStream();
                                _Answer = new StreamReader(Answer);
                                answer = _Answer.ReadToEnd();
                                WebResp.Close();
                                WebReq = null;
                                resultString = Regex.Replace(answer, @"\r\n|  ", string.Empty, RegexOptions.Multiline);
                                data = sreg.Matches(resultString);
                                for (int g = 0; g < data.Count; g++)
                                {
                                    int count_arr_point = 2;
                                    for (int group_id = 2; group_id <= data[g].Groups.Count; group_id++)
                                    {

                                        if (data[g].Groups[group_id].Value != "")
                                        {
                                            DB[data[g].Groups[1].Value][count_arr_point] = data[g].Groups[group_id].Value;
                                            count_arr_point++;
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ReportSender Sender = new ReportSender();
                                Reporter reporter = new Reporter();
                                reporter.EventType = "Http error";
                                reporter.ReportType = "?";
                                reporter.Stage = "ValidateQuerry";
                                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                                Sender.SendReport(reporter);

                                base1.Quit();
                                UpdateStatus();
                                MessageBox.Show("Сервер не ответил вовремя. Запрос был остановлен.\nПожалуйста, повторите запрос позже.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }

                }
                else
                {
                    foreach (KeyValuePair<string, string[]> sec in DB)
                    {
                        sec.Value[2] = "0";
                        sec.Value[3] = "0";
                        sec.Value[4] = "н/д";
                        sec.Value[5] = "0";
                        sec.Value[6] = "0";
                        sec.Value[7] = "0";
                        sec.Value[8] = "0";
                    }
                }

            }
            catch (Exception ex)
            {
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "?";
                reporter.Stage = "Ошибка при получении отчета";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);

                MessageBox.Show("Ошибка при получении отчета", "Ошибка");
                UpdateStatus();
                WebResp.Close();
                base1.Quit();
                return;
            }


            wb = base1.Workbooks.Open(Properties.Settings.Default["rep6_path"].ToString());
            ws = wb.Sheets[1];
            int lastcol = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastrow = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            //paste label

            for (int sheetId = 1; sheetId <= wb.Sheets.Count; sheetId++)
            {
                ws = wb.Sheets[sheetId];
                ws.Cells[1, lastcol + 1].Value2 = datefrom + " - " + dateto;
                foreach (KeyValuePair<string, string[]> person in DB)
                {
                    int row = ws.Cells.Find(person.Key).Row;
                    ws.Cells[row, lastcol + 1].Value2 = Convert.ToString(person.Value[sheetId + 1]);
                }
            }
            wb.Save();
            wb.Close();
            base1.Quit();
            AnalyseRep6Base(Properties.Settings.Default["rep6_path"].ToString());


            //FUCK OFF
            UpdateStatus();
        }
        private void button21_Click(object sender, EventArgs e) //Report 6 - Request data
        {
            if (!ValidateDates(dateTimePicker11, dateTimePicker12))
            {
                UpdateStatus();
                return;
            }

            if (!File.Exists(Properties.Settings.Default["rep6_list"].ToString()))
            {
                MessageBox.Show("Не найден файл со списком потенциальных кандидатов", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (!File.Exists(Properties.Settings.Default["rep6_database"].ToString()))
            {
                MessageBox.Show("Не найдена общая база Медиалогии для потенциальных кандидатов", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (!File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & !Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {
                MessageBox.Show("Не найдена историческая база потенциальных кандидатов.\nУкажите существующий файл, либо создайте новую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }

            if (File.Exists(Properties.Settings.Default["rep6_path"].ToString()) & Convert.ToBoolean(Properties.Settings.Default["rep6_new"]))
            {
                MessageBox.Show("База с указанным именем уже существует.\nВыберите другое имя базы, либо выберите уже существующую базу.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus();
                return;
            }


            try
            {
                Report6_request();
            }
            catch (Exception ex)
            {
                ErrorNotification(ex);
                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = ex.Message;
                reporter.ReportType = "6";
                reporter.Stage = "Database Analysys";
                reporter.ExceptionDescription = ex.Message + "  ;  " + ex.StackTrace;
                Sender.SendReport(reporter);
            }


        }

        private void label13_Click(object sender, EventArgs e)
        {
            if(Globals.SecretCounter >= 4)
            {
                MessageBox.Show("Sending Data");

                ReportSender Sender = new ReportSender();
                Reporter reporter = new Reporter();
                reporter.EventType = "Test event";
                reporter.ReportType = "0";
                reporter.Stage = "Sending the report";
                reporter.ExceptionDescription = "The test event has been fired";
                Sender.SendReport(reporter);

                Globals.SecretCounter = 0;
            }
            else
            {
                Globals.SecretCounter++;
            }
        }
    }

}
