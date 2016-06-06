using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using HMCSynthLib;
using OfficeOpenXml;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using Agilent.CommandExpert.ScpiNet.Ag34410_2_35;
using Agilent.CommandExpert.ScpiNet.AgMXG_A_01_80;
using Agilent.CommandExpert.ScpiNet.Ag90x0_SA_A_08_03;

namespace Mix_Fu
{
    public class Instrument
    {
        public string Location { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
    }
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<Instrument> instruments = new List<Instrument>();
        DataTable tbl;
        //string xslx_path;
        string xlsx_path = "";
        int delay = 300;
        decimal attenuation = 30;
        decimal maxfreq = 26500;
        string measure_type = "DSB down mixer";

        public MainWindow()
        {
            InitializeComponent();
            comboBox_instruments_LO.ItemsSource = instruments;
            comboBox_instruments_RF.ItemsSource = instruments;
            comboBox_instruments_IF.ItemsSource = instruments;
        }

        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            measure_type = listBox.SelectedItem.ToString();
        }

        public static void SearchInstruments(List<Instrument> instruments, int max, int gpib)
        {
            instruments.Clear();
            string idn;
            for (int i = 0; i <= max; i++)
            {
                string location = "GPIB" + gpib.ToString() + "::" + i.ToString() + "::INSTR";
                try
                {
                    using (Ag34410 mm = new Ag34410(location))
                    {
                        mm.SCPI.IDN.Query(out idn);
                        string[] idn_cut = idn.Split(',');
                        instruments.Add(new Instrument { Location = location, Name = idn_cut[1], FullName = idn });
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                    //instruments.Add(new Instrument { Location = location, Name = "no_responce", FullName = "no_responce" });
                }
            }
        }

        private void b_search_clk(object sender, RoutedEventArgs e)
        {
            SearchInstruments(instruments, Convert.ToInt32(textBox_number.Text), Convert.ToInt32(textBox_number_GPIB.Text));
            comboBox_instruments_LO.Items.Refresh();
            comboBox_instruments_RF.Items.Refresh();
            comboBox_instruments_IF.Items.Refresh();
            //tabControl.Items.Refresh();
        }

        private void run_quiery_Click(object sender, RoutedEventArgs e)
        {
            Instrument IF_instrument = new Instrument();
            IF_instrument = (Instrument)comboBox_instruments_IF.SelectedItem;

            using (Ag90x0_SA sa = new Ag90x0_SA(IF_instrument.Location))
            {
                string answer = "";
                string question = textBox_quiery.Text;
                try { sa.Transport.Query.Invoke(question, out answer); }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                textBox_answer.Text = answer;
            }
        }

        private void run_command_Click(object sender, RoutedEventArgs e)
        {
            Instrument IF_instrument = new Instrument();
            IF_instrument = (Instrument)comboBox_instruments_IF.SelectedItem;

            using (Ag90x0_SA sa = new Ag90x0_SA(IF_instrument.Location))
            {
                string question = textBox_quiery.Text;
                try { sa.Transport.Command.Invoke(question); }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public static void send(string location, string com)
        {
            using (AgSCPI99 instrument = new AgSCPI99(location))
            {
                try { instrument.Transport.Command.Invoke(com); }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public static string quiery(string location, string com)
        {
            string answer = "";
            using (AgSCPI99 instrument = new AgSCPI99(location))
            {
                try { instrument.Transport.Query.Invoke(com, out answer); }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                return answer;
            }
        }

        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;//!!! WOW
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        private void b_import_xlsx_Clk(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "xlsx файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            if ((bool)openFileDialog.ShowDialog())
            {
                try
                {
                    if ((xlsx_path = openFileDialog.FileName) != null)
                    {
                        //MessageBox.Show("файл выбран");
                        //FileInfo xlsx_file = new FileInfo(xlsx_path);
                        tbl = GetDataTableFromExcel(xlsx_path);

                        dataGrid.ItemsSource = tbl.AsDataView();
                    }
                    else { return; }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: файл не открыт. Сообщение об ошибке: " + ex.Message);
                }
            }
            //FileInfo xlsx_file = new FileInfo(xlsx_path);
        }

        private void b_save_xlsx_clk(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            //saveFileDialog.FileName = xlsx_path;
            saveFileDialog.Filter = "Таблица (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            //saveFileDialog.ShowDialog();
            if ((bool)saveFileDialog.ShowDialog())
            {
                //xlsxfile_path = saveFileDialog.FileName;
                FileInfo xlsx_file = new FileInfo(saveFileDialog.FileName);
                using (ExcelPackage pck = new ExcelPackage(xlsx_file))
                {
                    tbl = ((DataView)dataGrid.ItemsSource).ToTable();
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("1");
                    ws.Cells["A1"].LoadFromDataTable(tbl, true);

                    foreach (var cell in ws.Cells)
                    {
                        cell.Value = cell.Value.ToString().Replace('.', ',');
                        try { cell.Value = Convert.ToDecimal(cell.Value); }
                        catch
                        {
                            cell.Style.Font.Bold = true;
                        };
                    }

                    pck.Save();
                }
            }
        }

        private void textBox_delay_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                delay = Convert.ToInt32(textBox_delay.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Введите задержку в мсек");
            }
        }

        private void textBox_maxfreq_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                decimal.TryParse(textBox_maxfreq.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out maxfreq);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Введите максимальную рабочую частоту, на которой смогут работать и генератор, и анализатор спектра, в МГц");
            }
        }

        private void textBox_attenuation_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                decimal.TryParse(textBox_attenuation.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out attenuation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Введите значение входной аттенюации анализатора спектра в дБ");
            }
        }

        private void b_cal_LO_run(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboBox_instruments_LO.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("SOUR:POW " + "-35"));
            send(LO, ("OUTP:STAT ON"));

            if (measure_type == "Multiplier x2")
            {
                MessageBox.Show("Calibration error. Multiplier don't need LO.");
            }

            else
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_freq = row["FLO"].ToString();
                    string t_pow_goal = row["PLO-GOAL"].ToString();

                    if (!(string.IsNullOrEmpty(t_pow_goal)) && !(string.IsNullOrEmpty(t_freq)))
                    {
                        string t_pow = "";
                        string t_pow_temp = "";
                        decimal t_freq_dec = 0;
                        decimal t_pow_dec = 0;
                        decimal t_pow_goal_dec = 0;
                        decimal t_pow_temp_dec = 0;
                        decimal err = 1;

                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.00", CultureInfo.InvariantCulture);
                        decimal.TryParse(t_pow_goal.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);
                        t_pow_temp = t_pow_goal.Replace(',', '.');
                        t_pow_temp_dec = t_pow_goal_dec;

                        send(LO, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        int count = 0;
                        int delay_temp = delay;
                        while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10)
                        {
                            send(LO, ("SOUR:POW " + t_pow_temp));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture);
                            t_pow_temp_dec += err;
                            t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["PLO"] = t_pow_temp.Replace('.', ',');
                            count++;
                            delay += 50;
                        }
                        delay = delay_temp;

                        //if (Math.Abs(err) > 10) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                    }

                }

            }
            send(LO, ("OUTP:STAT OFF"));
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_cal_IN_run(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            //send(RF, ("SOUR:POW " + "-100"));
            send(RF, ("OUTP:STAT ON"));

            if (measure_type == "DSB down mixer")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_freq = row["FRF"].ToString();
                    string t_pow_goal = row["PRF-GOAL"].ToString();

                    if (!(string.IsNullOrEmpty(t_pow_goal)) && !(string.IsNullOrEmpty(t_freq)))
                    {
                        string t_pow = "";
                        string t_pow_temp = "";
                        decimal t_freq_dec = 0;
                        decimal t_pow_dec = 0;
                        decimal t_pow_goal_dec = 0;
                        decimal t_pow_temp_dec = 0;
                        decimal err = 1;

                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.00", CultureInfo.InvariantCulture);
                        decimal.TryParse(t_pow_goal.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);
                        t_pow_temp = t_pow_goal.Replace(',', '.');
                        t_pow_temp_dec = t_pow_goal_dec;

                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        int count = 0;
                        int delay_temp = delay;
                        while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10)
                        {
                            send(RF, ("SOUR:POW " + t_pow_temp));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture);
                            t_pow_temp_dec += err;
                            t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["PRF"] = t_pow_temp.Replace('.', ',');
                            count++;
                            delay += 50;
                        }
                        delay = delay_temp;
                        //if (Math.Abs(err) > 10) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                    }
                }
            }

            else if (measure_type == "DSB down mixer")
            {
                MessageBox.Show("Calibration error. Function is not realized in program yet.");
            }

            else if (measure_type == "SSB down mixer")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_freq = row["FLSB"].ToString();
                    string t_pow_goal = row["PLSB-GOAL"].ToString();

                    if (!(string.IsNullOrEmpty(t_pow_goal)) && !(string.IsNullOrEmpty(t_freq)))
                    {
                        string t_pow = "";
                        string t_pow_temp = "";
                        decimal t_freq_dec = 0;
                        decimal t_pow_dec = 0;
                        decimal t_pow_goal_dec = 0;
                        decimal t_pow_temp_dec = 0;
                        decimal err = 1;

                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.00", CultureInfo.InvariantCulture);
                        decimal.TryParse(t_pow_goal.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);
                        t_pow_temp = t_pow_goal.Replace(',', '.');
                        t_pow_temp_dec = t_pow_goal_dec;

                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        int count = 0;
                        int delay_temp = delay;
                        while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10)
                        {
                            send(RF, ("SOUR:POW " + t_pow_temp));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture);
                            t_pow_temp_dec += err;
                            t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["PLSB"] = t_pow_temp.Replace('.', ',');
                            count++;
                            delay += 50;
                        }
                        delay = delay_temp;
                        //if (Math.Abs(err) > 10) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                    }

                    t_freq = row["FUSB"].ToString();
                    t_pow_goal = row["PUSB-GOAL"].ToString();

                    if (!(string.IsNullOrEmpty(t_pow_goal)) && !(string.IsNullOrEmpty(t_freq)))
                    {
                        string t_pow = "";
                        string t_pow_temp = "";
                        decimal t_freq_dec = 0;
                        decimal t_pow_dec = 0;
                        decimal t_pow_goal_dec = 0;
                        decimal t_pow_temp_dec = 0;
                        decimal err = 1;

                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.00", CultureInfo.InvariantCulture);
                        decimal.TryParse(t_pow_goal.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);
                        t_pow_temp = t_pow_goal.Replace(',', '.');
                        t_pow_temp_dec = t_pow_goal_dec;

                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        int count = 0;
                        int delay_temp = delay;
                        while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10)
                        {
                            send(RF, ("SOUR:POW " + t_pow_temp));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture);
                            t_pow_temp_dec += err;
                            t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["PUSB"] = t_pow_temp.Replace('.', ',');
                            count++;
                            delay += 50;
                        }
                        delay = delay_temp;
                        //if (Math.Abs(err) > 10) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                    }

                }
            }

            else if (measure_type == "SSB up mixer")
            {
                MessageBox.Show("Calibration error. Function is not realized in program yet.");
            }

            else if (measure_type == "Multiplier x2")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_freq = row["FH1"].ToString();
                    string t_pow_goal = row["PIN-GOAL"].ToString();

                    if (!(string.IsNullOrEmpty(t_pow_goal)) && !(string.IsNullOrEmpty(t_freq)))
                    {
                        string t_pow = "";
                        string t_pow_temp = "";
                        decimal t_freq_dec = 0;
                        decimal t_pow_dec = 0;
                        decimal t_pow_goal_dec = 0;
                        decimal t_pow_temp_dec = 0;
                        decimal err = 1;

                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.00", CultureInfo.InvariantCulture);
                        decimal.TryParse(t_pow_goal.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);
                        t_pow_temp = t_pow_goal.Replace(',', '.');
                        t_pow_temp_dec = t_pow_goal_dec;

                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        int count = 0;
                        int delay_temp = delay;
                        while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10)
                        {
                            send(RF, ("SOUR:POW " + t_pow_temp));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture);
                            t_pow_temp_dec += err;
                            t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["PIN-GEN"] = t_pow_temp.Replace('.', ',');
                            count++;
                            delay += 50;
                        }
                        delay = delay_temp;
                        //if (Math.Abs(err) > 10) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                    }
                }
            }


            send(RF, ("OUTP:STAT OFF"));
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_cal_OUT_run(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            //send(RF, ("SOUR:POW " + "-100"));
            send(RF, ("OUTP:STAT ON"));

            if (measure_type == "DSB down mixer")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_pow = "-20";
                    string t_pow_temp = "-20";
                    string t_freq = "";
                    decimal t_freq_dec = 0;
                    decimal t_pow_dec = -20;
                    decimal t_pow_goal_dec = -20;
                    decimal err = 1;
                    string err_str = "";

                    send(RF, ("SOUR:POW " + t_pow_temp));

                    //ATT-IF
                    t_freq = row["FIF"].ToString();
                    err_str = row["ATT-IF"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-IF"] = err_str.Replace('.', ',');
                    }

                    //ATT-RF
                    t_freq = row["FRF"].ToString();
                    err_str = row["ATT-RF"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-RF"] = err_str.Replace('.', ',');
                    }

                    //ATT-LO
                    t_freq = row["FLO"].ToString();
                    err_str = row["ATT-LO"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-LO"] = err_str.Replace('.', ',');
                    }
                    //if (Math.Abs(err) > 3) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                }
            }

            else if (measure_type == "DSB down mixer")
            {
                MessageBox.Show("Calibration error. Function is not realized in program yet.");
            }

            else if (measure_type == "SSB down mixer" | measure_type == "SSB up mixer")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_pow = "-20";
                    string t_pow_temp = "-20";
                    string t_freq = "";
                    decimal t_freq_dec = 0;
                    decimal t_pow_dec = -20;
                    decimal t_pow_goal_dec = -20;
                    decimal err = 1;
                    string err_str = "";

                    send(RF, ("SOUR:POW " + t_pow_temp));

                    //ATT-IF
                    t_freq = row["FIF"].ToString();
                    err_str = row["ATT-IF"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-IF"] = err_str.Replace('.', ',');
                    }

                    //ATT-LSB
                    t_freq = row["FLSB"].ToString();
                    err_str = row["ATT-LSB"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-LSB"] = err_str.Replace('.', ',');
                    }

                    //ATT-USB
                    t_freq = row["FUSB"].ToString();
                    err_str = row["ATT-USB"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-USB"] = err_str.Replace('.', ',');
                    }

                    //ATT-LO
                    t_freq = row["FLO"].ToString();
                    err_str = row["ATT-LO"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)) && err_str != "-")
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);
                        t_freq = (t_freq_dec * 1000000000).ToString("0.000", CultureInfo.InvariantCulture);
                        send(RF, ("SOUR:FREQ " + t_freq));
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        err = t_pow_goal_dec - t_pow_dec;
                        err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                        row["ATT-LO"] = err_str.Replace('.', ',');
                    }
                    //if (Math.Abs(err) > 3) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                }

            }

            else if (measure_type == "Multiplier x2")
            {
                foreach (DataRow row in tbl.Rows)
                {
                    string t_pow = "-20";
                    string t_pow_temp = "-20";
                    string t_freq = "";
                    decimal t_freq_dec = 0;
                    decimal t_pow_dec = -20;
                    decimal t_pow_goal_dec = -20;
                    decimal err = 1;
                    string err_str = "";

                    send(RF, ("SOUR:POW " + t_pow_temp));

                    t_freq = row["FH1"].ToString();
                    if (!(string.IsNullOrEmpty(t_freq)))
                    {
                        decimal.TryParse(t_freq.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_dec);

                        //ATT-H1
                        if (t_freq_dec * 1000000000 < maxfreq * 1000000)
                        {
                            t_freq = (t_freq_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                            send(RF, ("SOUR:FREQ " + t_freq));
                            send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                            send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                            row["ATT-H1"] = err_str.Replace('.', ',');
                        }

                        //ATT-H2
                        if (t_freq_dec * 2000000000 < maxfreq * 1000000)
                        {
                            t_freq = (t_freq_dec * 2000000000).ToString("0", CultureInfo.InvariantCulture);
                            send(RF, ("SOUR:FREQ " + t_freq));
                            send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                            send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                            row["ATT-H2"] = err_str.Replace('.', ',');
                        }

                        //ATT-H3
                        if (t_freq_dec * 3000000000 < maxfreq * 1000000)
                        {
                            t_freq = (t_freq_dec * 3000000000).ToString("0", CultureInfo.InvariantCulture);
                            send(RF, ("SOUR:FREQ " + t_freq));
                            send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                            send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                            row["ATT-H3"] = err_str.Replace('.', ',');
                        }

                        //ATT-H4
                        if (t_freq_dec * 4000000000 < maxfreq * 1000000)
                        {
                            t_freq = (t_freq_dec * 4000000000).ToString("0", CultureInfo.InvariantCulture);
                            send(RF, ("SOUR:FREQ " + t_freq));
                            send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                            send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq));
                            Thread.Sleep(delay);
                            t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                            err = t_pow_goal_dec - t_pow_dec;
                            err_str = err.ToString("0.000", CultureInfo.InvariantCulture);
                            row["ATT-H4"] = err_str.Replace('.', ',');
                        }
                    }
                    //if (Math.Abs(err) > 3) { MessageBox.Show("Calibration error. Please check amplitude reference level on alayzer"); }
                }
            }

            send(RF, ("OUTP:STAT OFF"));
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_measure_mix_DSB_down_clk(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboBox_instruments_LO.SelectedItem).Location;
            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));

            send(RF, ("OUTP:STAT ON"));

            foreach (DataRow row in tbl.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_RF = row["PRF"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_RF)))
                {
                    string att_str = "";
                    string t_pow = "";
                    string t_freq_LO = "";
                    string t_freq_RF = "";
                    string t_freq_IF = "";
                    decimal t_freq_LO_dec = 0;
                    decimal t_freq_RF_dec = 0;
                    decimal t_freq_IF_dec = 0;
                    decimal t_pow_IF_dec = 0;
                    decimal t_pow_RF_dec = 0;
                    decimal t_pow_LO_dec = 0;
                    decimal t_pow_goal_RF_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PRF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_RF_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_RF_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_RF = (t_freq_RF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(RF, ("SOUR:FREQ " + t_freq_RF));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(RF, ("SOUR:POW " + t_pow_RF.Replace(',', '.')));

                    // IF
                    send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                    send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                    Thread.Sleep(delay);
                    t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                    decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                    decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                    t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["POUT-IF"] = t_pow.Replace('.', ',');
                    decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_RF_dec;
                    string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["CONV"] = t_conv.Replace('.', ',');

                    //ISO-RF
                    att_str = row["ATT-RF"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_RF));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_RF));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-RF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_RF_dec);
                        t_pow = t_pow_RF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-RF"] = t_pow.Replace('.', ',');
                        decimal t_iso_rf_dec = t_pow_goal_RF_dec - att - t_pow_RF_dec;
                        string t_iso_rf = t_iso_rf_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-RF"] = t_iso_rf.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LO_dec);
                        t_pow = t_pow_LO_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LO"] = t_pow.Replace('.', ',');
                        decimal t_iso_lo_dec = t_pow_goal_LO_dec - att - t_pow_LO_dec;
                        string t_iso_lo = t_iso_lo_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LO"] = t_iso_lo.Replace('.', ',');
                    }
                }
            }
            send(LO, "OUTP:STAT OFF");
            send(RF, "OUTP:STAT OFF");
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_measure_mix_SSB_down_clk(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboBox_instruments_LO.SelectedItem).Location;
            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));

            send(RF, ("OUTP:STAT ON"));

            foreach (DataRow row in tbl.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_LSB = row["PLSB"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_LSB)))
                {
                    string att_str = "";
                    string t_pow = "";
                    string t_freq_LO = "";
                    string t_freq_LSB = "";
                    string t_freq_IF = "";
                    decimal t_freq_LO_dec = 0;
                    decimal t_freq_LSB_dec = 0;
                    decimal t_freq_IF_dec = 0;
                    decimal t_pow_IF_dec = 0;
                    decimal t_pow_LSB_dec = 0;
                    decimal t_pow_LO_dec = 0;
                    decimal t_pow_goal_LSB_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PLSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LSB_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LSB_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_LSB = (t_freq_LSB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(RF, ("SOUR:FREQ " + t_freq_LSB));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(RF, ("SOUR:POW " + t_pow_LSB.Replace(',', '.')));

                    // IF
                    send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                    send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                    Thread.Sleep(delay);
                    t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                    decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                    decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                    t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["POUT-IF"] = t_pow.Replace('.', ',');
                    decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_LSB_dec;
                    string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["CONV-LSB"] = t_conv.Replace('.', ',');

                    //ISO-LSB
                    att_str = row["ATT-LSB"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_LSB));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_LSB));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LSB_dec);
                        t_pow = t_pow_LSB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LSB"] = t_pow.Replace('.', ',');
                        decimal t_iso_LSB_dec = t_pow_goal_LSB_dec - att - t_pow_LSB_dec;
                        string t_iso_LSB = t_iso_LSB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LSB"] = t_iso_LSB.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LO_dec);
                        t_pow = t_pow_LO_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LO"] = t_pow.Replace('.', ',');
                        decimal t_iso_lo_dec = t_pow_goal_LO_dec - att - t_pow_LO_dec;
                        string t_iso_lo = t_iso_lo_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LO"] = t_iso_lo.Replace('.', ',');
                    }
                }
            }

            foreach (DataRow row in tbl.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_USB = row["PUSB"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_USB)))
                {
                    string att_str = "";
                    string t_pow = "";
                    string t_freq_LO = "";
                    string t_freq_USB = "";
                    string t_freq_IF = "";
                    decimal t_freq_LO_dec = 0;
                    decimal t_freq_USB_dec = 0;
                    decimal t_freq_IF_dec = 0;
                    decimal t_pow_IF_dec = 0;
                    decimal t_pow_USB_dec = 0;
                    decimal t_pow_LO_dec = 0;
                    decimal t_pow_goal_USB_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PUSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_USB_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_USB_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_USB = (t_freq_USB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(RF, ("SOUR:FREQ " + t_freq_USB));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(RF, ("SOUR:POW " + t_pow_USB.Replace(',', '.')));

                    // IF
                    send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                    send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                    Thread.Sleep(delay);
                    t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                    decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                    decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                    t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["POUT-IF"] = t_pow.Replace('.', ',');
                    decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_USB_dec;
                    string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["CONV-USB"] = t_conv.Replace('.', ',');

                    //ISO-LSB
                    att_str = row["ATT-USB"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_USB));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_USB));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-USB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_USB_dec);
                        t_pow = t_pow_USB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-USB"] = t_pow.Replace('.', ',');
                        decimal t_iso_USB_dec = t_pow_goal_USB_dec - att - t_pow_USB_dec;
                        string t_iso_USB = t_iso_USB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-USB"] = t_iso_USB.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LO_dec);
                        t_pow = t_pow_LO_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LO"] = t_pow.Replace('.', ',');
                        decimal t_iso_lo_dec = t_pow_goal_LO_dec - att - t_pow_LO_dec;
                        string t_iso_lo = t_iso_lo_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LO"] = t_iso_lo.Replace('.', ',');
                    }
                }
            }

            send(LO, "OUTP:STAT OFF");
            send(RF, "OUTP:STAT OFF");
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();

        }

        private void b_measure_mix_SSB_up_clk(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboBox_instruments_LO.SelectedItem).Location;
            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));

            send(RF, ("OUTP:STAT ON"));

            foreach (DataRow row in tbl.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_RF = row["PRF"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_RF)))
                {
                    string att_str = "";
                    string t_pow = "";
                    string t_freq_LO = "";
                    string t_freq_RF = "";
                    string t_freq_IF = "";
                    decimal t_freq_LO_dec = 0;
                    decimal t_freq_RF_dec = 0;
                    decimal t_freq_IF_dec = 0;
                    decimal t_pow_IF_dec = 0;
                    decimal t_pow_RF_dec = 0;
                    decimal t_pow_LO_dec = 0;
                    decimal t_pow_goal_RF_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PRF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_RF_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_RF_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_RF = (t_freq_RF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(RF, ("SOUR:FREQ " + t_freq_RF));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(RF, ("SOUR:POW " + t_pow_RF.Replace(',', '.')));

                    // IF
                    send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                    send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                    Thread.Sleep(delay);
                    t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                    decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                    decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                    t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["POUT-IF"] = t_pow.Replace('.', ',');
                    decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_RF_dec;
                    string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["CONV"] = t_conv.Replace('.', ',');

                    //ISO-RF
                    att_str = row["ATT-RF"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_RF));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_RF));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-RF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_RF_dec);
                        t_pow = t_pow_RF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-RF"] = t_pow.Replace('.', ',');
                        decimal t_iso_rf_dec = t_pow_goal_RF_dec - att - t_pow_RF_dec;
                        string t_iso_rf = t_iso_rf_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-RF"] = t_iso_rf.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LO_dec);
                        t_pow = t_pow_LO_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LO"] = t_pow.Replace('.', ',');
                        decimal t_iso_lo_dec = t_pow_goal_LO_dec - att - t_pow_LO_dec;
                        string t_iso_lo = t_iso_lo_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LO"] = t_iso_lo.Replace('.', ',');
                    }
                }
            }
            send(LO, "OUTP:STAT OFF");
            send(RF, "OUTP:STAT OFF");
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_measure_mult_clk(object sender, RoutedEventArgs e)
        {
            tbl = ((DataView)dataGrid.ItemsSource).ToTable();

            string RF = ((Instrument)comboBox_instruments_RF.SelectedItem).Location;
            string IF = ((Instrument)comboBox_instruments_IF.SelectedItem).Location;

            send(IF, ":CAL:AUTO OFF");
            send(IF, ":SENS:FREQ:SPAN 10000000");
            send(IF, ":CALC:MARK1:MODE POS");
            send(IF, ":POW:ATT " + attenuation.ToString());
            //send(IF, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(RF, ("OUTP:STAT ON"));

            foreach (DataRow row in tbl.Rows)
            {
                string t_pow_gen = row["PIN-GEN"].ToString();
                string att = row["ATT-H1"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_gen)) && !(string.IsNullOrEmpty(att)))
                {
                    string t_pow = row["PIN-GOAL"].ToString();
                    send(RF, ("SOUR:POW " + t_pow_gen.Replace(',', '.')));

                    decimal t_pow_goal_dec = 0;
                    decimal.TryParse(t_pow.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);

                    //harm1
                    string t_freq_H1 = row["FH1"].ToString();
                    decimal t_freq_H1_dec = 0;
                    decimal.TryParse(t_freq_H1.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_H1_dec);
                    t_freq_H1 = (t_freq_H1_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    send(RF, ("SOUR:FREQ " + t_freq_H1));
                    if (t_freq_H1_dec * 1000 < maxfreq)
                    {
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_H1));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_H1));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H1_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H1_dec);
                        t_pow = t_pow_H1_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H1"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H1_dec = t_pow_H1_dec + att_dec - t_pow_goal_dec;
                        string conv_H1 = conv_H1_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H1"] = conv_H1.Replace('.', ',');
                    }

                    //harm2
                    decimal t_freq_H2_dec = t_freq_H1_dec * 2;
                    att = row["ATT-H2"].ToString();
                    if (!(string.IsNullOrEmpty(att)) && t_freq_H2_dec * 1000 < maxfreq)
                    {
                        string t_freq_H2 = (t_freq_H2_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_H2));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_H2));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H2_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H2_dec);
                        t_pow = t_pow_H2_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H2"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H2_dec = t_pow_H2_dec + att_dec - t_pow_goal_dec;
                        string conv_H2 = conv_H2_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H2"] = conv_H2.Replace('.', ',');
                    }

                    //harm3
                    decimal t_freq_H3_dec = t_freq_H1_dec * 3;
                    att = row["ATT-H3"].ToString();
                    if (!(string.IsNullOrEmpty(att)) && t_freq_H3_dec * 1000 < maxfreq)
                    {
                        string t_freq_H3 = (t_freq_H3_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_H3));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_H3));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H3_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H3_dec);
                        t_pow = t_pow_H3_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H3"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H3_dec = t_pow_H3_dec + att_dec - t_pow_goal_dec;
                        string conv_H3 = conv_H3_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H3"] = conv_H3.Replace('.', ',');
                    }

                    //harm4
                    decimal t_freq_H4_dec = t_freq_H1_dec * 4;
                    att = row["ATT-H4"].ToString();
                    if (!(string.IsNullOrEmpty(att)) && t_freq_H4_dec * 1000 < maxfreq)
                    {
                        string t_freq_H4 = (t_freq_H4_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(IF, (":SENSe:FREQuency:RF:CENTer " + t_freq_H4));
                        send(IF, (":CALCulate:MARKer1:X:CENTer " + t_freq_H4));
                        Thread.Sleep(delay);
                        t_pow = quiery(IF, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H4_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H4_dec);
                        t_pow = t_pow_H4_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H4"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H4_dec = t_pow_H4_dec + att_dec - t_pow_goal_dec;
                        string conv_H4 = conv_H4_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H4"] = conv_H4.Replace('.', ',');
                    }
                }
            }

            send(RF, "OUTP:STAT OFF");
            send(IF, ":CAL:AUTO ON");
            dataGrid.ItemsSource = tbl.AsDataView();
        }

        private void b_measure_clk(object sender, RoutedEventArgs e)
        {

            if (measure_type == "DSB down mixer")
            {
                MessageBox.Show("Function is not realized yet");
            }

            else if (measure_type == "DSB up mixer")
            {
                MessageBox.Show("Function is not realized yet");
            }

            else if (measure_type == "SSB down mixer")
            {
                MessageBox.Show("Function is not realized yet");
            }

            else if (measure_type == "SSB up mixer")
            {
                MessageBox.Show("Function is not realized yet");
            }

            else if (measure_type == "Multiplier x2")
            {
                MessageBox.Show("Function is not realized yet");
            }

        }
    }
}
