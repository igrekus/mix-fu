using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using Agilent.CommandExpert.ScpiNet.Ag34410_2_35;
using Agilent.CommandExpert.ScpiNet.Ag90x0_SA_A_08_03;

namespace Mix_Fu
{
    public struct Constants {
        public const string decimalFormat = "0.00";
        public const int GHz = 1000000000;
        public const int MHz = 1000000;
    }

    enum MeasureMode : byte {
        modeDSBDown = 0,
        modeDSBUp,
        modeSSBDown,
        modeSSBUp,
        modeMultiplier
    };

    public struct ParameterStruct {
        public string colFreq;
        public string colPow;
        public string colPowGoal;
    }

    public class Instrument {
        public string Location { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public override string ToString() {
            return base.ToString() + ": " + "loc: " + Location +
                                            " name:" + Name +
                                            " fname:" + FullName;
        }
    }

    public class CalibrationPoint {
        public string freq { get; set; } = "";
        public string pow { get; set; } = "";
        public string calPow { get; set; } = "";

        public decimal freqD { get; set; } = 0;
        public decimal powD { get; set; } = 0;
        public decimal calPowD { get; set; } = 0;
        public decimal error { get; set; } = 0;

        public override string ToString() {
            return base.ToString() + ": " + "freq=" + freqD.ToString() + 
                                            " pow=" + powD.ToString() + 
                                            " calPow=" + calPowD.ToString() + 
                                            " err=" + error.ToString();
        }

        public override bool Equals(object obj) {
            return this.Equals(obj as CalibrationPoint);
        }

        // TODO: modify for decimal parameters
        public bool Equals(CalibrationPoint rhs) {
            if (Object.ReferenceEquals(rhs, null)) {
                return false;
            }
            if (Object.ReferenceEquals(this, rhs)) {
                return true;
            }
            if (this.GetType() != rhs.GetType()) {
                return false;
            }
            return (freq == rhs.freq) && (pow == rhs.pow);
        }

        public override int GetHashCode() {
            return freq.GetHashCode() + pow.GetHashCode();
        }

        public static bool operator ==(CalibrationPoint lhs, CalibrationPoint rhs) {
            if (Object.ReferenceEquals(lhs, null)) {
                if (Object.ReferenceEquals(rhs, null)) {
                    return true;
                }
                return false;
            }
            return lhs.Equals(rhs);
        }

        public static bool operator !=(CalibrationPoint lhs, CalibrationPoint rhs) {
            return !(lhs == rhs);
        }
    }

    public class MeasurePoint {
        public List<string> freqs { get; set; }
        public List<string> powers { get; set; }
        public List<string> freqs_meas { get; set; }
        public List<string> powers_meas { get; set; }
    }
    
    public partial class MainWindow : Window {

        #region regDataMembers

        List<Instrument> listInstruments = new List<Instrument>();
        // List<CalPoint> listCalData = new List<CalPoint>();
        //        List<MeasPoint> listMeasData = new List<MeasPoint>();

        DataTable dataTable;
        string xlsx_path = "";
        int delay = 300;
        decimal attenuation = 30;
        decimal maxfreq = 26500;
        decimal span = 10*Constants.MHz;
        bool verbose = false;
        MeasureMode measureMode = MeasureMode.modeDSBDown;

        Task searchTask;
        Task calibrationTask;

        InstrumentManager instrumentManager = null;

        Dictionary<MeasureMode, ParameterStruct> inParameters = null;
        ParameterStruct loParameters;

    #endregion regDataMembers

    public MainWindow()
        {
            instrumentManager = new InstrumentManager(log);

            inParameters = new Dictionary<MeasureMode, ParameterStruct>();
            inParameters.Add(MeasureMode.modeDSBDown,    new ParameterStruct { colFreq = "FRF", colPow = "PRF", colPowGoal = "PRF-GOAL" });
            inParameters.Add(MeasureMode.modeDSBUp,      new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" });
            inParameters.Add(MeasureMode.modeSSBDown,    new ParameterStruct { colFreq = "FUSB", colPow = "PUSB", colPowGoal = "PUSB-GOAL" });
            inParameters.Add(MeasureMode.modeSSBUp,      new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" });
            inParameters.Add(MeasureMode.modeMultiplier, new ParameterStruct { colFreq = "FH1", colPow = "PIN", colPowGoal = "PIN-GOAL" });
            loParameters = new ParameterStruct { colFreq = "FLO", colPow = "PLO", colPowGoal = "PLO-GOAL" };

            InitializeComponent();

            comboLO.ItemsSource = listInstruments;
            comboIN.ItemsSource = listInstruments;
            comboOUT.ItemsSource = listInstruments;

            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        }

        #region regUiEvents

        // options
        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            measureMode = (MeasureMode)listBox.SelectedIndex;
            //log("mode: " + measureMode.ToString() + "|" + ((int)measureMode).ToString());
            //log("modestr: " + listBox.SelectedItem.ToString());
        }

        private void LogVerboseToggled(object sender, RoutedEventArgs e) {
            verbose = LogVerbose.IsChecked ?? false;
        }

        private void textBox_delay_TextChanged(object sender, TextChangedEventArgs e) {
            //log("text: " + tmptext);
            try {
                delay = Convert.ToInt32(textBox_delay.Text);
                instrumentManager.delay = delay;
            }
            catch (Exception ex) {
                MessageBox.Show("Введите задержку в мсек");
            }
        }

        private void textBox_maxfreq_TextChanged(object sender, TextChangedEventArgs e) {
            if (!decimal.TryParse(textBox_maxfreq.Text.Replace(',', '.'), 
                                  NumberStyles.Any, CultureInfo.InvariantCulture, out maxfreq)) {
                MessageBox.Show("Введите максимальную рабочую частоту в МГц, на которой смогут работать и генератор, и анализатор спектра");
            } else {
                maxfreq *= Constants.MHz;
                instrumentManager.maxfreq = maxfreq;
            }
        }

        private void textBox_span_TextChanged(object sender, TextChangedEventArgs e) {
            if (!decimal.TryParse(textBox_span.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out span)) {
                MessageBox.Show("Введите требуемое значение Span анализатора спектра в МГц");
            } else { 
                span = span * Constants.MHz;
                instrumentManager.span = span;
            }
        }

        private void textBox_attenuation_TextChanged(object sender, TextChangedEventArgs e) {
            if (!decimal.TryParse(textBox_attenuation.Text.Replace(',', '.'), 
                                  NumberStyles.Any, CultureInfo.InvariantCulture, out attenuation)) {
                MessageBox.Show("Введите значение входной аттенюации анализатора спектра в дБ");
            } else {
                instrumentManager.attenuation = attenuation;
            }
        }

        // comboboxes
        private void comboIN_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            instrumentManager.m_IN = (Instrument)((ComboBox)sender).SelectedItem;
            log(instrumentManager.m_IN.ToString(), true);
        }

        private void comboOUT_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            instrumentManager.m_OUT = (Instrument)((ComboBox)sender).SelectedItem;
            log(instrumentManager.m_OUT.ToString(), true);
        }

        private void comboLO_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            instrumentManager.m_LO = (Instrument)((ComboBox)sender).SelectedItem;
            log(instrumentManager.m_LO.ToString(), true);
        }

        // misc buttons
        private async void btnSearchClicked(object sender, RoutedEventArgs e) {
            if (searchTask != null && !searchTask.IsCompleted) {
                MessageBox.Show("Instrument search is already running.");
                return;
            }
            int max_port = Convert.ToInt32(textBox_number_maxport.Text);
            int gpib = Convert.ToInt32(textBox_number_GPIB.Text);

            listInstruments.Clear();

            searchTask = Task.Factory.StartNew(
                () => instrumentManager.searchInstruments(listInstruments, max_port, gpib));
            await searchTask;

            listInstruments.Add(new Instrument { Location = "IN_location", Name = "IN", FullName = "IN at IN_location" });
            listInstruments.Add(new Instrument { Location = "OUT_location", Name = "OUT", FullName = "OUT at OUT_location" });
            listInstruments.Add(new Instrument { Location = "LO_location", Name = "LO", FullName = "LO at LO_location" });

            comboIN.Items.Refresh();
            comboOUT.Items.Refresh();
            comboLO.Items.Refresh();

            comboIN.SelectedIndex = 2;
            comboOUT.SelectedIndex = 1;
            comboLO.SelectedIndex = 0;
        }

        private void btnRunQueryClicked(object sender, RoutedEventArgs e) {
            if (comboOUT.SelectedIndex == -1) {
                MessageBox.Show("Error: no OUT instrument set");
                return;
            }

            // TODO: move query syntax into InstrumentManager class
            string question = textBox_query.Text;
            string answer = "";
            log(">>> query: " + question);
            try {
                answer = instrumentManager.query(instrumentManager.m_OUT.Location, question);
            }
            catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message);
                log("error: " + ex.Message);
                answer = "error querying instrument";
            }
            log("> " + answer);
        }

        private void btnRunCommandClicked(object sender, RoutedEventArgs e) {
            if (comboOUT.SelectedIndex == -1) {
                MessageBox.Show("Error: no OUT instrument set");
                return;
            }

            string command = textBox_query.Text;
            log(">>> c: " + command);
            try {
                instrumentManager.send(instrumentManager.m_OUT.Location, command);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
                log("error: " + ex.Message);
            }
        }

        // data buttons
        private void btnImportXlsxClicked(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "xlsx файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            if ((bool)openFileDialog.ShowDialog()) {
                try {
                    if ((xlsx_path = openFileDialog.FileName) != null) {
                        dataTable = getDataTableFromExcel(xlsx_path);
                        dataGrid.ItemsSource = dataTable.AsDataView();
                    } else { return; }
                }
                catch (Exception ex) {
                    MessageBox.Show("Error: can't open file, check log");
                    log("error: " + ex.Message);
                }
            }
        }

        private void btnSaveXlsxClicked(object sender, RoutedEventArgs e) {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Таблица (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            if ((bool)saveFileDialog.ShowDialog()) {
                FileInfo xlsx_file = new FileInfo(saveFileDialog.FileName);
                using (ExcelPackage package = new ExcelPackage(xlsx_file)) {
                    dataTable = ((DataView)dataGrid.ItemsSource).ToTable();
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("1");
                    ws.Cells ["A1"].LoadFromDataTable(dataTable, true);

                    foreach (var cell in ws.Cells) {
                        cell.Value = cell.Value.ToString().Replace('.', ',');
                        try {
                            cell.Value = Convert.ToDecimal(cell.Value);
                        }
                        catch {
                            cell.Style.Font.Bold = true;
                        };
                    }

                    package.Save();
                }
            }
        }

        // calibration buttons
        private void btnCalibrateInClicked(object sender, RoutedEventArgs e) {
            // TODO: check for option input errors
            if (!canCalibrateIN_OUT()) {
                MessageBox.Show("Error: check log");
                return;
            }
            calibrateIN();
        }

        private void btnCalibrateOutClicked(object sender, RoutedEventArgs e) {
            // TODO: check for option input errors
            if (!canCalibrateIN_OUT()) {
                MessageBox.Show("Error: check log");
                return;
            }
            calibrateOUT();
        }

        private void btnCalibrateLoClicked(object sender, RoutedEventArgs e) {
            if (!canCalibrateLO()) {
                MessageBox.Show("Error: check log");
                return;
            }
            if (measureMode == MeasureMode.modeMultiplier) {
                MessageBox.Show("Error: check log");
                log("error: multiplier doesn't need LO");
                return;
            }
            calibrateLO();
        }

        // measure button
        private void btnMeasureClicked(object sender, RoutedEventArgs e) {
            if (!canMeasure()) {
                MessageBox.Show("Error: check log");
                return;
            }
            switch (measureMode) {
            case MeasureMode.modeDSBDown:
                measure_mix_DSB_down();
                break;
            case MeasureMode.modeDSBUp:
                measure_mix_DSB_up();
                break;
            case MeasureMode.modeSSBDown:
                measure_mix_SSB_down();
                break;
            case MeasureMode.modeSSBUp:
                measure_mix_SSB_up();
                break;
            case MeasureMode.modeMultiplier:
                measure_mult();
                break;
            default:
                return;
            }
        }

        #endregion regUiEvents

        #region regUtility

        public void log(string mes, bool onlyVerbose = false) {
            if (!onlyVerbose | onlyVerbose & verbose) {
                Dispatcher.Invoke((Action)delegate () {
                    textLog.Text += "\n" + "[" + DateTime.Now.ToShortTimeString() + "]: " + mes;
                    scrollviewer.ScrollToBottom();
                });
            }
        }

        public object cell2dec(string value)
        {
            decimal a;
            try {
                a = decimal.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture);
            }
            catch (Exception ex) {
                log("error: decimal parse error:" + ex.Message);
                return value;
            }
            return a;
        }

        #endregion regUtility

        #region regInstrumentManager

        public void send(string location, string com) {
            try {
                using (AgSCPI99 instrument = new AgSCPI99(location)) {
                    instrument.Transport.Command.Invoke(com);
                }
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public string query(string location, string question) {
            string answer = "";
            try {
                using (AgSCPI99 instrument = new AgSCPI99(location)) {
                    instrument.Transport.Query.Invoke(question, out answer);
                }
            }
            catch (Exception ex) {
                throw ex;
            }
            return answer;
        }

        #endregion regInstrumentManager

        #region regDataManager

        public static DataTable getDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var package = new OfficeOpenXml.ExcelPackage()) {
                using (var stream = File.OpenRead(path)) {
                    package.Load(stream);
                }
                var ws = package.Workbook.Worksheets.First();
                DataTable dataTable = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column]) {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;   //!!! WOW
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++) {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dataTable.Rows.Add();
                    foreach (var cell in wsRow) {
                        //row[cell.Start.Column - 1] = cell.Text;
                        row[cell.Start.Column - 1] = cell.Value.ToString().Replace(".", ",");
                    }
                }
                return dataTable;
            }
        }

        #endregion regDataManager

        #region regCalibrationManager

        // TODO: move to a module (instrumentManager?)
        // TODO: use instrumentManager properties as IN OUT
        public bool canCalibrateIN_OUT() {
            if (dataTable == null) {
                log("error: no table open");
                return false;
            }
            if (comboIN.SelectedIndex == -1) {
                log("error: no IN instrument selected");
                return false;
            }
            if (comboOUT.SelectedIndex == -1) {
                log("error: no OUT instrument selected");
                return false;
            }
            if ((calibrationTask != null) && (!calibrationTask.IsCompleted)) {
                log("error: calibration is already running");
                return false;
            }
            return true;
        }

        public bool canCalibrateLO() {
            if (dataTable == null) {
                log("error: no table open");
                return false;
            }
            if (comboLO.SelectedIndex == -1) {
                log("error: no LO instrument selected");
                return false;
            }
            if (comboOUT.SelectedIndex == -1) {
                log("error: no OUT instrument selected");
                return false;
            }
            if ((calibrationTask != null) && (!calibrationTask.IsCompleted)) {
                log("error: calibration is already running");
                return false;
            }
            return true;
        }

        public void calibrate(Action func) {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            calibrationTask = Task.Factory.StartNew(func);

            dataGrid.ItemsSource = dataTable.AsDataView();
        }

        public void calibrateIN() {
            calibrate(() => instrumentManager.calibrateInExec(dataTable, inParameters[measureMode]));
        }

        public void calibrateLO() {
            calibrate(() => instrumentManager.calibrateLoExec(dataTable, loParameters));
        }

        public void calibrateOUT() {
            calibrate(() => instrumentManager.calibrateOutExec(dataTable, measureMode));
        }

        #endregion regCalibrationManager

        #region regMeasurementManager

        public bool canMeasure() {
            if (dataTable == null) {
                log("error: no table open");
                return false;
            }
            if (comboIN.SelectedIndex == -1) {
                log("error: no IN instrument selected");
                return false;
            }
            if (comboOUT.SelectedIndex == -1) {
                log("error: no OUT instrument selected");
                return false;
            }
            if (comboLO.SelectedIndex == -1) {
                log("error: no LO instrument selected");
                return false;
            }
            return true;
        }

        public void measure_mix_DSB_down()
        {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboLO.SelectedItem).Location;
            string IN = ((Instrument)comboIN.SelectedItem).Location;
            string OUT = ((Instrument)comboOUT.SelectedItem).Location;

            send(OUT, ":CAL:AUTO OFF");
            send(OUT, ":SENS:FREQ:SPAN span");
            send(OUT, ":CALC:MARK1:MODE POS");
            send(OUT, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));

            send(IN, ("OUTP:STAT ON"));

            foreach (DataRow row in dataTable.Rows)
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
                    send(IN, ("SOUR:FREQ " + t_freq_RF));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(IN, ("SOUR:POW " + t_pow_RF.Replace(',', '.')));

                    // IF
                    send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                    send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                    Thread.Sleep(delay);
                    t_pow = query(OUT, ":CALCulate:MARKer:Y?");
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
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_RF));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_RF));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
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
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
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
            send(OUT, ":CAL:AUTO ON");
            send(LO, "OUTP:STAT OFF");
            send(IN, "OUTP:STAT OFF");
            dataGrid.ItemsSource = dataTable.AsDataView();

        }

        public void measure_mix_DSB_up()
        {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboLO.SelectedItem).Location;
            string IN = ((Instrument)comboIN.SelectedItem).Location;
            string OUT = ((Instrument)comboOUT.SelectedItem).Location;

            send(OUT, ":CAL:AUTO OFF");
            send(OUT, ":SENS:FREQ:SPAN span");
            send(OUT, ":CALC:MARK1:MODE POS");
            send(OUT, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));
            send(IN, ("OUTP:STAT ON"));

            foreach (DataRow row in dataTable.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_IF = row["PIF"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_IF)))
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
                    decimal t_pow_goal_IF_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_IF_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_RF_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_RF = (t_freq_RF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(IN, ("SOUR:FREQ " + t_freq_IF));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    send(IN, ("SOUR:POW " + t_pow_IF.Replace(',', '.')));

                    //CONV
                    send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_RF));
                    send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_RF));
                    Thread.Sleep(delay);
                    t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                    decimal.TryParse(row["ATT-RF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                    decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_RF_dec);
                    t_pow = t_pow_RF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["POUT-RF"] = t_pow.Replace('.', ',');
                    decimal t_conv_dec = t_pow_RF_dec + att - t_pow_goal_IF_dec;
                    string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                    row["CONV"] = t_conv.Replace('.', ',');

                    //ISO-IF
                    att_str = row["ATT-IF"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                        t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-IF"] = t_pow.Replace('.', ',');
                        decimal t_iso_if_dec = t_pow_goal_IF_dec - att - t_pow_IF_dec;
                        string t_iso_if = t_iso_if_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-IF"] = t_iso_if.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
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
            send(OUT, ":CAL:AUTO ON");
            send(LO, "OUTP:STAT OFF");
            send(IN, "OUTP:STAT OFF");
            dataGrid.ItemsSource = dataTable.AsDataView();

        }

        public void measure_mix_SSB_down()
        {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboLO.SelectedItem).Location;
            string IN = ((Instrument)comboIN.SelectedItem).Location;
            string OUT = ((Instrument)comboOUT.SelectedItem).Location;

            send(OUT, ":CAL:AUTO OFF");
            send(OUT, ":SENS:FREQ:SPAN span");
            send(OUT, ":CALC:MARK1:MODE POS");
            send(OUT, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));
            send(IN, ("OUTP:STAT ON"));

            foreach (DataRow row in dataTable.Rows)
            {
                string att_str = "";
                string t_pow = "";
                string t_freq_LO = "";
                string t_freq_LSB = "";
                string t_freq_USB = "";
                string t_freq_IF = "";
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_LSB = row["PLSB"].ToString();
                string t_pow_USB = row["PUSB"].ToString();
                decimal t_freq_LO_dec = 0;
                decimal t_freq_LSB_dec = 0;
                decimal t_freq_USB_dec = 0;
                decimal t_freq_IF_dec = 0;
                decimal t_pow_IF_dec = 0;
                decimal t_pow_LSB_dec = 0;
                decimal t_pow_USB_dec = 0;
                decimal t_pow_LO_dec = 0;
                decimal t_pow_goal_LSB_dec = 0;
                decimal t_pow_goal_USB_dec = 0;
                decimal t_pow_goal_LO_dec = 0;
                decimal att = 0;

                if (!(string.IsNullOrEmpty(t_pow_LO)))
                {

                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LO_dec);
                        t_pow = t_pow_LO_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LO"] = t_pow.Replace('.', ',');
                        decimal t_iso_lo_dec = t_pow_goal_LO_dec - att - t_pow_LO_dec;
                        string t_iso_lo = t_iso_lo_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-LO"] = t_iso_lo.Replace('.', ',');
                    }

                    if (!(string.IsNullOrEmpty(t_pow_LSB)) && t_pow_LSB != "-" )
                    {
                        decimal.TryParse(row["PLSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LSB_dec);
                        decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LSB_dec);
                        t_freq_LSB = (t_freq_LSB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                        send(IN, ("SOUR:FREQ " + t_freq_LSB));
                        send(IN, ("SOUR:POW " + t_pow_LSB.Replace(',', '.')));

                        //CONV-LSB
                        string att_if = row["ATT-IF"].ToString();
                        if (!(string.IsNullOrEmpty(att_if)) && att_if != "-")
                        {
                            send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                            send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                            Thread.Sleep(delay);
                            t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                            t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["POUT-IF"] = t_pow.Replace('.', ',');
                            decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_LSB_dec + 3;
                            string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["CONV-LSB"] = t_conv.Replace('.', ',');
                        }

                        //ISO-LSB
                        string att_lsb = row["ATT-LSB"].ToString();
                        if (!(string.IsNullOrEmpty(att_if)) && att_if != "-")
                        {
                            send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LSB));
                            send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LSB));
                            Thread.Sleep(delay);
                            t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(row["ATT-LSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LSB_dec);
                            t_pow = t_pow_LSB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["POUT-LSB"] = t_pow.Replace('.', ',');
                            decimal t_conv_dec = t_pow_goal_LSB_dec - att - t_pow_LSB_dec;
                            string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["ISO-LSB"] = t_conv.Replace('.', ',');
                        }
                    }

                    //CONV-USB
                    if (!(string.IsNullOrEmpty(t_pow_USB)) && t_pow_USB != "-")
                    {
                        decimal.TryParse(row["PUSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_USB_dec);
                        decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_USB_dec);
                        t_freq_USB = (t_freq_USB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                        send(IN, ("SOUR:FREQ " + t_freq_USB));
                        send(IN, ("SOUR:POW " + t_pow_USB.Replace(',', '.')));

                        //ISO-USB
                        string att_usb = row["ATT-USB"].ToString();
                        if (!(string.IsNullOrEmpty(att_usb)) && att_usb != "-")
                        {
                            send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_USB));
                            send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_USB));
                            Thread.Sleep(delay);
                            t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(row["ATT-USB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_USB_dec);
                            t_pow = t_pow_USB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["POUT-USB"] = t_pow.Replace('.', ',');
                            decimal t_conv_dec = t_pow_goal_USB_dec - att - t_pow_USB_dec;
                            string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["ISO-USB"] = t_conv.Replace('.', ',');
                        }

                        //CONV-USB
                        string att_if = row["ATT-IF"].ToString();
                        if (!(string.IsNullOrEmpty(att_if)) && att_if != "-")
                        {
                            send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                            send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                            Thread.Sleep(delay);
                            t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                            decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                            decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                            t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["POUT-IF"] = t_pow.Replace('.', ',');
                            decimal t_conv_dec = t_pow_IF_dec + att - t_pow_goal_USB_dec + 3;
                            string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                            row["CONV-USB"] = t_conv.Replace('.', ',');
                        }
                    }
                }
            }
            send(OUT, ":CAL:AUTO ON");
            send(LO, "OUTP:STAT OFF");
            send(IN, "OUTP:STAT OFF");
            dataGrid.ItemsSource = dataTable.AsDataView();
        }

        public void measure_mix_SSB_up()
        {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            string LO = ((Instrument)comboLO.SelectedItem).Location;
            string IN = ((Instrument)comboIN.SelectedItem).Location;
            string OUT = ((Instrument)comboOUT.SelectedItem).Location;

            send(OUT, ":CAL:AUTO OFF");
            send(OUT, ":SENS:FREQ:SPAN 1000000");
            send(OUT, ":CALC:MARK1:MODE POS");
            send(OUT, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(LO, ("OUTP:STAT ON"));
            //send(IN, ("OUTP:STAT ON"));

            foreach (DataRow row in dataTable.Rows)
            {
                string t_pow_LO = row["PLO"].ToString();
                string t_pow_IF = row["PIF"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_LO)) && !(string.IsNullOrEmpty(t_pow_IF)))
                {
                    string att_str = "";
                    string t_pow = "";
                    string t_freq_LO = "";
                    string t_freq_LSB = "";
                    string t_freq_USB = "";
                    string t_freq_IF = "";
                    decimal t_freq_LO_dec = 0;
                    decimal t_freq_LSB_dec = 0;
                    decimal t_freq_USB_dec = 0;
                    decimal t_freq_IF_dec = 0;
                    decimal t_pow_IF_dec = 0;
                    decimal t_pow_LSB_dec = 0;
                    decimal t_pow_USB_dec = 0;
                    decimal t_pow_LO_dec = 0;
                    decimal t_pow_goal_IF_dec = 0;
                    decimal t_pow_goal_LO_dec = 0;
                    decimal att = 0;

                    decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_IF_dec);
                    decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_LO_dec);
                    decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LO_dec);
                    decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_LSB_dec);
                    decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_USB_dec);
                    decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_IF_dec);
                    t_freq_LO = (t_freq_LO_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_LSB = (t_freq_LSB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_USB = (t_freq_USB_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    t_freq_IF = (t_freq_IF_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);

                    send(LO, ("SOUR:FREQ " + t_freq_LO));
                    //send(IN, ("SOUR:FREQ " + t_freq_IF));
                    send(LO, ("SOUR:POW " + t_pow_LO.Replace(',', '.')));
                    //send(IN, ("SOUR:POW " + t_pow_IF.Replace(',', '.')));

                    // CONV-LSB
                    att_str = row["ATT-LSB"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LSB));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LSB));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-LSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_LSB_dec);
                        t_pow = t_pow_LSB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-LSB"] = t_pow.Replace('.', ',');
                        decimal t_conv_dec = t_pow_LSB_dec + att - t_pow_goal_IF_dec - 3;
                        string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-LSB"] = t_conv.Replace('.', ',');
                    }

                    // CONV-USB
                    att_str = row["ATT-USB"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_USB));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_USB));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-USB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_USB_dec);
                        t_pow = t_pow_USB_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-USB"] = t_pow.Replace('.', ',');
                        decimal t_conv_dec = t_pow_USB_dec + att - t_pow_goal_IF_dec - 3;
                        string t_conv = t_conv_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-USB"] = t_conv.Replace('.', ',');
                    }

                    //ISO-IF
                    att_str = row["ATT-IF"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_IF));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_IF));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(row["ATT-IF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_IF_dec);
                        t_pow = t_pow_IF_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-IF"] = t_pow.Replace('.', ',');
                        decimal t_iso_if_dec = t_pow_goal_IF_dec - att - t_pow_IF_dec - 3;
                        string t_iso_if = t_iso_if_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["ISO-IF"] = t_iso_if.Replace('.', ',');
                    }

                    //ISO-LO
                    att_str = row["ATT-LO"].ToString();
                    if (att_str != "-")
                    {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_LO));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_LO));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
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
            send(OUT, ":CAL:AUTO ON");
            send(LO, "OUTP:STAT OFF");
            //send(IN, "OUTP:STAT OFF");
            dataGrid.ItemsSource = dataTable.AsDataView();
        }

        public void measure_mult()
        {
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            string IN = ((Instrument)comboIN.SelectedItem).Location;
            string OUT = ((Instrument)comboOUT.SelectedItem).Location;

            send(OUT, ":CAL:AUTO OFF");
            send(OUT, ":SENS:FREQ:SPAN span");
            send(OUT, ":CALC:MARK1:MODE POS");
            send(OUT, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(IN, ("OUTP:STAT ON"));

            foreach (DataRow row in dataTable.Rows)
            {
                string t_pow_gen = row["PIN-GEN"].ToString();
                string att1 = row["ATT-H1"].ToString();
                string att2 = row["ATT-H2"].ToString();
                string att3 = row["ATT-H3"].ToString();
                string att4 = row["ATT-H4"].ToString();
                if (!(string.IsNullOrEmpty(t_pow_gen)) & ((!(string.IsNullOrEmpty(att1)) & att1 != "-") || (!(string.IsNullOrEmpty(att2)) & att2 != "-") || (!(string.IsNullOrEmpty(att3)) & att3 != "-") || (!(string.IsNullOrEmpty(att4)) & att4 != "-")))
                {
                    string t_pow = row["PIN-GOAL"].ToString();
                    send(IN, ("SOUR:POW " + t_pow_gen.Replace(',', '.')));

                    decimal t_pow_goal_dec = 0;
                    decimal.TryParse(t_pow.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_goal_dec);

                    //harm1
                    string t_freq_H1 = row["FH1"].ToString();
                    decimal t_freq_H1_dec = 0;
                    decimal.TryParse(t_freq_H1.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out t_freq_H1_dec);
                    t_freq_H1 = (t_freq_H1_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                    send(IN, ("SOUR:FREQ " + t_freq_H1));
                    if (t_freq_H1_dec * 1000 < maxfreq && (!(string.IsNullOrEmpty(att1)) & att1 != "-"))
                        {
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_H1));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_H1));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H1_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H1_dec);
                        t_pow = t_pow_H1_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H1"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att1.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H1_dec = t_pow_H1_dec + att_dec - t_pow_goal_dec;
                        string conv_H1 = conv_H1_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H1"] = conv_H1.Replace('.', ',');
                        }

                    //harm2
                    decimal t_freq_H2_dec = t_freq_H1_dec * 2;
                    if (t_freq_H2_dec * 1000 < maxfreq && (!(string.IsNullOrEmpty(att2)) & att2 != "-"))
                    {
                        string t_freq_H2 = (t_freq_H2_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_H2));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_H2));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H2_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H2_dec);
                        t_pow = t_pow_H2_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H2"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att2.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H2_dec = t_pow_H2_dec + att_dec - t_pow_goal_dec;
                        string conv_H2 = conv_H2_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H2"] = conv_H2.Replace('.', ',');
                    }

                    //harm3
                    decimal t_freq_H3_dec = t_freq_H1_dec * 3;
                    if (t_freq_H3_dec * 1000 < maxfreq && (!(string.IsNullOrEmpty(att3)) & att3 != "-"))
                    {
                        string t_freq_H3 = (t_freq_H3_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_H3));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_H3));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H3_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H3_dec);
                        t_pow = t_pow_H3_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H3"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att3.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H3_dec = t_pow_H3_dec + att_dec - t_pow_goal_dec;
                        string conv_H3 = conv_H3_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H3"] = conv_H3.Replace('.', ',');
                    }

                    //harm4
                    decimal t_freq_H4_dec = t_freq_H1_dec * 4;
                    if (t_freq_H4_dec * 1000 < maxfreq && (!(string.IsNullOrEmpty(att4)) & att4 != "-"))
                    {
                        string t_freq_H4 = (t_freq_H4_dec * 1000000000).ToString("0", CultureInfo.InvariantCulture);
                        send(OUT, (":SENSe:FREQuency:RF:CENTer " + t_freq_H4));
                        send(OUT, (":CALCulate:MARKer1:X:CENTer " + t_freq_H4));
                        Thread.Sleep(delay);
                        t_pow = query(OUT, ":CALCulate:MARKer:Y?");
                        decimal t_pow_H4_dec = 0;
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_H4_dec);
                        t_pow = t_pow_H4_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["POUT-H4"] = t_pow.Replace('.', ',');
                        decimal att_dec = 0;
                        decimal.TryParse(att4.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out att_dec);
                        decimal conv_H4_dec = t_pow_H4_dec + att_dec - t_pow_goal_dec;
                        string conv_H4 = conv_H4_dec.ToString("0.00", CultureInfo.InvariantCulture);
                        row["CONV-H4"] = conv_H4.Replace('.', ',');
                    }
                }
            }
            send(OUT, ":CAL:AUTO ON");
            send(IN, "OUTP:STAT OFF");
            dataGrid.ItemsSource = dataTable.AsDataView();
        }

        #endregion regMeasurementManager

    }
}
