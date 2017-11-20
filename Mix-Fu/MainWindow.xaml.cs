//#define mock

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
using System.Diagnostics;
using System.Xml.Resolvers;

namespace Mix_Fu {

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
    
    public partial class MainWindow : Window {

#region regDataMembers
        // TODO: switch TryParse overloads
        // TODO: learn АКИП syntax

        List<Instrument> listInstruments = new List<Instrument>();

        DataTable dataTable;
        string xlsx_path = "";
        int delay = 300;
        decimal attenuation = 30;
        decimal maxfreq = 26500;
        decimal span = 10*Constants.MHz;
        bool verbose = false;
        MeasureMode measureMode = MeasureMode.modeDSBDown;

        NumberStyles style = NumberStyles.Any;
        CultureInfo culture = CultureInfo.InvariantCulture;

        private Task searchTask;
        private Task calibrationTask;
        private Task measureTask;
        private Progress<double> progressHandler;

        InstrumentManager instrumentManager;

        CancellationTokenSource searchTokenSource = new CancellationTokenSource();
        CancellationTokenSource calibrationTokenSource = new CancellationTokenSource();

#endregion regDataMembers

        public MainWindow() {
            instrumentManager = new InstrumentManager(log);

            InitializeComponent();

            comboLO.ItemsSource = listInstruments;
            comboIN.ItemsSource = listInstruments;
            comboOUT.ItemsSource = listInstruments;

            progressHandler = new Progress<double>();
            progressHandler.ProgressChanged += (sender, value) => { pbTaskStatus.Value = value; };

            //            listInstruments.Add(new Instrument { Location = "GPIB0::20::INSTR", Name = "IN", FullName = "GPIB0::20::INSTR" });
            //            listInstruments.Add(new Instrument { Location = "GPIB0::18::INSTR", Name = "OUT", FullName = "GPIB0::18::INSTR" });
            //            listInstruments.Add(new Instrument { Location = "GPIB0::1::INSTR", Name = "LO", FullName = "GPIB0::1:INSTR" });
#if mock
            listInstruments.Add(new Instrument { Location = "IN_location", Name = "IN", FullName = "IN at IN_location" });
            listInstruments.Add(new Instrument { Location = "OUT_location", Name = "OUT", FullName = "OUT at OUT_location" });
            listInstruments.Add(new Instrument { Location = "LO_location", Name = "LO", FullName = "LO at LO_location" });

            comboIN.SelectedIndex = 2;
            comboOUT.SelectedIndex = 1;
            comboLO.SelectedIndex = 0;
#endif

            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        }

        private void ProgressHandler_ProgressChanged(object sender, double e) {
            throw new NotImplementedException();
        }

        #region regUiEvents

        // options
        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            measureMode = (MeasureMode)listBox.SelectedIndex;
        }

        private void LogVerboseToggled(object sender, RoutedEventArgs e) {
            verbose = LogVerbose.IsChecked ?? false;
        }

        private void textBox_delay_TextChanged(object sender, TextChangedEventArgs e) {
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
            if (instrumentManager.m_IN != null) {
                log(instrumentManager.m_IN.ToString(), true);
            }
        }

        private void comboOUT_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            instrumentManager.m_OUT = (Instrument)((ComboBox)sender).SelectedItem;
            if (instrumentManager.m_IN != null) {
                log(instrumentManager.m_OUT.ToString(), true);
            }
        }

        private void comboLO_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            instrumentManager.m_LO = (Instrument)((ComboBox)sender).SelectedItem;
            if (instrumentManager.m_IN != null) {
                log(instrumentManager.m_LO.ToString(), true);
            }
        }

        // misc buttons
        private async void btnSearchClicked(object sender, RoutedEventArgs e) {
            if (searchTask != null && !searchTask.IsCompleted) {
                MessageBox.Show("Instrument search is already running.");
                return;
            }
            btnSearch.Visibility = Visibility.Hidden;
            btnStopSearch.Visibility = Visibility.Visible;

            int max_port = Convert.ToInt32(textBox_number_maxport.Text);
            int gpib = Convert.ToInt32(textBox_number_GPIB.Text);
            CancellationToken token = searchTokenSource.Token;

            listInstruments.Clear();

            try { 
                searchTask = Task.Factory.StartNew(
                    () => instrumentManager.searchInstruments(listInstruments, max_port, gpib, token), token);
                await searchTask;
            }
            catch (TaskCanceledException ex) {
                log(ex.Message);
            }

            comboIN.Items.Refresh();
            comboOUT.Items.Refresh();
            comboLO.Items.Refresh();

            btnStopSearch.Visibility = Visibility.Hidden;
            btnSearch.Visibility = Visibility.Visible;
        }

        private void btnStopSearchClicked(object sender, RoutedEventArgs e) {
            if (searchTask != null && !searchTask.IsCompleted) {
                searchTokenSource.Cancel();
            }
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
        // TODO: rewrite data handling
        private void btnImportXlsxClicked(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = false,
                Filter = "xlsx файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*"
            };
            // TODO: exception handling
            if (!(bool)openFileDialog.ShowDialog()) {
                return;
            }
            xlsx_path = openFileDialog.FileName;
            try {
                dataTable = getDataTableFromExcel(xlsx_path);
                dataGrid.ItemsSource = dataTable.AsDataView();
            }
            catch (Exception ex) {
                MessageBox.Show("Error: can't open file, check log");
                log("error: " + ex.Message);
            }
        }

        private void btnSaveXlsxClicked(object sender, RoutedEventArgs e) {
            // TODO: allow overwrite existing files
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog {
                Filter = "Таблица (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*"
            };
            if (!(bool) saveFileDialog.ShowDialog()) {
                return; 
            }
            var xlsx_file = new FileInfo(saveFileDialog.FileName);
            using (var package = new ExcelPackage(xlsx_file)) {
                dataTable = ((DataView)dataGrid.ItemsSource).ToTable();
                var ws = package.Workbook.Worksheets.Add("1");
                ws.Cells ["A1"].LoadFromDataTable(dataTable, true);
                foreach (var cell in ws.Cells) {
                    cell.Value = cell.Value.ToString().Replace(',', '.');
                    try {
                        cell.Value = Convert.ToDecimal(cell.Value);
                    }
                    catch {
                        cell.Style.Font.Bold = true;
                    }
                }
                package.Save();
            }
        }

        // calibration buttons
        private void btnCalibrateInClicked(object sender, RoutedEventArgs e) {
            // TODO: check for option input errors
            if (!canCalibrateIN_OUT()) {
                MessageBox.Show("Error: check log");
                return;
            }
            var progress = progressHandler as IProgress<double>;
            CancellationToken token = calibrationTokenSource.Token;
            calibrate(() => instrumentManager.calibrateIn(progress, dataTable, instrumentManager.inParameters[measureMode], token));
        }

        private void btnCalibrateOutClicked(object sender, RoutedEventArgs e) {
            // TODO: check for option input errors
            if (!canCalibrateIN_OUT()) {
                MessageBox.Show("Error: check log");
                return;
            }
            var progress = progressHandler as IProgress<double>;
            calibrate(() => instrumentManager.calibrateOut(progress, dataTable, instrumentManager.outParameters[measureMode], measureMode));
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
            var progress = progressHandler as IProgress<double>;
            CancellationToken token = calibrationTokenSource.Token;
            calibrate(() => instrumentManager.calibrateLo(progress, dataTable, instrumentManager.loParameters, token));
        }

        // measure button
        private void btnMeasureClicked(object sender, RoutedEventArgs e) {
            try {
                if (!canMeasure()) {
                    MessageBox.Show("Error: check log");
                    return;
                }

                measure();
            }
            catch (Exception ex) {
                log(ex.Message, false);
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

        public object cell2dec(string value) {
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

#region regDataManager

        public static DataTable getDataTableFromExcel(string path, bool hasHeader = true) {
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
            if (calibrationTask != null && !calibrationTask.IsCompleted) {
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

        public async void calibrate(Action func) {
            var stopwatch = Stopwatch.StartNew();

            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

//            calibrationTask = Task.Factory.StartNew(func);
            calibrationTask = Task.Run(func);
            await calibrationTask;

            dataGrid.ItemsSource = dataTable.AsDataView();

            stopwatch.Stop();
            log("run time: " + Math.Round(stopwatch.Elapsed.TotalMilliseconds / 1000, 2) + " sec", false);
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
            if (measureTask != null && !measureTask.IsCompleted) {
                log("error: measure is already running");
                return false;
            }
            return true;
        }

        public async void measure() {
            var stopwatch = Stopwatch.StartNew();

            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            var progress = progressHandler as IProgress<double>;
            switch (measureMode) {
                case MeasureMode.modeDSBDown:
                    measureTask = Task.Factory.StartNew(() => measure_mix_DSB_down(progress, dataTable));
                    break;
                case MeasureMode.modeDSBUp:
                    measureTask = Task.Factory.StartNew(() => measure_mix_DSB_up(progress, dataTable));
                    break;
                case MeasureMode.modeSSBDown:
                    measureTask = Task.Factory.StartNew(() => measure_mix_SSB_down(progress, dataTable));
                    break;
                case MeasureMode.modeSSBUp:
                    measureTask = Task.Factory.StartNew(() => measure_mix_SSB_up(progress, dataTable));
                    break;
                case MeasureMode.modeMultiplier:
                    measureTask = Task.Factory.StartNew(() => measure_mult(progress, dataTable));
                    break;
                default:
                    return;
            }

            dataGrid.ItemsSource = dataTable.AsDataView();

            await measureTask;
            stopwatch.Stop();
            log("run time: " + Math.Round(stopwatch.Elapsed.TotalMilliseconds / 1000, 2) + " sec", false);
        }

        private void measurePower(DataRow row, string SA, decimal powGoal, decimal freq, string colAtt, string colPow, string colConv, int coeff, int corr) {
            string attStr = row[colAtt].ToString().Replace(',', '.');
            if (string.IsNullOrEmpty(attStr) || attStr == "-") {
                log("error: measure: empty row, skipping: " + colPow + ": freq=" + freq + " powgoal=" + powGoal, true);
                row[colPow] = "-";
                row[colConv] = "-";
                return;
            }
            if (freq > maxfreq) {
                log("error: measure: freq is out of limits, skipping: " + colPow + ": freq=" + freq + " powgoal=" + powGoal, true);
                row[colPow] = "-";
                row[colConv] = "-";
                return;
            }
            // TODO: move to InstrumentManager
            try {
                instrumentManager.send(SA, ":SENSe:FREQuency:RF:CENTer " + freq);
                instrumentManager.send(SA, ":CALCulate:MARKer1:X:CENTer " + freq);
            }
            catch (Exception ex) {
                log("error: measure fail setting freq: " + ex.Message);
                row[colPow] = "-";
                row[colConv] = "-";
                return;
            }
            Thread.Sleep(delay);

            decimal att = 0;
            decimal.TryParse(attStr, NumberStyles.Any, CultureInfo.InvariantCulture, out att);

            decimal readPow = 0;
            try {
                decimal.TryParse(instrumentManager.query(SA, ":CALCulate:MARKer:Y?"), NumberStyles.Any,
                    CultureInfo.InvariantCulture, out readPow);
            }
            catch (Exception ex) {
                log("error: " + ex.Message, false);
            }
            decimal diff = coeff * (powGoal - att - readPow + corr);

            row[colPow] = readPow.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
            row[colConv] = diff.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void measure_mix_DSB_down(IProgress<double> prog, DataTable data) {
            log("start measure: " + measureMode);
            string IN = instrumentManager.m_IN.Location;
            string OUT = instrumentManager.m_OUT.Location;
            string LO = instrumentManager.m_LO.Location;

            // TODO: move all sends
            instrumentManager.prepareInstrument(IN, OUT);
            instrumentManager.send(LO, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                // TODO: convert do decimal?
                string inPowLO = row["PLO"].ToString().Replace(',', '.');
                string inPowRF = row["PRF"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLO) || inPowLO == "-" ||
                    string.IsNullOrEmpty(inPowRF) || inPowRF == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                // TODO: check for empty columns:
                decimal inPowLOGoalDec = 0;
                decimal inPowRFGoalDec = 0;
                decimal inFreqIFDec = 0;
                decimal inFreqRFDec = 0;
                decimal inFreqLODec = 0;

                // TODO: exception handling
                // TODO: need to check for empty cells?
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), style, culture, out inPowLOGoalDec);
                decimal.TryParse(row["PRF-GOAL"].ToString().Replace(',', '.'), style, culture, out inPowRFGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), style, culture, out inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), style, culture, out inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), style, culture, out inFreqIFDec);
                inFreqLODec *= Constants.GHz;
                inFreqRFDec *= Constants.GHz;
                inFreqIFDec *= Constants.GHz;

                // TODO: write "-" into corresponding column on fail
                try {
                    instrumentManager.send(IN, "SOUR:FREQ " + inFreqRFDec);
                    instrumentManager.send(LO, "SOUR:FREQ " + inFreqLODec);
                }
                catch (Exception ex) {
                    log("error: measure fail setting freq, skipping row: " +ex.Message);
                    continue;
                }
                try {
                    instrumentManager.send(IN, "SOUR:POW " + inPowRF);
                    instrumentManager.send(LO, "SOUR:POW " + inPowLO);
                }
                catch (Exception ex) {
                    log("error: measure fail setting pow, skipping row: " + ex.Message);
                    continue;
                }

                measurePower(row, OUT, inPowRFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV", -1, 0);
                measurePower(row, OUT, inPowRFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "ISO-RF", 1, 0);
                measurePower(row, OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instrumentManager.releaseInstrument(IN, OUT);
            // TODO: move sends to instrumentManager
            instrumentManager.send(LO, "OUTP:STAT OFF");

            log("end measure");
            prog?.Report(100);
        }

        public void measure_mix_DSB_up(IProgress<double> prog, DataTable data) {
            log("start measure: " + measureMode);

            string IN = instrumentManager.m_IN.Location;
            string OUT = instrumentManager.m_OUT.Location;
            string LO = instrumentManager.m_LO.Location;

            instrumentManager.prepareInstrument(IN, OUT);
            instrumentManager.send(LO, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                string inPowLO = row["PLO"].ToString().Replace(',', '.');
                string inPowIF = row["PIF"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLO) || inPowLO == "-" ||
                    string.IsNullOrEmpty(inPowIF) || inPowIF == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal inPowIFGoalDec = 0;
                decimal inPowLOGoalDec = 0;
                decimal inFreqLODec = 0;
                decimal inFreqRFDec = 0;
                decimal inFreqIFDec = 0;

                decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), style, culture, out inPowIFGoalDec);
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), style, culture, out inPowLOGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), style, culture, out inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), style, culture, out inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), style, culture, out inFreqIFDec);
                inFreqIFDec *= Constants.GHz;
                inFreqRFDec *= Constants.GHz;
                inFreqLODec *= Constants.GHz;

                // TODO: extract method
                try {
                    instrumentManager.send(LO, "SOUR:FREQ " + inFreqLODec);
                    instrumentManager.send(IN, "SOUR:FREQ " + inFreqIFDec);
                }
                catch (Exception ex) {
                    log("error: measure fail setting freq, skipping row: " + ex.Message);
                    continue;
                }
                try {
                    instrumentManager.send(LO, "SOUR:POW " + inPowLO);
                    instrumentManager.send(IN, "SOUR:POW " + inPowIF);
                }
                catch (Exception ex) {
                    log("error: measure fail setting pow, skipping row: " + ex.Message);
                    continue;
                }

                measurePower(row, OUT, inPowIFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "CONV", -1, 0);
                measurePower(row, OUT, inPowIFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "ISO-IF", 1, 0);
                measurePower(row, OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instrumentManager.releaseInstrument(IN, OUT);
            instrumentManager.send(LO, "OUTP:STAT OFF");
            log("end measure");
            prog?.Report(100);
        }

        public void measure_mix_SSB_down(IProgress<double> prog, DataTable data) {
            log("start measure: " + measureMode);

            string IN = instrumentManager.m_IN.Location;
            string OUT = instrumentManager.m_OUT.Location;
            string LO = instrumentManager.m_LO.Location;

            instrumentManager.prepareInstrument(IN, OUT);
            instrumentManager.send(LO, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                string inPowLOStr = row["PLO"].ToString().Replace(',', '.');
                string inPowLSBStr = row["PLSB"].ToString().Replace(',', '.');
                string inPowUSBStr = row["PUSB"].ToString().Replace(',', '.');
                decimal inFreqLODec = 0;
                decimal inFreqLSBDec = 0;
                decimal inFreqUSBDec = 0;
                decimal inFreqIFDec = 0;
                decimal inPowLSBGoalDec = 0;
                decimal inPowUSBGoalDec = 0;
                decimal inPowLOGoalDec = 0;

                if (string.IsNullOrEmpty(inPowLOStr) || inPowLOStr == "-" ||
                    string.IsNullOrEmpty(inPowLSBStr) || inPowLSBStr == "-" ||
                    string.IsNullOrEmpty(inPowUSBStr) || inPowUSBStr == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowLOGoalDec);
                decimal.TryParse(row["PUSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowUSBGoalDec);
                decimal.TryParse(row["PLSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowLSBGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLODec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqIFDec);
                decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqUSBDec);
                decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLSBDec);

                inFreqLODec *= Constants.GHz;
                inFreqIFDec *= Constants.GHz;
                inFreqLSBDec *= Constants.GHz;
                inFreqUSBDec *= Constants.GHz;

                // TODO: extract freq-pow setting method, add exception handling
                try {
                    instrumentManager.send(LO, "SOUR:FREQ " + inFreqLODec);
                    instrumentManager.send(LO, "SOUR:POW " + inPowLOStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting LO params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                measurePower(row, OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                try {
                    instrumentManager.send(IN, "SOUR:FREQ " + inFreqLSBDec);
                    instrumentManager.send(IN, "SOUR:POW " + inPowLSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN LSB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                measurePower(row, OUT, inPowLSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-LSB", -1, -3);
                measurePower(row, OUT, inPowLSBGoalDec, inFreqLSBDec, "ATT-LSB", "POUT-LSB", "ISO-LSB", 1, 0);

                try {
                    instrumentManager.send(IN, "SOUR:FREQ " + inFreqUSBDec);
                    instrumentManager.send(IN, "SOUR:POW " + inPowUSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN USB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                measurePower(row, OUT, inPowUSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-USB", -1, -3);
                measurePower(row, OUT, inPowUSBGoalDec, inFreqUSBDec, "ATT-USB", "POUT-USB", "ISO-USB", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instrumentManager.releaseInstrument(IN, OUT);
            instrumentManager.send(LO, "OUTP:STAT OFF");
            log("end measure");
            prog?.Report(100);
        }

        public void measure_mix_SSB_up(IProgress<double> prog, DataTable data) {
            log("start measure: " + measureMode);

//            string IN = instrumentManager.m_IN.Location;
            string OUT = instrumentManager.m_OUT.Location;
            string LO = instrumentManager.m_LO.Location;

            instrumentManager.send(OUT, ":CAL:AUTO OFF");
            instrumentManager.send(OUT, ":SENS:FREQ:SPAN 1000000");
            instrumentManager.send(OUT, ":CALC:MARK1:MODE POS");
            instrumentManager.send(OUT, ":POW:ATT " + attenuation);
            //instrumentManager.send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            instrumentManager.send(LO, "OUTP:STAT ON");
            //instrumentManager.send(IN, ("OUTP:STAT ON"));

            int i = 0;
            foreach (DataRow row in data.Rows) {
                string inPowLOStr = row["PLO"].ToString().Replace(',', '.');
                string inPowIFStr = row["PIF"].ToString().Replace(',', '.');

                if (string.IsNullOrEmpty(inPowLOStr) || inPowLOStr == "-" ||
                    string.IsNullOrEmpty(inPowIFStr) || inPowIFStr == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal inPowIFGoal = 0;
                decimal inPowLOGoal = 0;
                decimal inFreqLO = 0;
                decimal inFreqLSB = 0;
                decimal inFreqUSB = 0;
                decimal inFreqIF = 0;

                decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowIFGoal);
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowLOGoal);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLO);
                decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLSB);
                decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqUSB);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqIF);
                inFreqLO *= Constants.GHz;
                inFreqLSB *= Constants.GHz;
                inFreqUSB *= Constants.GHz;
                inFreqIF *= Constants.GHz;

                try {
                    instrumentManager.send(LO, "SOUR:FREQ " + inFreqLO);
                    instrumentManager.send(LO, "SOUR:POW " + inPowLOStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting LO params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                //send(IN, ("SOUR:FREQ " + t_freq_IF));
                //send(IN, ("SOUR:POW " + t_pow_IF.Replace(',', '.')));

                measurePower(row, OUT, inPowIFGoal, inFreqLSB, "ATT-LSB", "POUT-LSB", "CONV-LSB", -1, -3);
                measurePower(row, OUT, inPowIFGoal, inFreqUSB, "ATT-USB", "POUT-USB", "CONV-USB", -1, -3);
                measurePower(row, OUT, inPowIFGoal, inFreqIF, "ATT-IF", "POUT-IF", "ISO-IF", 1, -3);
                measurePower(row, OUT, inPowLOGoal, inFreqLO, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instrumentManager.send(OUT, ":CAL:AUTO ON");
            instrumentManager.send(LO, "OUTP:STAT OFF");
            //send(IN, "OUTP:STAT OFF");
            log("end measure");
            prog?.Report(100);
        }

        public void measure_mult(IProgress<double> prog, DataTable data) {
            log("start measure: " + measureMode);
            string IN = instrumentManager.m_IN.Location;
            string OUT = instrumentManager.m_OUT.Location;

            instrumentManager.prepareInstrument(IN, OUT);

            int i = 0;
            foreach (DataRow row in data.Rows) {
                string inPowGenStr = row["PIN-GEN"].ToString().Replace(',', '.');
                string inFreqH1Str = row["FH1"].ToString().Replace(',', '.');

                if (string.IsNullOrEmpty(inPowGenStr) || inPowGenStr == "-" ||
                    string.IsNullOrEmpty(inFreqH1Str) || inFreqH1Str == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal inPowGoal = 0;
                decimal inFreqH1 = 0;

                decimal.TryParse(row["PIN-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowGoal);
                decimal.TryParse(row["FH1"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqH1);
                inFreqH1 *= Constants.GHz;

                try {
                    instrumentManager.send(IN, "SOUR:POW " + inPowGenStr);
                    instrumentManager.send(IN, "SOUR:FREQ " + inFreqH1);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN params, skipping row (" + ex.Message + ")", false);
                    continue;
                }

                measurePower(row, OUT, inPowGoal, inFreqH1*1, "ATT-H1", "POUT-H1", "CONV-H1", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1*2, "ATT-H2", "POUT-H2", "CONV-H2", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1*3, "ATT-H3", "POUT-H3", "CONV-H3", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1*4, "ATT-H4", "POUT-H4", "CONV-H4", -1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instrumentManager.releaseInstrument(IN, OUT);
            log("start measure: " + measureMode);
            prog?.Report(100);
        }

#endregion regMeasurementManager

    }
}
