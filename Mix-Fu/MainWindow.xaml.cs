//#define mock

using System;
using System.CodeDom;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using Agilent.CommandExpert.ScpiNet.Ag34410_2_35;
using Agilent.CommandExpert.ScpiNet.Ag90x0_SA_A_08_03;
using System.Diagnostics;
using System.Media;
using System.Text;

namespace Mixer {

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
        // TODO: autoscroll table

        // TODO: progressbar for calibration and measurements
        // TODO: learn АКИП syntax

        List<Instrument> listInstruments = new List<Instrument>();

        DataTable dataTable;
        string inFile = "";
        int delay = 300;
        decimal attenuation = 30;
        decimal maxfreq = 26500;
        decimal span = 10*Constants.MHz;
        bool verbose = false;
        MeasureMode mode = MeasureMode.modeDSBDown;
        SoundPlayer sndAlert = new SoundPlayer(@".\alert.wav");

        private Task searchTask;
        private Task calibrationTask;
        private Task measureTask;
        private Progress<double> progressHandler;

        private FileInfo logFileInfo;

        InstrumentManager instrumentManager;

        private CancellationTokenSource searchTokenSource;
        private CancellationTokenSource calibrationTokenSource;
        private CancellationTokenSource measureTokenSource;

        #endregion regDataMembers

        public MainWindow() {
            instrumentManager = new InstrumentManager(log);

            InitializeComponent();

            comboLO.ItemsSource = listInstruments;
            comboIN.ItemsSource = listInstruments;
            comboOUT.ItemsSource = listInstruments;

            progressHandler = new Progress<double>();
            progressHandler.ProgressChanged += (sender, value) => { pbTaskStatus.Value = value; };

            sndAlert.LoadAsync();

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

            var logDir = new DirectoryInfo(@".\logs\");
            logFileInfo = new FileInfo(@".\logs\log-" + DateTime.Today.ToString("yyyy-MM-dd") + ".txt");
            try {
                if (!logDir.Exists) {
                    logDir.Create();
                }
                using (var f = new StreamWriter(logFileInfo.FullName, true)) {
                    f.Write("\n========================================================================");
                    f.Write("\nStart session: " + DateTime.Today.ToString("yyyy-MM-dd") +
                                DateTime.Now.ToShortTimeString());
                    f.Write("\n------------------------------------------------------------------------");
                }
            }
            catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message);
                Environment.Exit(1);
            }
        }

        #region regUiEvents

        // options
        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            mode = (MeasureMode)listBox.SelectedIndex;
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

        // search buttons
        private async void btnSearchClicked(object sender, RoutedEventArgs e) {
            if (searchTask != null && !searchTask.IsCompleted) {
                MessageBox.Show("Instrument search is already running.");
                return;
            }
            btnSearch.Visibility = Visibility.Hidden;
            btnStopSearch.Visibility = Visibility.Visible;

            int maxPort = Convert.ToInt32(textBox_number_maxport.Text);
            int gpib = Convert.ToInt32(textBox_number_GPIB.Text);
            searchTokenSource = new CancellationTokenSource();
            CancellationToken token = searchTokenSource.Token;

            listInstruments.Clear();

            try {
                var progress = progressHandler as IProgress<double>;
                searchTask = Task.Run(
                    () => instrumentManager.searchInstruments(progress, listInstruments, maxPort, gpib, token), token);
                await searchTask;
            }
            catch (Exception ex) {
                log("error: " + ex.Message);
            }
            finally {
//                searchTask?.Dispose();

                comboIN.Items.Refresh();
                comboOUT.Items.Refresh();
                comboLO.Items.Refresh();

                btnStopSearch.Visibility = Visibility.Hidden;
                btnSearch.Visibility = Visibility.Visible;
            }
        }

        private void btnStopSearchClicked(object sender, RoutedEventArgs e) {
            if (searchTask != null && !searchTask.IsCompleted) {
                searchTokenSource.Cancel();
            }
        }

        // manual command buttons
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
        private void btnImportXlsxClicked(object sender, RoutedEventArgs e) {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog {
                Multiselect = false,
                Filter = "xlsx файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*"
            };
            // TODO: exception handling
            if (!(bool)openFileDialog.ShowDialog()) {
                return;
            }
            inFile = openFileDialog.FileName;
            try {
                dataTable = getDataTableFromExcel(inFile);
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
            var outFile = new FileInfo(saveFileDialog.FileName);
            if (outFile.Exists) {
                outFile.Delete();
            }
            using (var package = new ExcelPackage(outFile)) {
                dataTable = ((DataView)dataGrid.ItemsSource).ToTable();
                var ws = package.Workbook.Worksheets.Add("1");
                ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                foreach (var cell in ws.Cells) {
                    string cellStr = cell.Value.ToString().Replace(',', '.');
                    cell.Value = cellStr;
                    decimal dummy;
                    if (!decimal.TryParse(cellStr, out dummy)) {
                        cell.Style.Font.Bold = true;
                    }
//                    try {
//                        cell.Value = Convert.ToDecimal(cell.Value);
//                    }
//                    catch {
//                        cell.Style.Font.Bold = true;
//                    }
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
            calibrationTokenSource = new CancellationTokenSource();
            CancellationToken token = calibrationTokenSource.Token;
            calibrate(
                () => instrumentManager.calibrateIn(progress, dataTable, instrumentManager.inParameters[mode], token),
                token);
        }

        private void btnCalibrateOutClicked(object sender, RoutedEventArgs e) {
            // TODO: check for option input errors
            if (!canCalibrateIN_OUT()) {
                MessageBox.Show("Error: check log");
                return;
            }
            var progress = progressHandler as IProgress<double>;
            calibrationTokenSource = new CancellationTokenSource();
            CancellationToken token = calibrationTokenSource.Token;
            calibrate(
                () => instrumentManager.calibrateOut(progress, dataTable, instrumentManager.outParameters[mode],
                    mode, token), token);
        }

        private void btnCalibrateLoClicked(object sender, RoutedEventArgs e) {
            if (!canCalibrateLO()) {
                MessageBox.Show("Error: check log");
                return;
            }
            if (mode == MeasureMode.modeMultiplier) {
                MessageBox.Show("Error: check log");
                log("error: multiplier doesn't need LO");
                return;
            }
            var progress = progressHandler as IProgress<double>;
            calibrationTokenSource = new CancellationTokenSource();
            CancellationToken token = calibrationTokenSource.Token;
            calibrate(() => instrumentManager.calibrateLo(progress, dataTable, instrumentManager.loParameters, token), token);
        }

        private void btnCancelCalibrationClicked(object sender, RoutedEventArgs e) {
            if (calibrationTask != null && !calibrationTask.IsCompleted) {
                calibrationTokenSource.Cancel();
            }
//            log("cancel calibrate", false);
        }

        // measure buttons
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

        private void btnCancelMeasureClicked(object sender, RoutedEventArgs e) {
            if (measureTask != null && !measureTask.IsCompleted) {
                measureTokenSource.Cancel();
            }
        }
        // main window closed
        private void onMainWindowClose(object sender, System.ComponentModel.CancelEventArgs e) {
            try {
                if (logFileInfo.Exists) {
                    using (var f = logFileInfo.AppendText()) {
                        f.Write("\n------------------------------------------------------------------------");
                        f.Write("\nEnd session: " + DateTime.Today.ToString("yyyy-MM-dd") +
                                DateTime.Now.ToShortTimeString());
                        f.Write("\n========================================================================");
                    }
                }
            }
            catch (Exception) {
//                Environment.Exit(1);
            }
        }

#endregion regUiEvents

#region regUtility

        public void log(string mes, bool onlyVerbose = false) {
            if (!onlyVerbose | onlyVerbose & verbose) {
                string logStr = "\n" + "[" + DateTime.Now.ToShortTimeString() + "]: " + mes;
                Dispatcher.Invoke((Action<string>)delegate(string str) {
                    textLog.Text += str;
                    scrollviewer.ScrollToBottom();
                    if (logFileInfo.Exists) {
                        using (var f = File.AppendText(logFileInfo.FullName)) {
                            f.Write(str);
                        }
                    }
                },  logStr);
            }
        }

#endregion regUtility

#region regDataManager

        public static DataTable getDataTableFromExcel(string path, bool hasHeader = true) {
            using (var package = new ExcelPackage()) {
                using (var stream = File.OpenRead(path)) {
                    package.Load(stream);
                }
                var ws = package.Workbook.Worksheets.First();
                DataTable dataTable = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column]) {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }
                var startRow = hasHeader ? 2 : 1;   //!!! WOW
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++) {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dataTable.Rows.Add();
                    foreach (var cell in wsRow) {
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

        public async void calibrate(Action func, CancellationToken token) {
            btnCancelCalibration.Visibility = Visibility.Visible;
            btnCalibrateIn.Visibility = Visibility.Hidden;
            btnCalibrateOut.Visibility = Visibility.Hidden;
            btnCalibrateLo.Visibility = Visibility.Hidden;
            log("start calibrate: " + mode);
            var stopwatch = Stopwatch.StartNew();

            dataTable = ((DataView) dataGrid.ItemsSource).ToTable();

            try {
                calibrationTask = Task.Run(func, token);
                dataGrid.ItemsSource = dataTable.AsDataView();
                await calibrationTask;
            }
            catch (Exception ex) {
                log("error: " + ex.Message);
            }
            finally {
                stopwatch.Stop();
                log("end calibrate, run time: " + Math.Round(stopwatch.Elapsed.TotalMilliseconds / 1000, 2) + " sec", false);
                sndAlert.Play();
                MessageBox.Show("Done.");
                btnCancelCalibration.Visibility = Visibility.Hidden;
                btnCalibrateIn.Visibility = Visibility.Visible;
                btnCalibrateOut.Visibility = Visibility.Visible;
                btnCalibrateLo.Visibility = Visibility.Visible;
            }
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
            btnMeasure.Visibility = Visibility.Hidden;
            btnCancelMeasure.Visibility = Visibility.Visible;
            log("start measure: " + mode);
            var stopwatch = Stopwatch.StartNew();
//            dataGrid.ScrollIntoView();
            dataTable = ((DataView)dataGrid.ItemsSource).ToTable();

            measureTokenSource = new CancellationTokenSource();
            CancellationToken token = measureTokenSource.Token;

            var progress = progressHandler as IProgress<double>;
            try {
                switch (mode) {
                    case MeasureMode.modeDSBDown:
                        measureTask = Task.Run(() =>
                            instrumentManager.measure_mix_DSB_down(progress, dataTable, token), token);
                        break;
                    case MeasureMode.modeDSBUp:
                        measureTask = Task.Run(() =>
                            instrumentManager.measure_mix_DSB_up(progress, dataTable, token), token);
                        break;
                    case MeasureMode.modeSSBDown:
                        measureTask = Task.Run(() => instrumentManager.measure_mix_SSB_down(progress, dataTable, token), token);
                        break;
                    case MeasureMode.modeSSBUp:
                        measureTask = Task.Run(() => instrumentManager.measure_mix_SSB_up(progress, dataTable, token), token);
                        break;
                    case MeasureMode.modeMultiplier:
                        measureTask = Task.Run(() => instrumentManager.measure_mult(progress, dataTable, token), token);
                        break;
                    default:
                        return;
                }
                dataGrid.ItemsSource = dataTable.AsDataView();
                await measureTask;
            }
            catch (Exception ex) {
                log("error: " + ex.Message);
            }
            finally {
                stopwatch.Stop();
                log("end measure, run time: " + Math.Round(stopwatch.Elapsed.TotalMilliseconds / 1000, 2) + " sec", false);
                sndAlert.Play();
                MessageBox.Show("Task complete.");
                btnMeasure.Visibility = Visibility.Visible;
                btnCancelMeasure.Visibility = Visibility.Hidden;
            }
        }

#endregion regMeasurementManager
    }
}
