#define mock

using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Resources;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Serialization;
using Agilent.CommandExpert.ScpiNet.Ag90x0_SA_A_08_03.SCPI.CONFigure.SANalyzer;
using OfficeOpenXml.FormulaParsing.Logging;

namespace Mixer {

    public struct ParameterStruct {
        public string colFreq;
        public string colPow;
        public string colPowGoal;
    }

    class InstrumentManager {

#region regDataMembers

        // instrument class registry
        private Dictionary<string, Func<string, string, Instrument>> instrumentRegistry =
            new Dictionary<string, Func<string, string, Instrument>> {
                                                                         {
                                                                             "N9030A",
                                                                             (loc, idn) => new Analyzer(loc, idn)
                                                                         }, {
                                                                             "GEN",
                                                                             (loc, idn) => new Generator(loc, idn)
                                                                         }, {
                                                                             "LO",
                                                                             (loc, idn) => new Generator(loc, idn)
                                                                         }
                                                                     };
        
        // instrument list
        public List<IInstrument> listInstruments = new List<IInstrument>();

        // parameter lists
        public Dictionary<MeasureMode, List<Tuple<string, string>>> outParameters;
        public Dictionary<MeasureMode, ParameterStruct> inParameters;
        public ParameterStruct loParameters;

        public IInstrument m_IN { get; set; }
        public IInstrument m_OUT { get; set; }
        public IInstrument m_LO { get; set; }

        public int delay { get; set; } = 300;
        public decimal attenuation { get; set; } = 30;
        public decimal maxfreq { get; set; } = 26500*(decimal)Constants.MHz;
        public decimal span { get; set; } = 10*(decimal)Constants.MHz;

        private Action<string, bool> log;   // TODO: make logger class

#endregion regDataMembers

        public InstrumentManager(Action<string, bool> logger) {
            log = logger;
            initParameterLists();
        }

        private void initParameterLists() {

            // IN calibration parameters
            inParameters = new Dictionary<MeasureMode, ParameterStruct> {
                { MeasureMode.modeDSBDown,    new ParameterStruct { colFreq = "FRF", colPow = "PRF", colPowGoal = "PRF-GOAL" } },
                { MeasureMode.modeDSBUp,      new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" } },
                { MeasureMode.modeSSBDown,    new ParameterStruct { colFreq = "FUSB", colPow = "PUSB", colPowGoal = "PUSB-GOAL" } },
                { MeasureMode.modeSSBUp,      new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" } },
                { MeasureMode.modeMultiplier, new ParameterStruct { colFreq = "FH1", colPow = "PIN", colPowGoal = "PIN-GOAL" } }
            };

            // LO calibration parameters
            loParameters = new ParameterStruct { colFreq = "FLO", colPow = "PLO", colPowGoal = "PLO-GOAL" };

            // OUT calibration parameters
            outParameters = new Dictionary<MeasureMode, List<Tuple<string, string>>>();

            var dsbdownup = new List<Tuple<string, string>> {
                new Tuple<string, string>("FIF", "ATT-IF"),
                new Tuple<string, string>("FRF", "ATT-RF"),
                new Tuple<string, string>("FLO", "ATT-LO")
            };
            outParameters.Add(MeasureMode.modeDSBDown, dsbdownup);
            outParameters.Add(MeasureMode.modeDSBUp, dsbdownup);

            var ssbdownup = new List<Tuple<string, string>> {
                new Tuple<string, string>("FIF", "ATT-IF"),
                new Tuple<string, string>("FLSB", "ATT-LSB"),
                new Tuple<string, string>("FUSB", "ATT-USB"),
                new Tuple<string, string>("FLO", "ATT-LO")
            };
            outParameters.Add(MeasureMode.modeSSBDown, ssbdownup);
            outParameters.Add(MeasureMode.modeSSBUp, ssbdownup);

            var multiplier = new List<Tuple<string, string>> {
                new Tuple<string, string>("FH1", "ATT-H1"),
                new Tuple<string, string>("FH1", "ATT-H2"),
                new Tuple<string, string>("FH1", "ATT-H3"),
                new Tuple<string, string>("FH1", "ATT-H4")
            };
            outParameters.Add(MeasureMode.modeMultiplier, multiplier);
        }

#if mock
        public void send(string location, string command) {
            log("send: " + command + " to: " + location, true);
        }
        public string query(string location, string question) {
            log("query: " + question + " to: " + location, true);
            return "-24.9";
        }
#else
        public void send(string location, string command) {
            log("debug: send: " + command + " to: " + location, true);
            try {
                using (AgSCPI99 instrument = new AgSCPI99(location)) {
                    instrument.Transport.Command.Invoke(command);
                }
            }
            catch (Exception ex) {
                throw ex;
            }
        }
        public string query(string location, string question) {
            log("debug: query: " + question + " to: " + location, true);
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
#endif

#region regInstrumentControl
        private string testLocation(string location) {
#if mock
            var rnd = new Random();
            var lst = new List<string> {
                                           "Agilent Technoligies,N9030A,MY49432146,A.11.04",
                                           "AAAA,GEN,1111",
                                           "BBBB,LO,2222",
                                       };
            return lst[rnd.Next(lst.Count)];
#else
            try {
                using (var instrument = new AgSCPI99(location)) {
                    instrument.Transport.Query.Invoke("*IDN?", out var answer);
                    return answer;
                }
            }
            catch (Exception ex) {
                log("search error: " + ex.Message, false);
            }
            return "";
#endif
        }
        
        public void searchInstruments(IProgress<double> prog, int maxPort, int gpib, CancellationToken token) {
            log("start instrument search...", false);
            for (int i = 0; i <= maxPort; i++) {
                if (token.IsCancellationRequested) {
                    log("error: task aborted", false);
                    return;
                }

                string location = "GPIB" + gpib + "::" + i + "::INSTR";
                log("trying " + location, false);

                // TODO: query dummy instrument
                // TODO: make a factory, which queries the ports through a dummy instrument and creates and returns and appropriate instrument instance?
                string idn = testLocation(location); 
                if (!string.IsNullOrEmpty(idn)) {
                    string name = idn.Split(',')[1];
                    listInstruments.Add(instrumentRegistry[name](location, idn));
                }
                log("found " + idn + " at " + location, false);
                prog?.Report((double)i / maxPort * 100);
#if mock
                if (listInstruments.Count == 3) {
                    break;
                }
#endif
            }
            prog?.Report(100);
            if (listInstruments.Count == 0) {
                log("error: no instruments found, check connection", false);
                return;
            }
            log("search done, found " + listInstruments.Count + " device(s)", false);
        }

        private bool instPrepareForCalib(Generator GEN, Analyzer SA) {
            try {
                SA.SetAutocalibration(Analyzer.AutocalState.AutocalOff);
                SA.SetFreqSpan(span);
                SA.SetMarkerMode(Analyzer.MarkerMode.ModePos);
                SA.SetPowerAttenuation(attenuation);
                //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString()) // TODO: add this line

                GEN.SetOutputModulation(Generator.OutputModulationState.ModulationOff);
                GEN.SetOutput(Generator.OutputState.OutputOn);                
            }
            catch (Exception ex) {
                log("error: can't prepare instruments: " + ex.Message, false);
                instReleaseFromCalib(GEN, SA);
                return false;
            }
            log("prepare instruments", false);
            return true;
        }

        private bool instPrepareForMeas(Generator GEN, Analyzer SA, Generator LO) {
            if (!instPrepareForCalib(GEN, SA)) 
                return false;
            
            try {
                LO.SetOutputModulation(Generator.OutputModulationState.ModulationOff);
                LO.SetOutput(Generator.OutputState.OutputOn);
            }
            catch (Exception ex) {
                log("error: fail turning LO generator on" + ex.Message, false);
                return false;
            }

            return true;
        }

        private bool instReleaseGen(Generator GEN) {
            try {
                GEN.SetOutput(Generator.OutputState.OutputOff);
            }
            catch (Exception ex) {
                log("error: can't release generator: " + ex.Message, false);
                return false;
            }

            return true;
        }

        private bool instReleaseAnalyzer(Analyzer SA) {
            try {
                SA.SetAutocalibration(Analyzer.AutocalState.AutocalOn);
            }
            catch (Exception ex) {
                log("error: can't release analyzer: " + ex.Message, false);
                return false;
            }

            return true;
        }

        private bool instReleaseFromCalib(Generator GEN, Analyzer SA) {
            if (!instReleaseGen(GEN))
                return false;
            if (!instReleaseAnalyzer(SA))
                return false;

            log("release instruments", false);
            return true;
        }

        private bool instReleaseFromMeas(Generator GEN, Analyzer SA, Generator LO) {
            if (!instReleaseGen(GEN))
                return false;
            if (!instReleaseAnalyzer(SA))
                return false;
            if (!instReleaseGen(LO))
                return false;

            log("release instruments", false);
            return true;
        }

        private bool instSetAnalyzerCenterFreq(Analyzer SA, decimal freq) {
            try {
                SA.SetMeasCenterFreq(freq);
                SA.SetMarker1XCenter(freq);
            }
            catch (Exception ex) {
                log("error: fail setting calibration freq= " + freq + ", skipping (" + ex.Message + ")", false);
                return false;
            }
            return true;
        }

        private bool instSetCalibFreq(Generator GEN, Analyzer SA, decimal freq) {
            if (!instSetAnalyzerCenterFreq(SA, freq))
                return false;
            try {
                GEN.SetSourceFreq(freq);
            }
            catch (Exception ex) {
                log("error: fail setting calibration freq= " + freq + ", skipping (" + ex.Message + ")", false);
                return false;
            }

            log("debug: set calibration freq=" + freq, true);
            return true;
        }

        private bool instSetGenPowGoal(Generator GEN, decimal pow) {
            try {
                GEN.SetSourcePow(pow);
            }
            catch (Exception ex) {
                log("error: fail setting calibration pow=" + pow + ", skipping (" + ex.Message + ")", false);
                return false;
            }

            log("debug: set calibration pow=" + pow, true);
            return true;
        }

        private decimal readPower(Analyzer SA) {
            // this function should always return correct decimal number, error handling is not necessary
            string readpow;
            try {
                readpow = SA.ReadMarker1Y();
            }
            catch (Exception ex) {
                log("error: can't read power: " + ex.Message, false);
                return 0;
            }            
            if (!decimal.TryParse(readpow, NumberStyles.Any, CultureInfo.InvariantCulture,
                                  out var readPowDec)) {
                log("error: can't parse read pow: " + readpow, false);
                return 0;
            }
            log("debug: read data:" + readPowDec, true);
            return readPowDec;
        }

#endregion regInstrumentControl

#region regCalibrations

        public void calibrate(Generator GEN, Analyzer SA, IProgress<double> prog, DataTable data, ParameterStruct paramDict, CancellationToken token) {

            // local helper fuction
            bool convertStrValues(string inFreqStr, string inPowGoalStr, out decimal inFreqDec, out decimal inPowGoalDec) {
                bool success = true;

                // TODO: extract str2dec method if used in other functions
                if (!decimal.TryParse(inFreqStr, NumberStyles.Any, CultureInfo.InvariantCulture,
                                      out inFreqDec)) {
                    log("error: fail parsing " + inFreqStr + ", skipping", false);
                    success = false;
                }

                if (!decimal.TryParse(inPowGoalStr, NumberStyles.Any, CultureInfo.InvariantCulture,
                                      out inPowGoalDec)) {
                    log("error: fail parsing " + inPowGoalStr + ", skipping", false);
                    success = false;
                }

                inFreqDec *= Constants.GHz;
                return success;
            }

            if (!instPrepareForCalib(GEN, SA))
                return;

            // TODO: if performance issue, write own key class, override Equals() and GetHash()
            var cache = new Dictionary<Tuple<decimal, decimal>, Tuple<decimal, decimal>>();

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    log("error: task aborted", false);
                    instReleaseFromCalib(GEN, SA);
                    return;
                }

                string inFreqStr = row[paramDict.colFreq].ToString().Replace(',', '.');
                string inPowGoalStr = row[paramDict.colPowGoal].ToString().Replace(',', '.');

                if (string.IsNullOrEmpty(inFreqStr) || inFreqStr == "-" ||
                    string.IsNullOrEmpty(inPowGoalStr) || inPowGoalStr == "-") {
                    continue;
                }

                if (!convertStrValues(inFreqStr, inPowGoalStr, out var inFreqDec, out var inPowGoalDec))
                    continue;

                log("debug: calibrate: freq=" + inFreqDec + " pgoal=" + inPowGoalDec, true);

                var freqPowPair = Tuple.Create(inFreqDec, inPowGoalDec);

                if (!cache.ContainsKey(freqPowPair)) {
                    decimal tempPowDec = inPowGoalDec;
                    decimal err = 1;

                    if (!instSetCalibFreq(GEN, SA, inFreqDec)) {
                        continue;
                    }

                    int count = 0;
                    int tmpDelay = delay;
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        if (!instSetGenPowGoal(GEN, tempPowDec))
                            break;

                        Thread.Sleep(tmpDelay);

                        err = inPowGoalDec - readPower(SA);   
                        tempPowDec += err;                        
                        ++count;
                        tmpDelay += 50;
                    }
                    cache.Add(freqPowPair, Tuple.Create(tempPowDec, err));
                }
                var powErrPair = cache[freqPowPair];
                row[paramDict.colPow] = powErrPair.Item1.ToString(Constants.decimalFormat).Replace('.', ',');
                row["ERR"] = powErrPair.Item2.ToString(Constants.decimalFormat).Replace('.', ',');

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instReleaseFromCalib(GEN, SA);
            prog?.Report(100);
        }

        public void calibrateIn(IProgress<double> prog, DataTable data, ParameterStruct paramDict, CancellationToken token) {
            calibrate((Generator)m_IN, (Analyzer)m_OUT, prog, data, paramDict, token);
        }

        public void calibrateLo(IProgress<double> prog, DataTable data, ParameterStruct paramDict, CancellationToken token) {
            calibrate((Generator)m_LO, (Analyzer)m_OUT, prog, data, paramDict, token);
        }

        private string getAttenuationError(Generator GEN, Analyzer SA, string freq, decimal powGoal, int harmonic = 1) {
            if (string.IsNullOrEmpty(freq) || freq == "-") {
                return "-";
            }

            // TODO: extract str2dec method, move logging there
            if (!decimal.TryParse(freq.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture,
                                  out var inFreqDec)) {
                log("error: fail parsing " + freq.Replace(",", "."), false);
                return "-";
            }
            inFreqDec *= Constants.GHz;

            if (inFreqDec * harmonic > maxfreq) {
                log("error: calibrate fail: frequency is out of limits: " + inFreqDec, false);
                return "-";
            }

            if (!instSetCalibFreq(GEN, SA, inFreqDec * harmonic))
                return "-";

            Thread.Sleep(delay);

            decimal errDec = powGoal - readPower(SA);

            if (errDec < 0)
                errDec = 0;

            return errDec.ToString("0.000", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void calibrateOut(IProgress<double> prog, DataTable data, List<Tuple<string, string>> parameters, MeasureMode mode, CancellationToken token) {
            // TODO: fail whole row on any error
            var GEN = (Generator)m_IN;
            var SA = (Analyzer)m_OUT;

            instPrepareForCalib(GEN, SA);

            const decimal tempPow = (decimal)-20.00;
            if (!instSetGenPowGoal(GEN, tempPow))
                return;

            var cache = new Dictionary<string, string>();

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    log("error: task aborted", false);
                    instReleaseFromCalib(GEN, SA);
                    return;
                }

                int harmonic = 1;   // hack

                foreach (var p in parameters) {
                    string freq = row[p.Item1].ToString();

                    if (!cache.ContainsKey(freq)) {
                        cache.Add(freq, getAttenuationError(GEN, SA, row[p.Item1].ToString(), tempPow, harmonic));
                    }

                    row[p.Item2] = cache[freq];

                    if (mode == MeasureMode.modeMultiplier)
                        ++harmonic;
                }

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            instReleaseFromCalib(GEN, SA);
            prog?.Report(100);
        }

#endregion regCalibrations

#region regMeasurement

        public void measurePower(DataRow row, Analyzer SA, decimal powGoal, decimal freq, string colAtt, string colPow, string colConv, int coeff, int corr) {
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

            if (!instSetAnalyzerCenterFreq(SA, freq)) {
                row[colPow]  = "-";
                row[colConv] = "-";
                return;
            }

            Thread.Sleep(delay);

            decimal.TryParse(attStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var att);

            decimal readPow = readPower(SA);

            decimal diff = coeff * (powGoal - att - readPow + corr);

            row[colPow] = readPow.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
            row[colConv] = diff.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void measure_mix_DSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string GET = m_LO.Location;

            if (!instPrepareForMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
                    return;
                }

                // TODO: convert do decimal?
                string inPowLO = row["PLO"].ToString().Replace(',', '.');
                string inPowRF = row["PRF"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLO) || inPowLO == "-" ||
                    string.IsNullOrEmpty(inPowRF) || inPowRF == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                // TODO: exctract to converson method
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowLOGoalDec);
                decimal.TryParse(row["PRF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowRFGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqIFDec);
                inFreqLODec *= Constants.GHz;
                inFreqRFDec *= Constants.GHz;
                inFreqIFDec *= Constants.GHz;

                // TODO: write "-" into corresponding column on fail
                try {
                    send(GEN, "SOUR:FREQ " + inFreqRFDec);
                    send(GET, "SOUR:FREQ " + inFreqLODec);
                }
                catch (Exception ex) {
                    log("error: measure fail setting freq, skipping row: " + ex.Message, false);
                    continue;
                }
                try {
                    send(GEN, "SOUR:POW " + inPowRF);
                    send(GET, "SOUR:POW " + inPowLO);
                }
                catch (Exception ex) {
                    log("error: measure fail setting pow, skipping row: " + ex.Message, false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowRFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowRFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "ISO-RF", 1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
            // TODO: move sends to instrumentManager
            send(GET, "OUTP:STAT OFF");

            prog?.Report(100);
        }

        public void measure_mix_DSB_up(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            instPrepareForCalib((Generator)m_IN, (Analyzer)m_OUT);
            send(GET, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(GET, "OUTP:STAT OFF");
                    instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
                    token.ThrowIfCancellationRequested();
                    return;
                }
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

                decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowIFGoalDec);
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowLOGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqIFDec);
                inFreqIFDec *= Constants.GHz;
                inFreqRFDec *= Constants.GHz;
                inFreqLODec *= Constants.GHz;

                // TODO: extract method
                try {
                    send(GET, "SOUR:FREQ " + inFreqLODec);
                    send(GEN, "SOUR:FREQ " + inFreqIFDec);
                }
                catch (Exception ex) {
                    log("error: measure fail setting freq, skipping row: " + ex.Message, false);
                    continue;
                }
                try {
                    send(GET, "SOUR:POW " + inPowLO);
                    send(GEN, "SOUR:POW " + inPowIF);
                }
                catch (Exception ex) {
                    log("error: measure fail setting pow, skipping row: " + ex.Message, false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowIFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "CONV", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "ISO-IF", 1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
            send(GET, "OUTP:STAT OFF");
            prog?.Report(100);
        }

        public void measure_mix_SSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            instPrepareForCalib((Generator)m_IN, (Analyzer)m_OUT);
            send(GET, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(GET, "OUTP:STAT OFF");
                    instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
                    token.ThrowIfCancellationRequested();
                    return;
                }
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
                    send(GET, "SOUR:FREQ " + inFreqLODec);
                    send(GET, "SOUR:POW " + inPowLOStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting LO params, skipping row (" + ex.Message + ")", false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                try {
                    send(GEN, "SOUR:FREQ " + inFreqLSBDec);
                    send(GEN, "SOUR:POW " + inPowLSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN LSB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowLSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-LSB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowLSBGoalDec, inFreqLSBDec, "ATT-LSB", "POUT-LSB", "ISO-LSB", 1, 0);

                try {
                    send(GEN, "SOUR:FREQ " + inFreqUSBDec);
                    send(GEN, "SOUR:POW " + inPowUSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN USB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowUSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-USB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowUSBGoalDec, inFreqUSBDec, "ATT-USB", "POUT-USB", "ISO-USB", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
            send(GET, "OUTP:STAT OFF");
            prog?.Report(100);
        }

        public void measure_mix_SSB_up(IProgress<double> prog, DataTable data, CancellationToken token) {
            //            string IN = instrumentManager.m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            send(SA, ":CAL:AUTO OFF");
            send(SA, ":SENS:FREQ:SPAN 1000000");
            send(SA, ":CALC:MARK1:MODE POS");
            send(SA, ":POW:ATT " + attenuation);
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            send(GET, "OUTP:STAT ON");
            //send(IN, ("OUTP:STAT ON"));

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(SA, ":CAL:AUTO ON");
                    send(GET, "OUTP:STAT OFF");
                    //send(IN, "OUTP:STAT OFF");
                    log("release instrument", false);
                    token.ThrowIfCancellationRequested();
                    return;
                }
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
                    send(GET, "SOUR:FREQ " + inFreqLO);
                    send(GET, "SOUR:POW " + inPowLOStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting LO params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                //send(IN, ("SOUR:FREQ " + t_freq_IF));
                //send(IN, ("SOUR:POW " + t_pow_IF.Replace(',', '.')));

                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqLSB, "ATT-LSB", "POUT-LSB", "CONV-LSB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqUSB, "ATT-USB", "POUT-USB", "CONV-USB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqIF, "ATT-IF", "POUT-IF", "ISO-IF", 1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoal, inFreqLO, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            send(SA, ":CAL:AUTO ON");
            send(GET, "OUTP:STAT OFF");
            //send(IN, "OUTP:STAT OFF");
            log("release instrument", false);
            prog?.Report(100);
        }

        public void measure_mult(IProgress<double> prog, DataTable data, CancellationToken token) {
            string IN = m_IN.Location;
            string OUT = m_OUT.Location;

            instPrepareForCalib((Generator)m_IN, (Analyzer)m_OUT);

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
                    token.ThrowIfCancellationRequested();
                    return;
                }
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
                    send(IN, "SOUR:POW " + inPowGenStr);
                    send(IN, "SOUR:FREQ " + inFreqH1);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN params, skipping row (" + ex.Message + ")", false);
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowGoal, inFreqH1 * 1, "ATT-H1", "POUT-H1", "CONV-H1", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowGoal, inFreqH1 * 2, "ATT-H2", "POUT-H2", "CONV-H2", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowGoal, inFreqH1 * 3, "ATT-H3", "POUT-H3", "CONV-H3", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowGoal, inFreqH1 * 4, "ATT-H4", "POUT-H4", "CONV-H4", -1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
            prog?.Report(100);
        }

#endregion regMeasurement
    }
}
