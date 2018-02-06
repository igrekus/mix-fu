#define mock

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading;
using NationalInstruments.VisaNS;

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
            new Dictionary<string, Func<string, string, Instrument>> { { "N9030A", (loc, idn) => new Analyzer(loc, idn) },
                                                                       { "GEN", (loc, idn)    => new Generator(loc, idn) },
                                                                       { "LO", (loc, idn)     => new Generator(loc, idn) } };
        
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

#region regInstrumentControl
        private string testLocation(string location) {
#if mock
            var rnd = new Random();
            var lst = new List<string> { "Agilent Technoligies,N9030A,MY49432146,A.11.04",
                                         "AAAA,GEN,1111",
                                         "BBBB,LO,2222" };
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

        private bool instPrepareAnalyzer(Analyzer SA) {
            try {
                SA.SetAutocalibration(Analyzer.AutocalState.AutocalOff);
                SA.SetFreqSpan(span);
                SA.SetMarkerMode(Analyzer.MarkerMode.ModePos);
                SA.SetPowerAttenuation(attenuation);
                //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString()) // TODO: add this line
            }
            catch (Exception ex) {
                log("error: can't prepare instruments: " + ex.Message, false);
                instReleaseAnalyzer(SA);
                return false;
            }

            log("prepare analyzer=" + SA, false);
            return true;
        }

        private bool instPrepareGenerator(Generator GEN) {
            try {
                GEN.SetOutputModulation(Generator.OutputModulationState.ModulationOff);
                GEN.SetOutput(Generator.OutputState.OutputOn);
            }
            catch (Exception ex) {
                log("error: can't prepare instruments: " + ex.Message, false);
                instReleaseGen(GEN);
                return false;
            }

            log("prepare gen=" + GEN, false);
            return true;
        }

        private bool instPrepareForCalib(Generator GEN, Analyzer SA) {
            return instPrepareAnalyzer(SA) && instPrepareGenerator(GEN);
        }

        private bool instPrepareForMeas(Generator GEN, Analyzer SA, Generator LO) {
            return instPrepareForCalib(GEN, SA) && instPrepareGenerator(LO);
        }

        private bool instPrepareForMeasSSBUp(Analyzer SA, Generator LO) {
            return instPrepareAnalyzer(SA) && instPrepareGenerator(LO);
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

        // TODO: rename this method
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

        // TODO: rename this method
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

        private bool instSetGenPow(Generator GEN, string pow) {
            try {
                GEN.SetSourcePow(pow);
            }
            catch (Exception ex) {
                log("error: fail setting pow=" + pow + " on gen=" + GEN + ", skipping (" + ex.Message + ")", false);
                return false;
            }

            log("debug: set pow=" + pow + " on gen=" + GEN, true);
            return true;
        }

        private bool instSetGenPow(Generator GEN, decimal pow) {
            return instSetGenPow(GEN, pow.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture));
        }

        private bool instSetGenFreq(Generator GEN, string freq) {
            try {
                GEN.SetSourceFreq(freq);
            }
            catch (Exception ex) {
                log("error: fail setting freq=" + freq + " on gen=" + GEN + ", skipping (" + ex.Message + ")", false);
                return false;
            }

            log("debug: set freq=" + freq + " on gen=" + GEN, true);
            return true;
        }

        private bool instSetGenFreq(Generator GEN, decimal freq) {
            return instSetGenFreq(GEN, freq.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture));
        }

        private bool instSetGenFreqPow(Generator GEN, decimal freq, string pow) {
            if (!instSetGenFreq(GEN, freq) || !instSetGenPow(GEN, pow)) {
                log("error: measure fail, skipping row", false);
                return false;
            }

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
                        if (!instSetGenPow(GEN, tempPowDec))
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
            if (!instSetGenPow(GEN, tempPow))
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

            // TODO: replace delay with wait for operaton complete from the instrument
            Thread.Sleep(delay);

            decimal.TryParse(attStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var att);

            decimal readPow = readPower(SA);

            decimal diff = coeff * (powGoal - att - readPow + corr);

            row[colPow] = readPow.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
            row[colConv] = diff.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void measure_mix_DSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            if (!instPrepareForMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
                    log("error: task aborted", false);
                    return;
                }

                // TODO: simplify check (extract method?)
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

                if (!instSetGenFreqPow((Generator)m_IN, inFreqRFDec, inPowRF) ||
                    !instSetGenFreqPow((Generator)m_LO, inFreqLODec, inPowLO)) {
                    // TODO: fix this crap
                    row["POUT-IF"] = "-"; row["CONV"]   = "-";
                    row["POUT-RF"] = "-"; row["ISO-RF"] = "-";
                    row["POUT-LO"] = "-"; row["ISO-LO"] = "-";
                    continue;
                }

                // TODO: remove code dupe
                measurePower(row, (Analyzer)m_OUT, inPowRFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowRFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "ISO-RF", 1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
            prog?.Report(100);
        }

        public void measure_mix_DSB_up(IProgress<double> prog, DataTable data, CancellationToken token) {
            if (!instPrepareForMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
                    log("error: task aborted", false);
                    return;
                }

                // TODO: convert to dec, simplify check
                string inPowLO = row["PLO"].ToString().Replace(',', '.');
                string inPowIF = row["PIF"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLO) || inPowLO == "-" ||
                    string.IsNullOrEmpty(inPowIF) || inPowIF == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowIFGoalDec);
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowLOGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqIFDec);
                inFreqIFDec *= Constants.GHz;
                inFreqRFDec *= Constants.GHz;
                inFreqLODec *= Constants.GHz;

                if (!instSetGenFreqPow((Generator)m_IN, inFreqIFDec, inPowIF) ||
                    !instSetGenFreqPow((Generator)m_LO, inFreqLODec, inPowLO)) {
                    row["POUT-RF"] = "-"; row["CONV"]   = "-";
                    row["POUT-IF"] = "-"; row["ISO-IF"] = "-";
                    row["POUT-LO"] = "-"; row["ISO-LO"] = "-";
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowIFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "CONV", -1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "ISO-IF", 1, 0);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
            prog?.Report(100);
        }

        public void measure_mix_SSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            if (!instPrepareForMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
                    log("error: task aborted", false);
                    return;
                }

                string inPowLOStr = row["PLO"].ToString().Replace(',', '.');
                string inPowLSBStr = row["PLSB"].ToString().Replace(',', '.');
                string inPowUSBStr = row["PUSB"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLOStr) || inPowLOStr == "-" ||
                    string.IsNullOrEmpty(inPowLSBStr) || inPowLSBStr == "-" ||
                    string.IsNullOrEmpty(inPowUSBStr) || inPowUSBStr == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowLOGoalDec);
                decimal.TryParse(row["PUSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowUSBGoalDec);
                decimal.TryParse(row["PLSB-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowLSBGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLODec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqIFDec);
                decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqUSBDec);
                decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLSBDec);
                inFreqLODec *= Constants.GHz;
                inFreqIFDec *= Constants.GHz;
                inFreqLSBDec *= Constants.GHz;
                inFreqUSBDec *= Constants.GHz;

                if (!instSetGenFreqPow((Generator)m_LO, inFreqLODec, inPowLOStr)) {
                    row["POUT-LO"]  = "-"; row["ISO-LO"]   = "-";
                    row["POUT-IF"]  = "-"; row["CONV-LSB"] = "-";
                    row["POUT-LSB"] = "-"; row["ISO-LSB"]  = "-";
                    row["POUT-IF"]  = "-"; row["CONV-USB"] = "-";
                    row["POUT-USB"] = "-"; row["ISO-USB"]  = "-";
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                if (!instSetGenFreqPow((Generator)m_IN, inFreqLSBDec, inPowLSBStr)) {
                    row["POUT-LO"]  = "-"; row["ISO-LO"]   = "-";
                    row["POUT-IF"]  = "-"; row["CONV-LSB"] = "-";
                    row["POUT-LSB"] = "-"; row["ISO-LSB"]  = "-";
                    row["POUT-IF"]  = "-"; row["CONV-USB"] = "-";
                    row["POUT-USB"] = "-"; row["ISO-USB"]  = "-";
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowLSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-LSB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowLSBGoalDec, inFreqLSBDec, "ATT-LSB", "POUT-LSB", "ISO-LSB", 1, 0);

                if (!instSetGenFreqPow((Generator)m_IN, inFreqUSBDec, inPowUSBStr)) {
                    row["POUT-LO"]  = "-"; row["ISO-LO"]   = "-";
                    row["POUT-IF"]  = "-"; row["CONV-LSB"] = "-";
                    row["POUT-LSB"] = "-"; row["ISO-LSB"]  = "-";
                    row["POUT-IF"]  = "-"; row["CONV-USB"] = "-";
                    row["POUT-USB"] = "-"; row["ISO-USB"]  = "-";
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowUSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-USB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowUSBGoalDec, inFreqUSBDec, "ATT-USB", "POUT-USB", "ISO-USB", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
            prog?.Report(100);
        }

        public void measure_mix_SSB_up(IProgress<double> prog, DataTable data, CancellationToken token) {
            if (!instPrepareForMeasSSBUp((Analyzer)m_OUT, (Generator)m_LO))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
                    log("error: task aborted", false);
                    return;
                }

                string inPowLOStr = row["PLO"].ToString().Replace(',', '.');
                string inPowIFStr = row["PIF"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowLOStr) || inPowLOStr == "-" ||
                    string.IsNullOrEmpty(inPowIFStr) || inPowIFStr == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal.TryParse(row["PIF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowIFGoal);
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowLOGoal);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLO);
                decimal.TryParse(row["FLSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqLSB);
                decimal.TryParse(row["FUSB"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqUSB);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqIF);
                inFreqLO *= Constants.GHz;
                inFreqLSB *= Constants.GHz;
                inFreqUSB *= Constants.GHz;
                inFreqIF *= Constants.GHz;

                if (!instSetGenFreqPow((Generator)m_LO, inFreqLO, inPowLOStr)) {
                    row["POUT-LSB"] = "-"; row["CONV-LSB"] = "-";
                    row["POUT-USB"] = "-"; row["CONV-USB"] = "-";
                    row["POUT-IF"]  = "-"; row["ISO-IF"]   = "-";
                    row["POUT-LO"]  = "-"; row["ISO-LO"]   = "-";
                    continue;
                }

                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqLSB, "ATT-LSB", "POUT-LSB", "CONV-LSB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqUSB, "ATT-USB", "POUT-USB", "CONV-USB", -1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowIFGoal, inFreqIF, "ATT-IF", "POUT-IF", "ISO-IF", 1, -3);
                measurePower(row, (Analyzer)m_OUT, inPowLOGoal, inFreqLO, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            instReleaseFromMeas((Generator)m_IN, (Analyzer)m_OUT, (Generator)m_LO);
            prog?.Report(100);
        }

        public void measure_mult(IProgress<double> prog, DataTable data, CancellationToken token) {
            if (!instPrepareForCalib((Generator)m_IN, (Analyzer)m_OUT))
                return;

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    instReleaseFromCalib((Generator)m_IN, (Analyzer)m_OUT);
                    log("error: task aborted", false);
                    return;
                }

                string inPowGenStr = row["PIN-GEN"].ToString().Replace(',', '.');
                string inFreqH1Str = row["FH1"].ToString().Replace(',', '.');
                if (string.IsNullOrEmpty(inPowGenStr) || inPowGenStr == "-" ||
                    string.IsNullOrEmpty(inFreqH1Str) || inFreqH1Str == "-") {
                    log("warning: empty row, skipping", false);
                    continue;
                }

                decimal.TryParse(row["PIN-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inPowGoal);
                decimal.TryParse(row["FH1"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var inFreqH1);
                inFreqH1 *= Constants.GHz;

                if (!instSetGenFreqPow((Generator)m_IN, inFreqH1, inPowGenStr)) {
                    row["POUT-H1"] = "-"; row["CONV-H1"] = "-";
                    row["POUT-H2"] = "-"; row["CONV-H2"] = "-";
                    row["POUT-H3"] = "-"; row["CONV-H3"] = "-";
                    row["POUT-H4"] = "-"; row["CONV-H4"] = "-";
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
