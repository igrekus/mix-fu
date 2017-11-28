//#define mock

using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading;

namespace Mixer {

    public struct ParameterStruct {
        public string colFreq;
        public string colPow;
        public string colPowGoal;
    }

    class InstrumentManager {
        // parameter lists
        public Dictionary<MeasureMode, List<Tuple<string, string>>> outParameters;
        public Dictionary<MeasureMode, ParameterStruct> inParameters;
        public ParameterStruct loParameters;

        // TODO: make instrument classes: Generator, Analyser, hide queries there into methods
        public Instrument m_IN { get; set; }
        public Instrument m_OUT { get; set; }
        public Instrument m_LO { get; set; }

        public int delay { get; set; } = 300;
        public decimal attenuation { get; set; } = 30;
        public decimal maxfreq { get; set; } = 26500*(decimal)Constants.MHz;
        public decimal span { get; set; } = 10*(decimal)Constants.MHz;

        private Action<string, bool> log;   // TODO: make logger class

        public InstrumentManager(Action<string, bool> logger) {
            log = logger;
            initParameterLists();
        }

        public InstrumentManager(Action<string, bool> logger, Instrument IN, Instrument OUT, Instrument LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
            log = logger;
            initParameterLists();
        }

        private void initParameterLists() {
            // IN calibration parameters
            inParameters = new Dictionary<MeasureMode, ParameterStruct> {
                {MeasureMode.modeDSBDown,
                 new ParameterStruct {colFreq = "FRF", colPow = "PRF", colPowGoal = "PRF-GOAL"}},
                {MeasureMode.modeDSBUp,
                 new ParameterStruct {colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL"}},
                {MeasureMode.modeSSBDown,
                 new ParameterStruct {colFreq = "FUSB", colPow = "PUSB", colPowGoal = "PUSB-GOAL"}},
                {MeasureMode.modeSSBUp,
                 new ParameterStruct {colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL"}},
                {MeasureMode.modeMultiplier,
                 new ParameterStruct {colFreq = "FH1", colPow = "PIN", colPowGoal = "PIN-GOAL"}}
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

        public void setInstruments(Instrument IN, Instrument OUT, Instrument LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
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
        public void searchInstruments(IProgress<double> prog, List<Instrument> instruments, int maxPort, int gpib, CancellationToken token) {
            log("start instrument search...", false);

            for (int i = 0; i <= maxPort; i++) {
                if (token.IsCancellationRequested) {
                    token.ThrowIfCancellationRequested();
                    return;
                }
                string location = "GPIB" + gpib + "::" + i + "::INSTR";
                log("try " + location, false);

                try {
                    string idn = query(location, "*IDN?");
                    instruments.Add(new Instrument { Location = location, Name = idn.Split(',')[1], FullName = idn });
                    log("found " + idn + " at " + location, false);
                }
                catch (Exception ex) {
                    log(ex.Message, true);
                }

                prog?.Report((double)i / maxPort * 100);
            }
            prog?.Report(100);
            if (instruments.Count == 0) {
                log("error: no instruments found, check connection", false);
                return;
            }
            log("search done, found " + instruments.Count + " device(s)", false);
        }

        public void prepareInstrument(string GEN, string SA) {
            // TODO: check SA & GEN assignment logic
            send(SA, ":CAL:AUTO OFF");                       // выключаем автокалибровку анализатора спектра
            send(SA, ":SENS:FREQ:SPAN " + span);             // выставляем спан
            send(SA, ":CALC:MARK1:MODE POS");                // выставляем режим маркера
            send(SA, ":POW:ATT " + attenuation);             // выставляем аттенюацию
            send(GEN, "OUTP:STAT ON");                       // включаем генератор
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString())
            // TODO: send GEN modulation off
            log("capture instrument", false);
        }

        private void setCalibrationFreq(string GEN, string SA, decimal inFreqDec) {
            // TODO: bind to instrument properties
            // exceptions handled in the calling method
            send(GEN, "SOUR:FREQ " + inFreqDec);
            send(SA, ":SENSe:FREQuency:RF:CENTer " + inFreqDec);
            send(SA, ":CALCulate:MARKer1:X:CENTer " + inFreqDec);
        }

        public void releaseInstrument(string GEN, string SA) {
            // TODO: check SA & GEN assignment logic
            send(SA, ":CAL:AUTO ON");     //включаем автокалибровку анализатора спектра
            send(GEN, "OUTP:STAT OFF");   //выключаем генератор
            log("release instrument", false);
        }

        public void calibrateIn(IProgress<double> prog, DataTable data, ParameterStruct paramDict, CancellationToken token) {
            // TODO: exception handling
            // TODO: split into methods
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            prepareInstrument(GEN, SA);

            // TODO: if performance issue, write own key class, override Equals() and GetHash()
            var cache = new Dictionary<Tuple<decimal, decimal>, Tuple<decimal, decimal>>();

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    releaseInstrument(GEN, SA);
                    token.ThrowIfCancellationRequested();
                    return;
                }
                string inFreqStr = row[paramDict.colFreq].ToString().Replace(',', '.');
                string inPowGoalStr = row[paramDict.colPowGoal].ToString().Replace(',', '.');

                if (string.IsNullOrEmpty(inFreqStr) || inFreqStr == "-" ||
                    string.IsNullOrEmpty(inPowGoalStr) || inPowGoalStr == "-") {
                    continue;
                }

                decimal inFreqDec;
                decimal inPowGoalDec;
                if (!decimal.TryParse(inFreqStr, NumberStyles.Any, CultureInfo.InvariantCulture,
                                      out inFreqDec)) {
                    log("error: fail parsing " + inFreqStr + ", skipping", false);
                    continue;
                }
                if (!decimal.TryParse(inPowGoalStr, NumberStyles.Any, CultureInfo.InvariantCulture,
                                      out inPowGoalDec)) {
                    log("error: fail parsing " + inPowGoalStr + ", skipping", false);
                    continue;
                }
                inFreqDec *= Constants.GHz;

                log("debug: calibrate: freq=" + inFreqDec + " pgoal=" + inPowGoalDec, true);

                var freqPowPair = Tuple.Create(inFreqDec, inPowGoalDec);

                if (!cache.ContainsKey(freqPowPair)) {
                    decimal tempPowDec = inPowGoalDec;
                    decimal err = 1;

                    try {
                        setCalibrationFreq(GEN, SA, inFreqDec);
                    }
                    catch (Exception ex) {
                        log("error: calibrate fail setting freq=" + inFreqDec + ", skipping (" + ex.Message + ")", false);
                        continue;
                    }

                    int count = 0;
                    int tmpDelay = delay;
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        // TODO: exception handling
                        try {
                            send(GEN, "SOUR:POW " + tempPowDec);
                        }
                        catch (Exception ex) {
                            log("error: calibrate fail setting pow=" + tempPowDec + ", skipping (" + ex.Message + ")", false);
                            break;
                        }
                        Thread.Sleep(delay);
                        // TODO: inline readPow
                        string readPow = query(SA, ":CALCulate:MARKer:Y?");
                        decimal readPowDec = 0;
                        decimal.TryParse(readPow, NumberStyles.Any,
                                         CultureInfo.InvariantCulture, out readPowDec);
                        log("read data:" + readPow + " " + readPowDec, true);

                        err = inPowGoalDec - readPowDec;
                        tempPowDec += err;
                        
                        ++count;
                        delay += 50;
                    }
                    delay = tmpDelay;
                    cache.Add(freqPowPair, Tuple.Create(tempPowDec, err));
                }
                var powErrPair = cache[freqPowPair];
                // ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',');
                row[paramDict.colPow] = powErrPair.Item1.ToString(Constants.decimalFormat).Replace('.', ',');
                row["ERR"] = powErrPair.Item2.ToString(Constants.decimalFormat).Replace('.', ',');

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            releaseInstrument(GEN, SA);
            prog?.Report(100);
        }

        public void calibrateLo(IProgress<double> prog, DataTable data, ParameterStruct paramDict, CancellationToken token) {
            Instrument tmpIn = m_IN;
            m_IN = m_LO;
            calibrateIn(prog, data, paramDict, token);
            m_IN = tmpIn;
        }

        public string getAttenuationError(string GEN, string SA, string freq, decimal powGoal, int harmonic = 1) {
            if (string.IsNullOrEmpty(freq) || freq == "-") {
                return "-";
            }

            decimal inFreqDec = 0;
            if (!decimal.TryParse(freq.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture,
                                  out inFreqDec)) {
                log("error: fail parsing " + freq.Replace(",", "."), false);
                return "-";
            }
            inFreqDec *= Constants.GHz;

            if (inFreqDec * harmonic > maxfreq) {
                log("error: calibrate fail: frequency is out of limits: " + inFreqDec, false);
                return "-";
            }

            try {
                setCalibrationFreq(GEN, SA, inFreqDec * harmonic);
            }
            catch (Exception ex) {
                log("error: calibrate fail setting freq=" + inFreqDec + " (" + ex.Message + ")", false);
                return "-";
            }
            
            Thread.Sleep(delay);

            string readPow = query(SA, ":CALCulate:MARKer:Y?");

            decimal readPowDec = 0;
            decimal.TryParse(readPow, NumberStyles.Any, CultureInfo.InvariantCulture, out readPowDec);
            decimal errDec = powGoal - readPowDec;

            if (errDec < 0)
                errDec = 0;

            return errDec.ToString("0.000", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void calibrateOut(IProgress<double> prog, DataTable data, List<Tuple<string, string>> parameters, MeasureMode mode, CancellationToken token) {
            // TODO: fail whole row on any error
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            prepareInstrument(GEN, SA);

            const decimal tempPow = (decimal)-20.00;
            try {
                send(GEN, "SOUR:POW " + tempPow);
            }
            catch (Exception ex) {
                log("error: calibrate fail setting pow=" + tempPow + ", aborting (" + ex.Message + ")", false);
                return;
            }

            var cache = new Dictionary<string, string>();

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    releaseInstrument(GEN, SA);
                    token.ThrowIfCancellationRequested();
                    return;
                }
                int harmonic = 1;   // hack

                foreach (var p in parameters) {
                    string freq = row[p.Item1].ToString();

                    if (!cache.ContainsKey(freq)) {
                        cache.Add(freq, getAttenuationError(GEN, SA, row[p.Item1].ToString(), tempPow, harmonic));
                    }

                    row[p.Item2] = cache[freq];

                    if (mode == MeasureMode.modeMultiplier) { 
                        ++harmonic;
                    }
                }

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            releaseInstrument(GEN, SA);
            prog?.Report(100);
        }

        public void measurePower(DataRow row, string SA, decimal powGoal, decimal freq, string colAtt, string colPow, string colConv, int coeff, int corr) {
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
                send(SA, ":SENSe:FREQuency:RF:CENTer " + freq);
                send(SA, ":CALCulate:MARKer1:X:CENTer " + freq);
            }
            catch (Exception ex) {
                log("error: measure fail setting freq: " + ex.Message, false);
                row[colPow] = "-";
                row[colConv] = "-";
                return;
            }
            Thread.Sleep(delay);

            decimal att = 0;
            decimal.TryParse(attStr, NumberStyles.Any, CultureInfo.InvariantCulture, out att);

            decimal readPow = 0;
            try {
                decimal.TryParse(query(SA, ":CALCulate:MARKer:Y?"), NumberStyles.Any,
                    CultureInfo.InvariantCulture, out readPow);
            }
            catch (Exception ex) {
                log("error: " + ex.Message, false);
            }
            decimal diff = coeff * (powGoal - att - readPow + corr);

            row[colPow] = readPow.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
            row[colConv] = diff.ToString(Constants.decimalFormat, CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void measure_mix_DSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            // TODO: move all sends
            prepareInstrument(GEN, SA);
            send(GET, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(GET, "OUTP:STAT OFF");
                    releaseInstrument(GEN, SA);
                    token.ThrowIfCancellationRequested();
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

                // TODO: check for empty columns:
                decimal inPowLOGoalDec = 0;
                decimal inPowRFGoalDec = 0;
                decimal inFreqIFDec = 0;
                decimal inFreqRFDec = 0;
                decimal inFreqLODec = 0;

                // TODO: exception handling
                // TODO: need to check for empty cells?
                decimal.TryParse(row["PLO-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowLOGoalDec);
                decimal.TryParse(row["PRF-GOAL"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inPowRFGoalDec);
                decimal.TryParse(row["FLO"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqLODec);
                decimal.TryParse(row["FRF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqRFDec);
                decimal.TryParse(row["FIF"].ToString().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out inFreqIFDec);
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

                measurePower(row, SA, inPowRFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV", -1, 0);
                measurePower(row, SA, inPowRFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "ISO-RF", 1, 0);
                measurePower(row, SA, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }

            releaseInstrument(GEN, SA);
            // TODO: move sends to instrumentManager
            send(GET, "OUTP:STAT OFF");

            prog?.Report(100);
        }

        public void measure_mix_DSB_up(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            prepareInstrument(GEN, SA);
            send(GET, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(GET, "OUTP:STAT OFF");
                    releaseInstrument(GEN, SA);
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

                measurePower(row, SA, inPowIFGoalDec, inFreqRFDec, "ATT-RF", "POUT-RF", "CONV", -1, 0);
                measurePower(row, SA, inPowIFGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "ISO-IF", 1, 0);
                measurePower(row, SA, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            releaseInstrument(GEN, SA);
            send(GET, "OUTP:STAT OFF");
            prog?.Report(100);
        }

        public void measure_mix_SSB_down(IProgress<double> prog, DataTable data, CancellationToken token) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;
            string GET = m_LO.Location;

            prepareInstrument(GEN, SA);
            send(GET, "OUTP:STAT ON");

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    send(GET, "OUTP:STAT OFF");
                    releaseInstrument(GEN, SA);
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
                measurePower(row, SA, inPowLOGoalDec, inFreqLODec, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

                try {
                    send(GEN, "SOUR:FREQ " + inFreqLSBDec);
                    send(GEN, "SOUR:POW " + inPowLSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN LSB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                measurePower(row, SA, inPowLSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-LSB", -1, -3);
                measurePower(row, SA, inPowLSBGoalDec, inFreqLSBDec, "ATT-LSB", "POUT-LSB", "ISO-LSB", 1, 0);

                try {
                    send(GEN, "SOUR:FREQ " + inFreqUSBDec);
                    send(GEN, "SOUR:POW " + inPowUSBStr);
                }
                catch (Exception ex) {
                    log("error: measure: fail setting IN USB params, skipping row (" + ex.Message + ")", false);
                    continue;
                }
                measurePower(row, SA, inPowUSBGoalDec, inFreqIFDec, "ATT-IF", "POUT-IF", "CONV-USB", -1, -3);
                measurePower(row, SA, inPowUSBGoalDec, inFreqUSBDec, "ATT-USB", "POUT-USB", "ISO-USB", 1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            releaseInstrument(GEN, SA);
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

                measurePower(row, SA, inPowIFGoal, inFreqLSB, "ATT-LSB", "POUT-LSB", "CONV-LSB", -1, -3);
                measurePower(row, SA, inPowIFGoal, inFreqUSB, "ATT-USB", "POUT-USB", "CONV-USB", -1, -3);
                measurePower(row, SA, inPowIFGoal, inFreqIF, "ATT-IF", "POUT-IF", "ISO-IF", 1, -3);
                measurePower(row, SA, inPowLOGoal, inFreqLO, "ATT-LO", "POUT-LO", "ISO-LO", 1, 0);

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

            prepareInstrument(IN, OUT);

            int i = 0;
            foreach (DataRow row in data.Rows) {
                if (token.IsCancellationRequested) {
                    releaseInstrument(IN, OUT);
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

                measurePower(row, OUT, inPowGoal, inFreqH1 * 1, "ATT-H1", "POUT-H1", "CONV-H1", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1 * 2, "ATT-H2", "POUT-H2", "CONV-H2", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1 * 3, "ATT-H3", "POUT-H3", "CONV-H3", -1, 0);
                measurePower(row, OUT, inPowGoal, inFreqH1 * 4, "ATT-H4", "POUT-H4", "CONV-H4", -1, 0);

                prog?.Report((double)i / data.Rows.Count * 100);
                ++i;
            }
            releaseInstrument(IN, OUT);
            prog?.Report(100);
        }
    }
}
