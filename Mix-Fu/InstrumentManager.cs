﻿#define mock

using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

// TODO: catch axception on calibration fail

namespace Mix_Fu {

    public struct ParameterStruct {
        public string colFreq;
        public string colPow;
        public string colPowGoal;
    }

    class InstrumentManager {
        // parameter lists
        public Dictionary<MeasureMode, List<Tuple<string, string>>> outParameters = null;
        public Dictionary<MeasureMode, ParameterStruct> inParameters = null;
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
            // INcalibration parameters
            inParameters = new Dictionary<MeasureMode, ParameterStruct>();
            inParameters.Add(MeasureMode.modeDSBDown, new ParameterStruct { colFreq = "FRF", colPow = "PRF", colPowGoal = "PRF-GOAL" });
            inParameters.Add(MeasureMode.modeDSBUp, new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" });
            inParameters.Add(MeasureMode.modeSSBDown, new ParameterStruct { colFreq = "FUSB", colPow = "PUSB", colPowGoal = "PUSB-GOAL" });
            inParameters.Add(MeasureMode.modeSSBUp, new ParameterStruct { colFreq = "FIF", colPow = "PIF", colPowGoal = "PIF-GOAL" });
            inParameters.Add(MeasureMode.modeMultiplier, new ParameterStruct { colFreq = "FH1", colPow = "PIN", colPowGoal = "PIN-GOAL" });

            // LO calibration parameters
            loParameters = new ParameterStruct { colFreq = "FLO", colPow = "PLO", colPowGoal = "PLO-GOAL" };

            // OUT calibration parameters
            outParameters = new Dictionary<MeasureMode, List<Tuple<string, string>>>();

            List<Tuple<string, string>> dsbdownup = new List<Tuple<string, string>>();
            dsbdownup.Add(new Tuple<string, string>("FIF", "ATT-IF"));
            dsbdownup.Add(new Tuple<string, string>("FRF", "ATT-RF"));
            dsbdownup.Add(new Tuple<string, string>("FLO", "ATT-LO"));
            outParameters.Add(MeasureMode.modeDSBDown, dsbdownup);
            outParameters.Add(MeasureMode.modeDSBUp, dsbdownup);

            List<Tuple<string, string>> ssbdownup = new List<Tuple<string, string>>();
            ssbdownup.Add(new Tuple<string, string>("FIF", "ATT-IF"));
            ssbdownup.Add(new Tuple<string, string>("FLSB", "ATT-LSB"));
            ssbdownup.Add(new Tuple<string, string>("FUSB", "ATT-USB"));
            ssbdownup.Add(new Tuple<string, string>("FLO", "ATT-LO"));
            outParameters.Add(MeasureMode.modeSSBDown, ssbdownup);
            outParameters.Add(MeasureMode.modeSSBUp, ssbdownup);

            List<Tuple<string, string>> multiplier = new List<Tuple<string, string>>();
            multiplier.Add(new Tuple<string, string>("FH1", "ATT-H1"));
            multiplier.Add(new Tuple<string, string>("FH1", "ATT-H2"));
            multiplier.Add(new Tuple<string, string>("FH1", "ATT-H3"));
            multiplier.Add(new Tuple<string, string>("FH1", "ATT-H4"));
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
        public void searchInstruments(List<Instrument> instruments, int maxPort, int gpib, CancellationToken token) {
            log("start instrument search...", false);

            for (int i = 0; i <= maxPort; i++) {
                if (token.IsCancellationRequested) {
                    log("error: aborted on user request", false);
                    token.ThrowIfCancellationRequested();
                    return;
                }
                string location = "GPIB" + gpib.ToString() + "::" + i.ToString() + "::INSTR";
                log("try " + location, false);

                try {
                    string idn = query(location, "*IDN?");
                    instruments.Add(new Instrument { Location = location, Name = idn.Split(',')[1], FullName = idn });
                    log("found " + idn + " at " + location, false);
                }
                catch (Exception ex) {
                    log(ex.Message, true);
                }
            }
            if (instruments.Count == 0) {
                log("error: no instruments found, check connection", false);
                return;
            }
            log("search done, found " + instruments.Count + " device(s)", false);
        }

        public void prepareInstrument(string GEN, string SA) {
            // TODO: check SA & GEN assignment logic
            send(SA, ":CAL:AUTO OFF");                       // выключаем автокалибровку анализатора спектра
            send(SA, ":SENS:FREQ:SPAN " + span.ToString());  // выставляем спан
            send(SA, ":CALC:MARK1:MODE POS");                // выставляем режим маркера
            send(SA, ":POW:ATT " + attenuation.ToString());  // выставляем аттенюацию
            send(GEN, "OUTP:STAT ON");                       // включаем генератор
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());

            // TODO: send GEN modulation off
        }

        private void setCalibrationFreq(string GEN, string SA, decimal inFreqDec) {
            // TODO: bind to instrument properties
            // TODO: exception handling
            send(GEN, "SOUR:FREQ " + inFreqDec);
            send(SA, ":SENSe:FREQuency:RF:CENTer " + inFreqDec);
            send(SA, ":CALCulate:MARKer1:X:CENTer " + inFreqDec);
        }

        public void releaseInstrument(string GEN, string SA) {
            // TODO: check SA & GEN assignment logic
            send(SA, ":CAL:AUTO ON");     //включаем автокалибровку анализатора спектра
            send(GEN, "OUTP:STAT OFF");   //выключаем генератор
        }

        public void calibrateIn(DataTable data, ParameterStruct paramDict, CancellationToken token) {
            // TODO: exception handling
            // TODO: split into methods
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            log("start calibrate IN: " + "GEN=" + GEN + " SA=" + SA, false);
            prepareInstrument(GEN, SA);

            // TODO: if performance issue, write own key class, override Equals() and GetHash()
            var cache = new Dictionary<Tuple<decimal, decimal>, Tuple<decimal, decimal>>();

            foreach (DataRow row in data.Rows) {
                string inFreqStr = row[paramDict.colFreq].ToString().Replace(',', '.');
                string inPowGoalStr = row[paramDict.colPowGoal].ToString().Replace(',', '.');

                if (string.IsNullOrEmpty(inFreqStr) || inFreqStr == "-" ||
                    string.IsNullOrEmpty(inPowGoalStr) || inPowGoalStr == "-") {
                    continue;
                }

                decimal inFreqDec = 0;
                decimal inPowGoalDec = 0;
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
                        log("error: fail setting freq, skipping: " + inFreqDec + " message: " + ex.Message, false);
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
                            log("error: fail setting pow, skipping: " + tempPowDec + " message: " + ex.Message, false);
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
            }
            releaseInstrument(GEN, SA);
            log("end calibrate IN", false);
        }

        public void calibrateLo(DataTable data, ParameterStruct paramDict, CancellationToken token) {
            Instrument tmpIn = m_IN;
            m_IN = m_LO;
            calibrateIn(data, paramDict, token);
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
                log("error: frequency is out of limits", false);
                return "-";
            }

            setCalibrationFreq(GEN, SA, inFreqDec * harmonic);
            Thread.Sleep(delay);

            string readPow = query(SA, ":CALCulate:MARKer:Y?");

            decimal readPowDec = 0;
            decimal.TryParse(readPow, NumberStyles.Any, CultureInfo.InvariantCulture, out readPowDec);
            decimal errDec = powGoal - readPowDec;

            if (errDec < 0) errDec = 0;

            return errDec.ToString("0.000", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void calibrateOut(DataTable data, List<Tuple<string, string>> parameters, MeasureMode mode) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            log("start calibrate IN: " + "GEN=" + GEN + " SA=" + SA, false);

            prepareInstrument(GEN, SA);

            decimal tempPow = (decimal)-20.00;
            send(GEN, "SOUR:POW " + tempPow);

            var cache = new Dictionary<string, string>();

            foreach (DataRow row in data.Rows) {
                int harmonic = 1;   // hack
                foreach (Tuple<string, string> p in parameters) {
                    string freq = row[p.Item1].ToString();

                    if (!cache.ContainsKey(freq)) {
                        cache.Add(freq, getAttenuationError(GEN, SA, row[p.Item1].ToString(), tempPow, harmonic));
                    }

                    row[p.Item2] = cache[freq];

                    if (mode == MeasureMode.modeMultiplier) { 
                        ++harmonic;
                    }
                }
            }

            releaseInstrument(GEN, SA);
        }

    }
}
