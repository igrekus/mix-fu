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

namespace Mix_Fu {

    class InstrumentManager {

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
        }

        public InstrumentManager(Action<string, bool> logger, Instrument IN, Instrument OUT, Instrument LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
            log = logger;
        }

        public void setInstruments(Instrument IN, Instrument OUT, Instrument LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
        }

        public void send(string location, string command) {
            //log("debug: send: " + command + " to: " + location, true);
            //// live
            //try
            //{
            //    using (AgSCPI99 instrument = new AgSCPI99(location)) {
            //        instrument.Transport.Command.Invoke(command);
            //    }
            //}
            //catch (Exception ex) {
            //    throw ex;
            //}

            // mock
            log("send: " + command + " to: " + location, true);
        }

        public string query(string location, string question) {
            //// live
            //log("debug: query: " + question + " to: " + location, true);
            //string answer = "";
            //try {
            //    using (AgSCPI99 instrument = new AgSCPI99(location)) {
            //        instrument.Transport.Query.Invoke(question, out answer);
            //    }
            //}
            //catch (Exception ex) {
            //    throw ex;
            //}
            //return answer;

            // mock
            log("query: " + question + " to: " + location, true);
            return "-25.5"; // TODO: test diff value
        }

        public void searchInstruments(List<Instrument> instruments, int maxPort, int gpib) {
            log("start instrument search...", false);

            for (int i = 0; i <= maxPort; i++) {
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
            send(GEN, "OUTP:STAT ON");                       // включаем генератор (redundant parenthesis around "")            
            send(SA, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            //send(IN, "SOUR:POW " + "-100");

            // TODO: send GEN modulation off
        }

        private void setupSA(string GEN, string SA, decimal inFreqDec) {
            // TODO: bind to instroment properties
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

        public void calibrateInExec(DataTable data, ParameterStruct paramDict) {
            // TODO: exception handling
            // TODO: split into methods
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            log("start calibrate IN: " + "GEN=" + GEN + " SA=" + SA, false);
            prepareInstrument(GEN, SA);

            List<CalibrationPoint> listCalData = new List<CalibrationPoint>();

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
                    log("error: fail parsing " + inFreqStr, false);
                    continue;
                }
                if (!decimal.TryParse(inPowGoalStr, NumberStyles.Any, CultureInfo.InvariantCulture,
                                      out inPowGoalDec)) {
                    log("error: fail parsing " + inPowGoalStr, false);
                    continue;
                }
                inFreqDec *= Constants.GHz;

                bool exists = listCalData.Exists(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);

                log("debug: calibrate: freq=" + inFreqDec + 
                                    " pgoal=" + inPowGoalDec + 
                                   " exists=" + exists, true);

                if (!exists) {
                    decimal tempPowDec = inPowGoalDec;
                    decimal err = 1;

                    setupSA(GEN, SA, inFreqDec);

                    int count = 0;
                    int tmpDelay = delay;
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        decimal readPowDec = 0;
                        send(GEN, "SOUR:POW " + tempPowDec);
                        Thread.Sleep(delay);
                        // TODO: inline readPow
                        string readPow = query(SA, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(readPow, NumberStyles.Any,
                                         CultureInfo.InvariantCulture, out readPowDec);
                        log("read data:" + readPow + " " + readPowDec, true);

                        err = inPowGoalDec - readPowDec;
                        tempPowDec += err;

                        ++count;
                        delay += 50;
                    }
                    delay = tmpDelay;
                    listCalData.Add(new CalibrationPoint {
                        freqD = inFreqDec,
                        powD = inPowGoalDec,
                        calPowD = tempPowDec,
                        error = err
                    });

                    log("debug: new " + listCalData.Last().ToString(), true);
                }
                CalibrationPoint updatedPoint = listCalData.Find(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);
                // ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',');
                row["ERR"] = updatedPoint.error.ToString(Constants.decimalFormat).Replace('.', ',');
                row[paramDict.colPow] = updatedPoint.calPowD.ToString(Constants.decimalFormat).Replace('.', ',');
            }
            releaseInstrument(GEN, SA);
            log("end calibrate IN", false);
        }

        public void calibrateLoExec(DataTable data, ParameterStruct paramDict) {
            Instrument tmpIn = m_IN;
            m_IN = m_LO;
            calibrateInExec(data, paramDict);
            m_IN = tmpIn;
        }

        public string getAttenuationError(string GEN, string SA, string freq, decimal powGoal, int harmonic = 1) {
            // TODO: refactor to match with in-lo calibration?            
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

            setupSA(GEN, SA, inFreqDec * harmonic);
            Thread.Sleep(delay);

            string readPow = query(SA, ":CALCulate:MARKer:Y?");

            decimal readPowDec = 0;
            decimal.TryParse(readPow, NumberStyles.Any, CultureInfo.InvariantCulture, out readPowDec);
            decimal errDec = powGoal - readPowDec;

            return errDec.ToString("0.000", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public void calibrateOutExec(DataTable data, MeasureMode mode) {
            string GEN = m_IN.Location;
            string SA = m_OUT.Location;

            log("start calibrate IN: " + "GEN=" + GEN + " SA=" + SA, false);

            prepareInstrument(GEN, SA);

            decimal tempPow = (decimal)-20.00;
            send(GEN, "SOUR:POW " + tempPow);

            List<Tuple<string, string>> paramList = new List<Tuple<string, string>>();

            // TODO: make parameter dict
            switch (mode) {
            case MeasureMode.modeDSBDown:
            case MeasureMode.modeDSBUp:
                paramList.Add(new Tuple<string, string>("FIF", "ATT-IF"));
                paramList.Add(new Tuple<string, string>("FRF", "ATT-RF"));
                paramList.Add(new Tuple<string, string>("FLO", "ATT-LO"));
                break;
            case MeasureMode.modeSSBDown:
            case MeasureMode.modeSSBUp:
                paramList.Add(new Tuple<string, string>("FIF", "ATT-IF"));
                paramList.Add(new Tuple<string, string>("FLSB", "ATT-LSB"));
                paramList.Add(new Tuple<string, string>("FUSB", "ATT-USB"));
                paramList.Add(new Tuple<string, string>("FLO", "ATT-LO"));
                break;
            case MeasureMode.modeMultiplier:
                paramList.Add(new Tuple<string, string>("FH1", "ATT-H1"));
                paramList.Add(new Tuple<string, string>("FH1", "ATT-H2"));
                paramList.Add(new Tuple<string, string>("FH1", "ATT-H3"));
                paramList.Add(new Tuple<string, string>("FH1", "ATT-H4"));
                break;
            default:
                return;
            }

            foreach (DataRow row in data.Rows) {
                foreach (Tuple<string, string> p in paramList) {
                    row[p.Item2] = getAttenuationError(GEN, SA, row[p.Item1].ToString(), tempPow, 4);
                }
            }

            releaseInstrument(GEN, SA);
        }

    }
}
