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

        private const string decimalFormat = "0.00";

        // TODO: make instrument classes: Generator, Analyser, hide queries there into methods
        public Instrument m_IN { get; set; }
        public Instrument m_OUT { get; set; }
        public Instrument m_LO { get; set; }

        public int delay { get; set; } = 300;
        public int relax { get; set; } = 70;
        public decimal attenuation { get; set; } = 30;
        public decimal maxfreq { get; set; } = 26500;
        public decimal span { get; set; } = 10000000;

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

        public void send(string location, string com) {
            //// live
            //try {
            //    using (AgSCPI99 instrument = new AgSCPI99(location)) {
            //        instrument.Transport.Command.Invoke(com);
            //    }
            //}
            //catch (Exception ex) {
            //    throw ex;
            //}

            // mock
            log("send: " + com + " to: " + location, false);
        }

        public string query(string location, string question) {
            //// live
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
            log("query: " + question + " to: " + location, false);
            return "-25.5"; // TODO: test diff value
        }

        public void searchInstruments(List<Instrument> instruments, int maxPort, int gpib) {
            log("start instrument search...", false);

            instruments.Clear();

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

        public void prepareCalibrateIn() {
            // TODO: check SA & GEN assignment logic
            string SA = m_IN.Location;
            string GEN = m_OUT.Location;

            send(SA, ":CAL:AUTO OFF");                       // выключаем автокалибровку анализатора спектра
            send(SA, ":SENS:FREQ:SPAN " + span.ToString());  // выставляем спан
            send(SA, ":CALC:MARK1:MODE POS");                // выставляем режим маркера
            send(GEN, "OUTP:STAT ON");                       // включаем генератор (redundant parenthesis around "")            
        }

        public void releaseInstrument() {
            // TODO: check SA & GEN assignment logic
            string SA = m_IN.Location;
            string GEN = m_OUT.Location;

            send(SA, ":CAL:AUTO ON");     //включаем автокалибровку анализатора спектра
            send(GEN, "OUTP:STAT OFF");   //выключаем генератор
        }

        public void calibrateIn(DataTable dataTable, string GEN, string SA, string freqCol, string powCol, string powGoalCol) {
            log("start calibrate IN", false);
            prepareCalibrateIn();

            List<CalPoint> listCalData = new List<CalPoint>();

            foreach (DataRow row in dataTable.Rows) {
                string inFreqStr = row[freqCol].ToString().Replace(',', '.');
                string inPowGoalStr = row[powGoalCol].ToString().Replace(',', '.');

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
                inFreqDec *= 1000000;

                bool exists = listCalData.Exists(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);

                log("debug: calibrate: freq=" + inFreqDec + 
                                    " pgoal=" + inPowGoalDec + 
                                   " exists=" + exists, false);

                if (!exists) {
                    decimal tempPowDec = inPowGoalDec;
                    decimal err = 1;

                    send(GEN, "SOUR:FREQ " + inFreqDec);
                    send(SA, ":SENSe:FREQuency:RF:CENTer " + inFreqDec);
                    send(SA, ":CALCulate:MARKer1:X:CENTer " + inFreqDec);

                    int count = 0;
                    int relax_temp = relax;
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        decimal readPowDec = 0;
                        send(GEN, "SOUR:POW " + tempPowDec);
                        Thread.Sleep(relax);
                        // TODO: inline readPow
                        string readPow = query(SA, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(readPow, NumberStyles.Any, 
                                         CultureInfo.InvariantCulture, out readPowDec);

                        err = inPowGoalDec - readPowDec;
                        tempPowDec += err;

                        ++count;
                        relax += 50;
                    }
                    relax = relax_temp;
                    listCalData.Add(new CalPoint { freqD = inFreqDec,
                                                   powD = inPowGoalDec,
                                                   calPowD = tempPowDec,
                                                   error = err });
                    
                    log("debug: new " + listCalData.Last().ToString(), false);
                }
                CalPoint updatedPoint = listCalData.Find(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);
                // ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',');
                row["ERR"] = updatedPoint.error.ToString(decimalFormat).Replace('.', ',');
                row[powCol] = updatedPoint.calPowD.ToString(decimalFormat).Replace('.', ',');
            }
            releaseInstrument();
            log("end calibrate IN", false);
        }

        public void calibrateInMult(DataTable dataTable, string GEN, string SA, string freqCol, string powCol, string powGoalCol) {
            // TODO: merge with calibrateIn(); diff: span adjustment const, additional SA command, delay instead of relax
            log("start calibrate IN mult", false);
            prepareCalibrateIn();

            send(SA, ":POW:ATT " + attenuation.ToString());
            //send(OUT, "DISP: WIND: TRAC: Y: RLEV " + (attenuation - 10).ToString());
            //send(IN, ("SOUR:POW " + "-100"));

            List<CalPoint> listCalData = new List<CalPoint>();

            foreach (DataRow row in dataTable.Rows) {
                string inFreqStr = row[freqCol].ToString().Replace(',', '.');
                string inPowGoalStr = row[powGoalCol].ToString().Replace(',', '.');

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
                inFreqDec *= 1000000000;  // TODO: check freq adjustment

                bool exists = listCalData.Exists(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);

                log("debug: calibrate: freq=" + inFreqDec +
                                    " pgoal=" + inPowGoalDec +
                                   " exists=" + exists, false);

                if (!exists) {
                    decimal tempPowDec = inPowGoalDec;
                    decimal err = 1;

                    send(GEN, "SOUR:FREQ " + inFreqDec);
                    send(SA, ":SENSe:FREQuency:RF:CENTer " + inFreqDec);
                    send(SA, ":CALCulate:MARKer1:X:CENTer " + inFreqDec);

                    int count = 0;
                    int delay_temp = delay;
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        decimal readPowDec = 0;
                        send(GEN, "SOUR:POW " + tempPowDec);
                        Thread.Sleep(delay);
                        string readPow = query(SA, ":CALCulate:MARKer:Y?");
                        decimal.TryParse(readPow, NumberStyles.Any, 
                                         CultureInfo.InvariantCulture, out readPowDec);

                        err = inPowGoalDec - readPowDec;
                        tempPowDec += err;

                        ++count;
                        delay += 50;
                    }
                    delay = delay_temp;

                    listCalData.Add(new CalPoint { freqD = inFreqDec,
                                                    powD = inPowGoalDec,
                                                 calPowD = tempPowDec,
                                                   error = err });
                    log("debug: new " + listCalData.Last().ToString(), false);
                }
                CalPoint updatedPoint = listCalData.Find(point => point.freqD == inFreqDec && point.powD == inPowGoalDec);
                // ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',');
                row["ERR"] = updatedPoint.error.ToString(decimalFormat).Replace('.', ',');
                row[powCol] = updatedPoint.calPowD.ToString(decimalFormat).Replace('.', ',');
            }
            releaseInstrument();
            log("end calibrate IN", false);
        }

    }
}
