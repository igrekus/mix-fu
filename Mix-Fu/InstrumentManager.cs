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

        public void calibrateIn(DataTable dataTable, string GEN, string SA, string freqCol, string powCol, string powGoalCol) {
            // TODO: refactor, read parameters into numbers, send ToString()
            prepareCalibrateIn();

            List<CalPoint> listCalData = new List<CalPoint>();
            //listCalData.Clear();

            foreach (DataRow row in dataTable.Rows) {
                string t_freq = (Convert.ToInt32(row[freqCol]) * 1000000).ToString();
                string t_pow_goal = row[powGoalCol].ToString().Replace(',', '.');

                // is calibration point already measured?
                bool exists = listCalData.Exists(Cal_Point =>
                    Cal_Point.freq == t_freq && Cal_Point.pow == t_pow_goal);

                if (!exists) {
                    string t_pow = "";
                    string t_pow_temp = "";
                    decimal t_pow_dec = 0;
                    decimal t_pow_goal_dec = 0;
                    decimal t_pow_temp_dec = 0;

                    decimal err = 1;

                    if (!decimal.TryParse(t_pow_goal, NumberStyles.Any, CultureInfo.InvariantCulture,
                        out t_pow_goal_dec)) {
                        log("error: can't parse: [" + t_pow_goal +
                            "] at row " + dataTable.Rows.IndexOf(row) + ", skipping", false /*true*/);
                        continue;
                    }

                    t_pow_temp = t_pow_goal;
                    t_pow_temp_dec = t_pow_goal_dec;

                    send(GEN, ("SOUR:FREQ " + t_freq));
                    send(SA, (":SENSe:FREQuency:RF:CENTer " + t_freq));
                    send(SA, (":CALCulate:MARKer1:X:CENTer " + t_freq));

                    int count = 0;
                    int relax_temp = relax;
                    // while err is within the bounds 0.05...10, 5 iterations max
                    while (count < 5 && Math.Abs(err) > (decimal)0.05 && Math.Abs(err) < 10) {
                        // set GEN source power level to POWGOAL column value
                        send(GEN, ("SOUR:POW " + t_pow_temp));
                        // wait before querying SA for response
                        Thread.Sleep(relax);
                        // read SA marker Y-value
                        t_pow = query(SA, ":CALCulate:MARKer:Y?");
                        log("marker y: " + t_pow, false);
                        decimal.TryParse(t_pow, NumberStyles.Any, CultureInfo.InvariantCulture, out t_pow_dec);
                        // calc diff between given goal pow and read pow
                        err = t_pow_goal_dec - t_pow_dec;

                        // new calibration point is measured pow + err
                        t_pow_temp_dec += err;
                        t_pow_temp = t_pow_temp_dec.ToString("0.00", CultureInfo.InvariantCulture);

                        count++;
                        relax += 50;
                    }
                    // measured pow, write down ERR
                    relax = relax_temp;
                    row["ERR"] = err.ToString("0.00", CultureInfo.InvariantCulture).Replace('.', ',');
                    // add measured calibration point
                    listCalData.Add(new CalPoint { freq = t_freq, pow = t_pow_goal, calPow = t_pow_temp.Replace('.', ',') });
                    log("debug: new CalPoint: freq=" + t_freq + " pow:" + t_pow_goal + " calpow:" + t_pow_temp, false);
                }
                // write pow column, use cached data if already measured
                row[powCol] = listCalData.Find(Cal_Point => Cal_Point.freq == t_freq && Cal_Point.pow == t_pow_goal).calPow;
                log("debug: add CalPoint from list:" + row[powCol], false);
            }
            send(SA, ":CAL:AUTO ON");     //включаем автокалибровку анализатора спектра
            send(GEN, "OUTP:STAT OFF");   //выключаем генератор
        }

    }
}
