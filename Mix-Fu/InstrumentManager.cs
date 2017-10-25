using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Mix_Fu {

    class InstrumentManager {

        //private string m_IN { get; set; }
        //private string m_OUT { get; set; }
        //private string m_LO { get; set; }

        public Instrument m_IN { get; set; }
        public Instrument m_OUT { get; set; }
        public Instrument m_LO { get; set; }

        public int delay { get; set; } = 300;
        public int relax { get; set; } = 70;
        public decimal attenuation { get; set; } = 30;
        public decimal maxfreq { get; set; } = 26500;
        public decimal span { get; set; } = 10000000;

        MeasureMode measureMode { get; set; } = MeasureMode.modeDSBDown;

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

                //try {
                //    using (AgSCPI99 instrument = new AgSCPI99(location)) {
                //        instrument.SCPI.IDN.Query(out idn);
                //        string[] idn_cut = idn.Split(',');
                //        instruments.Add(new Instrument { Location = location, Name = idn_cut[1], FullName = idn });
                //        log("found " + idn + " at " + location);
                //    }
                //}

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

        internal void prepareCalibrateIn(decimal span) {
            // TODO: check SA & GEN assignment logic
            string SA = m_IN.Location;
            string GEN = m_OUT.Location;

            send(SA, ":CAL:AUTO OFF");                       // выключаем автокалибровку анализатора спектра
            send(SA, ":SENS:FREQ:SPAN " + span.ToString());  // выставляем спан
            send(SA, ":CALC:MARK1:MODE POS");                // выставляем режим маркера
            send(GEN, "OUTP:STAT ON");                       // включаем генератор (redundant parenthesis around "")            
        }

    }
}
