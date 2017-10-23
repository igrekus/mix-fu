using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mix_Fu {

    class InstrumentManager {

        //private string m_IN { get; set; }
        //private string m_OUT { get; set; }
        //private string m_LO { get; set; }

        public string m_IN { get; set; } = "";
        public string m_OUT { get; set; } = "";
        public string m_LO { get; set; } = "";

        public int delay { get; set; } = 300;
        public int relax { get; set; } = 70;
        public decimal attenuation { get; set; } = 30;
        public decimal maxfreq { get; set; } = 26500;
        public decimal span { get; set; } = 10000000;

        MeasureMode measureMode { get; set; } = MeasureMode.modeDSBDown;

        Task searchTask = null;

        public InstrumentManager() {

        }

        public InstrumentManager(string IN, string OUT, string LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
        }

        public void setInstruments(string IN, string OUT, string LO) {
            m_IN = IN;
            m_OUT = OUT;
            m_LO = LO;
        }

    }

}
