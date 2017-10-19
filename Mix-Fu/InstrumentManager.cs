﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mix_Fu {

    class InstrumentManager {

        private string m_IN { get; set; }
        private string m_OUT { get; set; }
        private string m_LO { get; set; }

        public int delay = 300;
        public int relax = 70;
        public decimal attenuation = 30;
        public decimal maxfreq = 26500;
        public decimal span = 10000000;
        MeasureMode measureMode = MeasureMode.modeDSBDown;

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