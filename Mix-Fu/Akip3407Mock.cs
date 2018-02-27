using System;
using NationalInstruments.VisaNS;

namespace Mixer {
    internal class Akip3407Mock : Instrument {

        private UsbRaw _instrument;

        public Akip3407Mock(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name     = fullname.Split(',')[1];
        }

        public Akip3407Mock(string location) {
            throw new NotImplementedException();
        }

        public override string query(string question) {
            return "akip query success: " + question;
        }

        public override string send(string command) {
            return "generator command success: " + command;
        }

        public string SetOutput(string state) => send("OUTP:STAT " + state);

        public string SetOutputModulation(string state) => send(":OUTP:MOD:STAT " + state);

        public string SetSourceFreq(decimal freq) => send("SOUR:FREQ " + freq);

        public string SetSourceFreq(string freq) => send("SOUR:FREQ " + freq);

        public string SetSourcePow(decimal pow) => send("SOUR:POW " + pow);

        public string SetSourcePow(string pow) => send("SOUR:POW " + pow);
    }
}
