using System;
using NationalInstruments.VisaNS;

namespace Mixer {
    // TODO: inherit Generator
    internal class Akip3407 : Instrument, IGenerator {

        private UsbRaw _instrument;

        public Akip3407(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name     = fullname.Split(',')[1];

            try {
                _instrument = (UsbRaw)ResourceManager.GetLocalManager().Open(location);
            }
            catch (Exception ex) {
                // ignored
            }
        }

        protected virtual string query(string question) {
            string ans = _instrument.Query(question);
            return ans;
        }

        protected virtual string send(string command) {
            _instrument.Write(command);
            return "akip command success";
        }

        public override string RawQuery(string question) => query(question);

        public override string RawCommand(string command) => send(command);

        public string SetOutput(string state) => send("OUTP:STAT " + state);

        public string SetOutputModulation(string state) => send(":OUTP:MOD:STAT " + state);

        public string SetSourceFreq(decimal freq) => send("SOUR:FREQ " + freq);

        public string SetSourceFreq(string freq) => send("SOUR:FREQ " + freq);

        public string SetSourcePow(decimal pow) => send("SOUR:POW " + pow);

        public string SetSourcePow(string pow) => send("SOUR:POW " + pow);
    }
}
