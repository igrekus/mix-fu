#define mock
using System;
using NationalInstruments.VisaNS;

namespace Mixer {
    internal class Akip3407 : Instrument {

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

        public Akip3407(string location) {
            throw new NotImplementedException();
        }

        public override string query(string question) {
#if mock
            return "akip query success: " + question;
#else
            string ans = _instrument.Query(question);
            return ans;
#endif
        }

        public override string send(string command) {
#if mock
            return "generator command success: " + command;
#else
            _instrument.Write(command);
            return "akip command success";
#endif
        }

        public string SetOutput(string state) => send("OUTP:STAT " + state);

        public string SetOutputModulation(string state) => send(":OUTP:MOD:STAT " + state);

        public string SetSourceFreq(decimal freq) => send("SOUR:FREQ " + freq);

        public string SetSourceFreq(string freq) => send("SOUR:FREQ " + freq);

        public string SetSourcePow(decimal pow) => send("SOUR:POW " + pow);

        public string SetSourcePow(string pow) => send("SOUR:POW " + pow);
    }
}
