using System;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;

namespace Mixer
{
    internal class Generator : Instrument {

        public struct OutputModulationState {
            public const string ModulationOff = "OFF";
            public const string ModulationOn  = "ON";
        }

        public struct OutputState {
            public const string OutputOff = "OFF";
            public const string OutputOn = "ON";
        }

        private AgSCPI99 _instrument;

        public Generator(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name = fullname.Split(',')[1];

            try {
                _instrument = new AgSCPI99(location);
            }
            catch (Exception ex) {
                // ignored
            }
        }

        public Generator(string location) {
            throw new NotImplementedException();
        }

        public override string query(string question) {
            _instrument.Transport.Query.Invoke(question, out var ans);
            return ans;
        }

        public override string send(string command) {
            _instrument.Transport.Command.Invoke(command);
            return "generator command success";
        }

        public string SetOutput(string state) => send("OUTP:STAT " + state);

        public string SetOutputModulation(string state) => send(":OUTP:MOD:STAT " + state);

        public string SetSourceFreq(decimal freq) => send("SOUR:FREQ " + freq);

        public string SetSourceFreq(string freq) => send("SOUR:FREQ " + freq);

        public string SetSourcePow(decimal pow) => send("SOUR:POW " + pow);

        public string SetSourcePow(string pow) => send("SOUR:POW " + pow);
    }
}

