#define mock
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
#if mock
            return "generator query success: " + question;
#else
            _instrument.Transport.Query.Invoke(question, out var ans);
            return ans;
#endif
        }

        public override string send(string command) {
#if mock
            return "generator command success: " + command;
#else
            _instrument.Transport.Command.Invoke(command);
            return "generator command success";
#endif
        }

        public string SetOutput(string state) => send("OUTP:STAT " + state);

        public string SetOutputModulation(string state) => send(":OUTP:MOD:STAT " + state);

        public string SetSourceFreq(decimal inFreqDec) => send("SOUR:FREQ " + inFreqDec);

        public string SetSourcePow(decimal pow) => send("SOUR:POW " + pow);
    }
}

