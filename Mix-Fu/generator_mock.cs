using System;

namespace Mixer
{
    internal class GeneratorMock : Instrument {

        public struct OutputModulationState {
            public const string ModulationOff = "OFF";
            public const string ModulationOn  = "ON";
        }

        public struct OutputState {
            public const string OutputOff = "OFF";
            public const string OutputOn = "ON";
        }

        public GeneratorMock(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name = fullname.Split(',')[1];
        }

        public GeneratorMock(string location) {
            throw new NotImplementedException();
        }

        public override string query(string question) {
            // TODO: mock queries correctly
            return "generator query success: " + question;
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

