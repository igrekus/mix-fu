using System;

namespace Mixer
{
    class AnalyzerMock : Instrument {

        public struct AutocalState {
            public const string AutocalOff = "OFF";
            public const string AutocalOn = "ON";
        }

        public struct MarkerMode {
            public const string ModePos = "POS";
        }

        public AnalyzerMock(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name = fullname.Split(',')[1];
        }

        public AnalyzerMock(string location) {
            throw new NotImplementedException();
        }
        
        // TODO: mock queries correctly
        public override string query(string question) {
            return "analyzer query success: " + question;
        }

        public override string send(string command) {
            return "analyzer command success: " + command;
        }

        // TODO: make properties?
        public string SetAutocalibration(string state) => send(":CAL:AUTO " + state);

        public string SetFreqSpan(decimal span) => send(":SENS:FREQ:SPAN " + span);

        public string SetMarkerMode(string mode) => send(":CALC:MARK1:MODE " + mode);

        public string SetPowerAttenuation(decimal att) => send(":POW:ATT " + att);

        public string SetMeasCenterFreq(decimal freq) => send(":SENSe:FREQuency:RF:CENTer " + freq);

        public string SetMarker1XCenter(decimal freq) => send(":CALCulate:MARKer1:X:CENTer " + freq);

        public string ReadMarker1Y() => query(":CALCulate:MARKer:Y?");
    }
}
