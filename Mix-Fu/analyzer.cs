using System;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;

namespace Mixer
{
    class Analyzer : Instrument, IAnalyzer {

        public struct AutocalState {
            public const string AutocalOff = "OFF";
            public const string AutocalOn = "ON";
        }

        public struct MWPreselectorPath{
            public const string PathMPB = "MPB";
            public const string PathSTD = "STD";
        }

        public struct InternalPreampState {
            public const string InternalPreampOff = "OFF";
            public const string InternalPreampOn = "ON";
        }

        public struct MarkerMode {
            public const string ModePos = "POS";
        }

        private AgSCPI99 _instrument;

        public Analyzer() {
            // Empty ctor for mocking only!
        }

        public Analyzer(string location, string fullname) {
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

        protected virtual string query(string question) {
            _instrument.Transport.Query.Invoke(question, out var ans);
            return ans;
        }

        protected virtual string send(string command) {
            _instrument.Transport.Command.Invoke(command);
            return "analyzer command success";
        }

        public override string RawQuery(string question) => query(question);

        public override string RawCommand(string command) => send(command);

        // TODO: make properties?
        public string SetAutocalibration(string state) => send(":CAL:AUTO " + state);

        public string SetFreqSpan(decimal span) => send(":SENS:FREQ:SPAN " + span);

        //        public string SetMWPreselectorPath(string state) => send(":POW:MW:PRES " + state);
        //        public string SetMWPreselectorPath(string state) => send(":POW:MW:PATH " + state);
        public string SetMWPreselectorPath(string state) => "ok";

        //        public string SetInternalPreampState(string state) => send(":POW:GAIN " + state);
        public string SetInternalPreampState(string state) => "ok";

        public string SetMarkerMode(string mode) => send(":CALC:MARK1:MODE " + mode);

        public string SetPowerAttenuation(decimal att) => send(":POW:ATT " + att);

        public string SetMeasCenterFreq(decimal freq) => send(":SENSe:FREQuency:CENTer " + freq);

        public string SetMarker1XCenter(decimal freq) => send(":CALCulate:MARKer1:X:CENTer " + freq);

        public string ReadMarker1Y() => query(":CALCulate:MARKer:Y?");

        public string FindPeakAndReadMarker1Y() {
            send("CALC:MARK1:MAX");
            System.Threading.Thread.Sleep(500);
            return query(":CALC:MARK1:Y?");
        }
    }
}
