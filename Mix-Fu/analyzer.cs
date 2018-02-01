#define mock
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;

namespace Mixer
{
    class Analyzer : Instrument {

        public struct AutocalState {
            public const string AutocalOff = "OFF";
            public const string AutocalOn = "ON";
        }

        public struct MarkerMode {
            public const string ModePos = "POS";
        }

        private AgSCPI99 _instrument;

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

        public Analyzer(string location) {
            throw new NotImplementedException();
        }

        public override string query(string question) {
#if mock
            return "analyzer query success: " + question;
#else
            _instrument.Transport.Query.Invoke(question, out var ans);
            return ans;
#endif
        }

        public override string send(string command) {
#if mock
            return "analyzer command success: " + command;
#else
            _instrument.Transport.Command.Invoke(command);
            return "analyzer command success";
#endif
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
