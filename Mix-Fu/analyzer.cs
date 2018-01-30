using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Agilent.CommandExpert.ScpiNet.AgSCPI99_1_0;

namespace Mixer
{
    internal class Analyzer : Instrument {
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
            throw new NotImplementedException();
        }

        public override void send(string command) {
            throw new NotImplementedException();
        }
    }
}
