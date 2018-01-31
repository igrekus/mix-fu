#define mock
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

        public override QueryResult query(string question) {
#if mock
            return new QueryResult {code = 0, answer = "analyser query success"};
#else
            string ans;
            try {
                _instrument.Transport.Query.Invoke(question, out ans);
                return new QueryResult { code = 0, answer = ans };
            }
            catch (Exception ex) {
                return new QueryResult { code = -1, answer = ex.Message };
            }
#endif
        }

        public override CommandResult send(string command) {
            throw new NotImplementedException();
        }
    }
}
