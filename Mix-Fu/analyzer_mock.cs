using System;

namespace Mixer {
    class AnalyzerMock : Analyzer {

        public AnalyzerMock(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name     = fullname.Split(',')[1];
        }

        protected override string query(string question) {
            return "analyzer: " + Name + " query success: " + question;
        }

        protected override string send(string command) {
            return "analyzer: " + Name + " command success: " + command;
        }

//        public override string RawQuery(string question) => query(question);
//
//        public override string RawCommand(string command) => send(command);
    }
}
