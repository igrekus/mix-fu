using System;

namespace Mixer {
    internal class Akip3407Mock : Akip3407 {

        public Akip3407Mock(string location, string fullname) : base(location, fullname) { }

        protected override string query(string question) {
            return "AKIP query success: " + question;
        }

        protected override string send(string command) {
            return "AKIP command success: " + command;
        }

        public override string RawQuery(string question) => query(question);

        public override string RawCommand(string command) => send(command);
    }
}
