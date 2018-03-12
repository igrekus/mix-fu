namespace Mixer {
    internal class GeneratorMock : Generator {

        public GeneratorMock(string location, string fullname) {
            Location = location;
            FullName = fullname;
            Name     = fullname.Split(',')[1];
        }

        protected override string query(string question) {
            return "generator: " + Name + " query success: " + question;
        }

        protected override string send(string command) {
            return "generator: " + Name + " command success: " + command;
        }
    }
}

