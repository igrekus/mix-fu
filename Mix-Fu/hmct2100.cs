namespace Mixer
{
    internal class HMCT2100 : Generator {

        public HMCT2100(string location, string fullname) : base(location, fullname) {}

        public override string SetOutputModulation(string state) => "skipping: set modulation";
    }
}

