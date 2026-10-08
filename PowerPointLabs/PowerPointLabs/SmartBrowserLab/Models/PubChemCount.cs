using Newtonsoft.Json;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemCount
    {
        [JsonProperty("heavy_atom")]
        public int HeavyAtom { get; set; }

        [JsonProperty("atom_chiral")]
        public int AtomChiral { get; set; }

        [JsonProperty("atom_chiral_def")]
        public int AtomChiralDef { get; set; }

        [JsonProperty("atom_chiral_undef")]
        public int AtomChiralUndef { get; set; }

        [JsonProperty("bond_chiral")]
        public int BondChiral { get; set; }

        [JsonProperty("bond_chiral_def")]
        public int BondChiralDef { get; set; }

        [JsonProperty("bond_chiral_undef")]
        public int BondChiralUndef { get; set; }

        [JsonProperty("isotope_atom")]
        public int IsotopeAtom { get; set; }

        [JsonProperty("covalent_unit")]
        public int CovalentUnit { get; set; }

        [JsonProperty("tautomers")]
        public int Tautomers { get; set; }
    }
}
