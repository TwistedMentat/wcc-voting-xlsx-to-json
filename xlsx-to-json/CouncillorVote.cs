using System.Diagnostics;

namespace xlsx_to_json
{
    [DebuggerDisplay("Councillor: {CouncillorName}, Choice: {Choice}, Vote: {VoteName}")]
    public class CouncillorVote
    {
        public string CouncillorName { get; set; }
        public string VoteName { get; set; }
        public Choice Choice { get; set; }
    }
}
