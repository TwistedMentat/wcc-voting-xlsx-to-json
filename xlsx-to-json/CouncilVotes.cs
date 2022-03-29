using System.Diagnostics;

namespace xlsx_to_json
{
    [DebuggerDisplay("Votes: {VoteNames.Count}, Councillors: {Councillors.Count}")]
    public class CouncilVotes
    {
        public CouncilVotes()
        {
            VoteNames = new List<(string, CellReference)>();
            Councillors = new List<string>();
            Votes = new List<CouncillorVote>();
        }

        public IList<(string VoteName, CellReference CellReference)> VoteNames { get; set; }
        public IList<string> Councillors { get; set; }
        public IList<CouncillorVote> Votes { get; set; }
    }
}
