namespace xlsx_to_json
{
    public class CouncilVotes
    {
        public CouncilVotes()
        {
            VoteNames = new List<(string, CellReference)>();
            Councillors = new List<string>();
            Votes = new List<CouncilorVote>();
        }

        public IList<(string VoteName, CellReference CellReference)> VoteNames { get; set; }
        public IList<string> Councillors { get; set; }
        public IList<CouncilorVote> Votes { get; set; }
    }
}
