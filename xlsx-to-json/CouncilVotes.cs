namespace xlsx_to_json
{
    public class CouncilVotes
    {
        public CouncilVotes()
        {
            VoteNames = new List<string>();
            Councilors = new List<string>();
            Votes = new List<CouncilorVote>();
        }

        public IList<string> VoteNames { get; set; }
        public IList<string> Councilors { get; set; }
        public IList<CouncilorVote> Votes { get; set; }
    }
}