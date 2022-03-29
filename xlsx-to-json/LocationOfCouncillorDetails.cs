using System.Diagnostics;

namespace xlsx_to_json
{
    [DebuggerDisplay("Column: {ColumnName}, Cell: {CellReference}")]
    public class LocationOfCouncillorDetails
    {
        public string ColumnName { get; set; }
        public CellReference CellReference { get; set; }
        public string SharedStringIndex { get; set; }
    }
}
