using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace xlsx_to_json
{
    public class WccVotingSpreadsheet
    {
        private SharedStringTable sharedStringTable;
        private SheetData sheetData;
        private IEnumerable<Cell> allCells;

        public WccVotingSpreadsheet(SpreadsheetDocument spreadsheetDocument)
        {
            sharedStringTable = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;
            IEnumerable<WorksheetPart> worksheetParts = spreadsheetDocument.WorkbookPart.WorksheetParts;
            WorksheetPart worksheetPart = worksheetParts.First();
            sheetData = (SheetData)worksheetPart.RootElement.ChildElements.Single(w => w is SheetData);
            allCells = sheetData.ChildElements.SelectMany(r => r.ChildElements).Cast<Cell>();
        }

        public CouncilVotes TransformExcel()
        {
            // Why not just make an in memory table with all the properties you want? Then just use objects and LINQ to chop it up as needed :/
            
            List<CouncillorDetails?> councillorDetails = new()
            {
                GetSharedStringIndexOfHeaderName("Keypad SN", sharedStringTable),
                GetSharedStringIndexOfHeaderName("First Name", sharedStringTable),
                GetSharedStringIndexOfHeaderName("Last Name", sharedStringTable)
            };
            councillorDetails.RemoveAll(cd => cd == null);

            foreach (Cell? cell in sheetData.ChildElements.SelectMany(row => row.ChildElements.Where(cell => councillorDetails.Any(s => s.SharedStringIndex == cell.InnerText))))
            {
                if (cell == null)
                {
                    continue;
                }

                councillorDetails.Single(cd => cd.SharedStringIndex == cell.InnerText).CellReference = cell.CellReference.Value;
            }

            bool keepCheckingForVotingSectionStart = true;

            Row headerRow = allCells.Single(cell => cell.DataType?.Value == CellValues.SharedString && cell.InnerText == "5").Parent as Row;

            CouncilVotes councilVotes = new()
            {
                VoteNames = GetVoteNames(headerRow, sharedStringTable, councillorDetails.Select(cd => cd.SharedStringIndex).ToList())
            };

            // Need to find where the "Keypad SN" cell is. That will define the row to start.

            // Get rows with vote information in them. This will be every contigous row starting from the header row which we already know
            IList<Row> voteTableRows= sheetData.ChildElements.Cast<Row>().Where(row => row.RowIndex >= headerRow.RowIndex).ToList();

            for (int i = 0; i < sheetData.ChildElements.Count; i++)
            {
                Row row = (Row)sheetData.ChildElements[i];

                if (keepCheckingForVotingSectionStart && row.RowIndex != headerRow.RowIndex)
                {
                    continue;
                }
                else
                {
                    keepCheckingForVotingSectionStart = false;
                }

                string[] spanStartAndEnd = row.Spans.Items.First().Value.Split(":");
                int columnStart = int.Parse(spanStartAndEnd[0]) - 1;
                int worksheetColumnIndexEnd = int.Parse(spanStartAndEnd[1]) - 1;

                if (i > headerRow.RowIndex)
                {
                    Cell councilorCell = (Cell)row.ChildElements[0];
                    councilVotes.Councilors.Add(councilorCell.CellValue.Text);
                }

                for (int worksheetColumnIndex = 1; worksheetColumnIndex < worksheetColumnIndexEnd; worksheetColumnIndex++)
                {
                    ExtractVotesForCouncillor(councilVotes, row, worksheetColumnIndex, headerRow);
                }
            }

            return councilVotes;
        }

        private void ExtractVotesForCouncillor(CouncilVotes councilVotes, Row row, int worksheetColumnIndex, Row headerRow)
        {
            if (row == headerRow)
            {
                return;
            }

            string currentCell = GetWorksheetColumnReference(worksheetColumnIndex) + row.RowIndex;

            Cell cellsInRow = row.ChildElements.Cast<Cell>().SingleOrDefault(ce => ce.CellReference == currentCell);

            int choiceValue;
            if (cellsInRow == null)
            {
                choiceValue = 0;
            }
            else
            {
                choiceValue = int.Parse(cellsInRow.InnerText);
            }

            CouncilorVote councilorVote = new()
            {
                CouncilorName = row.ChildElements[0].InnerText,
                Choice = (Choice)choiceValue,
                VoteName = councilVotes.VoteNames[worksheetColumnIndex - 1]
            };
            councilVotes.Votes.Add(councilorVote);
        }

        private IList<string> GetVoteNames(Row headerRow, SharedStringTable sharedStringTable, ICollection<string> valuesThatAreNotVotes)
        {
            List<string> voteNames = headerRow.ChildElements.Cast<Cell>().Select(cell => GetSharedStringIndexOfVoteName(cell, sharedStringTable)).ToList();

            voteNames.RemoveAll(vn => valuesThatAreNotVotes.Contains(vn));

            return voteNames.ToList();
        }

        private string GetSharedStringIndexOfVoteName(Cell cell, SharedStringTable sharedStringTable)
        {
            return sharedStringTable.ChildElements.Cast<SharedStringItem>().ToList()[int.Parse(cell.CellValue.Text)].Text.Text;
        }

        private CouncillorDetails GetSharedStringIndexOfHeaderName(string columnName, SharedStringTable sharedStringTable)
        {
            OpenXmlElement keypadSnCell = sharedStringTable.ChildElements.SingleOrDefault(ce => ce.InnerText.Equals(columnName));
            if (keypadSnCell == null)
            {
                return null;
            }

            string sharedStringIndex = sharedStringTable.ChildElements.ToList().IndexOf(keypadSnCell).ToString();
            return new CouncillorDetails { Label = columnName, SharedStringIndex = sharedStringIndex };
        }

        // joinked from here https://stackoverflow.com/a/297214
        private string GetWorksheetColumnReference(int index)
        {
            index -= 1; //adjust so it matches 0-indexed array rather than 1-indexed column

            int quotient = index / 26;
            if (quotient > 0)
                return GetWorksheetColumnReference(quotient) + chars[index % 26].ToString();
            else
                return chars[index % 26].ToString();
        }
        private char[] chars = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
    }
}
