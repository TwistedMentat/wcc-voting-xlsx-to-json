using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace xlsx_to_json
{
    public class WccVotingSpreadsheet
    {
        private const string KeypadSnColumnName = "Keypad SN";
        private const string FirstNameColumnName = "First Name";
        private const string LastNameColumnName = "Last Name";
        private readonly SharedStringTable _sharedStringTable;
        private readonly SheetData _sheetData;
        private readonly IEnumerable<Cell> _allCellsInSheet;
        private readonly Regex _splitCellReferenceRegex = new("^([a-zA-Z]+)(\\d+)$");

        public WccVotingSpreadsheet(SpreadsheetDocument spreadsheetDocument)
        {
            _sharedStringTable = spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;
            IEnumerable<WorksheetPart> worksheetParts = spreadsheetDocument.WorkbookPart.WorksheetParts;
            WorksheetPart worksheetPart = worksheetParts.First();
            _sheetData = (SheetData)worksheetPart.RootElement.ChildElements.First(w => w is SheetData); // If we have multiple sheets it will only look at the first in the file. Which may not match in Excel.
            _allCellsInSheet = _sheetData.ChildElements.SelectMany(r => r.ChildElements).Cast<Cell>();
        }

        public CouncilVotes TransformExcel()
        {
            // Why not just make an in memory table with all the properties you want? Then just use objects and LINQ to chop it up as needed :/

            List<LocationOfCouncillorDetails> locationsOfCouncillorDetails = FindLocationsOfCouncillorDetails();

            Row headerRow = (Row)_allCellsInSheet.First(cell => IsCellInHeaderRow(cell, locationsOfCouncillorDetails)).Parent;

            CouncilVotes councilVotes = new()
            {
                VoteNames = GetVoteNames(headerRow, _sharedStringTable, locationsOfCouncillorDetails)
            };

            // Improvement option: Get every row starting from the header row until there's a blank row
            IList<Row> voteTableRows = _sheetData.ChildElements.Cast<Row>().Where(row => row.RowIndex > headerRow.RowIndex).OrderBy(vtr => vtr.RowIndex).ToList();

            foreach (Row voteTableRow in voteTableRows)
            {
                Cell councillorCell = (Cell)voteTableRow.ChildElements[0];
                councilVotes.Councillors.Add(councillorCell.CellValue.Text);

                ExtractVotesForCouncillor(councilVotes, voteTableRow, locationsOfCouncillorDetails);
            }


            return councilVotes;
        }

        private static bool IsCellInHeaderRow(Cell cell, List<LocationOfCouncillorDetails> locationsOfCouncillorDetails)
        {
            return cell.CellValue != null
                                && cell.DataType == CellValues.SharedString
                                && locationsOfCouncillorDetails.Any(cd => cd.SharedStringIndex == cell.CellValue.Text);
        }

        private List<LocationOfCouncillorDetails> FindLocationsOfCouncillorDetails()
        {
            List<LocationOfCouncillorDetails> councillorDetailLocations = new()
            {
                GetSharedStringIndexOfHeaderName(KeypadSnColumnName, _sharedStringTable),
                GetSharedStringIndexOfHeaderName(FirstNameColumnName, _sharedStringTable),
                GetSharedStringIndexOfHeaderName(LastNameColumnName, _sharedStringTable)
            };
            councillorDetailLocations.RemoveAll(cd => cd == null);

            foreach (Cell cell in _sheetData.ChildElements.SelectMany(row => row.ChildElements.Where(cell => councillorDetailLocations.Any(s => s.SharedStringIndex == cell.InnerText))))
            {
                if (cell == null)
                {
                    continue;
                }

                councillorDetailLocations.Single(cd => cd.SharedStringIndex == cell.InnerText).CellReference = new CellReference(cell.CellReference);
            }

            return councillorDetailLocations;
        }

        private void ExtractVotesForCouncillor(CouncilVotes councilVotes, Row row, IList<LocationOfCouncillorDetails> councillorDetailLocations)
        {
            string councillorName = GetCouncillorName(row, councillorDetailLocations);

            foreach (Cell cell in row.ChildElements.Cast<Cell>())
            {
                // split reference
                MatchCollection matchCollection = _splitCellReferenceRegex.Matches(cell.CellReference);
                string columnName = matchCollection[0].Groups[1].Value;
                // put excel column of header value in votes collection

                // Skip any of the columns that are the councillor names
                if (councillorDetailLocations.Any(cdl => cell.IsInColumn(cdl.CellReference.ColumnName)))
                {
                    continue;
                }

                string voteName = councilVotes.VoteNames.Single(vn => cell.IsInColumn(vn.CellReference.ColumnName)).VoteName;

                CouncillorVote councilorVote = new()
                {
                    CouncillorName = councillorName,
                    Choice = (Choice)ExtractVoteOption(cell),
                    VoteName = voteName
                };
                councilVotes.Votes.Add(councilorVote);

            }
        }

        private string GetCouncillorName(Row row, IList<LocationOfCouncillorDetails> councillorDetailLocations)
        {
            if (councillorDetailLocations.Any(cdl => cdl.ColumnName == FirstNameColumnName))
            {
                // assume first and last columns exist
                CellReference firstNameCellReference = councillorDetailLocations.Single(cdl => cdl.ColumnName == FirstNameColumnName).CellReference;
                CellReference lastNameCellReference = councillorDetailLocations.Single(cdl => cdl.ColumnName == LastNameColumnName).CellReference;

                // Need to pull these from shared strings
                Cell firstNameValue = row.ChildElements.Cast<Cell>().Single(c => c.IsInColumn(firstNameCellReference.ColumnName));
                Cell lastNameValue = row.ChildElements.Cast<Cell>().Single(c => c.IsInColumn(lastNameCellReference.ColumnName));

                return string.Concat(firstNameValue.GetCellValue(_sharedStringTable), " ", lastNameValue.GetCellValue(_sharedStringTable));
            }
            else
            {
                // use keypad sn
                LocationOfCouncillorDetails keypadSn = councillorDetailLocations.Single(cdl => cdl.ColumnName == KeypadSnColumnName);
                return row.ChildElements.Cast<Cell>().Single(c => c.IsInColumn(keypadSn.CellReference.ColumnName)).CellValue.Text;
            }
        }

        private int ExtractVoteOption(Cell cellWithVoteChoice)
        {
            if (cellWithVoteChoice == null)
            {
                return 0;
            }

            string voteChoice = cellWithVoteChoice.GetCellValue(_sharedStringTable);

            bool isStringValue = Enum.TryParse(voteChoice, out Choice option);
            if (isStringValue)
            {
                return (int)option;
            }

            return int.Parse(cellWithVoteChoice.InnerText);
        }

        private IList<(string VoteName, CellReference CellReference)> GetVoteNames(Row headerRow, SharedStringTable sharedStringTable, IEnumerable<LocationOfCouncillorDetails> locationsOfCouncillorDetails)
        {
            List<(string VoteName, CellReference CellReference)> voteNames = headerRow.ChildElements.Cast<Cell>().Select(cell => (GetSharedStringValueOfVoteName(cell, sharedStringTable), new CellReference(cell.CellReference))).ToList();

            voteNames.RemoveAll(vn => locationsOfCouncillorDetails.Any(cd => cd.ColumnName == vn.VoteName));

            return voteNames.ToList();
        }

        private string GetSharedStringValueOfVoteName(Cell cell, SharedStringTable sharedStringTable)
        {
            return sharedStringTable.ChildElements.Cast<SharedStringItem>().ToList()[int.Parse(cell.CellValue.Text)].Text.Text;
        }

        private LocationOfCouncillorDetails GetSharedStringIndexOfHeaderName(string columnName, SharedStringTable sharedStringTable)
        {
            OpenXmlElement keypadSnCell = sharedStringTable.ChildElements.SingleOrDefault(ce => ce.InnerText.Equals(columnName));
            if (keypadSnCell == null)
            {
                return null;
            }

            string sharedStringIndex = sharedStringTable.ChildElements.ToList().IndexOf(keypadSnCell).ToString();
            return new LocationOfCouncillorDetails { ColumnName = columnName, SharedStringIndex = sharedStringIndex };
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
