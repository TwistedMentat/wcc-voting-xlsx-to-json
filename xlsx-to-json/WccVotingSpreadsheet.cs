using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx_to_json
{
    public class WccVotingSpreadsheet
    {
        public CouncilVotes TransformExcel(WorkbookPart workbookPart)
        {
            CouncilVotes councilVotes = new CouncilVotes();

            IEnumerable<WorksheetPart> worksheetParts = workbookPart.WorksheetParts;
            WorksheetPart worksheetPart = worksheetParts.First();
            SheetData sheetData = (SheetData)worksheetPart.RootElement.ChildElements.Single(w => w is SheetData);
            SharedStringTablePart? sharedStringTablePart = workbookPart.SharedStringTablePart;

            // Why not just make an in memory table with all the properties you want? Then just use objects and LINQ to chop it up as needed :/

            IDictionary<string, string> sharedStringIndexesForValuesWanted = new Dictionary<string, string>();

            GetSharedStringIndex(sharedStringTablePart, sharedStringIndexesForValuesWanted, "Keypad SN");
            GetSharedStringIndex(sharedStringTablePart, sharedStringIndexesForValuesWanted, "First Name");
            GetSharedStringIndex(sharedStringTablePart, sharedStringIndexesForValuesWanted, "Last Name");


            Dictionary<string, string> cellReferencesForValuesWanted = new();

            foreach (Cell? cell in sheetData.ChildElements.SelectMany(row => row.ChildElements.Where(cell => sharedStringIndexesForValuesWanted.ContainsKey(cell.InnerText))))
            {
                if (cell == null)
                {
                    continue;
                }
                // the below will have the full cell reference
                cellReferencesForValuesWanted[sharedStringIndexesForValuesWanted[cell.InnerText]] = cell.CellReference.Value;
            }

            bool keepCheckingForVotingSectionStart = true;

            Row startingRow = sheetData.ChildElements.SelectMany(row => row.ChildElements.Where(cell => ((Cell)cell).DataType?.Value == CellValues.SharedString && ((Cell)cell).InnerText == "5")).Single().Parent as Row;

            // Need to find where the "Keypad SN" cell is. That will define the row to start.
            for (int i = 0; i < sheetData.ChildElements.Count; i++)
            {
                Row row = (Row)sheetData.ChildElements[i];

                string? rowNumberAttribute = row.GetAttribute("r", string.Empty).Value;
                if (keepCheckingForVotingSectionStart && row.RowIndex != startingRow.RowIndex)
                {
                    continue;
                }
                else
                {
                    keepCheckingForVotingSectionStart = false;
                }

                string[] spanStartAndEnd = row.Spans.Items.First().Value.Split(":");
                int columnStart = int.Parse(spanStartAndEnd[0]) - 1;
                int columnEnd = int.Parse(spanStartAndEnd[1]) - 1;

                if (i > 5)
                {
                    Cell councilorCell = (Cell)row.ChildElements[0];
                    councilVotes.Councilors.Add(councilorCell.CellValue.Text);
                }
                for (int j = 1; j < columnEnd; j++)
                {
                    string currentCell = ColumnName(j) + rowNumberAttribute;
                    if (i == 5)
                    {
                        SharedStringItem sharedStringItem = (SharedStringItem)sharedStringTablePart.SharedStringTable.ChildElements[int.Parse(row.ChildElements[j].InnerText)];
                        councilVotes.VoteNames.Add(sharedStringItem.InnerText);
                    }
                    else
                    {

                        OpenXmlElement? cell = row.ChildElements.SingleOrDefault(ce => ((Cell)ce).CellReference == currentCell);
                        
                        int choiceValue;
                        if (cell == null)
                        {
                            choiceValue = 0;
                        }
                        else
                        {
                            choiceValue = int.Parse(cell.InnerText);
                        }

                        CouncilorVote councilorVote = new()
                        {
                            CouncilorName = row.ChildElements[0].InnerText,
                            Choice = (Choice)choiceValue,
                            VoteName = councilVotes.VoteNames[j - 1]
                        };
                        councilVotes.Votes.Add(councilorVote);
                    }
                }
            }

            return councilVotes;
        }

        private static void GetSharedStringIndex(SharedStringTablePart? sharedStringTablePart, IDictionary<string, string> sharedStringIndexesForValuesWanted, string columnName)
        {
            OpenXmlElement keypadSnCell = sharedStringTablePart.SharedStringTable.ChildElements.SingleOrDefault(ce => ce.InnerText.Equals(columnName));
            if (keypadSnCell == null)
            {
                return;
            }

            string indexOfKeypadSnString = sharedStringTablePart.SharedStringTable.ChildElements.ToList().IndexOf(keypadSnCell).ToString();
            sharedStringIndexesForValuesWanted[indexOfKeypadSnString] = columnName;
        }

        // joinked from here https://stackoverflow.com/a/297214
        private string ColumnName(int index)
        {
            index -= 1; //adjust so it matches 0-indexed array rather than 1-indexed column

            int quotient = index / 26;
            if (quotient > 0)
                return ColumnName(quotient) + chars[index % 26].ToString();
            else
                return chars[index % 26].ToString();
        }
        private char[] chars = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
    }
}
