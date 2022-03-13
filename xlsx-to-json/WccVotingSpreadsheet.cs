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

            for (int i = 5; i < sheetData.ChildElements.Count; i++)
            {
                Row row = (Row)sheetData.ChildElements[i];
                string rowNumber = row.GetAttribute("r", string.Empty).Value;
                string[] spanStartAndEnd = row.GetAttribute("spans", string.Empty).Value.Split(":");
                int columnStart = int.Parse(spanStartAndEnd[0]) - 1;
                int columnEnd = int.Parse(spanStartAndEnd[1]) - 1;

                if (i > 5)
                {
                    Cell councilorCell = (Cell)row.ChildElements[0];
                    councilVotes.Councilors.Add(councilorCell.InnerText);
                }
                for (int j = 1; j < columnEnd; j++)
                {
                    string currentCell = ColumnName(j) + rowNumber;
                    if (i == 5)
                    {
                        SharedStringItem sharedStringItem = (SharedStringItem)sharedStringTablePart.SharedStringTable.ChildElements[int.Parse(row.ChildElements[j].InnerText)];
                        councilVotes.VoteNames.Add(sharedStringItem.InnerText);
                    }
                    else
                    {

                        DocumentFormat.OpenXml.OpenXmlElement? cell = row.ChildElements.SingleOrDefault(ce => ce.GetAttribute("r", "").Value == currentCell);
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
                            VoteName = councilVotes.VoteNames[j-1]
                        };
                        councilVotes.Votes.Add(councilorVote);
                    }
                }
            }

            return councilVotes;
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
