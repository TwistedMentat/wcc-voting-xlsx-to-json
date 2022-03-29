using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace xlsx_to_json
{
    public static class CellHelpers
    {
        private static readonly Regex _splitCellReferenceRegex = new("^([a-zA-Z]+)(\\d+)$");

        public static bool IsInSameColumn(this Cell thisCell, Cell cellToCompare)
        {
            MatchCollection thisCellMatchCollection = _splitCellReferenceRegex.Matches(thisCell.CellReference);
            string thisCellColumnName = thisCellMatchCollection[0].Groups[1].Value;

            MatchCollection compareCellMatchCollection = _splitCellReferenceRegex.Matches(cellToCompare.CellReference);
            string compareCellColumnName = compareCellMatchCollection[0].Groups[1].Value;

            return thisCellColumnName == compareCellColumnName;
        }

        public static bool IsInColumn(this Cell thisCell, string columnName)
        {
            MatchCollection thisCellMatchCollection = _splitCellReferenceRegex.Matches(thisCell.CellReference);
            string thisCellColumnName = thisCellMatchCollection[0].Groups[1].Value;

            return thisCellColumnName == columnName;
        }

        public static bool IsInSameRow(this Cell thisCell, Cell cellToCompare)
        {
            return ((Row)thisCell.Parent).RowIndex == ((Row)cellToCompare.Parent).RowIndex;
        }

        public static string GetCellValue(this Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                return ((SharedStringItem)sharedStringTable.ToList()[int.Parse(cell.CellValue.Text)]).Text.Text;
            }
            return cell.CellValue.Text;
        }

    }
}
