﻿using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace xlsx_to_json
{

    public class CellReference
    {
        public string ColumnName { get; private set; }
        public int RowNumber { get; private set; }

        private readonly Regex _splitCellReferenceRegex = new("^([a-zA-Z]+)(\\d+)$");
        public CellReference(string cellReference)
        {
            MatchCollection matchCollection = _splitCellReferenceRegex.Matches(cellReference);
            Console.WriteLine(matchCollection.Count);
            // get column
            ColumnName = matchCollection[0].Groups[1].Value;
            RowNumber = int.Parse(matchCollection[0].Groups[2].Value);
        }

        public bool IsInSameColumn(CellReference cellToCompare)
        {
            if (ColumnName == cellToCompare.ColumnName)
            {
                return true;
            }
            return false;
        }

        public bool IsInSameRow(CellReference cellToCompare)
        {
            if (RowNumber == cellToCompare.RowNumber)
            {
                return true;
            }
            return false;
        }
    }
}
