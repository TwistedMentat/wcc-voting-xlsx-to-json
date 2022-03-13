// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using xlsx_to_json;
using System.Text.Json;

Console.WriteLine("Hello, World!");

SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(args[0], false);

WccVotingSpreadsheet wccVotingSpreadsheet = new WccVotingSpreadsheet();


CouncilVotes councilVotes = wccVotingSpreadsheet.TransformExcel(spreadsheetDocument.WorkbookPart);

string outputJson = JsonSerializer.Serialize(councilVotes);

File.WriteAllText($"wcc-votes-{DateTime.UtcNow.Millisecond}", outputJson);