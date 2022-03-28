// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using xlsx_to_json;
using System.Text.Json;
using System.Text.Json.Serialization;

Console.WriteLine("Hello, World!");

SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(args[0], false);

WccVotingSpreadsheet wccVotingSpreadsheet = new WccVotingSpreadsheet(spreadsheetDocument);


CouncilVotes councilVotes = wccVotingSpreadsheet.TransformExcel();

JsonSerializerOptions options = new JsonSerializerOptions
{
    Converters =
    {
        new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
    }
};
string outputJson = JsonSerializer.Serialize(councilVotes, options);

File.WriteAllText($"wcc-votes-{DateTime.UtcNow.Ticks}.json", outputJson);