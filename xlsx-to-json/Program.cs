// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using xlsx_to_json;
using System.Text.Json;
using System.Text.Json.Serialization;
using CsvHelper;

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

// Write out a csv of just the votes

using StreamWriter streamWriter = new StreamWriter($"wcc-votes-{DateTime.UtcNow.Ticks}.csv");
using CsvWriter csvWriter = new CsvWriter(streamWriter, System.Globalization.CultureInfo.InvariantCulture);

csvWriter.WriteRecords<CouncillorVote>(councilVotes.Votes);
