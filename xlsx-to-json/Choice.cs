using CsvHelper.Configuration.Attributes;

namespace xlsx_to_json
{
    public enum Choice
    {
        [Name("Unset")]
        Unset,
        [Name("Yes")]
        Yes,
        [Name("No")]
        No,
        [Name("Abstain")]
        Abstain
    }
}
