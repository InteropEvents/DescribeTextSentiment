using Azure;
using Azure.AI.TextAnalytics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

string path = @"C:\source\tmp\reviews - Copy.xlsx";
string languageKey = Environment.GetEnvironmentVariable("LANGUAGE_KEY") ?? throw new ArgumentException("LANGUAGE_KEY");
string languageEndpoint = Environment.GetEnvironmentVariable("LANGUAGE_ENDPOINT") ?? throw new Exception("LANGUAGE_ENDPOINT");
AzureKeyCredential credentials = new(languageKey);
Uri endpoint = new(languageEndpoint);
TextAnalyticsClient client = new(endpoint, credentials);

using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, true))
{
    Workbook? workbook = spreadsheetDocument?.WorkbookPart?.Workbook;
    SheetData? sheetData = spreadsheetDocument?.WorkbookPart?.GetPartsOfType<WorksheetPart>().FirstOrDefault()?.Worksheet?.GetFirstChild<SheetData>();

    if (sheetData is not null)
    {
        SharedStringTable? sharedStringTable = spreadsheetDocument?.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

        IEnumerable<Row> rows = sheetData.ChildElements.OfType<Row>();

        foreach (Row row in rows)
        {
            bool breakOuter = false;
            IEnumerable<Cell> cells = row.ChildElements.OfType<Cell>();

            foreach (Cell cell in cells)
            {
                if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable is not null && cell.CellValue is not null)
                {
                    uint index = uint.Parse(cell.CellValue.InnerText);
                    string text = GetSharedString(sharedStringTable, index);

                    DocumentSentiment review = client.AnalyzeSentiment(text, options: new AnalyzeSentimentOptions()
                    {
                        IncludeOpinionMining = true
                    });

                    Console.WriteLine($"Document sentiment: {review.Sentiment}\n");

                    foreach (SentenceSentiment sentenceSentiment in review.Sentences)
                    {
                        Console.WriteLine($"\tSentence sentiment: {sentenceSentiment.Sentiment}");
                        Console.WriteLine($"\tText: {sentenceSentiment.Text}");

                        Console.WriteLine($"\t\tPositive: {sentenceSentiment.ConfidenceScores.Positive}");
                        Console.WriteLine($"\t\tNeutral: {sentenceSentiment.ConfidenceScores.Neutral}");
                        Console.WriteLine($"\t\tNegative: {sentenceSentiment.ConfidenceScores.Negative}\n");
                    }

                    uint idx = GetOrAddSharedStringIndex(sharedStringTable, review.Sentiment.ToString());
                    string column = !string.IsNullOrEmpty(cell.CellReference?.ToString())
                        ? string.Concat(cell.CellReference.ToString()!.TakeWhile(c => !char.IsDigit(c) && char.IsLetter(c)))
                        : string.Empty;

                    if (!string.IsNullOrEmpty(column))
                    {
                        Cell newCell = new(
                            new CellValue((int)idx))
                        { CellReference = string.Concat(GetNextLetter(column), row.RowIndex), DataType = CellValues.SharedString };

                        row.AppendChild(newCell);

                        breakOuter = true;
                    }
                }

                if (breakOuter)
                {
                    break;
                }
            }
        }
    }
}

uint GetOrAddSharedStringIndex(SharedStringTable sharedStringTable, string text)
{
    IEnumerable<SharedStringItem> sharedStringItems = sharedStringTable.Descendants<SharedStringItem>();

    for (int i = 0; i < sharedStringItems.Count(); i++)
    {
        if (sharedStringTable.ElementAt(i).InnerText == text)
        {
            return (uint)i;
        }
    }
    // If not found, add it to the shared string table
    sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
    return (uint)(sharedStringItems.Count() - 1);
}

string GetSharedString(SharedStringTable sharedStringTable, uint index)
{
    if (index < sharedStringTable.Count())
    {
        return sharedStringTable.Descendants<SharedStringItem>().ElementAt((int)index).InnerText;
    }

    throw new ArgumentOutOfRangeException(nameof(index), "Index is out of range of the shared string table.");
}

string GetNextLetter(string input)
{
    if (string.IsNullOrEmpty(input))
        throw new ArgumentException("Input must be a non-empty string.", nameof(input));

    // Convert to lowercase for consistency
    input = input.ToLowerInvariant();

    // Check if all characters are letters and all the same
    if (!input.All(char.IsLetter))
        throw new ArgumentException("Input must contain only English alphabet letters.", nameof(input));
    if (!input.All(c => c == input[0]))
        throw new ArgumentException("All characters in input must be the same letter for this sequence.", nameof(input));

    char current = input[0];
    int length = input.Length;

    // Find next letter
    if (current < 'z')
    {
        char next = (char)(current + 1);
        return new string(next, length);
    }
    else
    {
        // After 'z', increase length and start with 'a'
        return new string('a', length + 1);
    }
}