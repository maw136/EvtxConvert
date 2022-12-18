// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Xml.Linq;

const char Separator = ';';
const string RowJoinSeparator = "\";\"";

var inputPath = args[0];
var outputPath = args[1];

Console.WriteLine("Processing files:");
Console.WriteLine("Input: {0}", inputPath);
Console.WriteLine("Output: {0}", outputPath);

using var inputFile = File.OpenRead(inputPath);

var document = await XDocument.LoadAsync(inputFile, LoadOptions.SetLineInfo, CancellationToken.None);

var allEvents = document.Root!.Elements().ToList();

if (allEvents.Any(e => e.Name.LocalName != "Event"))
    throw new InvalidOperationException("Not All elements are 'Event' elements (under the root)");

var columns = allEvents.SelectMany(FlattenSubelementNames).Distinct().ToList();

Console.WriteLine("Collected data columns: {0}, rows: {1}", columns.Count, allEvents.Count);

var rows = allEvents.Select(ToRow);

//CreateExcelFile(outputPath, Path.GetFileNameWithoutExtension(outputPath), sheetData => PrepareData(sheetData, rows, columns));

CreateCSVFile(outputPath, rows, columns);

Console.WriteLine("Finished");

return 0;

static void CreateCSVFile(string outputPath, IEnumerable<SimpleRow> rows, IList<string> columns)
{
    using var outputWriter = new StreamWriter(outputPath, Encoding.UTF8,
        new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.ReadWrite, Share = FileShare.Read });

    var allColumnsAsHeader = string.Join(Separator, columns);
    outputWriter.WriteLine(allColumnsAsHeader);
    foreach (var row in rows)
    {
        outputWriter.WriteLine(BuildRowString(row, columns));
    }
}

static void CreateExcelFile(string filepath, string sheetName, Action<SheetData> manipulateData)
{
    // Create a spreadsheet document by supplying the filepath.
    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(
        filepath,
        SpreadsheetDocumentType.Workbook
    );
    // Add a WorkbookPart to the document.
    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
    workbookpart.Workbook = new Workbook();
    // Add a WorksheetPart to the WorkbookPart.
    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData();
    manipulateData(sheetData);
    worksheetPart.Worksheet = new Worksheet(sheetData);
    // Add Sheets to the Workbook.
    Sheets sheets = workbookpart.Workbook.AppendChild(new Sheets());
    // Append a new worksheet and associate it with the workbook.
    Sheet sheet = new ()
    {
        Id = workbookpart.GetIdOfPart(worksheetPart),
        SheetId = 1,
        Name = sheetName
    };
    sheets.Append(sheet);
    //Save & close
    workbookpart.Workbook.Save();
    spreadsheetDocument.Close();
}

static void PrepareData(SheetData sheetData, IEnumerable<SimpleRow> rows, IList<string> columns)
{
    Row headerRow = new();

    foreach (string column in columns)
    {
        Cell cell = new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(column)
        };
        headerRow.AppendChild(cell);
    }

    sheetData.AppendChild(headerRow);

    foreach (var row in rows)
    {
        Row newRow = new();
        foreach (string col in columns)
        {
            var newValue = row.Values.TryGetValue(col, out var value) ? value : string.Empty;
            Cell cell = new()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(newValue)
            };
            newRow.AppendChild(cell);
        }

        sheetData.AppendChild(newRow);
    }
}

static bool CanBeFlattened(XElement eventSubelement)
{
    return !eventSubelement.HasElements;
}

static IEnumerable<string> FlattenSubelementNames(XElement eventElement)
{
    return eventElement.Descendants().Where(CanBeFlattened).SelectMany(FlattenSingleElement);

    static IEnumerable<string> FlattenSingleElement(XElement eventSubelement)
    {
        if (eventSubelement.IsEmpty)
        {
            foreach (XAttribute xAttribute in eventSubelement.Attributes())
            {
                yield return $"{eventSubelement.Name.LocalName}_{xAttribute.Name.LocalName}";
            }
        }
        else
        {
            if (eventSubelement.Parent.Elements(eventSubelement.Name).Count() > 1)
            {
                //       if (StringComparer.OrdinalIgnoreCase.Equals(eventSubelement.Name.LocalName, "data"))
                var nameAttr = eventSubelement.Attribute("Name")!;
                yield return $"{eventSubelement!.Parent!.Name.LocalName}_{nameAttr.Value}";
            }
            else
            {
                yield return $"{eventSubelement!.Parent!.Name.LocalName}_{eventSubelement.Name.LocalName}";
            }
        }
    }
}

static SimpleRow ToRow(XElement eventElement)
{
    var values = eventElement.Descendants().Where(CanBeFlattened).Aggregate(new Dictionary<string, string>(), FlattenSingleElement);

    return new SimpleRow(values);

    static Dictionary<string, string> FlattenSingleElement(Dictionary<string, string> values, XElement eventSubelement)
    {
        if (eventSubelement.IsEmpty)
        {
            foreach (XAttribute xAttribute in eventSubelement.Attributes())
            {
                values.Add($"{eventSubelement.Name.LocalName}_{xAttribute.Name.LocalName}", SanitizeValue(xAttribute.Value));
            }
        }
        else
        {
            if (eventSubelement.Parent.Elements(eventSubelement.Name).Count() > 1)
            {
                //       if (StringComparer.OrdinalIgnoreCase.Equals(eventSubelement.Name.LocalName, "data"))
                var nameAttr = eventSubelement.Attribute("Name")!;
                values.Add($"{eventSubelement.Parent!.Name.LocalName}_{nameAttr.Value}", SanitizeValue(eventSubelement.Value));
            }
            else
            {
                values.Add($"{eventSubelement!.Parent!.Name.LocalName}_{eventSubelement.Name.LocalName}", SanitizeValue(eventSubelement.Value));
            }
        }

        return values;
    }
}

static StringBuilder BuildRowString(SimpleRow row, IEnumerable<string> columns)
{
    var valuesOfColumns = columns.Select(c => row.Values.TryGetValue(c, out var value) ? value : String.Empty);

    var sb = new StringBuilder();
    sb.Append('"');
    sb.AppendJoin(RowJoinSeparator, valuesOfColumns);
    sb.Append('"');

    return sb;
}


static string SanitizeValue(string value)
{
    //return value;
    return value.ReplaceLineEndings(" ");
}


readonly record struct SimpleRow(Dictionary<string, string> Values);
