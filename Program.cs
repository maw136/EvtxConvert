// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

const char Separator = ';';
const string RowJoinSeparator = "\";\"";

var inputPath = args[0];
var outputPath = args[1];

Console.WriteLine("Processing files:");
Console.WriteLine("Input: {0}", inputPath);
Console.WriteLine("Output: {0}", outputPath);

using var reader = XmlReader.Create(File.OpenRead(inputPath),
    new XmlReaderSettings
    {
        CloseInput = true,
        Async = true,
        ConformanceLevel = ConformanceLevel.Fragment,
        CheckCharacters = false,
        ValidationType = ValidationType.None,
        ValidationFlags = XmlSchemaValidationFlags.None,
        IgnoreProcessingInstructions = true,
    });

var document = await XDocument.LoadAsync(reader, LoadOptions.SetLineInfo, CancellationToken.None);

var allElements = document.Root!.Elements().ToList();

HashSet<string> validElements = new(new[] { "Event", "Record" }, StringComparer.OrdinalIgnoreCase);

var invalidElements = allElements.Where(e => !validElements.Contains(e.Name.LocalName)).ToList();
if (invalidElements.Count > 0)
{
    throw new InvalidOperationException(
        $"Not All elements are {string.Join(',', validElements)} elements (under the root): {string.Join(',', invalidElements.Select(e => $"{e.Name.LocalName}: {((IXmlLineInfo)e).LineNumber}, {((IXmlLineInfo)e).LinePosition}"))}");
}

var columns = allElements.SelectMany(FlattenSubelementNames).Distinct().ToList();

Console.WriteLine("Collected data columns: {0}, rows: {1}", columns.Count, allElements.Count);

var rows = allElements.Select(ToRow);

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
    Sheet sheet = new()
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

    static IEnumerable<string> FlattenSingleElement(XElement subelement)
    {
        if (subelement.IsEmpty)
        {
            foreach (XAttribute xAttribute in subelement.Attributes())
            {
                yield return $"{subelement.Name.LocalName}_{xAttribute.Name.LocalName}";
            }
        }
        else
        {
            if (subelement.Parent.Elements(subelement.Name).Count() > 1)
            {
                if (StringComparer.OrdinalIgnoreCase.Equals(subelement.Name.LocalName, "Data"))
                {
                    var nameAttr = subelement.Attribute("Name")!;
                    yield return $"{subelement!.Parent!.Name.LocalName}_{nameAttr.Value}";
                }
                else
                {
                    yield return $"{subelement!.Parent!.Name.LocalName}_{subelement.Name.LocalName}";
                }
            }
            else
            {
                if (StringComparer.OrdinalIgnoreCase.Equals(subelement.Parent.Name.LocalName, "Substitution"))
                {
                    var indexAttr = subelement.Parent.Attribute("index")!;
                    yield return $"{subelement!.Parent!.Name.LocalName}_index_{indexAttr.Value}_{subelement.Name.LocalName}";
                }
                else
                {
                    yield return $"{subelement!.Parent!.Name.LocalName}_{subelement.Name.LocalName}";
                }
            }
        }
    }
}

static SimpleRow ToRow(XElement eventElement)
{
    var values = eventElement.Descendants().Where(CanBeFlattened).Aggregate(new Dictionary<string, string>(), FlattenSingleElement);

    return new SimpleRow(values);

    static Dictionary<string, string> FlattenSingleElement(Dictionary<string, string> values, XElement subelement)
    {
        if (subelement.IsEmpty)
        {
            foreach (XAttribute xAttribute in subelement.Attributes())
            {
                values.Add($"{subelement.Name.LocalName}_{xAttribute.Name.LocalName}", SanitizeValue(xAttribute.Value));
            }
        }
        else
        {
            if (subelement.Parent!.Elements(subelement.Name).Count() > 1)
            {
                if (StringComparer.OrdinalIgnoreCase.Equals(subelement.Name.LocalName, "Data"))
                {
                    var nameAttr = subelement.Attribute("Name")!;
                    values.Add($"{subelement.Parent!.Name.LocalName}_{nameAttr.Value}",
                        SanitizeValue(subelement.Value));
                }
                else
                {
                    values.Add($"{subelement!.Parent!.Name.LocalName}_{subelement.Name.LocalName}", SanitizeValue(subelement.Value));
                }
            }
            else
            {
                if (StringComparer.OrdinalIgnoreCase.Equals(subelement.Parent.Name.LocalName, "Substitution"))
                {
                    var indexAttr = subelement.Parent.Attribute("index")!;
                    values.Add(
                        $"{subelement!.Parent!.Name.LocalName}_index_{indexAttr.Value}_{subelement.Name.LocalName}",
                        SanitizeValue(subelement.Value));
                }
                else
                {
                    values.Add($"{subelement!.Parent!.Name.LocalName}_{subelement.Name.LocalName}",
                        SanitizeValue(subelement.Value));
                }
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
