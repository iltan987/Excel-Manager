using System.IO.Compression;
using System.Xml;

namespace ExcelManager
{
    public class Workbook
    {
        public List<Worksheet> workSheets { get; private set; } = new List<Worksheet>();

        public Worksheet DefaultWorkSheet => workSheets[0];

        public string? FilePath { get; private set; }

        public static Workbook Create(string defaultSheetName = "Sheet1")
        {
            if (string.IsNullOrWhiteSpace(defaultSheetName))
                throw new ArgumentException($"'{nameof(defaultSheetName)}' cannot be null or whitespace.", nameof(defaultSheetName));

            Workbook workbook = new()
            {
                workSheets = new List<Worksheet> { new Worksheet(defaultSheetName) }
            };
            return workbook;
        }

        public static Workbook LoadExcel(string path)
        {
            Workbook workBook;
            ZipArchive archive;
            XmlDocument? xmlDoc;
            XmlNamespaceManager namespaceManager;
            XmlNodeList? nodeList;
            XmlNodeList? cols;
            List<string> SharedStrings;
            XmlNode? valueNode;
            string attr;
            Worksheet worksheet;



            workBook = new()
            {
                FilePath = path
            };

            using (archive = ZipFile.OpenRead(path))
            {
                xmlDoc = EntryReader(archive, "xl/sharedStrings.xml");
                if (xmlDoc != null)
                {
                    namespaceManager = new(xmlDoc.NameTable);
                    namespaceManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    nodeList = xmlDoc.SelectNodes("//ns:si/ns:t", namespaceManager);

                    if (nodeList is null)
                        throw new Exception($"Shared Strings are not found.{Environment.NewLine}File: \"{path}\"");

                    SharedStrings = new List<string>();
                    foreach (XmlNode item in nodeList)
                        SharedStrings.Add(item.InnerText);
                }
                else
                    SharedStrings = new List<string>();

                xmlDoc = EntryReader(archive, "xl/workbook.xml");
                namespaceManager = new(xmlDoc!.NameTable);
                namespaceManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                nodeList = xmlDoc.SelectNodes("//ns:sheets/ns:sheet", namespaceManager);

                if (nodeList is null)
                    throw new Exception($"Sheets are not found.{Environment.NewLine}File: \"{path}\"");

                foreach (XmlElement sheet in nodeList)
                    workBook.workSheets.Add(new Worksheet(sheet.GetAttribute("name")) { xmlDocument = EntryReader(archive, "xl/worksheets/sheet" + sheet.GetAttribute("sheetId") + ".xml") });

                for (int i = 0; i < workBook.workSheets.Count; i++)
                {
                    worksheet = workBook.workSheets[i];
                    if (worksheet.xmlDocument is null)
                        throw new Exception($"WorkBook -> WorkSheet[\"{i}\"] (\"{worksheet.Name}\") -> XmlDocument is not found.{Environment.NewLine}File: \"{workBook.FilePath}\"");

                    namespaceManager = new(worksheet.xmlDocument.NameTable);
                    namespaceManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                    nodeList = worksheet.xmlDocument.SelectNodes("//ns:sheetData/ns:row", namespaceManager);

                    if (nodeList is null)
                        throw new Exception($"WorkBook -> WorkSheet[\"{i}\"] (\"{worksheet.Name}\") -> XmlDocument -> Row Information is not found.{Environment.NewLine}File: \"{workBook.FilePath}\"");

                    foreach (XmlElement rowElement in nodeList)
                    {
                        attr = rowElement.GetAttribute("r");
                        worksheet.Rows[attr] = new Dictionary<string, Cell>();

                        cols = rowElement.SelectNodes("ns:c", namespaceManager);

                        if (cols is null)
                            throw new Exception($"WorkBook -> Worksheet[\"{i}\"] (\"{worksheet.Name}\") -> XmlDocument -> Row \"{attr}\" -> Column Information is not found.{Environment.NewLine}File: \"{workBook.FilePath}\"");

                        foreach (XmlElement col in cols)
                        {
                            valueNode = col.SelectSingleNode("ns:v", namespaceManager);
                            if (valueNode is null)
                                throw new Exception($"WorkBook -> Worksheet[\"{i}\"] (\"{worksheet.Name}\") -> XmlDocument -> Row \"{attr}\" -> Column -> Value is not found.{Environment.NewLine}File: \"{workBook.FilePath}\"");
                            worksheet.Rows[attr][col.GetAttribute("r")] = col.GetAttribute("t") == "s"
                                ? new Cell() { type = "s", StringValue = SharedStrings[int.Parse(valueNode.InnerText)] }
                                : new Cell() { DecimalValue = decimal.Parse(valueNode.InnerText) };
                        }
                    }
                }
            }

            return workBook;
        }

        public void SaveAs(string path)
        {
            using var fileStream = new FileStream(path, FileMode.Create);
            using ZipArchive archive = new(fileStream, ZipArchiveMode.Create);

            #region Done or nothing to edit
            // Nothing to edit
            EntrySaver(archive, "_rels\\.rels", @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml"" />
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml"" />
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" />
</Relationships>");

            // No need to edit. Already done
            EntrySaver(archive, "docProps\\app.xml", $@"<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size=""2"" baseType=""variant"">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>{workSheets.Count}</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size=""{workSheets.Count}"" baseType=""lpstr"">
      {string.Join($"{Environment.NewLine}      ", workSheets.ConvertAll(f => $"<vt:lpstr>{f.Name}</vt:lpstr>"))}
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>");

            // No need to edit. Already done
            EntrySaver(archive, "docProps\\core.xml", $@"<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>iltan</dc:creator>
  <dcterms:created xsi:type=""dcterms:W3CDTF"">{DateTime.Now:yyyy\-MM\-dd\THH:mm:ss\Z}</dcterms:created>
</cp:coreProperties>");

            // No need to edit. Already done
            EntrySaver(archive, "xl\\_rels\\workbook.xml.rels", $@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId{workSheets.Count + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml"" />
  {string.Join(Environment.NewLine + "  ", Enumerable.Range(1, workSheets.Count).Reverse().ToList().ConvertAll(f => $"<Relationship Id=\"rId{f}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{f}.xml\" />"))}
  <Relationship Id=""rId{workSheets.Count + 2}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""sharedStrings.xml"" />
</Relationships>");
            #endregion

            // One of the most important parts! Edit here!
            List<string> SharedStrings = new();
            #region Just Declaring Part
            Worksheet workSheet;
            string res;
            string? strVal;
            int i, index, sharedStringCounter = 0;
            #endregion
            for (i = 0; i < workSheets.Count; i++)
            {
                workSheet = workSheets[i];
                res = $@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac xr xr2 xr3"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"" xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision"" xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"" xmlns:xr3=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"" xr:uid=""{{00000000-0001-0000-0000-000000000000}}"">
  <dimension ref=""{workSheet.Rows.First().Value.First().Key}"" />
  <sheetViews>
    <sheetView tabSelected=""1"" workbookViewId=""0"" />
  </sheetViews>
  <sheetFormatPr defaultRowHeight=""15"" x14ac:dyDescent=""0.25"" />
  <sheetData>";
                foreach (var row in workSheet.Rows)
                {
                    res += $@"
    <row r=""{row.Key}"">";
                    foreach (var item in row.Value)
                    {
                        if (item.Value.type == "s")
                        {
                            sharedStringCounter++;
                            strVal = item.Value.StringValue;
                            if (strVal is null)
                                throw new Exception("String value is null");
                            if ((index = SharedStrings.IndexOf(strVal)) == -1)
                            {
                                index = SharedStrings.Count;
                                SharedStrings.Add(strVal);
                            }
                            res += $@"
      <c r=""{item.Key}"" t=""s"">
        <v>{index}</v>
      </c>";
                        }
                        else
                            res += $@"
      <c r=""{item.Key}"">
        <v>{item.Value.Value}</v>
      </c>";
                    }
                    res += @"
    </row>";
                }

                EntrySaver(archive, "xl\\worksheets\\sheet" + (i + 1) + ".xml", res + @"
  </sheetData>
  <pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3"" />
</worksheet>");
            }

            // One of the most important parts! Edit here!
            EntrySaver(archive, "xl\\sharedStrings.xml", $@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""{sharedStringCounter}"" uniqueCount=""{SharedStrings.Count}"">
{string.Join(Environment.NewLine + "  ", SharedStrings.ConvertAll(f => $@"<si>
    <t>{f}</t>
  </si>"))}
</sst>");

            #region Done or nothing to edit
            // Nothing to edit
            EntrySaver(archive, "xl\\styles.xml", @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac x16r2"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"" xmlns:x16r2=""http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"">
  <fonts count=""1"" x14ac:knownFonts=""1"">
    <font>
      <sz val=""11"" />
      <color theme=""1"" />
      <name val=""Calibri"" />
      <family val=""2"" />
      <scheme val=""minor"" />
    </font>
  </fonts>
  <fills count=""2"">
    <fill>
      <patternFill patternType=""none"" />
    </fill>
    <fill>
      <patternFill patternType=""gray125"" />
    </fill>
  </fills>
  <borders count=""1"">
    <border>
      <left />
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
  </borders>
  <cellStyleXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" />
  </cellStyleXfs>
  <cellXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" />
  </cellXfs>
  <cellStyles count=""1"">
    <cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
  </cellStyles>
  <dxfs count=""0"" />
  <tableStyles count=""0"" defaultTableStyle=""TableStyleMedium2"" defaultPivotStyle=""PivotStyleLight16"" />
  <extLst>
    <ext uri=""{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">
      <x14:slicerStyles defaultSlicerStyle=""SlicerStyleLight1"" />
    </ext>
    <ext uri=""{9260A510-F301-46a8-8635-F512D64BE5F5}"" xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"">
      <x15:timelineStyles defaultTimelineStyle=""TimeSlicerStyleLight1"" />
    </ext>
  </extLst>
</styleSheet>");

            // No need to edit. Already done.
            EntrySaver(archive, "xl\\workbook.xml", $@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x15 xr xr6 xr10 xr2"" xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"" xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision"" xmlns:xr6=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"" xmlns:xr10=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"" xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"">
  <workbookPr />
  <bookViews>
    <workbookView tabRatio=""600"" />
  </bookViews>
  <sheets>
    {string.Join(Environment.NewLine + "    ", Enumerable.Range(0, workSheets.Count).ToList().ConvertAll(f => $"<sheet name=\"{workSheets[f].Name}\" sheetId=\"{f + 1}\" r:id=\"rId{f + 1}\" />"))}
  </sheets>
</workbook>");

            // No need to edit. Already done.
            EntrySaver(archive, "[Content_Types].xml", $@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
  <Default Extension=""xml"" ContentType=""application/xml"" />
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" />
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
  <Override PartName=""/xl/sharedStrings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"" />
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml"" />
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml"" />
  {string.Join($"{Environment.NewLine}  ", Enumerable.Range(0, workSheets.Count).ToList().ConvertAll(f => "<Override PartName=\"/xl/worksheets/sheet" + (f + 1) + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />"))}
</Types>");
            #endregion
        }

        private static void EntrySaver(ZipArchive archive, string entryName, string val)
        {
            using StreamWriter sw = new(archive.CreateEntry(entryName, CompressionLevel.Optimal).Open());
            sw.Write(val);
        }

        private static XmlDocument? EntryReader(ZipArchive archive, string entryName)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            if (string.IsNullOrEmpty(entryName))
                throw new ArgumentException($"'{nameof(entryName)}' cannot be null or empty.", nameof(entryName));

            var entry = archive.GetEntry(entryName);
            if (entry is null)
            {
                return entryName == "xl/sharedStrings.xml" ? null : throw new Exception($"No entry found called {entryName}");
            }
            XmlDocument xmlDocument = new();
            using StreamReader sr = new(entry.Open());
            xmlDocument.LoadXml(sr.ReadToEnd());

            return xmlDocument;
        }
    }
}