using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;
using System.Globalization;

namespace ClasseursRouges
{
    public class ExcelCell
    {
        public string Value;
        public bool IsNumeric;
        public int StyleIndex;
        // Styles:
        // 0 = Normal left text
        // 1 = Header dark blue/red bold white text centered
        // 2 = Subheader gray bold centered
        // 3 = Grade normal centered (>= 4.0)
        // 4 = Grade insufficient (< 4.0) light red background, dark red bold text
        // 5 = Grade parenthesized / C.M / Disp italic gray centered
        // 6 = Moyenne bold centered (>= 4.0)
        // 7 = Moyenne insufficient (< 4.0) bold red background
        // 8 = Section title bold left
        // 9 = Principal discipline name (bold left)
    }

    public class ExcelRow
    {
        public List<ExcelCell> Cells = new List<ExcelCell>();
        public double Height = 20.0;
    }

    public class ExcelSheet
    {
        public string Name;
        public List<ExcelRow> Rows = new List<ExcelRow>();
        public List<double> ColumnWidths = new List<double>();
        public List<string> MergedCells = new List<string>(); // e.g. "A1:F1"
    }

    public class XlsxWriter
    {
        public static void SaveWorkbook(string filePath, List<ExcelSheet> sheets)
        {
            if (File.Exists(filePath)) File.Delete(filePath);

            using (FileStream fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml", BuildContentTypes(sheets.Count));
                WriteEntry(zip, "_rels/.rels", BuildRootRels());
                WriteEntry(zip, "xl/workbook.xml", BuildWorkbookXml(sheets));
                WriteEntry(zip, "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheets.Count));
                WriteEntry(zip, "xl/styles.xml", BuildStylesXml());

                for (int i = 0; i < sheets.Count; i++)
                {
                    WriteEntry(zip, string.Format("xl/worksheets/sheet{0}.xml", i + 1), BuildSheetXml(sheets[i]));
                }
            }
        }

        private static void WriteEntry(ZipArchive zip, string path, string content)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (Stream s = entry.Open())
            using (StreamWriter sw = new StreamWriter(s, new UTF8Encoding(false)))
            {
                sw.Write(content);
            }
        }

        private static string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;")
                    .Replace("'", "&apos;");
        }

        public static string GetColName(int colIndexZeroBased)
        {
            int dividend = colIndexZeroBased + 1;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static string BuildContentTypes(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.AppendFormat("<Override PartName=\"/xl/worksheets/sheet{0}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i);
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string BuildRootRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                   "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                   "</Relationships>";
        }

        private static string BuildWorkbookXml(List<ExcelSheet> sheets)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheets>");
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < sheets.Count; i++)
            {
                string safeName = sheets[i].Name;
                if (string.IsNullOrWhiteSpace(safeName)) safeName = "Feuille" + (i + 1);
                foreach (char c in new char[] { '\\', '/', '?', '*', '[', ']', ':' })
                    safeName = safeName.Replace(c.ToString(), "-");
                safeName = safeName.Trim();
                if (string.IsNullOrWhiteSpace(safeName)) safeName = "Feuille" + (i + 1);
                if (safeName.Length > 31) safeName = safeName.Substring(0, 31).Trim();

                string finalName = safeName;
                int dupIdx = 2;
                while (usedNames.Contains(finalName))
                {
                    string suffix = " (" + dupIdx + ")";
                    int maxBase = 31 - suffix.Length;
                    string basePart = safeName.Length > maxBase ? safeName.Substring(0, maxBase).Trim() : safeName;
                    finalName = basePart + suffix;
                    dupIdx++;
                }
                usedNames.Add(finalName);

                sb.AppendFormat("<sheet name=\"{0}\" sheetId=\"{1}\" r:id=\"rId{1}\"/>", EscapeXml(finalName), i + 1);
            }
            sb.Append("</sheets></workbook>");
            return sb.ToString();
        }

        private static string BuildWorkbookRels(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.AppendFormat("<Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{0}.xml\"/>", i);
            }
            sb.AppendFormat("<Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>", sheetCount + 1);
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string BuildStylesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<numFmts count=\"1\">" +
                "<numFmt numFmtId=\"164\" formatCode=\"0.0\"/>" +
            "</numFmts>" +
            "<fonts count=\"7\">" +
                "<font><sz val=\"10\"/><color rgb=\"FF1E293B\"/><name val=\"Segoe UI\"/></font>" + // 0: Normal
                "<font><b/><sz val=\"10\"/><color rgb=\"FFFFFFFF\"/><name val=\"Segoe UI\"/></font>" + // 1: White bold
                "<font><b/><sz val=\"10\"/><color rgb=\"FF1E293B\"/><name val=\"Segoe UI\"/></font>" + // 2: Dark bold
                "<font><b/><sz val=\"10\"/><color rgb=\"FFB91C1C\"/><name val=\"Segoe UI\"/></font>" + // 3: Red bold (<4.0)
                "<font><i/><sz val=\"9.5\"/><color rgb=\"FF64748B\"/><name val=\"Segoe UI\"/></font>" + // 4: Italic gray ((5.0)/C.M)
                "<font><b/><sz val=\"12\"/><color rgb=\"FF991B1B\"/><name val=\"Segoe UI\"/></font>" + // 5: Section title red
                "<font><b/><sz val=\"10\"/><color rgb=\"FF0F172A\"/><name val=\"Segoe UI\"/></font>" + // 6: Principal discipline bold
            "</fonts>" +
            "<fills count=\"7\">" +
                "<fill><patternFill patternType=\"none\"/></fill>" + // 0
                "<fill><patternFill patternType=\"gray125\"/></fill>" + // 1
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF991B1B\"/></patternFill></fill>" + // 2: Classeur Rouge header
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE2E8F0\"/></patternFill></fill>" + // 3: Subheader light slate
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFEE2E2\"/></patternFill></fill>" + // 4: Light red (<4.0)
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF8FAFC\"/></patternFill></fill>" + // 5: Soft zebra/moyenne bg
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFECACA\"/></patternFill></fill>" + // 6: Stronger red for insufficient moyenne
            "</fills>" +
            "<borders count=\"2\">" +
                "<border><left/><right/><top/><bottom/><diagonal/></border>" + // 0
                "<border>" + // 1: Thin gray border
                    "<left style=\"thin\"><color rgb=\"FFCBD5E1\"/></left>" +
                    "<right style=\"thin\"><color rgb=\"FFCBD5E1\"/></right>" +
                    "<top style=\"thin\"><color rgb=\"FFCBD5E1\"/></top>" +
                    "<bottom style=\"thin\"><color rgb=\"FFCBD5E1\"/></bottom>" +
                    "<diagonal/>" +
                "</border>" +
            "</borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"10\">" +
                // 0: Normal text left bordered
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"left\" wrapText=\"1\"/></xf>" +
                // 1: Header dark red bold white centered
                "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\" wrapText=\"1\"/></xf>" +
                // 2: Subheader slate bold centered
                "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"3\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\" wrapText=\"1\"/></xf>" +
                // 3: Grade normal centered (>= 4.0)
                "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\"/></xf>" +
                // 4: Grade insufficient (< 4.0) light red
                "<xf numFmtId=\"164\" fontId=\"3\" fillId=\"4\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\"/></xf>" +
                // 5: Grade indicative (5.0) or C.M / Disp italic gray
                "<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\"/></xf>" +
                // 6: Moyenne bold centered (>= 4.0)
                "<xf numFmtId=\"164\" fontId=\"2\" fillId=\"5\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\"/></xf>" +
                // 7: Moyenne insufficient (< 4.0) bold red
                "<xf numFmtId=\"164\" fontId=\"3\" fillId=\"6\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"center\"/></xf>" +
                // 8: Section title bold red left no border
                "<xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"left\"/></xf>" +
                // 9: Principal discipline name bold left bordered
                "<xf numFmtId=\"0\" fontId=\"6\" fillId=\"5\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" horizontal=\"left\" wrapText=\"1\"/></xf>" +
            "</cellXfs>" +
            "</styleSheet>";
        }

        private static string BuildSheetXml(ExcelSheet sheet)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            if (sheet.ColumnWidths.Count > 0)
            {
                sb.Append("<cols>");
                for (int c = 0; c < sheet.ColumnWidths.Count; c++)
                {
                    sb.AppendFormat(CultureInfo.InvariantCulture, "<col min=\"{0}\" max=\"{0}\" width=\"{1:0.0}\" customWidth=\"1\"/>", c + 1, sheet.ColumnWidths[c]);
                }
                sb.Append("</cols>");
            }

            sb.Append("<sheetData>");
            for (int r = 0; r < sheet.Rows.Count; r++)
            {
                ExcelRow row = sheet.Rows[r];
                int rowNum = r + 1;
                sb.AppendFormat(CultureInfo.InvariantCulture, "<row r=\"{0}\" ht=\"{1:0.0}\" customHeight=\"1\">", rowNum, row.Height);
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    ExcelCell cell = row.Cells[c];
                    if (cell == null) continue;
                    string cellRef = GetColName(c) + rowNum.ToString();
                    string val = cell.Value ?? "";

                    if (cell.IsNumeric && !string.IsNullOrWhiteSpace(val))
                    {
                        double num;
                        if (double.TryParse(val.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out num))
                        {
                            sb.AppendFormat(CultureInfo.InvariantCulture, "<c r=\"{0}\" s=\"{1}\"><v>{2}</v></c>", cellRef, cell.StyleIndex, num);
                            continue;
                        }
                    }
                    if ((cell.StyleIndex == 0 || cell.StyleIndex == 9) && val.Contains("\n"))
                    {
                        int nlIdx = val.IndexOf('\n');
                        string line1 = val.Substring(0, nlIdx).TrimEnd('\r');
                        string line2 = val.Substring(nlIdx + 1).TrimStart('\r');
                        string boldTag = (cell.StyleIndex == 9) ? "<b/>" : "";
                        string color1 = (cell.StyleIndex == 9) ? "FF0F172A" : "FF1E293B";
                        sb.AppendFormat(
                            "<c r=\"{0}\" s=\"{1}\" t=\"inlineStr\"><is>" +
                            "<r><rPr>{2}<sz val=\"10\"/><color rgb=\"{3}\"/><rFont val=\"Segoe UI\"/></rPr><t>{4}</t></r>" +
                            "<r><rPr><i/><sz val=\"8.5\"/><color rgb=\"FF64748B\"/><rFont val=\"Segoe UI\"/></rPr><t xml:space=\"preserve\">&#10;{5}</t></r>" +
                            "</is></c>",
                            cellRef, cell.StyleIndex, boldTag, color1, EscapeXml(line1), EscapeXml(line2)
                        );
                        continue;
                    }
                    sb.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"inlineStr\"><is><t>{2}</t></is></c>", cellRef, cell.StyleIndex, EscapeXml(val));
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");

            if (sheet.MergedCells.Count > 0)
            {
                sb.AppendFormat("<mergeCells count=\"{0}\">", sheet.MergedCells.Count);
                foreach (string mc in sheet.MergedCells)
                {
                    sb.AppendFormat("<mergeCell ref=\"{0}\"/>", mc);
                }
                sb.Append("</mergeCells>");
            }

            sb.Append("</worksheet>");
            return sb.ToString();
        }
    }
}
