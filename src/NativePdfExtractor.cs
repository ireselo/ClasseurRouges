using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;

namespace ClasseursRouges
{
    public class PdfWord
    {
        public string Text;
        public double X;
        public double Y;
    }

    public class NativePdfTextExtractor
    {
        // Extracts positioned text tokens from digital PDF pages using FlateDecode content streams
        public static List<List<PdfWord>> ExtractWordsPerPage(string pdfPath)
        {
            List<List<PdfWord>> pages = new List<List<PdfWord>>();
            try
            {
                byte[] bytes = File.ReadAllBytes(pdfPath);
                string raw = Encoding.GetEncoding("iso-8859-1").GetString(bytes);

                // Find all stream ... endstream blocks
                Regex streamRx = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled);
                MatchCollection matches = streamRx.Matches(raw);

                foreach (Match m in matches)
                {
                    int startIdx = m.Groups[1].Index;
                    int len = m.Groups[1].Length;
                    if (len <= 2) continue;

                    byte[] streamBytes = new byte[len];
                    Array.Copy(bytes, startIdx, streamBytes, 0, len);

                    string content = TryDecompressFlate(streamBytes);
                    if (string.IsNullOrEmpty(content))
                    {
                        // Check if uncompressed text stream
                        if (m.Groups[1].Value.Contains("BT") && m.Groups[1].Value.Contains("ET"))
                            content = m.Groups[1].Value;
                    }

                    if (!string.IsNullOrEmpty(content) && content.Contains("BT") && content.Contains("ET"))
                    {
                        List<PdfWord> words = ParseContentStream(content);
                        if (words.Count > 5)
                        {
                            pages.Add(words);
                        }
                    }
                }
            }
            catch
            {
                // Fallback to WinRT rendering + OCR
            }
            return pages;
        }

        private static string TryDecompressFlate(byte[] data)
        {
            try
            {
                // Skip 2-byte zlib header (0x78 0x9C / 0x01 / 0xDA)
                int offset = (data.Length > 2 && data[0] == 0x78) ? 2 : 0;
                using (MemoryStream ms = new MemoryStream(data, offset, data.Length - offset))
                using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                using (StreamReader sr = new StreamReader(ds, Encoding.GetEncoding("iso-8859-1")))
                {
                    return sr.ReadToEnd();
                }
            }
            catch
            {
                return null;
            }
        }

        private static List<PdfWord> ParseContentStream(string content)
        {
            List<PdfWord> words = new List<PdfWord>();
            Regex btEt = new Regex(@"BT([\s\S]*?)ET", RegexOptions.Compiled);
            Regex tmRx = new Regex(@"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Tm", RegexOptions.Compiled);
            Regex tdRx = new Regex(@"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+T[dD]", RegexOptions.Compiled);
            Regex tjStrRx = new Regex(@"\((?:\\.|[^\\)])*\)\s*Tj|\[(?:[^\]]*)\]\s*TJ", RegexOptions.Compiled);
            Regex parenRx = new Regex(@"\((?:\\.|[^\\)])*\)", RegexOptions.Compiled);

            foreach (Match block in btEt.Matches(content))
            {
                string b = block.Groups[1].Value;
                double curX = 0, curY = 0;
                string[] lines = b.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    Match tm = tmRx.Match(line);
                    if (tm.Success)
                    {
                        curX = double.Parse(tm.Groups[5].Value, CultureInfo.InvariantCulture);
                        curY = double.Parse(tm.Groups[6].Value, CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        Match td = tdRx.Match(line);
                        if (td.Success)
                        {
                            curX += double.Parse(td.Groups[1].Value, CultureInfo.InvariantCulture);
                            curY += double.Parse(td.Groups[2].Value, CultureInfo.InvariantCulture);
                        }
                    }

                    foreach (Match tj in tjStrRx.Matches(line))
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (Match p in parenRx.Matches(tj.Value))
                        {
                            string inner = p.Value.Substring(1, p.Value.Length - 2);
                            inner = DecodePdfString(inner);
                            sb.Append(inner);
                        }
                        string txt = sb.ToString().Trim();
                        if (!string.IsNullOrEmpty(txt))
                        {
                            // Convert PDF bottom-left Y (0..842) to top-down 2000px scale (0..2806)
                            double scaledX = curX * (1984.0 / 595.0);
                            double scaledY = (842.0 - curY) * (2806.0 / 842.0);
                            words.Add(new PdfWord { Text = txt, X = scaledX, Y = scaledY });
                        }
                    }
                }
            }
            return words;
        }

        private static string DecodePdfString(string s)
        {
            return s.Replace(@"\(", "(").Replace(@"\)", ")").Replace(@"\\", @"\");
        }
    }
}
