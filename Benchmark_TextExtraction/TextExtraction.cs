using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using RichEditShapeCollection = DevExpress.XtraRichEdit.API.Native.ShapeCollection;
using RichEditCommentCollection = DevExpress.XtraRichEdit.API.Native.CommentCollection;
using DevExpress.Spreadsheet;
using SpreadsheetShapeType = DevExpress.Spreadsheet.ShapeType;
using DevExpress.Pdf;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;

namespace Benchmark_TextExtraction {
    [MemoryDiagnoser]
    //[SimpleJob(RuntimeMoniker.Mono)]
    //[SimpleJob(RuntimeMoniker.CoreRt31)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    public class TextExtraction_GcWorkstation_RichEdit {
        RichEditDocumentServer richEditDocumentServer;
        string[] richEditFilePaths;
        public TextExtraction_GcWorkstation_RichEdit() {
            richEditDocumentServer = new RichEditDocumentServer();
            string richEditFiles = "path/to/file";
            richEditFilePaths = Directory.GetFiles(richEditFiles, "*.*", SearchOption.AllDirectories);
        }

        [Benchmark]
        public void RichEdit() {
            StringBuilder text;
            foreach (var item in richEditFilePaths) {
                text = GetText(richEditDocumentServer, item);
            }
        }
        #region RichEditTextExtractionMethods
        StringBuilder GetText(RichEditDocumentServer server, string filePath) {
            richEditDocumentServer.LoadDocument(filePath);
            StringBuilder builder = new StringBuilder();
            builder.Append(server.Document.GetText(server.Document.Range));

            NoteCollection footnotes = server.Document.Footnotes;
            foreach (var footnote in footnotes) {
                var footnoteContent = footnote.BeginUpdate();
                builder.Append(footnoteContent.GetText(footnote.Range));
                footnoteContent.EndUpdate();
            }

            NoteCollection endnotes = server.Document.Footnotes;
            foreach (var endnote in endnotes) {
                var endnoteContent = endnote.BeginUpdate();
                builder.Append(endnoteContent.GetText(endnote.Range));
                endnoteContent.EndUpdate();
            }

            RichEditCommentCollection comments = server.Document.Comments;
            foreach (var comment in comments) {
                var commentContent = comment.BeginUpdate();
                builder.Append(commentContent.GetText(comment.Range));
                commentContent.EndUpdate();
            }

            RichEditShapeCollection shapes = server.Document.Shapes;
            foreach (var shape in shapes) {
                if (shape.ShapeFormat != null && shape.ShapeFormat.HasText) {
                    var document = shape.ShapeFormat.TextBox.Document;
                    builder.Append(document.GetText(document.Range));
                }
            }

            SectionCollection sections = server.Document.Sections;
            foreach (var section in sections) {
                builder.Append(GetHeaderText(section, HeaderFooterType.Even));
                builder.Append(GetHeaderText(section, HeaderFooterType.First));
                builder.Append(GetHeaderText(section, HeaderFooterType.Odd));
                builder.Append(GetFooterText(section, HeaderFooterType.Even));
                builder.Append(GetFooterText(section, HeaderFooterType.First));
                builder.Append(GetFooterText(section, HeaderFooterType.Odd));
            }
            return builder;
        }

        string GetHeaderText(Section section, HeaderFooterType headerType) {
            string text = string.Empty;
            if (section.HasHeader(headerType)) {
                var header = section.BeginUpdateHeader(headerType);
                text = header.GetText(header.Range);
                header.EndUpdate();
            }
            return text;
        }

        string GetFooterText(Section section, HeaderFooterType footerType) {
            string text = string.Empty;
            if (section.HasFooter(footerType)) {
                var header = section.BeginUpdateFooter(footerType);
                text = header.GetText(header.Range);
                header.EndUpdate();
            }
            return text;
        }
        #endregion
    }
    [MemoryDiagnoser]
    //[SimpleJob(RuntimeMoniker.Mono)]
    //[SimpleJob(RuntimeMoniker.CoreRt31)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    public class TextExtraction_GcWorkstation_Spreadsheet {
        Workbook workbook;
        string[] spreadsheetFilePaths;
        public TextExtraction_GcWorkstation_Spreadsheet() {
            workbook = new Workbook();
            string spreadsheetFiles = "path/to/file";
            spreadsheetFilePaths = Directory.GetFiles(spreadsheetFiles, "*.*", SearchOption.AllDirectories);
        }

        [Benchmark]
        public void Spreadsheet() {
            StringBuilder text;
            foreach (var item in spreadsheetFilePaths) {
                text = GetText(workbook, item);
            }
        }
        #region SpreadsheetTextExtractionMethods
        static IEnumerable<string> GetCellTextOnly(Workbook workbook) =>
workbook.Worksheets.SelectMany(x => x.GetExistingCells()
.Where(c => c.Value.IsText)
.Select(c => c.Value.TextValue));

        static IEnumerable<string> GetCellDisplayText(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.GetExistingCells().Select(c => c.DisplayText));

        static IEnumerable<string> GetShapeText(Workbook workbook) =>
             workbook.Worksheets.SelectMany(x => x.Shapes
                .Flatten()
                .Where(s => s.ShapeType == SpreadsheetShapeType.Shape && s.ShapeText.HasText)
                .Select(s => s.ShapeText.Characters().Text));

        static IEnumerable<string> GetChartTitles(Workbook workbook) =>
            workbook.Worksheets.SelectMany(x => x.Charts.Select(c => c.Title.PlainText));

        StringBuilder GetText(Workbook workbook, string filePath) {
            StringBuilder builder = new StringBuilder();
            workbook.LoadDocument(filePath);
            var query = GetCellDisplayText(workbook)
                .Concat(GetChartTitles(workbook))
                .Concat(GetShapeText(workbook));
            foreach (string str in query)
                builder.Append(str);
            return builder;
        }
        #endregion
    }

    [MemoryDiagnoser]
    //[SimpleJob(RuntimeMoniker.Mono)]
    //[SimpleJob(RuntimeMoniker.CoreRt31)]
    [SimpleJob(RuntimeMoniker.NetCoreApp31)]
    public class TextExtraction_GcWorkstation_Pdf {
        PdfDocumentProcessor pdfDocumentProcessor;
        string[] pdfFilePaths;
        public TextExtraction_GcWorkstation_Pdf() {
            pdfDocumentProcessor = new PdfDocumentProcessor();
            string pdfFiles = "path/to/file";
            pdfFilePaths = Directory.GetFiles(pdfFiles, "*.*", SearchOption.AllDirectories);
        }

        [Benchmark]
        public void Pdf() {
            string text;
            foreach (var item in pdfFilePaths) {
                pdfDocumentProcessor.LoadDocument(item);
                text = pdfDocumentProcessor.GetText();
            }
        }
    }
}
