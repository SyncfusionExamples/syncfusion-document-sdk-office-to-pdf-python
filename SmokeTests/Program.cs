using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using EShare.Documents;

var key = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
if (!string.IsNullOrWhiteSpace(key)) DocumentService.RegisterLicense(key);

string root = Path.GetFullPath(args.Length > 0 ? args[0] : "test-output");
Directory.CreateDirectory(root);
string source = Path.Combine(root, "input.docx");
string output = Path.Combine(root, "marked.docx");
using (var document = new WordDocument())
{
    var section = document.AddSection();
    section.AddParagraph().AppendText("Original body must survive.");
    section.HeadersFooters.Header.AddParagraph().AppendText("Existing header");
    section.PageSetup.DifferentFirstPage = true;
    section.PageSetup.DifferentOddAndEvenPages = true;
    section.HeadersFooters.FirstPageHeader.AddParagraph().AppendText("First page");
    var next = document.AddSection();
    next.HeadersFooters.LinkToPrevious = true;
    next.AddParagraph().AppendText("Second section");
    document.Save(source, FormatType.Docx);
}
DocumentService.MarkWord(source, output, "CONFIDENTIAL");
using (var document = new WordDocument(output, FormatType.Docx))
{
    Check(document.GetText().Contains("Original body must survive."), "Word body preserved");
    Check(document.Sections[0].HeadersFooters.Header.Paragraphs.Cast<WParagraph>().Select(p => p.Text).Aggregate("", (a, b) => a + b).Contains("Existing header"), "Existing header preserved");
    Check(document.Sections[0].HeadersFooters.FirstPageHeader.Paragraphs.Cast<WParagraph>().Select(p => p.Text).Aggregate("", (a, b) => a + b).Contains("CONFIDENTIAL"), "First page marked");
    Check(document.Sections[0].HeadersFooters.EvenFooter.Paragraphs.Cast<WParagraph>().Select(p => p.Text).Aggregate("", (a, b) => a + b).Contains("CONFIDENTIAL"), "Even footer marked");

}
using (var zip = System.IO.Compression.ZipFile.OpenRead(output))
{
    using var xml = zip.GetEntry("word/document.xml")!.Open();
    var docXml = System.Xml.Linq.XDocument.Load(xml);
    System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    Check(!docXml.Descendants(w + "sectPr").Last().Elements(w + "headerReference").Any(), "Header inheritance preserved in DOCX");
}
string xlsx = Path.Combine(root, "input.xlsx");
using (var engine = new Syncfusion.XlsIO.ExcelEngine())
{
    engine.Excel.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
    var book = engine.Excel.Workbooks.Create(2);
    book.Worksheets[0].Range["A1"].Text = "Original cell";
    book.Worksheets[0].PageSetup.LeftHeader = "Existing left header";
    book.Worksheets[0].PageSetup.DifferentFirstPageHF = true;
    book.Worksheets[0].PageSetup.DifferentOddAndEvenPagesHF = true;
    book.Worksheets[1].Range["A1"].Text = "Second sheet";
    book.SaveAs(xlsx);
    book.Close();
}
string markedXlsx = Path.Combine(root, "marked.xlsx");
DocumentService.MarkExcel(xlsx, markedXlsx, "R&D CONFIDENTIAL");
using (var engine = new Syncfusion.XlsIO.ExcelEngine())
{
    using var stream = File.OpenRead(markedXlsx);
    var book = engine.Excel.Workbooks.Open(stream);
    Check(book.Worksheets[0].Range["A1"].Text == "Original cell", "Excel cell preserved");
    Check(book.Worksheets[0].PageSetup.LeftHeader == "Existing left header", "Excel left header preserved");
    foreach (Syncfusion.XlsIO.IWorksheet sheet in book.Worksheets)
    {
        Check(sheet.PageSetup.CenterHeader == "R&&D CONFIDENTIAL", "Excel label escapes ampersand");
        Check(sheet.PageSetup.FirstPage.CenterFooter == "R&&D CONFIDENTIAL", "Excel first footer marked");
        Check(sheet.PageSetup.EvenPage.CenterHeader == "R&&D CONFIDENTIAL", "Excel even header marked");
    }
    book.Close();
}
string pptx = Path.Combine(root, "input.pptx");
using (var deck = Syncfusion.Presentation.Presentation.Create())
{
    var slide = deck.Slides.Add(Syncfusion.Presentation.SlideLayoutType.Blank);
    slide.Shapes.AddTextBox(40, 100, 400, 100).TextBody.AddParagraph("Original slide text");
    deck.Slides.Add(Syncfusion.Presentation.SlideLayoutType.Blank);
    deck.Save(pptx);
}
string markedPptx = Path.Combine(root, "marked.pptx");
DocumentService.MarkPowerPoint(pptx, markedPptx, "CONFIDENTIAL");
string reclassified = Path.Combine(root, "reclassified.pptx");
DocumentService.MarkPowerPoint(markedPptx, reclassified, "INTERNAL");
using (var deck = Syncfusion.Presentation.Presentation.Open(reclassified))
{
    Check(deck.Slides[0].Shapes.Count == 3, "PowerPoint retains original and exactly two markings");
    Check(deck.Slides[0].Shapes.Cast<Syncfusion.Presentation.IShape>().Any(s => s.TextBody.Text.Contains("Original slide text")), "PowerPoint content preserved");
    foreach (var slide in deck.Slides)
        Check(slide.Shapes.Cast<Syncfusion.Presentation.IShape>().Count(s => s.ShapeName.StartsWith("eSHARE.Classification.") && s.TextBody.Text.Contains("INTERNAL")) == 2, "PowerPoint labels updated without duplication");
}
DocumentService.WordToPdf(output, Path.Combine(root, "word.pdf"));
DocumentService.ExcelToPdf(markedXlsx, Path.Combine(root, "excel.pdf"));
DocumentService.PowerPointToPdf(markedPptx, Path.Combine(root, "powerpoint.pdf"));
foreach (var name in new[] { "word", "excel", "powerpoint" })
{
    using var pdf = new Syncfusion.Pdf.Parsing.PdfLoadedDocument(Path.Combine(root, name + ".pdf"));
    Check(pdf.Pages.Count > 0, name + " conversion creates PDF pages");
}
DocumentService.WatermarkPdf(Path.Combine(root, "powerpoint.pdf"), Path.Combine(root, "watermarked.pdf"), "RESTRICTED");
using (var pdf = new Syncfusion.Pdf.Parsing.PdfLoadedDocument(Path.Combine(root, "watermarked.pdf")))
{
    Check(pdf.Pages.Count == 2, "PDF watermark preserves page count");
    foreach (Syncfusion.Pdf.PdfLoadedPage page in pdf.Pages)
        Check(page.ExtractText().Contains("RESTRICTED"), "PDF watermark appears on each page");
}
static void Check(bool condition, string name)
{
    if (!condition) throw new Exception("FAIL: " + name);
    Console.WriteLine("PASS: " + name);
}
