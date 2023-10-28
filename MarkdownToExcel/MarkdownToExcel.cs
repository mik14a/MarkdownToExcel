using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

using ConsoleAppFramework;

using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

using Excel = Microsoft.Office.Interop.Excel;

namespace MarkdownToExcel;

public class MarkdownToExcel : ConsoleAppBase
{
    [Command("convert", "Converts a markdown file to an excel file.")]
    public async Task ConvertAsync([Option("s", "source markdown file path.")] string sourcePath,
                                   [Option("d", "destination excel file path.")] string destinationPath) {
        Excel.Application? excel = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;

        try {
            excel = DynamicExcel.CreateApplication();
            workbooks = excel.Workbooks;
            workbook = workbooks.Add();

            var text = await File.ReadAllTextAsync(sourcePath);
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            var markdown = Markdown.Parse(text, pipeline);
            TraverseDocument(markdown, workbook);

            workbook.SaveAs(destinationPath);
        } finally {
            workbook?.Close();
            if (workbook is not null) Marshal.ReleaseComObject(workbook);
            workbooks?.Close();
            if (workbooks is not null) Marshal.ReleaseComObject(workbooks);
            excel?.Quit();
            if (excel is not null) Marshal.ReleaseComObject(excel);
        }
    }

    void TraverseDocument(MarkdownDocument markdown, Excel.Workbook workbook) {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        try {
            sheets = workbook.Sheets;
            sheet = sheets[1];
            var row = 1; var col = 1;
            foreach (var block in markdown) {
                TraverseBlock(block, sheet, ref row, ref col);
            }
        } finally {
            if (sheet is not null) Marshal.ReleaseComObject(sheet);
            if (sheets is not null) Marshal.ReleaseComObject(sheets);
        }
    }

    void TraverseBlock(Block block, Excel.Worksheet sheet, ref int row, ref int col) {
        switch (block) {
        case HeadingBlock heading:
            var headingText = string.Join("", heading.Inline.Select(x => x.ToString()));
            sheet.Cells[row++, col] = headingText;
            break;
        default:
            System.Diagnostics.Debug.WriteLine($"Not support: {block}");
            break;
        }
    }

    void TraverseInline(ContainerInline containerInline, Excel.Worksheet sheet, ref int row, ref int col) {
        var inline = containerInline.FirstChild;
        while (inline is not null) {
            switch (inline) {
            default:
                System.Diagnostics.Debug.WriteLine($"Not support: {inline}");
                break;
            }
            inline = inline.NextSibling;
        }
    }
}
