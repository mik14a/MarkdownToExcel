using ConsoleAppFramework;

namespace MarkdownToExcel;

class Program
{
    static void Main(string[] args) {
        var app = ConsoleApp.Create(args);
        app.AddCommands<MarkdownToExcel>();
        app.Run();
    }
}
