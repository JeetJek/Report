using System;
using Word = Microsoft.Office.Interop.Word;


namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.Application app = new Word.Application();
            Object fileName = @"D:\тест.docx";
            Object missing = Type.Missing;
            app.Documents.Open(ref fileName);
            Word.Find find = app.Selection.Find;
            find.Text = "2";
            find.Replacement.Text = "два";
            Object wrap = Word.WdFindWrap.wdFindContinue;
            Object replace = Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: missing, Replace: replace);
            app.ActiveDocument.Save();
            app.ActiveDocument.Close();
            app.Quit();
        }
    }
}
