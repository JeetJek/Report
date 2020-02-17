using System;
using Word = Microsoft.Office.Interop.Word;


namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
                return;

            Word.Application app = new Word.Application();
            Object fileName = args[0];
            Object missing = Type.Missing;
            app.Documents.Open(ref fileName);
            Word.Find find = app.Selection.Find;
            find.Text = "#Date#";
            find.Replacement.Text = DateTime.Today.ToString();
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
