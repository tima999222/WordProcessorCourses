#region

using DocumentFormat.OpenXml.Wordprocessing;

#endregion

namespace ASTepanov.Docx
{
    public static class xWordProcessing
    {
        public static IEnumerable<TableRow> Rows(
            this Table table)
        {
            foreach (var tline in table)
                if (tline is TableRow)
                    yield return tline as TableRow;
        }
    }
}