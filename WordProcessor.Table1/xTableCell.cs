#region

using DocumentFormat.OpenXml.Wordprocessing;

#endregion

namespace UNTI.RegistryOfAccepted
{
    public static class xTableCell
    {
        public static void SetText(this TableCell cell, string text)
        {
            var p = cell.Elements<Paragraph>().First();

            // Find the first run in the paragraph.
            var r = p.Elements<Run>().First();

            // Set the text for the run.
            var t = r.Elements<Text>().First();
            t.Text = text;
            var texts = r.Elements<Text>().Skip(1);
            foreach (var txt in texts) txt.Remove();
        }
    }
}