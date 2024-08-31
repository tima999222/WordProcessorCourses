#region

using AStepanov.Core.Ex;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordProcessor.Table1.Entities;

#endregion

namespace ASTepanov.Docx
{
    public class DocumentProcessor : IDisposable
    {
        private readonly MemoryStream memoryStream; //File.Open(filepath, FileMode.Open);

        // Open a WordProcessingDocument based on a stream.
        private readonly WordprocessingDocument wordprocessingDocument;

        public DocumentProcessor(byte[] data)
        {
            memoryStream = new MemoryStream(data, true);

            wordprocessingDocument =
                WordprocessingDocument.Open(memoryStream, true);
        }

        public string FileName { get; }

        #region IDisposable Members

        public void Dispose()
        {
            wordprocessingDocument.Close();
            memoryStream.Close();
        }

        #endregion

        public void Map<T>(T value)
        {
            var props = value.GetProperties();
            var values = new Dictionary<string, string>();

            foreach (var keyValuePair in props)
            {
                var propKey = $"{value.GetType().Name}.{keyValuePair.Key}"; //ключ для замены
                var strValue = string.Empty;


                if (keyValuePair.Value != null) strValue = keyValuePair.Value?.ToString();

                values.Add(propKey, strValue); //данные для замены
            }

            ProcessText(p =>
            {
                foreach (var keyValuePair in values)
                {
                    if (!p.InnerText.Contains(keyValuePair.Key)) continue;

                    p.Text = p.Text.Replace(keyValuePair.Key, keyValuePair.Value);
                }
            });
        }

        public void MapErrorCounters(string name, int? count)
        {
            if (count == null)
            {
                count = 0;
            }

            var propKey = $"{name}"; //ключ для замены

            ProcessText(p =>
            {
                if (p.InnerText.Contains(propKey))
                    p.Text = p.Text.Replace(propKey, count.ToString());
            });
        }
        
        public void MapErrorNames(string name, string? errName)
        {
            if (errName == null)
            {
                errName = "";
            }

            var propKey = $"{name}"; //ключ для замены

            ProcessText(p =>
            {
                if (p.InnerText.Contains(propKey))
                    p.Text = p.Text.Replace(propKey, errName);
            });
        }

        /// <summary>
        ///     Обойти все дерево документа с заданным действием к нему
        /// </summary>
        /// <param name="action"></param>
        public void ProcessEach(Action<OpenXmlElement> action)
        {
            foreach (var node in wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements)
                ProcessNode(action, node);
        }

        /// <summary>
        ///     Обработать все текстовые ноды
        /// </summary>
        /// <param name="textAction"></param>
        public void ProcessText(Action<Text> textAction)
        {
            ProcessEach(action =>
            {
                foreach (var actionChildElement in action.ChildElements)
                {
                    if (!(actionChildElement is Text)) continue;

                    var text = (Text)actionChildElement;
                    textAction(text);
                }
            });
        }

        private static void ProcessNode(Action<OpenXmlElement> action, OpenXmlElement node)
        {
            ProcessNode<OpenXmlElement>(action, node);
        }

        private static void ProcessNode<T>(Action<OpenXmlElement> action, OpenXmlElement node)
        {
            action(node);

            foreach (var nodeChildElement in node.ChildElements) ProcessNode<T>(action, nodeChildElement);
        }

        public void Save()
        {
            wordprocessingDocument.Save();
        }


        public void SaveAs(string path)
        {
            wordprocessingDocument.SaveAs(path);
        }


        public byte[] Bytes()
        {
            using (var memoryStream = new MemoryStream())
            {
                wordprocessingDocument.Clone(memoryStream);
                return memoryStream.ToArray();
            }

            //throw new NotImplementedException();
        }

        public void MapItems<T>(IEnumerable<T> items, int rowSkips)
        {
            if (!items.Any()) return;

            var itemType = items.First().GetType();
            var itemTypeBeginMappingKey = "{" + itemType.Name + ".";

            ProcessEach(e =>
            {
                if (!(e is Table table)) return;

                var secondRow = table.Elements<TableRow>().Skip(rowSkips).FirstOrDefault();
                if (secondRow == null || !secondRow.InnerText.Contains(itemTypeBeginMappingKey)) return;

                var templateRow = secondRow.CloneNode(true) as TableRow;
                secondRow.Remove();

                foreach (var item in items)
                {
                    var newRow = (TableRow)templateRow.CloneNode(true);
                    var propertyValues = itemType.GetProperties()
                        .ToDictionary(
                            prop => itemTypeBeginMappingKey + prop.Name + "}",
                            prop => prop.GetValue(item)?.ToString() ?? string.Empty
                        );

                    foreach (var cell in newRow.Elements<TableCell>())
                    {
                        var cellText = cell.InnerText.Trim();
                        if (propertyValues.TryGetValue(cellText, out var value))
                        {
                            cell.RemoveAllChildren<Paragraph>();
                            var newParagraph = new Paragraph();

                            ParagraphProperties paraProperties = new ParagraphProperties();
                            paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                            newParagraph.PrependChild(paraProperties);

                            var newRun = new Run();

                            RunProperties runProperties = new RunProperties();
                            runProperties.Append(new FontSize() { Val = "16" });
                            newRun.Append(runProperties);

                            newRun.Append(new Text(value));
                            newParagraph.Append(newRun);
                            cell.Append(newParagraph);
                        }
                    }

                    table.Append(newRow);
                }
            });
        }

        public void MapItemsOther<T>(IEnumerable<T> items, int rowSkips)
        {
            if (!items.Any()) return;

            var itemType = items.First().GetType();
            var itemTypeBeginMappingKey = "{" + itemType.Name + ".";

            ProcessEach(e =>
            {
                if (!(e is Table table)) return;

                var secondRow = table.Elements<TableRow>().Skip(rowSkips).FirstOrDefault();
                if (secondRow == null || !secondRow.InnerText.Contains(itemTypeBeginMappingKey)) return;

                var templateRow = secondRow.CloneNode(true) as TableRow;
                secondRow.Remove();

                foreach (var item in items)
                {
                    var partCount = 0;
                    var propertyValues = itemType.GetProperties()
                        .Where(prop => prop.Name != "Participants")
                        .ToDictionary(
                            prop => itemTypeBeginMappingKey + prop.Name + "}",
                            prop => prop.GetValue(item)?.ToString() ?? string.Empty
                        );

                    // Создание строки для каждого участника
                    if (item is Startup startup)
                    {
                        if (startup.Participants != null)
                            partCount = startup.Participants.Count();
                        else
                        {
                            partCount = 0;
                        }

                        if (partCount != 0)
                        {
                            foreach (var participant in startup.Participants)
                            {
                                var newRow = (TableRow)templateRow.CloneNode(true);

                                // Обработка обычных столбцов
                                for (int i = 0; i < 5; i++)
                                {
                                    var cell = newRow.Elements<TableCell>().ElementAt(i);
                                    var cellText = cell.InnerText.Trim();
                                    if (propertyValues.TryGetValue(cellText, out var value))
                                    {
                                        Paragraph para = cell.Elements<Paragraph>().First();

                                        ParagraphProperties paraProperties = new ParagraphProperties();
                                        paraProperties.Justification = new Justification()
                                            { Val = JustificationValues.Center };
                                        para.PrependChild(paraProperties);

                                        RunProperties runProperties = new RunProperties();
                                        runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 8pt
                                        Run run = new Run();
                                        run.Append(runProperties);
                                        Text text = new Text(value);
                                        run.Append(text);
                                        para.RemoveAllChildren<Run>();
                                        para.Append(run);
                                    }
                                }

                                // Обработка столбцов с участниками
                                for (int i = 5; i < 8; i++)
                                {
                                    var cell = newRow.Elements<TableCell>().ElementAt(i);
                                    var cellText = cell.InnerText.Trim();

                                    var cleanCellText = cellText.Replace(itemTypeBeginMappingKey, "")
                                        .Replace("Participants.", "").Replace("{", "").Replace("}", "");
                                    var participantProperty = typeof(Participant).GetProperty(cleanCellText);
                                    if (participantProperty != null)
                                    {
                                        var participantValue = participantProperty.GetValue(participant)?.ToString() ??
                                                               string.Empty;

                                        Paragraph para = cell.Elements<Paragraph>().First();
                                        ParagraphProperties paraProperties = new ParagraphProperties();
                                        paraProperties.Justification = new Justification()
                                            { Val = JustificationValues.Center };
                                        para.PrependChild(paraProperties);

                                        RunProperties runProperties = new RunProperties();
                                        runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 8pt
                                        Run run = new Run();
                                        run.Append(runProperties);
                                        Text text = new Text(participantValue);
                                        run.Append(text);
                                        para.RemoveAllChildren<Run>();
                                        para.Append(run);
                                    }
                                }

                                table.Append(newRow);
                            }
                        }
                        else
                        {
                            var newRow = (TableRow)templateRow.CloneNode(true);

                            // Обработка обычных столбцов
                            for (int i = 0; i < 5; i++)
                            {
                                var cell = newRow.Elements<TableCell>().ElementAt(i);
                                var cellText = cell.InnerText.Trim();
                                if (propertyValues.TryGetValue(cellText, out var value))
                                {
                                    Paragraph para = cell.Elements<Paragraph>().First();

                                    ParagraphProperties paraProperties = new ParagraphProperties();
                                    paraProperties.Justification = new Justification()
                                        { Val = JustificationValues.Center };
                                    para.PrependChild(paraProperties);

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 8pt
                                    Run run = new Run();
                                    run.Append(runProperties);
                                    Text text = new Text(value);
                                    run.Append(text);
                                    para.RemoveAllChildren<Run>();
                                    para.Append(run);
                                }
                            }

                            // Обработка столбцов с участниками
                            for (int i = 5; i < 8; i++)
                            {
                                var cell = newRow.Elements<TableCell>().ElementAt(i);
                                var cellText = cell.InnerText.Trim();

                                var cleanCellText = cellText.Replace(itemTypeBeginMappingKey, "")
                                    .Replace("Participants.", "").Replace("{", "").Replace("}", "");
                                var participantProperty = typeof(Participant).GetProperty(cleanCellText);
                                if (participantProperty != null)
                                {
                                    Paragraph para = cell.Elements<Paragraph>().First();
                                    ParagraphProperties paraProperties = new ParagraphProperties();
                                    paraProperties.Justification = new Justification()
                                        { Val = JustificationValues.Center };
                                    para.PrependChild(paraProperties);

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 8pt
                                    Run run = new Run();
                                    run.Append(runProperties);
                                    Text text = new Text("-");
                                    run.Append(text);
                                    para.RemoveAllChildren<Run>();
                                    para.Append(run);
                                }
                            }

                            table.Append(newRow);
                        }
                    }

                    if (partCount > 1)
                    {
                        int currentRow = table.Elements<TableRow>().Count() - partCount;
                        MergeCells(table, currentRow, currentRow + partCount, 0, 4);
                    }
                }
            });
        }

        private void MergeCells(Table table, int startRowIndex, int endRowIndex, int startColumnIndex,
            int endColumnIndex)
        {
            for (int columnIndex = startColumnIndex; columnIndex <= endColumnIndex; columnIndex++)
            {
                var rows = table.Elements<TableRow>().ToList();

                var startCell = rows[startRowIndex].Elements<TableCell>().ElementAt(columnIndex);
                startCell.Append(new VerticalMerge() { Val = MergedCellValues.Restart });

                for (int rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex++)
                {
                    var cell = rows[rowIndex].Elements<TableCell>().ElementAt(columnIndex);
                    cell.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
                }
            }
        }
    }
}