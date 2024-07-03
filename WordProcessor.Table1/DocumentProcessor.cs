#region

using AStepanov.Core.Ex;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordProcessor.Table1.Entities;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using RowFields = DocumentFormat.OpenXml.Spreadsheet.RowFields;

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

        public void MapValuesPerTable<T>(T value)
        {
            Dictionary<string, object> properties = value.GetProperties();
            var values = new Dictionary<string, string>();
            foreach (var keyValuePair in properties)
            {
                var key = "{" + value.GetType().Name + "." + keyValuePair.Key + "}";
                var empty = string.Empty;
                if (keyValuePair.Value != null)
                    empty = keyValuePair.Value?.ToString();
                values.Add(key, empty);
            }

            ProcessText(o =>
            {
                foreach (var e in values)
                {
                    var innertext = o.InnerText;
                    if (innertext.Contains(e.Key)) o.Text = o.Text.Replace(e.Key, e.Value);
                }
            });
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
                            runProperties.Append(new FontSize(){ Val = "16" });
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

        //старый вариант
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
                    var newRow = (TableRow)templateRow.CloneNode(true);
                    var propertyValues = itemType.GetProperties()
                        .Where(prop => prop.Name != "Participants")
                        .ToDictionary(
                            prop => itemTypeBeginMappingKey + prop.Name + "}",
                            prop => prop.GetValue(item)?.ToString() ?? string.Empty
                        );

                    // Обработка обычных столбцов
                    for (int i = 0; i < 5; i++)
                    {
                        var cell = newRow.Elements<TableCell>().ElementAt(i);
                        var cellText = cell.InnerText.Trim();
                        if (propertyValues.TryGetValue(cellText, out var value))
                        {
                            Paragraph para = cell.Elements<Paragraph>().First();

                            ParagraphProperties paraProperties = new ParagraphProperties();
                            paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                            para.PrependChild(paraProperties);

                            RunProperties runProperties = new RunProperties();
                            runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
                            Run run = new Run();
                            run.Append(runProperties);
                            Text text = new Text(value);
                            run.Append(text);
                            para.RemoveAllChildren<Run>();
                            para.Append(run);
                        }
                    }

                    // Обработка столбцов с участниками
                    if (item is Startup startup && startup.Participants != null)
                    {
                        var participantsProperties = typeof(Participant).GetProperties()
                            .ToDictionary(
                                prop => itemTypeBeginMappingKey + "Participants." + prop.Name + "}",
                                prop => string.Empty
                            );

                        for (int i = 5; i < 8; i++)
                        {
                            var cell = newRow.Elements<TableCell>().ElementAt(i);
                            var cellText = cell.InnerText.Trim();

                            // Проверяем, содержится ли текст ячейки в ключах словаря
                            if (participantsProperties.ContainsKey(cellText))
                            {
                                cell.RemoveAllChildren<Paragraph>();

                                foreach (var participant in startup.Participants)
                                {
                                    var newParagraph = new Paragraph();

                                    ParagraphProperties paraProperties = new ParagraphProperties();
                                    paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                    newParagraph.PrependChild(paraProperties);

                                    var newRun = new Run();

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.Append(new FontSize() { Val = "16" });
                                    newRun.Append(runProperties);

                                    var cleanCellText = cellText.Replace(itemTypeBeginMappingKey, "").Replace("Participants.", "").Replace("{", "").Replace("}", "");
                                    var participantProperty = typeof(Participant).GetProperty(cleanCellText);
                                    if (participantProperty != null)
                                    {
                                        var participantValue = participantProperty.GetValue(participant)?.ToString() ??
                                                               string.Empty;
                                        newRun.Append(new Text(participantValue));
                                    }

                                    //TODO: ДОБАВЛЯТЬ НОВУЮ ЯЧЕЙКУ В ЯЧЕЙКУ (по умолчанию так нельзя)
                                    newParagraph.Append(newRun);
                                    cell.Append(newParagraph);
                                }
                            }
                        }
                    }
                    table.Append(newRow);
                }
            });
        }

        #region Вариант где создается куча строк, остается придумать как мёрджить 1-5 столбцы

        /*public void MapItemsOther<T>(IEnumerable<T> items, int rowSkips)
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
                    var propertyValues = itemType.GetProperties()
                        .Where(prop => prop.Name != "Participants")
                        .ToDictionary(
                            prop => itemTypeBeginMappingKey + prop.Name + "}",
                            prop => prop.GetValue(item)?.ToString() ?? string.Empty
                        );

                    // Создание строки для каждого участника
                    if (item is Startup startup && startup.Participants != null)
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
                                    paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                    para.PrependChild(paraProperties);

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
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

                                var cleanCellText = cellText.Replace(itemTypeBeginMappingKey, "").Replace("Participants.", "").Replace("{", "").Replace("}", "");
                                var participantProperty = typeof(Participant).GetProperty(cleanCellText);
                                if (participantProperty != null)
                                {
                                    var participantValue = participantProperty.GetValue(participant)?.ToString() ?? string.Empty;

                                    Paragraph para = cell.Elements<Paragraph>().First();
                                    ParagraphProperties paraProperties = new ParagraphProperties();
                                    paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                    para.PrependChild(paraProperties);

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
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
                                paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                para.PrependChild(paraProperties);

                                RunProperties runProperties = new RunProperties();
                                runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
                                Run run = new Run();
                                run.Append(runProperties);
                                Text text = new Text(value);
                                run.Append(text);
                                para.RemoveAllChildren<Run>();
                                para.Append(run);
                            }
                        }

                        table.Append(newRow);
                    }
                }
            });
        }*/

        #endregion

        #region Кривой вариант из скрина который я кидал в ЛС Ане
        /*        public void MapItemsOther<T>(IEnumerable<T> items, int rowSkips)
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
                            var propertyValues = itemType.GetProperties()
                                .Where(prop => prop.Name != "Participants")
                                .ToDictionary(
                                    prop => itemTypeBeginMappingKey + prop.Name + "}",
                                    prop => prop.GetValue(item)?.ToString() ?? string.Empty
                                );

                            TableRow headerRow = (TableRow)templateRow.CloneNode(true);

                            // Обработка обычных столбцов (1-5)
                            for (int i = 0; i < 5; i++)
                            {
                                var cell = headerRow.Elements<TableCell>().ElementAt(i);
                                var cellText = cell.InnerText.Trim();
                                if (propertyValues.TryGetValue(cellText, out var value))
                                {
                                    Paragraph para = cell.Elements<Paragraph>().First();

                                    ParagraphProperties paraProperties = new ParagraphProperties();
                                    paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                    para.PrependChild(paraProperties);

                                    RunProperties runProperties = new RunProperties();
                                    runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
                                    Run run = new Run();
                                    run.Append(runProperties);
                                    Text text = new Text(value);
                                    run.Append(text);
                                    para.RemoveAllChildren<Run>();
                                    para.Append(run);
                                }
                            }
                            table.Append(headerRow);

                            // Обработка столбцов с участниками (6-8)
                            if (item is Startup startup && startup.Participants != null)
                            {
                                foreach (var participant in startup.Participants)
                                {
                                    var participantRow = (TableRow)templateRow.CloneNode(true);

                                    for (int i = 5; i < 8; i++)
                                    {
                                        var cell = participantRow.Elements<TableCell>().ElementAt(i);
                                        var cellText = cell.InnerText.Trim();

                                        var cleanCellText = cellText.Replace(itemTypeBeginMappingKey, "").Replace("Participants.", "").Replace("{", "").Replace("}", "");
                                        var participantProperty = typeof(Participant).GetProperty(cleanCellText);
                                        if (participantProperty != null)
                                        {
                                            var participantValue = participantProperty.GetValue(participant)?.ToString() ?? string.Empty;

                                            Paragraph para = cell.Elements<Paragraph>().First();
                                            ParagraphProperties paraProperties = new ParagraphProperties();
                                            paraProperties.Justification = new Justification() { Val = JustificationValues.Center };
                                            para.PrependChild(paraProperties);

                                            RunProperties runProperties = new RunProperties();
                                            runProperties.FontSize = new FontSize() { Val = "16" }; // Размер шрифта 12pt
                                            Run run = new Run();
                                            run.Append(runProperties);
                                            Text text = new Text(participantValue);
                                            run.Append(text);
                                            para.RemoveAllChildren<Run>();
                                            para.Append(run);
                                        }
                                    }
                                    table.Append(participantRow);
                                }

                                // Объединение ячеек с 1 по 5 столбцы для всех строк стартапа
                                int rowsCount = startup.Participants.Count();
                                MergeHeaderCells(table, rowsCount);
                            }
                            else
                            {
                                table.Append(headerRow);
                            }
                        }
                    });
                }

                private void MergeHeaderCells(Table table, int rowsCount)
                {
                    var rows = table.Elements<TableRow>().ToList();
                    var startIndex = rows.Count - rowsCount - 1; // -1 для headerRow

                    for (int i = startIndex; i <= startIndex + rowsCount; i++)
                    {
                        var row = rows[i];
                        for (int j = 0; j < 5; j++)
                        {
                            var cell = row.Elements<TableCell>().ElementAt(j);
                            if (i == startIndex)
                            {
                                cell.GetFirstChild<TableCellProperties>().Append(new HorizontalMerge { Val = MergedCellValues.Restart });
                            }
                            else
                            {
                                cell.RemoveAllChildren<Paragraph>();
                                cell.GetFirstChild<TableCellProperties>().Append(new HorizontalMerge { Val = MergedCellValues.Continue });
                            }
                        }
                    }
                }*/
        #endregion

        private void ReplaceCellText(TableCell cell, string value)
        {
            cell.RemoveAllChildren<Paragraph>();
            var newParagraph = new Paragraph();
            var newRun = new Run(new Text(value));
            newParagraph.Append(newRun);
            cell.Append(newParagraph);
        }
    }
}