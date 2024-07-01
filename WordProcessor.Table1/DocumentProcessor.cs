#region

using AStepanov.Core.Ex;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
                            // Очистка содержимого ячейки
                            cell.RemoveAllChildren<Paragraph>();
                            // Создание нового параграфа и Run
                            var newParagraph = new Paragraph();
                            var newRun = new Run();
                            newRun.Append(new Text(value));
                            newParagraph.Append(newRun);
                            cell.Append(newParagraph);
                        }
                    }

                    table.Append(newRow);
                }
            });
        }
    }
}