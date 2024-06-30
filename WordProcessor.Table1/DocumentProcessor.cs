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
                var propKey = "{" + $"{value.GetType().Name}.{keyValuePair.Key}" + "}"; //ключ для замены
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

                    var text = (Text) actionChildElement;
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

        public void MapItems<T>(IEnumerable<T> value)
        {
            if (!value.Any()) return;

            var itemType = value.First().GetType();
            var itemTypeBeginMappingKey = "{" + itemType.Name;
            ProcessEach(e =>
            {
                var table = e as Table;
                var isTable = table is Table;
                if (!isTable) return;

                var secondOrDefault = table.Rows().SecondOrDefault();
                if (secondOrDefault == null) return;
                if (!secondOrDefault.InnerText.Contains(itemTypeBeginMappingKey)) return; //

                var hasTail = table.Rows().Count() > 2;

                var rowTail = table.Rows().Skip(2)
                    .Where(r => !r.InnerText.Contains("{"))
                    .Select(r => r.CloneNode(true)).ToList();

                if (hasTail)
                    foreach (var tableRow in table.Rows().Skip(2))
                        //удалить хвост
                        tableRow.Remove();
                //todo: если строка содержит элементы для маппинга-только тогда маппим!

                var secondLine = secondOrDefault?.CloneNode(true) as TableRow;
                if (!secondLine.InnerText.Contains(itemTypeBeginMappingKey)) return; //
                secondOrDefault.Remove();

                var bindedCount = 0;

                foreach (var v in value)
                {
                    //Для кажждой строки - создай строку в таблице 
                    var propertyValues = new Dictionary<string, string>();
                    //свойства и их значения для конкретной строки данных

                    v.GetProperties()
                        .Select(p =>
                            new KeyValuePair<string, string>("{" + itemType.Name + "." + p.Key + "}",
                                p.Value?.ToString())).ToList()
                        .ForEach(kvp => propertyValues.Add(kvp.Key, kvp.Value));

                    var tr = new TableRow();
                    foreach (var cells in secondLine)
                    {
                        //Для каждой ячейки-создать ячейку с данными из строки dataline

                        var cell = cells as TableCell;
                        var innerText = cell.InnerText;

                        if (!propertyValues.ContainsKey(innerText))
                        {
                            tr.Append(cell.CloneNode(true));
                            //скопировать исходное состояние в новую ячейку, если свойства под нее нет
                            continue;
                        }

                        ++bindedCount;
                        //если указано имя своЙства
                        var tableCellProps = cell.TableCellProperties.CloneNode(true);
                        var newCellParps = new Paragraph(new Run(new Text(propertyValues[innerText])));
                        // //todo:Здесь будет тull reference tсли свойство null
                        var newCell = new TableCell(tableCellProps, newCellParps);

                        tr.Append(newCell);
                    }

                    table.Append(tr);
                }

                if (hasTail)
                    foreach (TableRow tailRow in rowTail)
                        //удалить хвост
                        table.Append(tailRow);
            });
        }
    }
}