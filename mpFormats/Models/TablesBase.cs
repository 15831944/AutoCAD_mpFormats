namespace mpFormats.Models
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Linq;
    using JetBrains.Annotations;

    /// <summary>
    /// Класс описывает текущую базу таблиц (зависит от страны: Россия, Украина и т.п.)
    /// </summary>
    public class TablesBase
    {
        /// <summary>Таблицы в базе</summary>
        public List<TableDocumentInBase> Stamps;

        /// <summary>
        /// Инициализация класса
        /// </summary>
        public TablesBase()
        {
            Stamps = LoadTables(XElement.Parse(GetResourceTextFile("Stamps.xml")));
        }

        /// <summary>Чтение текстового внедренного ресурса</summary>
        /// <param name="filename">File name</param>
        /// <returns></returns>
        private static string GetResourceTextFile(string filename)
        {
            var result = string.Empty;
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("mpFormats.Resources." + filename))
            {
                if (stream != null)
                {
                    using (var sr = new StreamReader(stream))
                    {
                        result = sr.ReadToEnd();
                    }
                }
            }

            return result;
        }
        
        /// <summary>Получение списка таблиц из базы</summary>
        /// <param name="xmlFile">xml-файл базы из ресурсов</param>
        /// <returns></returns>
        private static List<TableDocumentInBase> LoadTables(XElement xmlFile)
        {
            var tables = new List<TableDocumentInBase>();
            foreach (var xElement in xmlFile.Elements("Stamp"))
            {
                var newDocumentInBase = new TableDocumentInBase
                {
                    Name = xElement.Attribute("name")?.Value,
                    TableStyleName = xElement.Attribute("tablestylename")?.Value,
                    Document = xElement.Attribute("document")?.Value,
                    Description = xElement.Attribute("description")?.Value,
                    HasSureNames = bool.TryParse(xElement.Attribute("hassurenames")?.Value, out var b) && b, // false
                    Surenames = xElement.Attribute("surnames")?.Value.Split(',').ToList(),
                    SurenamesKeys = xElement.Attribute("surnameskeys")?.Value.Split(',').ToList(),
                    SurenamesCol = xElement.Attribute("surnamescol")?.Value,
                    FirstCell = GetCellCoordinate(xElement.Attribute("firstcell")?.Value),
                    HasFields = bool.TryParse(xElement.Attribute("hasfields")?.Value, out b) && b, // false
                    CellCount = int.TryParse(xElement.Attribute("cellcount")?.Value, out var i) ? i : 0,
                    FieldsNames = xElement.Attribute("fieldsnames")?.Value.Split(',').ToList(),
                    FieldsCoordinates = GetCellCoordinates(xElement.Attribute("fieldscoordinates")?.Value),
                    Logo = bool.TryParse(xElement.Attribute("logo")?.Value, out b) && b, // false
                    LogoCoordinates = GetCellCoordinate(xElement.Attribute("logocoordinates")?.Value),
                    Img = xElement.Attribute("img")?.Value,
                    BigCells = GetCellCoordinates(xElement.Attribute("bigCells")?.Value),
                    NullMargin = GetCellCoordinates(xElement.Attribute("nullMargin")?.Value),
                    SmallCells = GetCellCoordinates(xElement.Attribute("smallCells")?.Value),
                    SmallTextHeight = double.TryParse(xElement.Attribute("smallTextHeight")?.Value, out var d) ? d : 0.0,
                    FixLeftBorder = GetCellCoordinates(xElement.Attribute("fixLeftBorder")?.Value),
                    FixTopBorder = GetCellCoordinates(xElement.Attribute("fixTopBorder")?.Value),
                    FixTopThickBorder = GetCellCoordinates(xElement.Attribute("fixTopThickBorder")?.Value),
                    DateCoordinates = GetCellCoordinates(xElement.Attribute("dateCoordinates")?.Value)
                };

                if (newDocumentInBase.NullMargin != null && newDocumentInBase.DateCoordinates != null)
                    newDocumentInBase.NullMargin.AddRange(newDocumentInBase.DateCoordinates);

                tables.Add(newDocumentInBase);
            }

            return tables;
        }

        /// <summary>
        /// Получение списка координат из атрибута
        /// </summary>
        /// <param name="attribute">Значение атрибута</param>
        /// <returns></returns>
        private static List<CellCoordinate> GetCellCoordinates(string attribute)
        {
            var coordinates = new List<CellCoordinate>();
            if (!string.IsNullOrWhiteSpace(attribute))
            {
                var splitted = attribute.Split(';');
                foreach (var s in splitted)
                {
                    if (string.IsNullOrWhiteSpace(s))
                        continue;
                    var cc = GetCellCoordinate(s);
                    if (cc != null)
                        coordinates.Add(cc);
                }
            }

            return coordinates;
        }

        /// <summary>
        /// Получение координаты ячейки из строкового представления
        /// </summary>
        /// <param name="stringValue">Строковое представление координаты вида "c,r"</param>
        [CanBeNull]
        private static CellCoordinate GetCellCoordinate(string stringValue)
        {
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                var splitted = stringValue.Split(',');
                if (splitted.Length == 2)
                {
                    if (int.TryParse(splitted[0], out var column) &&
                        int.TryParse(splitted[1], out var row))
                    {
                        return new CellCoordinate
                        {
                            Column = column,
                            Row = row
                        };
                    }
                }
            }

            return null;
        }
    }
}