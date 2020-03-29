namespace mpFormats.Models
{
    using System.Collections.Generic;
    using JetBrains.Annotations;

    /// <summary>
    /// Класс описывает документ в базе (то, что в xml-файле)
    /// </summary>
    public class TableDocumentInBase
    {
        /// <summary>
        /// Название таблицы
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Название табличного стиля
        /// </summary>
        public string TableStyleName { get; set; }
        
        /// <summary>
        /// Нормативный документ
        /// </summary>
        public string Document { get; set; }
        
        /// <summary>
        /// Описание таблицы (ссылка на таблицу в нормативном документе)
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// Есть ли в штампе фамилии
        /// </summary>
        public bool HasSureNames { get; set; }
        
        /// <summary>
        /// Список должностей, доступных для штампа
        /// </summary>
        public List<string> Surenames { get; set; }
        
        /// <summary>
        /// Список ключей, соответствующий списку доступных должностей
        /// </summary>
        public List<string> SurenamesKeys { get; set; }
        
        /// <summary>
        /// Количество должностей, доступных для размещения в штампе
        /// </summary>
        public string SurenamesCol { get; set; }
        
        /// <summary>
        /// Номер ячейки (колонка, строка) с которой начинается заполнение должностей вниз
        /// </summary>
        [CanBeNull]
        public CellCoordinate FirstCell { get; set; }
        
        /// <summary>
        /// Доступность полей для штампа
        /// </summary>
        public bool HasFields { get; set; }
        
        /// <summary>
        /// Количество ячеек в штампе
        /// </summary>
        public int CellCount { get; set; }
        
        /// <summary>
        /// Ключи полей для вставки
        /// </summary>
        public List<string> FieldsNames { get; set; }
        
        /// <summary>
        /// Координаты для полей (колонка, строка)
        /// </summary>
        public List<CellCoordinate> FieldsCoordinates { get; set; }
        
        /// <summary>
        /// Есть ли ячейка для логотипа
        /// </summary>
        public bool Logo { get; set; }
        
        /// <summary>
        /// Координаты ячейка (колонка, строка) с логотипом
        /// </summary>
        [CanBeNull]
        public CellCoordinate LogoCoordinates { get; set; }
        
        /// <summary>
        /// Имя файла изображения
        /// </summary>
        public string Img { get; set; }
        
        /// <summary>
        /// Координаты ячеек (колонка, строка) у которых делать большой шрифт
        /// </summary>
        public List<CellCoordinate> BigCells { get; set; }
        
        /// <summary>
        /// Координаты ячеек (колонка, строка) у которых делать нулевой отступ
        /// </summary>
        public List<CellCoordinate> NullMargin { get; set; }

        /// <summary>
        /// Координаты ячеек у которых делать малый шрифт
        /// </summary>
        public List<CellCoordinate> SmallCells { get; set; }

        /// <summary>
        /// Высота малого шрифта
        /// </summary>
        public double SmallTextHeight { get; set; }

        /// <summary>
        /// Ячейки у которых нужно "исправить" левую границу
        /// </summary>
        public List<CellCoordinate> FixLeftBorder { get; set; }

        /// <summary>
        /// Ячейки у которых нужно "исправить" левую границу
        /// </summary>
        public List<CellCoordinate> FixTopBorder { get; set; }

        /// <summary>
        /// Ячейки у которых нужно "исправить" верхнюю границу на толстую (0,4)
        /// </summary>
        public List<CellCoordinate> FixTopThickBorder { get; set; }

        /// <summary>
        /// Координаты ячеек (колонка, строка) для вставки поля "Дата"
        /// </summary>
        [CanBeNull]
        public List<CellCoordinate> DateCoordinates { get; set; }
    }
}