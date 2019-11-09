namespace mpFormats
{
    using System;
    using System.Globalization;
    using System.Linq;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.Colors;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    /// <summary>
    /// Запуск функции и создание блока
    /// </summary>
    public class MpFormatsAdd
    {
        private const string LangItem = "mpFormats";
        private MpFormats _mpFormats;

        [CommandMethod("ModPlus", "mpFormats", CommandFlags.Modal)]
        public void StartMpFormats()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());

            if (_mpFormats == null)
            {
                _mpFormats = new MpFormats();
                _mpFormats.Closed += (sender, args) => _mpFormats = null;
            }

            if (_mpFormats.IsLoaded)
                _mpFormats.Activate();
            else
                Application.ShowModelessWindow(Application.MainWindow.Handle, _mpFormats, false);
        }

        /// <summary>
        /// Draw block
        /// </summary>
        /// <param name="format">формат</param>
        /// <param name="multiplicity">кратность</param>
        /// <param name="side">Сторона кратности</param>
        /// <param name="orientation">Ориентация</param>
        /// <param name="number">Номер страницы (да, нет)</param>
        /// <param name="copy">Копировал</param>
        /// <param name="bottomFrame">Нижняя рамка</param>
        /// <param name="hasFpt">Есть ли начальная точка</param>
        /// <param name="insertPt">Начальная точка (для замены)</param>
        /// <param name="txtStyle">TextStyle name</param>
        /// <param name="scale">масштаб</param>
        /// <param name="rotation">Поворот</param>
        /// <param name="setCurrentLayer">Установить текущий слой</param>
        /// <param name="accordingToGost">True - размеры некоторых форматок будут соответствовать таблице в ГОСТ,
        /// False - размеры будут получаться математически</param>
        /// <param name="bottomLeftPt">Нижняя верхняя точка</param>
        /// <param name="topLeftPt">Верхняя левая точка</param>
        /// <param name="bottomRightPt">Нижняя правая точка</param>
        /// <param name="replaceVector3D">Вектор смещения</param>
        /// <param name="blockInsertionPoint">Точка вставки блока</param>
        /// <returns></returns>
        public static bool DrawBlock(
            string format,
            string multiplicity,
            string side,
            string orientation,
            bool number,
            bool copy,
            string bottomFrame,
            bool hasFpt,
            Point3d insertPt,
            string txtStyle,
            double scale,
            double? rotation,
            bool setCurrentLayer,
            bool accordingToGost,
            out Point3d bottomLeftPt,
            out Point3d topLeftPt,
            out Point3d bottomRightPt,
            out Vector3d replaceVector3D,
            out Point3d blockInsertionPoint)
        {
            bottomLeftPt = bottomRightPt = topLeftPt = blockInsertionPoint = new Point3d(0.0, 0.0, 0.0);
            replaceVector3D = new Vector3d(0.0, 0.0, 0.0);
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var returned = false;

            try
            {
                using (doc.LockDocument())
                {
                    // Задаем значение ширины и высоты в зависимости формата, кратности и стороны кратности
                    var formatSize = GetFormatSize(format, orientation, side, multiplicity, accordingToGost);

                    #region points

                    var pt1 = new Point3d(0.0, 0.0, 0.0);
                    var pt2 = new Point3d(0.0 + formatSize.Width, 0.0, 0.0);

                    // Для форматов А4 и А3 нижняя рамка 10мм (по ГОСТ)
                    Point3d pt11;
                    Point3d pt22;
                    if (format.Equals("A4") || format.Equals("A3"))
                    {
                        if (bottomFrame.Equals(Language.GetItem(LangItem, "h19")))
                        {
                            pt11 = new Point3d(pt1.X + 20, pt1.Y + 10, 0.0);
                            pt22 = new Point3d(pt2.X - 5, pt2.Y + 10, 0.0);
                        }
                        else
                        {
                            pt11 = new Point3d(pt1.X + 20, pt1.Y + 5, 0.0);
                            pt22 = new Point3d(pt2.X - 5, pt2.Y + 5, 0.0);
                        }
                    }
                    else
                    {
                        pt11 = new Point3d(pt1.X + 20, pt1.Y + 5, 0.0);
                        pt22 = new Point3d(pt2.X - 5, pt2.Y + 5, 0.0);
                    }

                    var pt3 = new Point3d(0.0 + formatSize.Width, 0.0 + formatSize.Height, 0.0);
                    var pt33 = new Point3d(pt3.X - 5, pt3.Y - 5, 0.0);
                    var pt4 = new Point3d(0.0, 0.0 + formatSize.Height, 0.0);
                    var pt44 = new Point3d(pt4.X + 20, pt4.Y - 5, 0.0);
                    var ptt1 = new Point3d(pt2.X - 55, pt22.Y - 4, 0.0);
                    var ptt2 = new Point3d(pt2.X - 125, pt22.Y - 4, 0.0);
                    var pts1 = new Point3dCollection { pt1, pt2, pt3, pt4 };
                    var pts2 = new Point3dCollection { pt11, pt22, pt33, pt44 };

                    // points for stamps
                    bottomLeftPt = pt11;
                    topLeftPt = pt44;
                    bottomRightPt = pt22;

                    #endregion

                    #region block name

                    var isNumber = number ? "N" : "NN";
                    var isCopy = copy ? "C" : "NC";
                    var isAcc = accordingToGost ? "A" : "NA";

                    var blockName = multiplicity.Equals("1") 
                        ? $"{format}_{orientation}_{side}_{isNumber}_{isCopy}_{isAcc}" 
                        : $"{format}x{multiplicity}_{orientation}_{side}_{isNumber}_{isCopy}_{isAcc}";

                    if (format.Equals("A4") || format.Equals("A3"))
                        blockName = $"{blockName}_{bottomFrame}";

                    #endregion

                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        // Если в базе есть такой блок - вставляем его
                        if (blockTable.Has(blockName))
                        {
                            var blockId = blockTable[blockName];
                            var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);

                            // Если нет начальной точки, то рисуем через джигу
                            if (!hasFpt)
                            {
                                var pt = new Point3d(0, 0, 0);
                                var blockReference = new BlockReference(pt, blockId);

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                var entJig = new BlockJig(blockReference);
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    blockInsertionPoint = blockReference.Position;
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            }

                            // Если есть начальная точка - то вставляем в нее
                            else
                            {
                                var blockReference = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = blockReference.Position;

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                btr.AppendEntity(blockReference);
                                tr.AddNewlyCreatedDBObject(blockReference, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        }

                        // Если блока нет - создаем и вставляем
                        else
                        {
                            try
                            {
                                SymbolUtilityServices.ValidateSymbolName(
                                    blockName, false);
                            }
                            catch
                            {
                                MessageBox.Show(Language.GetItem(LangItem, "err14"));
                            }

                            var btr = new BlockTableRecord { Name = blockName };

                            // Add the new block to the block table
                            blockTable.UpgradeOpen();
                            blockTable.Add(btr);
                            tr.AddNewlyCreatedDBObject(btr, true);

                            // Рисуем примитивы и добавляем в блок
                            var objectCollection = new DBObjectCollection();

                            // внешняя рамка
                            var pLine1 = new Polyline
                            {
                                LineWeight = LineWeight.LineWeight020,
                                Layer = "0",
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                Closed = true
                            };
                            for (var i = 0; i < pts1.Count; i++)
                            {
                                var pp = new Point2d(pts1[i].X, pts1[i].Y);
                                pLine1.AddVertexAt(i, pp, 0, 0, 0);
                            }

                            // внутренняя рамка
                            var pLine2 = new Polyline
                            {
                                LineWeight = LineWeight.LineWeight050,
                                Layer = "0",
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                Closed = true
                            };
                            for (var i = 0; i < pts2.Count; i++)
                            {
                                var pp = new Point2d(pts2[i].X, pts2[i].Y);
                                pLine2.AddVertexAt(i, pp, 0, 0, 0);
                            }

                            // Формат
                            var txt1 = new DBText
                            {
                                Height = 3,
                                Position = ptt1,
                                Layer = "0",
                                Annotative = AnnotativeStates.False,
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                TextStyleId = tst[txtStyle],
                                TextString = !multiplicity.Equals("1")
                                    ? "Формат" + " " + format + "x" + multiplicity
                                    : "Формат" + " " + format
                            };

                            // Копировал
                            if (copy)
                            {
                                var txt2 = new DBText
                                {
                                    Height = 3,
                                    TextString = "Копировал:",
                                    Position = ptt2,
                                    Layer = "0",
                                    Annotative = AnnotativeStates.False,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                    TextStyleId = tst[txtStyle]
                                };
                                objectCollection.Add(txt2);
                            }

                            // Номер листа
                            if (number)
                            {
                                var ptn1 = new Point3d(pt33.X - 10, pt33.Y, 0.0);
                                var ptn2 = new Point3d(ptn1.X, ptn1.Y - 7, 0.0);
                                var ptn3 = new Point3d(pt33.X, pt33.Y - 7, 0.0);
                                var line1 = new Line
                                {
                                    StartPoint = ptn1,
                                    EndPoint = ptn2,
                                    Layer = "0",
                                    LineWeight = LineWeight.LineWeight050,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                                };
                                var line2 = new Line
                                {
                                    StartPoint = ptn2,
                                    EndPoint = ptn3,
                                    Layer = "0",
                                    LineWeight = LineWeight.LineWeight050,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                                };
                                objectCollection.Add(line1);
                                objectCollection.Add(line2);
                            }

                            objectCollection.Add(pLine1);
                            objectCollection.Add(pLine2);
                            objectCollection.Add(txt1);

                            foreach (Entity ent in objectCollection)
                            {
                                btr.AppendEntity(ent);
                                tr.AddNewlyCreatedDBObject(ent, true);
                            }

                            // Добавляем расширенные данные для возможности замены
                            ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                            btr.XData = new ResultBuffer(
                                new TypedValue(1001, "MP_FORMAT"),
                                new TypedValue(1000, "MP_FORMAT"));

                            // annotative state
                            btr.Annotative = AnnotativeStates.False;

                            var blockId = blockTable[blockName];

                            var cbtr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);

                            // Если начальной точки нет - рисуем через джигу
                            if (!hasFpt)
                            {
                                var pt = new Point3d(0, 0, 0);
                                var blockReference = new BlockReference(pt, blockId);

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                blockReference.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));

                                var entJig = new BlockJig(blockReference);

                                // Perform the jig operation
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    blockInsertionPoint = blockReference.Position;
                                    var ent = entJig.GetEntity();

                                    ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                    ent.XData = new ResultBuffer(
                                        new TypedValue(1001, "MP_FORMAT"),
                                        new TypedValue(1000, "MP_FORMAT"));
                                    cbtr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            }

                            // Если начальная точка есть - вставляем в нее
                            else
                            {
                                var blockReference = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = blockReference.Position;

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                blockReference.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));
                                cbtr.AppendEntity(blockReference);
                                tr.AddNewlyCreatedDBObject(blockReference, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        }

                        tr.Commit();
                    }
                }

                return returned;
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
                return false;
            }
        }

        public static bool DrawBlockHand(
            double width,
            double height,
            bool number, // Номер страницы (да, нет)
            bool copy, // Копировал
            bool hasFpt, // Есть ли начальная точка
            Point3d insertPt, // Начальная точка (для замены)
            string txtStyle,
            double scale, // масштаб
            double? rotation,
            bool setCurrentLayer,
            out Point3d bottomLeftPt,
            out Point3d topLeftPt,
            out Point3d bottomRightPt,
            out Vector3d replaceVector3D,
            out Point3d blockInsertionPoint)
        {
            bottomLeftPt = bottomRightPt = topLeftPt = blockInsertionPoint = new Point3d(0.0, 0.0, 0.0);
            replaceVector3D = new Vector3d(0.0, 0.0, 0.0);
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var returned = false;

            try
            {
                using (doc.LockDocument())
                {
                    #region block name
                    var isNumber = number ? "N" : "NN";
                    var isCopy = copy ? "C" : "NC";
                    var blockName = width.ToString(CultureInfo.InvariantCulture) + "x" +
                                    height.ToString(CultureInfo.InvariantCulture) + "_" +
                                    isNumber + "_" + isCopy;
                    #endregion

                    #region points
                    var pt1 = new Point3d(0.0, 0.0, 0.0);
                    var pt2 = new Point3d(0.0 + width, 0.0, 0.0);
                    var pt11 = new Point3d(pt1.X + 20, pt1.Y + 5, 0.0);
                    var pt22 = new Point3d(pt2.X - 5, pt2.Y + 5, 0.0);
                    var pt3 = new Point3d(0.0 + width, 0.0 + height, 0.0);
                    var pt33 = new Point3d(pt3.X - 5, pt3.Y - 5, 0.0);
                    var pt4 = new Point3d(0.0, 0.0 + height, 0.0);
                    var pt44 = new Point3d(pt4.X + 20, pt4.Y - 5, 0.0);
                    var ptt1 = new Point3d(pt2.X - 55, pt22.Y - 4, 0.0);
                    var ptt2 = new Point3d(pt2.X - 125, pt22.Y - 4, 0.0);

                    // points for stamps
                    bottomLeftPt = pt11;
                    topLeftPt = pt44;
                    bottomRightPt = pt22;

                    var pts1 = new Point3dCollection();
                    var pts2 = new Point3dCollection();
                    pts1.Add(pt1);
                    pts1.Add(pt2);
                    pts1.Add(pt3);
                    pts1.Add(pt4);
                    pts2.Add(pt11);
                    pts2.Add(pt22);
                    pts2.Add(pt33);
                    pts2.Add(pt44);
                    #endregion

                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        // Если в базе есть такой блок - вставляем его
                        if (bt.Has(blockName))
                        {
                            var blockId = bt[blockName];
                            var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);

                            // Если нет начальной точки, то рисуем через джигу
                            if (!hasFpt)
                            {
                                var pt = new Point3d(0, 0, 0);
                                var blockReference = new BlockReference(pt, blockId);

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                var entJig = new BlockJig(blockReference);

                                // Perform the jig operation
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    blockInsertionPoint = blockReference.Position;
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            }

                            // Если есть начальная точка - то вставлем в нее
                            else
                            {
                                var blockReference = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = blockReference.Position;

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                btr.AppendEntity(blockReference);
                                tr.AddNewlyCreatedDBObject(blockReference, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        }

                        // Если блока нет - создаем и вставляем
                        else
                        {
                            try
                            {
                                SymbolUtilityServices.ValidateSymbolName(blockName, false);
                            }
                            catch
                            {
                                MessageBox.Show(Language.GetItem(LangItem, "err14"));
                            }

                            var btr = new BlockTableRecord { Name = blockName };

                            // Add the new block to the block table
                            bt.UpgradeOpen();
                            bt.Add(btr);
                            tr.AddNewlyCreatedDBObject(btr, true);

                            // Рисуем примитивы и добавляем в блок
                            var objectCollection = new DBObjectCollection();

                            // внешняя рамка
                            var pLine1 = new Polyline
                            {
                                LineWeight = LineWeight.LineWeight020,
                                Layer = "0",
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                Closed = true
                            };
                            for (var i = 0; i < pts1.Count; i++)
                            {
                                var pp = new Point2d(pts1[i].X, pts1[i].Y);
                                pLine1.AddVertexAt(i, pp, 0, 0, 0);
                            }

                            // внутренняя рамка
                            var pLine2 = new Polyline
                            {
                                LineWeight = LineWeight.LineWeight050,
                                Layer = "0",
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                Closed = true
                            };
                            for (var i = 0; i < pts2.Count; i++)
                            {
                                var pp = new Point2d(pts2[i].X, pts2[i].Y);
                                pLine2.AddVertexAt(i, pp, 0, 0, 0);
                            }

                            // Формат
                            var txt1 = new DBText
                            {
                                Height = 3,
                                TextString = "Формат",
                                Position = ptt1,
                                Layer = "0",
                                Annotative = AnnotativeStates.False,
                                Linetype = "Continuous",
                                Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                TextStyleId = tst[txtStyle]
                            };

                            // Копировал
                            if (copy)
                            {
                                var txt2 = new DBText
                                {
                                    Height = 3,
                                    TextString = "Копировал:",
                                    Position = ptt2,
                                    Layer = "0",
                                    Annotative = AnnotativeStates.False,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0),
                                    TextStyleId = tst[txtStyle]
                                };
                                objectCollection.Add(txt2);
                            }

                            if (number)
                            {
                                var ptn1 = new Point3d(pt33.X - 10, pt33.Y, 0.0);
                                var ptn2 = new Point3d(ptn1.X, ptn1.Y - 7, 0.0);
                                var ptn3 = new Point3d(pt33.X, pt33.Y - 7, 0.0);
                                var line1 = new Line
                                {
                                    StartPoint = ptn1,
                                    EndPoint = ptn2,
                                    Layer = "0",
                                    LineWeight = LineWeight.LineWeight050,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                                };
                                var line2 = new Line
                                {
                                    StartPoint = ptn2,
                                    EndPoint = ptn3,
                                    Layer = "0",
                                    LineWeight = LineWeight.LineWeight050,
                                    Linetype = "Continuous",
                                    Color = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                                };
                                objectCollection.Add(line1);
                                objectCollection.Add(line2);
                            }

                            objectCollection.Add(pLine1);
                            objectCollection.Add(pLine2);
                            objectCollection.Add(txt1);

                            foreach (Entity ent in objectCollection)
                            {
                                btr.AppendEntity(ent);
                                tr.AddNewlyCreatedDBObject(ent, true);
                            }

                            // Добавляем расширенные данные для возможности замены
                            ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                            btr.XData = new ResultBuffer(
                                new TypedValue(1001, "MP_FORMAT"),
                                new TypedValue(1000, "MP_FORMAT"));

                            // annotative state
                            btr.Annotative = AnnotativeStates.False;

                            var blockId = bt[blockName];

                            var cbtr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);

                            // Если начальной точки нет - рисуем через джигу
                            if (!hasFpt)
                            {
                                var pt = new Point3d(0, 0, 0);
                                var blockReference = new BlockReference(pt, blockId);
                                blockInsertionPoint = blockReference.Position;

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                blockReference.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));

                                var entJig = new BlockJig(blockReference);
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                    ent.XData = new ResultBuffer(
                                        new TypedValue(1001, "MP_FORMAT"),
                                        new TypedValue(1000, "MP_FORMAT"));
                                    cbtr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            }

                            // Если начальная точка есть - вставляем в нее
                            else
                            {
                                var blockReference = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = blockReference.Position;

                                SetFormatBlockProperties(blockReference, insertPt, scale, rotation, setCurrentLayer, db);

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                blockReference.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));
                                cbtr.AppendEntity(blockReference);
                                tr.AddNewlyCreatedDBObject(blockReference, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        }

                        tr.Commit();
                    }
                }

                return returned;
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
                return false;
            }
        }

        private static void SetFormatBlockProperties(
            BlockReference blockReference,
            Point3d insertPt,
            double scale,
            double? rotation,
            bool setCurrentLayer,
            Database db)
        {
            // scale
            var mat = Matrix3d.Scaling(scale, insertPt);
            blockReference.TransformBy(mat);

            // set current layer
            if (setCurrentLayer)
                blockReference.LayerId = db.Clayer;
            else
                blockReference.Layer = "0";

            // rotation
            if (rotation != null)
            {
                var rm = Matrix3d.Rotation(rotation.Value, blockReference.Normal, insertPt);
                blockReference.TransformBy(rm);
            }
        }

        public static void ReplaceBlock(
            string format, // формат
            string multiplicity, // кратность
            string side, // Сторона кратности
            string orientation, // Ориентация
            bool number, // Номер страницы (да, нет)
            bool copy, // Копировал
            string bottomFrame, // Нижняя рамка
            string txtStyle,
            double scale,
            bool setCurrentLayer,
            bool accordingToGost)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                var peo = new PromptEntityOptions("\n" + Language.GetItem(LangItem, "msg9"));
                peo.SetRejectMessage("\n" + Language.GetItem(LangItem, "msg10"));
                peo.AddAllowedClass(typeof(BlockReference), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                    return;
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        var obj = tr.GetObject(per.ObjectId, OpenMode.ForWrite);
                        var blockReference = (BlockReference)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                        if (bt.Has(blockReference.Name))
                        {
                            var dbObject = bt[blockReference.Name].GetObject(OpenMode.ForWrite);
                            var rb = dbObject.XData;

                            if (rb != null)
                            {
                                if (rb.Cast<TypedValue>().Any(tv => tv.Value.Equals("MP_FORMAT")))
                                {
                                    var pt = blockReference.Position;
                                    var rotation = blockReference.Rotation;
                                    blockReference.Erase(true);
                                    DrawBlock(
                                        format,
                                        multiplicity,
                                        side,
                                        orientation,
                                        number,
                                        copy,
                                        bottomFrame,
                                        true,
                                        pt,
                                        txtStyle,
                                        scale,
                                        rotation,
                                        setCurrentLayer,
                                        accordingToGost,
                                        out _,
                                        out _,
                                        out _,
                                        out _,
                                        out pt
                                    );
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        public static void ReplaceBlockHand(
            double width,
            double height,
            bool number, // Номер страницы (да, нет)
            bool copy, // Копировал
            string txtStyle,
            double scale,
            bool setCurrentLayer)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                var peo = new PromptEntityOptions("\n" + Language.GetItem(LangItem, "msg9"));
                peo.SetRejectMessage("\n" + Language.GetItem(LangItem, "msg10"));
                peo.AddAllowedClass(typeof(BlockReference), false);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                    return;
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        var obj = tr.GetObject(per.ObjectId, OpenMode.ForWrite);
                        var blk = (BlockReference)tr.GetObject(obj.ObjectId, OpenMode.ForWrite);
                        if (bt.Has(blk.Name))
                        {
                            var dbblock = bt[blk.Name].GetObject(OpenMode.ForWrite);
                            var rb = dbblock.XData;

                            if (rb != null)
                            {
                                if (rb.Cast<TypedValue>().Any(tv => tv.Value.Equals("MP_FORMAT")))
                                {
                                    var pt = blk.Position;
                                    var rotation = blk.Rotation;
                                    blk.Erase(true);

                                    DrawBlockHand(
                                        width,
                                        height,
                                        number,
                                        copy,
                                        true,
                                        pt,
                                        txtStyle,
                                        scale,
                                        rotation,
                                        setCurrentLayer,
                                        out _,
                                        out _,
                                        out _,
                                        out _,
                                        out pt
                                    );
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        /// <summary>
        /// Get format size
        /// </summary>
        /// <param name="format">Format name</param>
        /// <param name="orientation">Sheet orientation</param>
        /// <param name="side">Side of multiplicity</param>
        /// <param name="multiplicityString">Multiplicity string value</param>
        /// <param name="accordingToGost">True - размеры некоторых форматок будут соответствовать таблице в ГОСТ,
        /// False - размеры будут получаться математически</param>
        /// <returns>Instance of <see cref="FormatSize"/></returns>
        public static FormatSize GetFormatSize(
            string format,
            string orientation,
            string side,
            string multiplicityString,
            bool accordingToGost)
        {
            var isLandscape = orientation.Equals(Language.GetItem(LangItem, "h8"));
            var isShortSizeMultiplicity = side.Equals(Language.GetItem(LangItem, "h11"));
            var multiplicity = int.Parse(multiplicityString);

            if (format.Equals("A0"))
            {
                if (multiplicity > 1)
                {
                    if (isLandscape)
                    {
                        return isShortSizeMultiplicity
                            ? new FormatSize(841 * multiplicity, 1189)
                            : new FormatSize(1189 * multiplicity, 841);
                    }

                    return isShortSizeMultiplicity
                        ? new FormatSize(1189, 841 * multiplicity)
                        : new FormatSize(841, 1189 * multiplicity);
                }

                return isLandscape ? new FormatSize(1189, 841) : new FormatSize(841, 1189);
            }

            if (format.Equals("A1"))
            {
                if (multiplicity > 1)
                {
                    if (isLandscape)
                    {
                        if (isShortSizeMultiplicity)
                        {
                            if (accordingToGost)
                            {
                                if (multiplicity == 3)
                                    return new FormatSize(1783, 841);
                                if (multiplicity == 4)
                                    return new FormatSize(2378, 841);
                            }

                            return new FormatSize(594 * multiplicity, 841);
                        }

                        return new FormatSize(841 * multiplicity, 594);
                    }

                    if (isShortSizeMultiplicity)
                    {
                        if (accordingToGost)
                        {
                            if (multiplicity == 3)
                                return new FormatSize(841, 1783);
                            if (multiplicity == 4)
                                return new FormatSize(841, 2378);
                        }

                        return new FormatSize(841, 594 * multiplicity);
                    }

                    return new FormatSize(594, 841 * multiplicity);
                }

                return isLandscape ? new FormatSize(841, 594) : new FormatSize(594, 841);
            }

            if (format.Equals("A2"))
            {
                if (multiplicity > 1)
                {
                    if (isLandscape)
                    {
                        if (isShortSizeMultiplicity)
                        {
                            if (accordingToGost)
                            {
                                if (multiplicity == 3)
                                    return new FormatSize(1261, 594);
                                if (multiplicity == 4)
                                    return new FormatSize(1682, 594);
                                if (multiplicity == 5)
                                    return new FormatSize(2102, 594);
                            }

                            return new FormatSize(420 * multiplicity, 594);
                        }

                        return new FormatSize(594 * multiplicity, 420);
                    }

                    if (isShortSizeMultiplicity)
                    {
                        if (accordingToGost)
                        {
                            if (multiplicity == 3)
                                return new FormatSize(594, 1261);
                            if (multiplicity == 4)
                                return new FormatSize(594, 1682);
                            if (multiplicity == 5)
                                return new FormatSize(594, 2102);
                        }

                        return new FormatSize(594, 420 * multiplicity);
                    }

                    return new FormatSize(420, 594 * multiplicity);
                }

                return isLandscape ? new FormatSize(594, 420) : new FormatSize(420, 594);
            }

            if (format.Equals("A3"))
            {
                if (multiplicity > 1)
                {
                    if (isLandscape)
                    {
                        if (isShortSizeMultiplicity)
                        {
                            if (accordingToGost)
                            {
                                if (multiplicity == 4)
                                    return new FormatSize(1189, 420);
                                if (multiplicity == 5)
                                    return new FormatSize(1486, 420);
                                if (multiplicity == 6)
                                    return new FormatSize(1783, 420);
                                if (multiplicity == 7)
                                    return new FormatSize(2080, 420);
                            }

                            return new FormatSize(297 * multiplicity, 420);
                        }

                        return new FormatSize(420 * multiplicity, 297);
                    }

                    if (isShortSizeMultiplicity)
                    {
                        if (accordingToGost)
                        {
                            if (multiplicity == 4)
                                return new FormatSize(420, 1189);
                            if (multiplicity == 5)
                                return new FormatSize(420, 1486);
                            if (multiplicity == 6)
                                return new FormatSize(420, 1783);
                            if (multiplicity == 7)
                                return new FormatSize(420, 2080);
                        }

                        return new FormatSize(420, 297 * multiplicity);
                    }

                    return new FormatSize(297, 420 * multiplicity);
                }

                return isLandscape ? new FormatSize(420, 297) : new FormatSize(297, 420);
            }

            if (format.Equals("A4"))
            {
                if (multiplicity > 1)
                {
                    if (isLandscape)
                    {
                        if (isShortSizeMultiplicity)
                        {
                            if (accordingToGost)
                            {
                                if (multiplicity == 4)
                                    return new FormatSize(841, 297);
                                if (multiplicity == 5)
                                    return new FormatSize(1051, 297);
                                if (multiplicity == 6)
                                    return new FormatSize(1261, 297);
                                if (multiplicity == 7)
                                    return new FormatSize(1471, 297);
                                if (multiplicity == 8)
                                    return new FormatSize(1682, 297);
                                if (multiplicity == 9)
                                    return new FormatSize(1892, 297);
                            }

                            return new FormatSize(210 * multiplicity, 297);
                        }

                        return new FormatSize(297 * multiplicity, 210);
                    }

                    if (isShortSizeMultiplicity)
                    {
                        if (accordingToGost)
                        {
                            if (multiplicity == 4)
                                return new FormatSize(297, 841);
                            if (multiplicity == 5)
                                return new FormatSize(297, 1051);
                            if (multiplicity == 6)
                                return new FormatSize(297, 1261);
                            if (multiplicity == 7)
                                return new FormatSize(297, 1471);
                            if (multiplicity == 8)
                                return new FormatSize(297, 1682);
                            if (multiplicity == 9)
                                return new FormatSize(297, 1892);
                        }

                        return new FormatSize(297, 210 * multiplicity);
                    }

                    return new FormatSize(210, 297 * multiplicity);
                }

                return isLandscape ? new FormatSize(297, 210) : new FormatSize(210, 297);
            }

            if (format.Equals("A5"))
                return isLandscape ? new FormatSize(210, 148) : new FormatSize(148, 210);

            if (format.Equals("A6"))
                return isLandscape ? new FormatSize(148, 105) : new FormatSize(105, 148);

            if (format.Equals("B0"))
                return isLandscape ? new FormatSize(1414, 1000) : new FormatSize(1000, 1414);

            if (format.Equals("B1"))
                return isLandscape ? new FormatSize(1000, 707) : new FormatSize(707, 1000);

            if (format.Equals("B2"))
                return isLandscape ? new FormatSize(707, 500) : new FormatSize(500, 707);

            if (format.Equals("B3"))
                return isLandscape ? new FormatSize(500, 353) : new FormatSize(353, 500);

            if (format.Equals("B4"))
                return isLandscape ? new FormatSize(353, 250) : new FormatSize(250, 353);

            if (format.Equals("B5"))
                return isLandscape ? new FormatSize(250, 176) : new FormatSize(176, 250);

            if (format.Equals("B6"))
                return isLandscape ? new FormatSize(250, 176) : new FormatSize(176, 250);

            if (format.Equals("C0"))
                return isLandscape ? new FormatSize(1297, 917) : new FormatSize(917, 1297);

            if (format.Equals("C1"))
                return isLandscape ? new FormatSize(917, 648) : new FormatSize(648, 917);

            if (format.Equals("C2"))
                return isLandscape ? new FormatSize(648, 458) : new FormatSize(458, 648);

            if (format.Equals("C3"))
                return isLandscape ? new FormatSize(458, 324) : new FormatSize(324, 458);

            if (format.Equals("C4"))
                return isLandscape ? new FormatSize(324, 229) : new FormatSize(229, 324);

            if (format.Equals("C5"))
                return isLandscape ? new FormatSize(229, 162) : new FormatSize(162, 229);

            throw new ArgumentOutOfRangeException(nameof(format), "Can't get format size");
        }
    }
}