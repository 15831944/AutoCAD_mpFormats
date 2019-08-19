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
                _mpFormats.Closed += win_Closed;
            }

            if (_mpFormats.IsLoaded)
                _mpFormats.Activate();
            else
                Application.ShowModelessWindow(Application.MainWindow.Handle, _mpFormats, false);
        }

        private void win_Closed(object sender, EventArgs e)
        {
            _mpFormats = null;
        }

        public static bool DrawBlock
        (
            string format, // формат
            string multiplicity, // кратность
            string side, // Сторона кратности
            string orientation, // Ориентация
            bool number, // Номер страницы (да, нет)

            bool copy, // Копировал
            string bottomFrame, // Нижняя рамка
            bool hasFpt, // Есть ли начальная точка
            Point3d insertPt, // Начальная точка (для замены)
            string txtStyle, // TextStyle name
            double scale, // масштаб
            double? rotation,
            out Point3d bottomLeftPt,
            out Point3d topLeftPt,
            out Point3d bottomRightPt,
            out Vector3d replaceVector3D,
            out Point3d blockInsertionPoint
        )
        {
            bottomLeftPt = bottomRightPt = topLeftPt = blockInsertionPoint = new Point3d(0.0, 0.0, 0.0);
            replaceVector3D = new Vector3d(0.0, 0.0, 0.0);
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var returned = false;

            double dlina = 0.0, visota = 0.0;
            try
            {
                using (doc.LockDocument())
                {
                    // Задаем значение ширины и высоты в зависимости формата, кратности и стороны кратности

                    #region format size

                    //if (format.Equals("A0"))
                    //{
                    //    if (int.Parse(multiplicity) > 1)
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 1189;
                    //                visota = 841 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 841;
                    //                visota = 1189 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 841 * int.Parse(multiplicity);
                    //                visota = 1189;
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 1189 * int.Parse(multiplicity);
                    //                visota = 841;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            dlina = 841;
                    //            visota = 1189;
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            dlina = 1189;
                    //            visota = 841;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A1"))
                    //{
                    //    if (int.Parse(multiplicity) > 1)
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 841;
                    //                visota = 594 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 594;
                    //                visota = 841 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 594 * int.Parse(multiplicity);
                    //                visota = 841;
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 841 * int.Parse(multiplicity);
                    //                visota = 594;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            dlina = 594;
                    //            visota = 841;
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            dlina = 841;
                    //            visota = 594;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A2"))
                    //{
                    //    if (int.Parse(multiplicity) > 1)
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 594;
                    //                visota = 420 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 420;
                    //                visota = 594 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 420 * int.Parse(multiplicity);
                    //                visota = 594;
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 594 * int.Parse(multiplicity);
                    //                visota = 420;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            dlina = 420;
                    //            visota = 594;
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            dlina = 594;
                    //            visota = 420;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A3"))
                    //{
                    //    if (int.Parse(multiplicity) > 1)
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 420;
                    //                visota = 297 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 297;
                    //                visota = 420 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 297 * int.Parse(multiplicity);
                    //                visota = 420;
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 420 * int.Parse(multiplicity);
                    //                visota = 297;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            dlina = 297;
                    //            visota = 420;
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            dlina = 420;
                    //            visota = 297;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A4"))
                    //{
                    //    if (int.Parse(multiplicity) > 1)
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 297;
                    //                visota = 210 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 210;
                    //                visota = 297 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            if (side.Equals(Language.GetItem(LangItem, "h11")))
                    //            {
                    //                dlina = 210 * int.Parse(multiplicity);
                    //                visota = 297;
                    //            }
                    //            if (side.Equals(Language.GetItem(LangItem, "h12")))
                    //            {
                    //                dlina = 297 * int.Parse(multiplicity);
                    //                visota = 210;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //        {
                    //            dlina = 210;
                    //            visota = 297;
                    //        }
                    //        if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //        {
                    //            dlina = 297;
                    //            visota = 210;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A5"))
                    //{
                    //    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    //    {
                    //        dlina = 148;
                    //        visota = 210;
                    //    }
                    //    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    //    {
                    //        dlina = 210;
                    //        visota = 148;
                    //    }
                    //}

                    #endregion

                    GetFormatSize(format, orientation, side, multiplicity, out dlina, out visota);

                    #region points

                    var pt1 = new Point3d(0.0, 0.0, 0.0);
                    var pt2 = new Point3d(0.0 + dlina, 0.0, 0.0);
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

                    var pt3 = new Point3d(0.0 + dlina, 0.0 + visota, 0.0);
                    var pt33 = new Point3d(pt3.X - 5, pt3.Y - 5, 0.0);
                    var pt4 = new Point3d(0.0, 0.0 + visota, 0.0);
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

                    var isnumber = number ? "N" : "NN";
                    var iscopy = copy ? "C" : "NC";
                    string blockname;
                    if (!multiplicity.Equals("1"))
                        blockname = format + "x" + multiplicity + "_" + orientation + "_"
                                    + side + "_" + isnumber + "_" + iscopy;
                    else
                        blockname = format + "_" + orientation + "_" + side + "_" + isnumber
                                    + "_" + iscopy;
                    if (format.Equals("A4") || format.Equals("A3"))
                        blockname = blockname + "_" + bottomFrame;

                    #endregion

                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        // Если в базе есть такой блок - вставляем его
                        if (bt.Has(blockname))
                        {
                            var blockId = bt[blockname];
                            var btr =
                                (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                            if (!hasFpt) // Если нет начальной точки, то рисуем через джигу
                            {
                                var pt = new Point3d(0, 0, 0);
                                var br = new BlockReference(pt, blockId);
                                // scale
                                var mat = Matrix3d.Scaling(scale, pt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, pt);
                                    br.TransformBy(rm);
                                }
                                //==================

                                var entJig = new BlockJig(br);
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    blockInsertionPoint = br.Position;
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            } //
                            else // Если есть начальная точка - то вставлем в нее
                            {
                                var br = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = br.Position;
                                // scale
                                var mat = Matrix3d.Scaling(scale, insertPt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, insertPt);
                                    br.TransformBy(rm);
                                }
                                //==================

                                btr.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);
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
                                    blockname, false);
                            }
                            catch
                            {
                                ModPlusAPI.Windows.MessageBox.Show(Language.GetItem(LangItem, "err14"));
                            }
                            var btr = new BlockTableRecord { Name = blockname };

                            ////////////////////////////////////////////
                            // Add the new block to the block table
                            bt.UpgradeOpen();
                            bt.Add(btr);
                            tr.AddNewlyCreatedDBObject(btr, true);
                            //*******************************

                            // Рисуем примитивы и добавляем в блок
                            var ents = new DBObjectCollection();
                            // внешняя рамка
                            var pline1 = new Polyline
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
                                pline1.AddVertexAt(i, pp, 0, 0, 0);
                            }
                            // внутренняя рамка
                            var pline2 = new Polyline
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
                                pline2.AddVertexAt(i, pp, 0, 0, 0);
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
                                ents.Add(txt2);
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
                                ents.Add(line1);
                                ents.Add(line2);
                            }

                            ents.Add(pline1);
                            ents.Add(pline2);
                            ents.Add(txt1);

                            foreach (Entity ent in ents)
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

                            var blockId = bt[blockname];

                            var cbtr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                            if (!hasFpt) // Если начальной точки нет - рисуем через джигу
                            {
                                var pt = new Point3d(0, 0, 0);
                                var br = new BlockReference(pt, blockId);
                                // scale
                                var mat = Matrix3d.Scaling(scale, pt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, pt);
                                    br.TransformBy(rm);
                                }
                                //==================

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                br.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));

                                var entJig = new BlockJig(br);
                                // Perform the jig operation
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    blockInsertionPoint = br.Position;
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
                            else // Если начальная точка есть - вставляем в нее
                            {
                                var br = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = br.Position;
                                // scale
                                var mat = Matrix3d.Scaling(scale, insertPt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, insertPt);
                                    br.TransformBy(rm);
                                }
                                //==================

                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                br.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));
                                cbtr.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        } // else
                        tr.Commit();
                    } // tr
                }
                return returned;
            } // try
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
                return false;
            }
        }

        public static bool DrawBlockHand
        (
            double dlina,
            double visota,
            bool number, // Номер страницы (да, нет)

            bool copy, // Копировал
            bool hasFpt, // Есть ли начальная точка
            Point3d insertPt, // Начальная точка (для замены)
            string txtStyle,
            double scale, // масштаб
            double? rotation,
            out Point3d bottomLeftPt,
            out Point3d topLeftPt,
            out Point3d bottomRightPt,
            out Vector3d replaceVector3D,
            out Point3d blockInsertionPoint
        )
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
                    var isnumber = number ? "N" : "NN";
                    var iscopy = copy ? "C" : "NC";
                    var blockname = dlina.ToString(CultureInfo.InvariantCulture) + "x" +
                                    visota.ToString(CultureInfo.InvariantCulture) + "_" +
                                    isnumber + "_" + iscopy;
                    #endregion

                    #region points
                    var pt1 = new Point3d(0.0, 0.0, 0.0);
                    var pt2 = new Point3d(0.0 + dlina, 0.0, 0.0);
                    var pt11 = new Point3d(pt1.X + 20, pt1.Y + 5, 0.0);
                    var pt22 = new Point3d(pt2.X - 5, pt2.Y + 5, 0.0);
                    var pt3 = new Point3d(0.0 + dlina, 0.0 + visota, 0.0);
                    var pt33 = new Point3d(pt3.X - 5, pt3.Y - 5, 0.0);
                    var pt4 = new Point3d(0.0, 0.0 + visota, 0.0);
                    var pt44 = new Point3d(pt4.X + 20, pt4.Y - 5, 0.0);
                    var ptt1 = new Point3d(pt2.X - 55, pt22.Y - 4, 0.0);
                    var ptt2 = new Point3d(pt2.X - 125, pt22.Y - 4, 0.0);
                    // points for stamps
                    bottomLeftPt = pt11;
                    topLeftPt = pt44;
                    bottomRightPt = pt22;
                    //
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
                        if (bt.Has(blockname))
                        {
                            var blockId = bt[blockname];
                            var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                            if (!hasFpt) // Если нет начальной точки, то рисуем через джигу
                            {
                                var pt = new Point3d(0, 0, 0);
                                var br = new BlockReference(pt, blockId);
                                // scale
                                var mat = Matrix3d.Scaling(scale, pt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, pt);
                                    br.TransformBy(rm);
                                }
                                //==================
                                var entJig = new BlockJig(br);
                                // Perform the jig operation
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    blockInsertionPoint = br.Position;
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                    doc.TransactionManager.QueueForGraphicsFlush();
                                    returned = true;
                                }
                            } //
                            else // Если есть начальная точка - то вставлем в нее
                            {
                                var br = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = br.Position;
                                // scale
                                var mat = Matrix3d.Scaling(scale, insertPt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, insertPt);
                                    br.TransformBy(rm);
                                }
                                //==================
                                btr.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        }
                        // Если блока нет - создаем и вставляем
                        else
                        {
                            try
                            {
                                SymbolUtilityServices.ValidateSymbolName(blockname, false);
                            }
                            catch
                            {
                                ModPlusAPI.Windows.MessageBox.Show(Language.GetItem(LangItem, "err14"));
                            }

                            var btr = new BlockTableRecord { Name = blockname };
                            // Add the new block to the block table
                            bt.UpgradeOpen();
                            bt.Add(btr);
                            tr.AddNewlyCreatedDBObject(btr, true);
                            //*******************************

                            // Рисуем примитивы и добавляем в блок
                            var ents = new DBObjectCollection();

                            // внешняя рамка
                            var pline1 = new Polyline
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
                                pline1.AddVertexAt(i, pp, 0, 0, 0);
                            }
                            // внутренняя рамка
                            var pline2 = new Polyline
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
                                pline2.AddVertexAt(i, pp, 0, 0, 0);
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
                                ents.Add(txt2);
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
                                ents.Add(line1);
                                ents.Add(line2);
                            }
                            ents.Add(pline1);
                            ents.Add(pline2);
                            ents.Add(txt1);

                            foreach (Entity ent in ents)
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

                            var blockId = bt[blockname];

                            var cbtr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, false);
                            if (!hasFpt) // Если начальной точки нет - рисуем через джигу
                            {
                                var pt = new Point3d(0, 0, 0);
                                var br = new BlockReference(pt, blockId);
                                blockInsertionPoint = br.Position;
                                // scale
                                var mat = Matrix3d.Scaling(scale, pt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, pt);
                                    br.TransformBy(rm);
                                }
                                //==================
                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                br.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));

                                var entJig = new BlockJig(br);
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
                            else // Если начальная точка есть - вставляем в нее
                            {
                                var br = new BlockReference(insertPt, blockId);
                                blockInsertionPoint = br.Position;
                                // scale
                                var mat = Matrix3d.Scaling(scale, insertPt);
                                br.TransformBy(mat);
                                // rotation
                                if (rotation != null)
                                {
                                    var rm = Matrix3d.Rotation(rotation.Value, br.Normal, insertPt);
                                    br.TransformBy(rm);
                                }
                                //==================
                                ModPlus.Helpers.XDataHelpers.AddRegAppTableRecord("MP_FORMAT");
                                br.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));
                                cbtr.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);
                                doc.TransactionManager.QueueForGraphicsFlush();
                                returned = true;
                            }
                        } // else
                        tr.Commit();
                    } // tr
                }
                return returned;
            } // try
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
                return false;
            }
        }

        public static void ReplaceBlock
        (
            string format, // формат
            string multiplicity, // кратность
            string side, // Сторона кратности
            string orientation, // Ориентация
            bool number, // Номер страницы (да, нет)
            bool copy, // Копировал
            string bottomFrame, // Нижняя рамка
            string txtStyle,
            double scale
        )
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
                if (per.Status != PromptStatus.OK) return;
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
                                    Point3d bottomLeftPt;
                                    Point3d topLeftPt;
                                    Point3d bottomRightPt;
                                    Vector3d replaceVector3D;
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
                                        out bottomLeftPt,
                                        out topLeftPt,
                                        out bottomRightPt,
                                        out replaceVector3D,
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

        public static void ReplaceBlockHand
        (
            double dlina,
            double visota,
            bool number, // Номер страницы (да, нет)
            bool copy, // Копировал
            string txtStyle,
            double scale
        )
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

                                    Point3d bottomLeftPt;
                                    Point3d topLeftPt;
                                    Point3d bottomRightPt;
                                    Vector3d replaceVector3D;
                                    DrawBlockHand(
                                        dlina,
                                        visota,
                                        number,
                                        copy,
                                        true,
                                        pt,
                                        txtStyle,
                                        scale,
                                        rotation,
                                        out bottomLeftPt,
                                        out topLeftPt,
                                        out bottomRightPt,
                                        out replaceVector3D,
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

        public static void GetFormatSize(
            string format,
            string orientation,
            string side,
            string multiplicity,
            out double dlina,
            out double visota
        )
        {
            dlina = visota = 0;

            if (format.Equals("A0"))
            {
                if (int.Parse(multiplicity) > 1)
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 1189;
                            visota = 841 * int.Parse(multiplicity);
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 841;
                            visota = 1189 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 841 * int.Parse(multiplicity);
                            visota = 1189;
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 1189 * int.Parse(multiplicity);
                            visota = 841;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        dlina = 841;
                        visota = 1189;
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        dlina = 1189;
                        visota = 841;
                    }
                }
            }
            if (format.Equals("A1"))
            {
                if (int.Parse(multiplicity) > 1)
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 841;
                            visota = 594 * int.Parse(multiplicity);
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 594;
                            visota = 841 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 594 * int.Parse(multiplicity);
                            visota = 841;
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 841 * int.Parse(multiplicity);
                            visota = 594;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        dlina = 594;
                        visota = 841;
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        dlina = 841;
                        visota = 594;
                    }
                }
            }
            if (format.Equals("A2"))
            {
                if (int.Parse(multiplicity) > 1)
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 594;
                            visota = 420 * int.Parse(multiplicity);
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 420;
                            visota = 594 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 420 * int.Parse(multiplicity);
                            visota = 594;
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 594 * int.Parse(multiplicity);
                            visota = 420;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        dlina = 420;
                        visota = 594;
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        dlina = 594;
                        visota = 420;
                    }
                }
            }
            if (format.Equals("A3"))
            {
                if (int.Parse(multiplicity) > 1)
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 420;
                            visota = 297 * int.Parse(multiplicity);
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 297;
                            visota = 420 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 297 * int.Parse(multiplicity);
                            visota = 420;
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 420 * int.Parse(multiplicity);
                            visota = 297;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        dlina = 297;
                        visota = 420;
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        dlina = 420;
                        visota = 297;
                    }
                }
            }
            if (format.Equals("A4"))
            {
                if (int.Parse(multiplicity) > 1)
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 297;
                            visota = 210 * int.Parse(multiplicity);
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 210;
                            visota = 297 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        if (side.Equals(Language.GetItem(LangItem, "h11")))
                        {
                            dlina = 210 * int.Parse(multiplicity);
                            visota = 297;
                        }
                        if (side.Equals(Language.GetItem(LangItem, "h12")))
                        {
                            dlina = 297 * int.Parse(multiplicity);
                            visota = 210;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                    {
                        dlina = 210;
                        visota = 297;
                    }
                    if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                    {
                        dlina = 297;
                        visota = 210;
                    }
                }
            }
            if (format.Equals("A5"))
            {
                if (orientation.Equals(Language.GetItem(LangItem, "h9")))
                {
                    dlina = 148;
                    visota = 210;
                }
                if (orientation.Equals(Language.GetItem(LangItem, "h8")))
                {
                    dlina = 210;
                    visota = 148;
                }
            }
        }
    }
}