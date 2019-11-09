namespace mpFormats
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media.Imaging;
    using System.Xml.Linq;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Internal;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Windows;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using MessageBox = ModPlusAPI.Windows.MessageBox;
    using Visibility = System.Windows.Visibility;

    public partial class MpFormats
    {
        private const string LangItem = "mpFormats";
        public static int Namecol;// Количество допустимых фамилий
        public static int Row;// Строка начала заполнения должностей
        public static int Col;// Столбец начала заполнения должностей
        public static bool HasSureNames;// Есть ли должности

        // Текущий документ с таблицами
        private XElement _xmlTblsDoc;

        // Текущая таблица в виде XElement
        private XElement _currentTblXml;

        // Переменные для хранения координат вставки имени и номера в штамп
        // выносим отдельно, т.к. окно открывается в другом методе
        private static string _lnamecoord = string.Empty;
        private static string _lnumbercoord = string.Empty;
        private static string _lname = string.Empty;
        private static string _lnumber = string.Empty;

        public MpFormats()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h4");

            // Настройки видимости для штампа отключаем тут, чтобы видеть в редакторе окна
            GridStamp.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Collapsed;
        }

        #region window basic

        private void MetroWindow_MouseEnter(object sender, MouseEventArgs e)
        {
            Focus();
        }

        private void MetroWindow_MouseLeave(object sender, MouseEventArgs e)
        {
            Utils.SetFocusToDwgView();
        }

        private void Tb_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Запрет нажатия пробела
            if (e.Key == Key.Space)
                e.Handled = true;
        }

        private void _PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Ввод только цифр и точки
            short val;
            if (!short.TryParse(e.Text, out val))
                e.Handled = true;
        }
        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;

                _xmlTblsDoc = XElement.Parse(GetResourceTextFile("Stamps.xml"));

                // Заполнение списка масштабов
                var ocm = db.ObjectContextManager;
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                foreach (var ans in occ.Cast<AnnotationScale>())
                {
                    CbScales.Items.Add(ans.Name);
                }

                // Начальное значение масштаба
                var cans = occ.CurrentContext as AnnotationScale;
                if (cans != null)
                    CbScales.SelectedItem = cans.Name;

                // Заполняем список текстовых стилей
                string txtstname;
                using (var acTrans = doc.TransactionManager.StartTransaction())
                {
                    var tst = (TextStyleTable)acTrans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                    foreach (var id in tst)
                    {
                        var tstr = (TextStyleTableRecord)acTrans.GetObject(id, OpenMode.ForRead);
                        if (!string.IsNullOrEmpty(tstr.Name))
                            CbTextStyle.Items.Add(tstr.Name);
                    }

                    var curtxt = (TextStyleTableRecord)acTrans.GetObject(db.Textstyle, OpenMode.ForRead);
                    if (CbTextStyle.Items.Contains(curtxt.Name))
                        txtstname = curtxt.Name;
                    else
                        txtstname = CbTextStyle.Items[0].ToString();
                    acTrans.Commit();
                }

                CbTextStyle.SelectedItem = txtstname;

                // Логотип
                CbLogo.Items.Clear();
                CbLogo.Items.Add(ModPlusAPI.Language.GetItem(LangItem, "no"));
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    using (doc.LockDocument())
                    {
                        var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                        foreach (var id in bt)
                        {
                            var btRecord = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if (!btRecord.IsLayout & !btRecord.IsAnonymous)
                            {
                                if (!CbLogo.Items.Contains(btRecord.Name))
                                    CbLogo.Items.Add(btRecord.Name);
                            }
                        }
                    }
                }

                CbLogo.SelectedIndex = 0;

                // Загрузка из настроек
                LoadFromSettings();

                // Проверка файла со штампами
                if (!CheckTableFileExist())
                {
                    MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err5"), MessageBoxIcon.Alert);

                    // Видимость
                    GridStamp.Visibility =
                        CbLogo.Visibility =
                        ChkStamp.Visibility = ChkB1.Visibility = ChkB2.Visibility = ChkB3.Visibility = GridSplitterStamp.Visibility =
                        Image_stamp.Visibility = Image_b1.Visibility = Image_b2.Visibility = Image_top.Visibility =
                        Visibility.Collapsed;
                    ChkStamp.IsChecked = ChkB1.IsChecked = ChkB2.IsChecked = ChkB3.IsChecked = false;
                }

                // Включаем обработчики событий
                CbMultiplicity.SelectionChanged += CbMultiplicity_OnSelectionChanged;
                RbVertical.Checked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbVertical.Unchecked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbHorizontal.Checked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbHorizontal.Unchecked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbLong.Checked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbLong.Unchecked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbShort.Checked += RadioButton_FormatSettings_OnChecked_OnUnchecked;
                RbShort.Unchecked += RadioButton_FormatSettings_OnChecked_OnUnchecked;

                // show format size
                ShowFormatSize();
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            SaveToSettings();
        }

        // Выбор базы таблиц
        private void CbDocumentsFor_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (sender is ComboBox cb && cb.SelectedItem is ComboBoxItem && cb.SelectedIndex != -1)
                {
                    UserConfigFile.SetValue("mpFormats", "CbDocumentsFor",
                        cb.SelectedIndex.ToString(CultureInfo.InvariantCulture), true);

                    switch (cb.SelectedIndex)
                    {
                        case 0: // RU
                            {
                            }

                            break;
                        case 1: // UA
                            {
                            }

                            break;
                        case 2: // BY
                            {
                            }

                            break;
                    }

                    FillStamps();
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        private static bool CheckTableFileExist()
        {
            var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
            using (key)
            {
                if (key != null)
                {
                    // Директория расположения файла
                    var dir = Path.Combine(key.GetValue("TopDir").ToString(), "Data", "Dwg");

                    // Имя файла из которого берем таблицу
                    var sourceFileName = Path.Combine(dir, "Stamps.dwg");
                    return File.Exists(sourceFileName);
                }
            }

            return false;
        }

        // Загрузка из настроек
        private void LoadFromSettings()
        {
            var li = ModPlusConnector.Instance.Name;

            CbDocumentsFor.SelectedIndex =
                    int.TryParse(UserConfigFile.GetValue(li, "CbDocumentsFor"), out int index)
                        ? index
                        : 0;

            // format
            CbFormat.SelectedIndex = int.TryParse(UserConfigFile.GetValue(li, "CbFormat"), out int i) ? i : 3;
            CbMultiplicity.SelectedIndex = int.TryParse(UserConfigFile.GetValue(li, "CbMultiplicity"), out i) ? i : 0;
            CbBottomFrame.SelectedIndex = int.TryParse(UserConfigFile.GetValue(li, "CbBottomFrame"), out i) ? i : 0;

            // Выбранный штамп
            CbTables.SelectedIndex = int.TryParse(UserConfigFile.GetValue(li, "CbTables"), out i) ? i : 0;

            // Масштаб
            var scale = UserConfigFile.GetValue(li, "CbScales");
            CbScales.SelectedIndex = CbScales.Items.Contains(scale)
                ? CbScales.Items.IndexOf(scale)
                : 0;
            ChkB1.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChkB1"), out var b) && b;
            ChkB2.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChkB2"), out b) && b;
            ChkB3.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChkB3"), out b) && b;
            ChbCopy.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChbCopy"), out b) && b;
            ChbNumber.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChbNumber"), out b) && b;
            ChkStamp.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChkStamp"), out b) && b;
            RbVertical.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "RbVertical"), out b) && b;
            RbLong.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "RbLong"), out b) && b;

            // set current layer
            ChkSetCurrentLayer.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "SetCurrentLayer"), out b) && b;

            TbFormatHeight.Value = int.TryParse(UserConfigFile.GetValue(li, "TbFormatHeight"), out i)
                ? i : 10;
            TbFormatLength.Value = int.TryParse(UserConfigFile.GetValue(li, "TbFormatLength"), out i)
                ? i : 10;

            // Текстовый стиль (меняем, если есть в настройках, а иначе оставляем текущий)
            var textStyle = UserConfigFile.GetValue(li, "CbTextStyle");
            if (CbTextStyle.Items.Contains(textStyle))
                CbTextStyle.SelectedIndex = CbTextStyle.Items.IndexOf(textStyle);

            // Логотип
            var logo = UserConfigFile.GetValue(li, "CbLogo");
            CbLogo.SelectedIndex = CbLogo.Items.Contains(logo) ? CbLogo.Items.IndexOf(logo) : 0;

            // Поля
            ChbHasFields.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "ChbHasFields"), out b) && b;

            // Высота текста
            TbMainTextHeight.Text = double.TryParse(
                UserConfigFile.GetValue(li, "MainTextHeight"),
                out var d)
                ? d.ToString(CultureInfo.InvariantCulture)
                : "2.5";
            TbBigTextHeight.Text = double.TryParse(
                UserConfigFile.GetValue(li, "BigTextHeight"),
                out d)
                ? d.ToString(CultureInfo.InvariantCulture)
                : "3.5";

            // logo from
            ChkLogoFromBlock.IsChecked = !bool.TryParse(UserConfigFile.GetValue(li, "LogoFromBlock"), out b) || b;
            ChkLogoFromFile.IsChecked = bool.TryParse(UserConfigFile.GetValue(li, "LogoFromFile"), out b) && b;

            // logo file
            var ffs = UserConfigFile.GetValue(li, "LogoFile");
            if (!string.IsNullOrEmpty(ffs))
            {
                if (File.Exists(ffs))
                    TbLogoFile.Text = ffs;
            }
        }

        // Сохранение в настройки
        private void SaveToSettings()
        {
            try
            {
                var li = ModPlusConnector.Instance.Name;
                UserConfigFile.SetValue(li, "CbFormat", CbFormat.SelectedIndex.ToString(), false);
                UserConfigFile.SetValue(li, "CbMultiplicity", CbMultiplicity.SelectedIndex.ToString(),
                    false);
                UserConfigFile.SetValue(li, "CbBottomFrame", CbBottomFrame.SelectedIndex.ToString(),
                    false);

                UserConfigFile.SetValue(li, "ChkB1",
                    (ChkB1.IsChecked != null && ChkB1.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "ChkB2",
                    (ChkB2.IsChecked != null && ChkB2.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "ChkB3",
                    (ChkB3.IsChecked != null && ChkB3.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "ChbCopy",
                    (ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "ChbNumber",
                    (ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "ChkStamp",
                    (ChkStamp.IsChecked != null && ChkStamp.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "RbVertical",
                    (RbVertical.IsChecked != null && RbVertical.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "RbLong",
                    (RbLong.IsChecked != null && RbLong.IsChecked.Value).ToString(), false);

                UserConfigFile.SetValue(li, "TbFormatHeight", TbFormatHeight.Value.ToString(), false);
                UserConfigFile.SetValue(li, "TbFormatLength", TbFormatLength.Value.ToString(), false);

                // Текстовый стиль
                UserConfigFile.SetValue(li, "CbTextStyle", CbTextStyle.SelectedItem.ToString(), false);

                // Логотип
                UserConfigFile.SetValue(li, "CbLogo", CbLogo.SelectedItem.ToString(), false);

                // Поля
                UserConfigFile.SetValue(li, "ChbHasFields",
                    (ChbHasFields.IsChecked != null && ChbHasFields.IsChecked.Value).ToString(), false);

                // Выбранный штамп
                UserConfigFile.SetValue(li, "CbTables",
                    CbTables.SelectedIndex.ToString(CultureInfo.InvariantCulture), false);

                // Масштаб
                UserConfigFile.SetValue(li, "CbScales",
                                        CbScales.SelectedItem.ToString(), false);
                UserConfigFile.SetValue(li, "MainTextHeight", TbMainTextHeight.Text, false);
                UserConfigFile.SetValue(li, "BigTextHeight", TbBigTextHeight.Text, false);
                UserConfigFile.SetValue(li, "LogoFromBlock", (ChkLogoFromBlock.IsChecked != null && ChkLogoFromBlock.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "LogoFromFile", (ChkLogoFromFile.IsChecked != null && ChkLogoFromFile.IsChecked.Value).ToString(), false);
                UserConfigFile.SetValue(li, "LogoFile",
                    File.Exists(TbLogoFile.Text) ? TbLogoFile.Text : string.Empty, false);

                // set current layer
                UserConfigFile.SetValue(li, "SetCurrentLayer", (ChkSetCurrentLayer.IsChecked != null && ChkSetCurrentLayer.IsChecked.Value).ToString(), false);

                UserConfigFile.SaveConfigFile();
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static double Scale(string scaleName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ocm = db.ObjectContextManager;
            var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
            var ansc = occ.GetContext(scaleName) as AnnotationScale;
            return ansc.DrawingUnits / ansc.PaperUnits;
        }

        private void FillSurenames()
        {
            // Очистка списков
            LbStampSurnames.Items.Clear();
            LbSurnames.Items.Clear();

            // Все должности в список должностей. В список должностей штампа - ничего
            foreach (var s in _currentTblXml.Attribute("surnames").Value.Split(','))
            {
                LbSurnames.Items.Add(s);
            }

            // Заполняем список пользовательских должностей
            if (UserConfigFile.ConfigFileXml.Element("Settings") != null)
            {
                var setXml = UserConfigFile.ConfigFileXml.Element("Settings");
                if (setXml?.Element("UserSurnames") != null)
                {
                    var element = setXml.Element("UserSurnames");
                    if (element != null)
                    {
                        foreach (var sn in element.Elements("Surname"))
                        {
                            LbSurnames.Items.Add(sn.Attribute("Surname").Value);
                        }
                    }
                }
            }

            // Заполняем значения из файла настроек. Про совпадении в "левом" списке - удаляем
            if (UserConfigFile.ConfigFileXml.Element("Settings") != null)
            {
                var setXml = UserConfigFile.ConfigFileXml.Element("Settings");
                if (setXml?.Element("mpStampTblSaves") != null)
                {
                    var element = setXml.Element("mpStampTblSaves");
                    if (element?.Attribute(_currentTblXml.Attribute("tablestylename").Value) != null)
                    {
                        LbStampSurnames.Items.Clear();
                        var xElement = setXml.Element("mpStampTblSaves");
                        if (xElement != null)
                        {
                            foreach (var item in xElement.Attribute(_currentTblXml.Attribute("tablestylename").Value).Value.Split('$'))
                            {
                                if (LbSurnames.Items.Contains(item))
                                    LbSurnames.Items.Remove(item);
                                LbStampSurnames.Items.Add(item);
                            }
                        }
                    }
                    else
                    {
                        FillStampSurnamesEmpty();
                    }
                }
                else
                {
                    FillStampSurnamesEmpty();
                }
            }
            else
            {
                FillStampSurnamesEmpty();
            }
        }

        private void FillStampSurnamesEmpty()
        {
            // Список должностей в штампе заполняем пустыми значениями
            for (var i = 0; i < Namecol; i++)
            {
                LbStampSurnames.Items.Add(string.Empty);
            }
        }

        private void FillStamps()
        {
            // Очищаем список таблиц
            CbTables.ItemsSource = null;
            try
            {
                var docFor = ((ComboBoxItem)CbDocumentsFor.SelectedItem).Tag;
                var stamps = new List<Stamp>();
                foreach (var tbl in _xmlTblsDoc.Elements("Stamp"))
                {
                    var docForAttr = tbl.Attribute("DocFor");
                    if (docForAttr != null && tbl.Attribute("noShow") == null &
                        docForAttr.Value.Equals(docFor))
                    {
                        stamps.Add(new Stamp
                        {
                            Name = tbl.Attribute("name")?.Value,
                            TableStyle = tbl.Attribute("tablestylename")?.Value,
                            Description = tbl.Attribute("document")?.Value + " " + tbl.Attribute("description")?.Value
                        });
                    }
                }

                CbTables.ItemsSource = stamps;

                // Устанавливаем первое значение
                CbTables.SelectedIndex = 0;
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        // Выбор таблицы
        private void CbTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // Если не пустой список
                if (CbTables.Items.Count > 0)
                {
                    var stamp = CbTables.SelectedItem as Stamp;

                    // Изменяем значения элементов формы в зависимости от выбранной таблицы
                    foreach (var tbl in _xmlTblsDoc.Elements("Stamp"))
                    {
                        if (stamp != null)
                        {
                            if (stamp.Name.Equals(tbl.Attribute("name").Value) &
                                stamp.TableStyle.Equals(tbl.Attribute("tablestylename").Value))
                            {
                                _currentTblXml = tbl;

                                // Изображение
                                try
                                {
                                    var uriSource = new Uri(
                                        @"/mpFormats_" + ModPlusConnector.Instance.AvailProductExternalVersion + ";component/Resources/Preview/" +
                                        tbl.Attribute("img")?.Value + ".png", UriKind.Relative);
                                    Image_stamp.Source = new BitmapImage(uriSource);
                                }
                                catch
                                {
                                    // ignore
                                }

                                // Если можно использовать поля
                                var hasFieldsAttr = tbl.Attribute("hasfields");
                                if (hasFieldsAttr != null && hasFieldsAttr.Value.Equals("true"))
                                {
                                    ChbHasFields.IsEnabled = true;
                                    ChbHasFields.IsChecked = true;
                                    BtFields.IsEnabled = true;
                                }
                                else
                                {
                                    ChbHasFields.IsEnabled = false;
                                    ChbHasFields.IsChecked = false;
                                    BtFields.IsEnabled = false;
                                }

                                var hasSurenamesAttr = tbl.Attribute("hassurenames");
                                if (hasSurenamesAttr != null)
                                    HasSureNames = hasSurenamesAttr.Value.Equals("true");

                                // Если есть фамилии
                                if (HasSureNames)
                                {
                                    LbStampSurnames.Visibility = Visibility.Visible;
                                    LbSurnames.Visibility = Visibility.Visible;

                                    // Максимальное количество имен
                                    var sureNamesColAttr = tbl.Attribute("surnamescol");
                                    if (sureNamesColAttr != null)
                                        Namecol = int.Parse(sureNamesColAttr.Value);

                                    // Заполняем список должностей в отдельной функции
                                    FillSurenames();

                                    // Строка и столбец
                                    Row = int.Parse(tbl.Attribute("firstcell").Value.Split(',').GetValue(1).ToString());
                                    Col = int.Parse(tbl.Attribute("firstcell").Value.Split(',').GetValue(0).ToString());
                                }
                                else
                                {
                                    // Очистка списков
                                    LbStampSurnames.Items.Clear();
                                    LbSurnames.Items.Clear();
                                    LbStampSurnames.Visibility = Visibility.Hidden;
                                    LbSurnames.Visibility = Visibility.Hidden;
                                }

                                // Логотип
                                CbLogo.IsEnabled = BtGetFileForLogo.IsEnabled = ChkLogoFromFile.IsEnabled = ChkLogoFromBlock.IsEnabled = bool.Parse(tbl.Attribute("logo").Value);

                                // Большой текст
                                TbBigTextHeight.IsEnabled = tbl.Attribute("bigCells") != null;

                                // Координаты вставки номера и имени
                                var lnameattr = tbl.Attribute("lnamecoordinates");
                                _lnamecoord = lnameattr?.Value ?? string.Empty;
                                var lnumberattr = tbl.Attribute("lnumbercooridnates");
                                _lnumbercoord = lnumberattr?.Value ?? string.Empty;
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        // Вызов функции "Поля"
        private void BtFields_Click(object sender, RoutedEventArgs e)
        {
            // Проверка полной версии
            if (!Registration.IsFunctionBought("mpStamps", ModPlusConnector.Instance.AvailProductExternalVersion))
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg4"));
            }
            else
            {
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_MPSTAMPFIELDS ", true, false, false);
            }
        }

        // Вызов диалогового окна "Редактирование пользовательских фамилий"
        private void BtAddUserSurname_OnClick(object sender, RoutedEventArgs e)
        {
            // Проверка полной версии
            if (!Registration.IsFunctionBought("mpStamps", ModPlusConnector.Instance.AvailProductExternalVersion))
            {
                MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg4"));
            }
            else
            {
                var window = new EditUserSurnames
                {
                    Topmost = true
                };
                window.ShowDialog();

                // Перезаполняем список фамилий
                FillSurenames();
            }
        }

        // Действие, если выбран элемент в списке должностей
        private void LbSurnames_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BtAddSurname.IsEnabled = LbSurnames.SelectedIndex != -1;
        }

        // Действие, если выбран элемент в списке должностей штампа
        private void LbStampSurnames_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LbStampSurnames.SelectedIndex != -1)
            {
                BtRemoveSurname.IsEnabled = true;
                if (LbStampSurnames.SelectedIndex >= 0)
                {
                    BtUpSurname.IsEnabled = true;
                    BtDownSurename.IsEnabled = true;
                }
                else
                {
                    BtUpSurname.IsEnabled = false;
                    BtDownSurename.IsEnabled = false;
                }
            }
            else
            {
                BtRemoveSurname.IsEnabled = false;
            }
        }

        // Нажатие на кнопку "Добавить должность"
        private void BtAddSurname_Click(object sender, RoutedEventArgs e)
        {
            if (LbSurnames.SelectedIndex != -1)
            {
                // Получаем кол-во "пустых" значений
                var emptedCount = (from object item in LbStampSurnames.Items where string.IsNullOrEmpty(item.ToString()) select item.ToString()).ToList().Count;

                // Если кол-во пустых больше 0, то заменяем первое пустое выбранным значением
                if (emptedCount > 0)
                {
                    foreach (var item in LbStampSurnames.Items)
                    {
                        if (string.IsNullOrEmpty(item.ToString()))
                        {
                            var index = LbStampSurnames.Items.IndexOf(item);
                            LbStampSurnames.Items.Remove(item);
                            LbStampSurnames.Items.Insert(index, LbSurnames.SelectedItem);
                            LbSurnames.Items.Remove(LbSurnames.SelectedItem);
                            SaveNamesForCurrentTableInConfigFile();
                            break;
                        }
                    }
                }

                // Если кол-во пустых равно 0, значит мы достигли предела
                else
                {
                    MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg5") + " " +
                        Namecol.ToString(CultureInfo.InvariantCulture) + " " + ModPlusAPI.Language.GetItem(LangItem, "msg6"));
                }
            }
        }

        // Нажатие кнопки "Удалить должность"
        private void BtRemoveSurname_Click(object sender, RoutedEventArgs e)
        {
            if (LbStampSurnames.SelectedIndex != -1)
            {
                // Если выбранное значение пустое, то ничего не делаем
                // Это нужно для того, чтобы сохранять требуемое кол-во значений в списке
                var selected = LbStampSurnames.SelectedItem;
                if (!string.IsNullOrEmpty(selected.ToString()))
                {
                    LbSurnames.Items.Add(selected);
                    LbStampSurnames.Items.Remove(selected);

                    // Добавляем пустое значение!!!
                    LbStampSurnames.Items.Add(string.Empty);
                    SaveNamesForCurrentTableInConfigFile();
                }
            }
        }

        // Позиция вверх
        private void BtUpSurname_Click(object sender, RoutedEventArgs e)
        {
            if (LbStampSurnames.SelectedIndex != -1 &
                LbStampSurnames.SelectedIndex != 0)
            {
                var temp = LbStampSurnames.SelectedItem;
                var place = LbStampSurnames.SelectedIndex;
                LbStampSurnames.Items.Remove(LbStampSurnames.SelectedItem);
                LbStampSurnames.Items.Insert(place - 1, temp);
                LbStampSurnames.Focus();
                LbStampSurnames.SelectedIndex = place - 1;
                SaveNamesForCurrentTableInConfigFile();
            }
        }

        // Позиция вниз
        private void BtDownSurename_Click(object sender, RoutedEventArgs e)
        {
            if (LbStampSurnames.SelectedIndex != -1 &
                LbStampSurnames.SelectedIndex != Namecol - 1)
            {
                var temp = LbStampSurnames.SelectedItem;
                var place = LbStampSurnames.SelectedIndex;
                LbStampSurnames.Items.Remove(LbStampSurnames.SelectedItem);
                LbStampSurnames.Items.Insert(place + 1, temp);
                LbStampSurnames.Focus();
                LbStampSurnames.SelectedIndex = place + 1;
                SaveNamesForCurrentTableInConfigFile();
            }
        }

        // Сохранение должностей в штампе для текущего штампа в файл настроек
        private void SaveNamesForCurrentTableInConfigFile()
        {
            if (bool.Parse(_currentTblXml.Attribute("hassurenames").Value))
            {
                var str = LbStampSurnames.Items.Cast<object>()
                    .Aggregate(string.Empty, (current, item) => current + (item.ToString() + "$"));
                if (UserConfigFile.ConfigFileXml.Element("Settings") != null)
                {
                    var setXml = UserConfigFile.ConfigFileXml.Element("Settings");
                    if (setXml != null && setXml.Element("mpStampTblSaves") == null)
                    {
                        var tblXml = new XElement("mpStampTblSaves");
                        tblXml.SetAttributeValue(_currentTblXml.Attribute("tablestylename").Value, str.Substring(0, str.Length - 1));
                        setXml.Add(tblXml);
                    }
                    else
                    {
                        var element = setXml?.Element("mpStampTblSaves");
                        element?.SetAttributeValue(_currentTblXml.Attribute("tablestylename").Value, str.Substring(0, str.Length - 1));
                    }

                    UserConfigFile.SaveConfigFile();
                }
            }
        }

        private void CbFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var arr = new ArrayList();
            if (CbFormat.SelectedIndex == 0) // A0
            {
                arr = new ArrayList { "1", "2", "3" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }

            if (CbFormat.SelectedIndex == 1) // A1
            {
                arr = new ArrayList { "1", "3", "4" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }

            if (CbFormat.SelectedIndex == 2) // A2
            {
                arr = new ArrayList { "1", "3", "4", "5" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }

            if (CbFormat.SelectedIndex == 3) // A3
            {
                arr = new ArrayList { "1", "3", "4", "5", "6", "7" };
                PanelBottomFrame.Visibility = Visibility.Visible;
                int i;
                CbBottomFrame.SelectedIndex = int.TryParse(UserConfigFile.GetValue("mpFormats", "CbBottomFrame"), out i) ? i : 0;
            }

            if (CbFormat.SelectedIndex == 4) // A4
            {
                arr = new ArrayList { "1", "3", "4", "5", "6", "7", "8", "9" };
                PanelBottomFrame.Visibility = Visibility.Visible;
                int i;
                CbBottomFrame.SelectedIndex = int.TryParse(UserConfigFile.GetValue("mpFormats", "CbBottomFrame"), out i) ? i : 0;
            }

            if (CbFormat.SelectedIndex == 5) // A5
            {
                arr = new ArrayList { "1" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }

            CbMultiplicity.ItemsSource = arr;
            CbMultiplicity.SelectedIndex = 0;// Установить первое значение

            // show format size
            ShowFormatSize();
        }

        private void BtAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Point3d bottomLeftPt;
                Point3d topLeftPt;
                Point3d bottomRightPt;
                Vector3d replaceVector3D;
                Point3d blockInsertionPoint3D;

                if (Tabs.SelectedIndex == 0)
                {
                    string side, orientation;
                    if (RbShort.IsChecked != null && RbShort.IsChecked.Value)
                        side = ModPlusAPI.Language.GetItem(LangItem, "h11");
                    else
                        side = ModPlusAPI.Language.GetItem(LangItem, "h12");
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h8");
                    else
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h9");
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                    var multiplicity = CbMultiplicity.SelectedItem.ToString();

                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        // Форматка
                        if (MpFormatsAdd.DrawBlock(
                                format,
                                multiplicity,
                                side,
                                orientation,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbBottomFrame.SelectionBoxItem.ToString(),
                                false,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString()),
                                null,
                                ChkSetCurrentLayer.IsChecked ?? false,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D))
                        {
                            AddStamps(
                                bottomLeftPt,
                                topLeftPt,
                                bottomRightPt,
                                replaceVector3D,
                                Scale(CbScales.SelectedItem.ToString()),
                                blockInsertionPoint3D);
                        }
                    }
                    finally
                    {
                        Show();
                    }
                }

                if (Tabs.SelectedIndex == 1)
                {
                    if (!TbFormatLength.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err6"));
                        return;
                    }

                    if (!TbFormatHeight.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err7"));
                        return;
                    }

                    if (TbFormatLength.Value < 30)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err8"));
                        return;
                    }

                    if (TbFormatHeight.Value < 15)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err9"));
                        return;
                    }

                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        if (MpFormatsAdd.DrawBlockHand(
                                TbFormatLength.Value.Value,
                                TbFormatHeight.Value.Value,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                false,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString()),
                                null,
                                ChkSetCurrentLayer.IsChecked ?? false,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D))

                        {
                            AddStamps(
                                bottomLeftPt,
                                topLeftPt,
                                bottomRightPt,
                                replaceVector3D,
                                Scale(CbScales.SelectedItem.ToString()),
                                blockInsertionPoint3D);
                        }
                    }
                    finally
                    {
                        Show();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        private void BtReplace_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Tabs.SelectedIndex == 0)
                {
                    string side, orientation;

                    if (RbShort.IsChecked != null && RbShort.IsChecked.Value)
                        side = ModPlusAPI.Language.GetItem(LangItem, "h11");
                    else
                        side = ModPlusAPI.Language.GetItem(LangItem, "h12");
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h8");
                    else
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h9");
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                    var multiplicity = CbMultiplicity.SelectedItem.ToString();

                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        MpFormatsAdd.ReplaceBlock(
                                format,
                                multiplicity,
                                side,
                                orientation,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbBottomFrame.SelectionBoxItem.ToString(),
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString()),
                                ChkSetCurrentLayer.IsChecked ?? false);
                    }
                    finally
                    {
                        Show();
                    }
                }

                if (Tabs.SelectedIndex == 1)
                {
                    if (!TbFormatLength.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err6"));
                        return;
                    }

                    if (!TbFormatHeight.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err7"));
                        return;
                    }

                    if (TbFormatLength.Value < 30)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err8"));
                        return;
                    }

                    if (TbFormatHeight.Value < 15)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err9"));
                        return;
                    }

                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    Show();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        MpFormatsAdd.ReplaceBlockHand(
                                TbFormatLength.Value.Value,
                                TbFormatHeight.Value.Value,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString()),
                                ChkSetCurrentLayer.IsChecked ?? false);
                    }
                    finally
                    {
                        Show();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        private void BtCreateLayout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Point3d bottomLeftPt;
                Point3d topLeftPt;
                Point3d bottomRightPt;
                Vector3d replaceVector3D;
                Point3d blockInsertionPoint3D;

                if (Tabs.SelectedIndex == 0)
                {
                    string side, orientation;
                    if (RbShort.IsChecked != null && RbShort.IsChecked.Value)
                        side = ModPlusAPI.Language.GetItem(LangItem, "h11");
                    else
                        side = ModPlusAPI.Language.GetItem(LangItem, "h12");
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h8");
                    else
                        orientation = ModPlusAPI.Language.GetItem(LangItem, "h9");
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                    var multiplicity = CbMultiplicity.SelectedItem.ToString();

                    // Переменная указывает, следует ли оставить масштаб 1:1
                    // Создаем лист
                    if (!CreateLayout(out var layoutScaleOneToOne))
                        return;
                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        if (MpFormatsAdd.DrawBlock(
                                format,
                                multiplicity,
                                side,
                                orientation,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbBottomFrame.SelectionBoxItem.ToString(),
                                true,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                null,
                                ChkSetCurrentLayer.IsChecked ?? false,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D))

                        {
                            AddStamps(
                                bottomLeftPt,
                                topLeftPt,
                                bottomRightPt,
                                replaceVector3D,
                                layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                blockInsertionPoint3D);
                        }
                    }
                    finally
                    {
                        Show();
                    }
                }

                if (Tabs.SelectedIndex == 1)
                {
                    if (!TbFormatLength.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err6"));
                        return;
                    }

                    if (!TbFormatHeight.Value.HasValue)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err7"));
                        return;
                    }

                    if (TbFormatLength.Value < 30)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err8"));
                        return;
                    }

                    if (TbFormatHeight.Value < 15)
                    {
                        MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err9"));
                        return;
                    }

                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    // Создаем лист
                    if (!CreateLayout(out var layoutScaleOneToOne))
                        return;
                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        if (MpFormatsAdd.DrawBlockHand(
                                TbFormatLength.Value.Value,
                                TbFormatHeight.Value.Value,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                true,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                null,
                                ChkSetCurrentLayer.IsChecked ?? false,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt, out replaceVector3D,
                                out blockInsertionPoint3D))

                        {
                            AddStamps(
                                  bottomLeftPt,
                                  topLeftPt,
                                  bottomRightPt,
                                  replaceVector3D,
                                  layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                  blockInsertionPoint3D);
                        }
                    }
                    finally
                    {
                        Show();
                    }
                }

                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_zoom _all ", false, false, false);
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        private static bool CreateLayout(out bool layoutScaleOneToOne)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var returned = false;
            layoutScaleOneToOne = true;

            try
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    using (doc.LockDocument())
                    {
                        var lm = LayoutManager.Current;
                        var lname = ModPlusAPI.Language.GetItem(LangItem, "h52") + lm.LayoutCount.ToString(CultureInfo.InvariantCulture);

                        var layoutName = new LayoutName(lname);

                        // Если оба значения отсутствуют, тогда скрываем вообще ввод
                        if (string.IsNullOrEmpty(_lnamecoord) & string.IsNullOrEmpty(_lnumbercoord))
                            layoutName.GridAddToStamp.Visibility = Visibility.Collapsed;
                        else
                            layoutName.GridAddToStamp.Visibility = Visibility.Visible;

                        // Если нет номера
                        if (string.IsNullOrEmpty(_lnumbercoord))
                        {
                            layoutName.ChkLNumber.Visibility = Visibility.Collapsed;
                            layoutName.TbLNumber.Visibility = Visibility.Collapsed;
                        }
                        else
                        {
                            layoutName.ChkLNumber.Visibility = Visibility.Visible;
                            layoutName.TbLNumber.Visibility = Visibility.Visible;
                            layoutName.TbLNumber.Text = lm.LayoutCount.ToString(CultureInfo.InvariantCulture);
                        }

                        // Если нет имени
                        layoutName.ChkAddNameToStamp.Visibility = string.IsNullOrEmpty(_lnamecoord)
                            ? Visibility.Collapsed
                            : Visibility.Visible;
                        layoutName.ChkAddNameToStamp.IsChecked =
                            bool.TryParse(UserConfigFile.GetValue("mpFormats", "ChkAddNameToStamp"), out var b) && b;
                        layoutName.ChkLNumber.IsChecked =
                            bool.TryParse(UserConfigFile.GetValue("mpFormats", "ChkLNumber"), out b) && b;

                        // scale
                        layoutName.ChkLayoutScaleOneToOne.IsChecked = !bool.TryParse(UserConfigFile.GetValue("mpFormats", "LayoutScaleOneToOne"), out b) || b;
                        if (layoutName.ShowDialog() == true)
                        {
                            UserConfigFile.SetValue("mpFormats", "ChkAddNameToStamp",
                                (layoutName.ChkAddNameToStamp.IsChecked != null && layoutName.ChkAddNameToStamp.IsChecked.Value).ToString(), false);
                            UserConfigFile.SetValue("mpFormats", "ChkLNumber",
                                (layoutName.ChkLNumber.IsChecked != null && layoutName.ChkLNumber.IsChecked.Value).ToString(), false);
                            UserConfigFile.SetValue("mpFormats", "LayoutScaleOneToOne",
                                (layoutName.ChkLayoutScaleOneToOne.IsChecked != null && layoutName.ChkLayoutScaleOneToOne.IsChecked.Value).ToString(), false);
                            if (layoutName.ChkLayoutScaleOneToOne.IsChecked != null)
                                layoutScaleOneToOne = layoutName.ChkLayoutScaleOneToOne.IsChecked.Value;

                            lname = layoutName.TbLayoutName.Text;
                            if (layoutName.ChkAddNameToStamp.IsChecked != null && layoutName.ChkAddNameToStamp.IsChecked.Value)
                                _lname = lname;
                            if (layoutName.ChkLNumber.IsChecked != null && layoutName.ChkLNumber.IsChecked.Value)
                                _lnumber = layoutName.TbLNumber.Text;
                            var loId = lm.CreateLayout(lname);
                            var lo = tr.GetObject(loId, OpenMode.ForWrite) as Layout;
                            lo?.Initialize();
                            lm.CurrentLayout = lo?.LayoutName;
                            db.TileMode = false;
                            ed.SwitchToPaperSpace();
                            returned = true;
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
                returned = false;
            }

            return returned;
        }

        private void AddStamps(
            Point3d bottomLeftPt,
            Point3d topLeftPt,
            Point3d bottomRightPt,
            Vector3d replaceVector3D,
            double scale,
            Point3d blockInsertionPoint3D)
        {
            var docFor = ((ComboBoxItem)CbDocumentsFor.SelectedItem).Tag;
            if (ChkB1.IsChecked != null && ChkB1.IsChecked.Value)
            {
                if (docFor.Equals("RU"))
                    BtAddTable("Mp_GOST_P_21.1101_F3L2", "BottomRight", bottomLeftPt, replaceVector3D, scale, blockInsertionPoint3D);
                if (docFor.Equals("UA"))
                    BtAddTable("Mp_DSTU_B_A.2.4-4_F3L2", "BottomRight", bottomLeftPt, replaceVector3D, scale, blockInsertionPoint3D);
                if (docFor.Equals("BY"))
                    BtAddTable("Mp_STB_2255_E1L2", "BottomRight", bottomLeftPt, replaceVector3D, scale, blockInsertionPoint3D);
            }

            if (ChkB2.IsChecked != null && ChkB2.IsChecked.Value)
            {
                var pt = new Point3d(bottomLeftPt.X, bottomLeftPt.Y + 85, bottomLeftPt.Z);
                if (docFor.Equals("RU"))
                    BtAddTable("Mp_GOST_P_21.1101_F3L", "BottomRight", pt, replaceVector3D, scale, blockInsertionPoint3D);
                if (docFor.Equals("UA"))
                    BtAddTable("Mp_DSTU_B_A.2.4-4_F3L", "BottomRight", pt, replaceVector3D, scale, blockInsertionPoint3D);
                if (docFor.Equals("BY"))
                    BtAddTable("Mp_STB_2255_E1L", "BottomRight", pt, replaceVector3D, scale, blockInsertionPoint3D);
            }

            if (ChkB3.IsChecked != null && ChkB3.IsChecked.Value)
            {
                BtAddTable("Mp_GOST_2.104_dop1", "TopLeft", topLeftPt, replaceVector3D, scale, blockInsertionPoint3D);
            }

            if (ChkStamp.IsChecked != null && ChkStamp.IsChecked.Value)
            {
                if (CbTables.SelectedIndex != -1)
                {
                    if (CbTables.SelectedItem is Stamp stamp)
                        BtAddTable(stamp.TableStyle, "BottomRight", bottomRightPt, replaceVector3D, scale, blockInsertionPoint3D);
                }
            }
        }

        // Вставка штампа
        private void BtAddTable(
            string tableStyleName,
            string pointAligin,
            Point3d insertPt,
            Vector3d replaceVector3D,
            double scale,
            Point3d blockInsertionPoint3D)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var sourceDb = new Database(false, true);

            try
            {
                // Блокируем документ
                using (doc.LockDocument())
                {
                    var tr = doc.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        foreach (var xmltbl in _xmlTblsDoc.Elements("Stamp"))
                        {
                            if (tableStyleName.Equals(xmltbl.Attribute("tablestylename").Value))
                            {
                                // Директория расположения файла
                                var dir = Path.Combine(Constants.CurrentDirectory, "Data", "Dwg");

                                // Имя файла из которого берем таблицу
                                var sourceFileName = Path.Combine(dir, "Stamps.dwg");

                                // Read the DWG into a side database
                                sourceDb.ReadDwgFile(sourceFileName, FileShare.Read, true, "");
                                var tblIds = new ObjectIdCollection();

                                // Создаем пустую таблицу
                                var tbl = new Table();

                                var tm = sourceDb.TransactionManager;

                                using (var myT = tm.StartTransaction())
                                {
                                    var sourceBtr = (BlockTableRecord)myT.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForWrite, false);

                                    foreach (var obj in sourceBtr)
                                    {
                                        var ent = (Entity)myT.GetObject(obj, OpenMode.ForWrite);
                                        if (ent is Table table &&
                                            table.TableStyleName.Equals(xmltbl.Attribute("tablestylename").Value))
                                        {
                                            tblIds.Add(table.ObjectId);
                                            var im = new IdMapping();
                                            sourceDb.WblockCloneObjects(tblIds, db.CurrentSpaceId, im,
                                                DuplicateRecordCloning.Ignore, false);
                                            tbl = (Table)tr.GetObject(im.Lookup(table.ObjectId).Value, OpenMode.ForWrite);
                                            if (ChkSetCurrentLayer.IsChecked == true)
                                                tbl.LayerId = db.Clayer;
                                            break;
                                        }
                                    }

                                    myT.Commit();
                                }

                                if (tbl.ObjectId == ObjectId.Null)
                                {
                                    MessageBox.Show(
                                        ModPlusAPI.Language.GetItem(LangItem, "err10") + " " +
                                        tableStyleName + Environment.NewLine +
                                        ModPlusAPI.Language.GetItem(LangItem, "err11"));
                                    return;
                                }

                                var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                                // Перемещаем 
                                var mInsertPt = tbl.Position;
                                var width = tbl.Width;
                                var height = tbl.Height;

                                if (pointAligin.Equals("TopLeft"))
                                    mInsertPt = insertPt;
                                if (pointAligin.Equals("TopRight"))
                                    mInsertPt = new Point3d(insertPt.X - width, insertPt.Y, insertPt.Z);
                                if (pointAligin.Equals("BottomLeft"))
                                    mInsertPt = new Point3d(insertPt.X, insertPt.Y + height, insertPt.Z);
                                if (pointAligin.Equals("BottomRight"))
                                    mInsertPt = new Point3d(insertPt.X - width, insertPt.Y + height, insertPt.Z);
                                tbl.Position = mInsertPt;

                                var mat = Matrix3d.Displacement(replaceVector3D.GetNormal() * replaceVector3D.Length);
                                tbl.TransformBy(mat);

                                // Масштабируем относительно точки вставки блока
                                mat = Matrix3d.Scaling(scale, blockInsertionPoint3D);
                                tbl.TransformBy(mat);

                                doc.TransactionManager.QueueForGraphicsFlush();
                                sourceDb.Dispose();

                                tbl.Cells.TextStyleId = tst[CbTextStyle.SelectedItem.ToString()];

                                // Отступ в ячейках
                                for (var i = 0; i < tbl.Columns.Count; i++)
                                {
                                    for (var j = 0; j < tbl.Rows.Count; j++)
                                    {
                                        var cell = tbl.Cells[j, i];
                                        cell.Borders.Horizontal.Margin = 0.5 * scale;
                                        cell.Borders.Vertical.Margin = 0.5 * scale;
                                    }
                                }

                                // Нулевой отступ в ячейках
                                var nullMargin = xmltbl.Attribute("nullMargin");
                                if (nullMargin != null)
                                {
                                    foreach (var cellCoord in nullMargin.Value.Split(';'))
                                    {
                                        var cell = tbl.Cells[int.Parse(cellCoord.Split(',')[1]), int.Parse(cellCoord.Split(',')[0])];
                                        cell.Borders.Horizontal.Margin = 0;
                                        cell.Borders.Vertical.Margin = 0;
                                    }
                                }

                                // surenames
                                HasSureNames = xmltbl.Attribute("hassurenames").Value.Equals("true");
                                if (HasSureNames)
                                {
                                    for (var i = 0; i < LbStampSurnames.Items.Count; i++)
                                    {
                                        tbl.Cells[Row + i, Col].TextString =
                                            LbStampSurnames.Items[i].ToString();
                                    }
                                }

                                if (ChbHasFields.IsChecked != null && ChbHasFields.IsChecked.Value)
                                {
                                    AddFieldsToStamp(tbl);
                                }

                                // Добавление логотипа
                                if (bool.Parse(xmltbl.Attribute("logo").Value))
                                {
                                    var logoName = string.Empty;
                                    if (ChkLogoFromBlock.IsChecked != null && ChkLogoFromBlock.IsChecked.Value)
                                    {
                                        if (CbLogo.SelectedIndex > 0)
                                        {
                                            logoName = CbLogo.SelectedItem.ToString();
                                        }
                                    }
                                    else if (ChkLogoFromFile.IsChecked != null && ChkLogoFromFile.IsChecked.Value)
                                    {
                                        logoName = CreateBlockFromFile();
                                    }

                                    if (!string.IsNullOrEmpty(logoName))
                                    {
                                        var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                                        if (bt.Has(logoName))
                                        {
                                            var coln =
                                                int.Parse(
                                                    xmltbl.Attribute("logocoordinates").Value.Split(',')
                                                        .GetValue(0)
                                                        .ToString());
                                            var rown =
                                                int.Parse(
                                                    xmltbl.Attribute("logocoordinates").Value.Split(',')
                                                        .GetValue(1)
                                                        .ToString());
                                            var cell = tbl.Cells[rown, coln];
                                            cell.Borders.Top.Margin = 1 * scale;
                                            cell.Borders.Bottom.Margin = 1 * scale;
                                            cell.Borders.Right.Margin = 1 * scale;
                                            cell.Borders.Left.Margin = 1 * scale;
                                            cell.Alignment = CellAlignment.MiddleCenter;
                                            cell.Contents[0].BlockTableRecordId = bt[logoName];
                                        }
                                    }
                                }

                                // Высота текста
                                // Сначала ставим для всех ячеек 2.5
                                for (var i = 0; i < tbl.Columns.Count; i++)
                                {
                                    for (var j = 0; j < tbl.Rows.Count; j++)
                                    {
                                        var cell = tbl.Cells[j, i];
                                        cell.TextHeight =
                                            double.TryParse(TbMainTextHeight.Text, out var d)
                                                ? d * scale
                                                : 2.5 * scale;
                                    }
                                }

                                // Теперь ставим дополнительную высоту
                                var bigCellAtr = xmltbl.Attribute("bigCells");
                                if (bigCellAtr != null)
                                {
                                    foreach (var cellCoord in bigCellAtr.Value.Split(';'))
                                    {
                                        var cell = tbl.Cells[int.Parse(cellCoord.Split(',')[1]), int.Parse(cellCoord.Split(',')[0])];
                                        cell.TextHeight =
                                            double.TryParse(TbBigTextHeight.Text, out var d)
                                                ? d * scale
                                                : 3.5 * scale;
                                    }
                                }

                                // Имя листа и номер
                                if (xmltbl.Attribute("lnamecoordinates") != null)
                                {
                                    if (!string.IsNullOrEmpty(_lnamecoord))
                                    {
                                        if (!string.IsNullOrEmpty(_lname))
                                        {
                                            var rown = -1;
                                            var coln = -1;
                                            int.TryParse(_lnamecoord.Split(',').GetValue(0).ToString(), out rown);
                                            int.TryParse(_lnamecoord.Split(',').GetValue(1).ToString(), out coln);
                                            if (rown != -1 & coln != -1)
                                                tbl.Cells[rown, coln].TextString = _lname;
                                        }
                                    }
                                }

                                if (xmltbl.Attribute("lnumbercooridnates") != null)
                                {
                                    if (!string.IsNullOrEmpty(_lnumbercoord))
                                    {
                                        if (!string.IsNullOrEmpty(_lnumber))
                                        {
                                            var rown = -1;
                                            var coln = -1;
                                            int.TryParse(_lnumbercoord.Split(',').GetValue(0).ToString(), out rown);
                                            int.TryParse(_lnumbercoord.Split(',').GetValue(1).ToString(), out coln);
                                            if (rown != -1 & coln != -1)
                                            {
                                                if (string.IsNullOrEmpty(tbl.Cells[rown, coln].TextString))
                                                    tbl.Cells[rown, coln].TextString = _lnumber;
                                                else
                                                    tbl.Cells[rown, coln].TextString = tbl.Cells[rown, coln].TextString + " " + _lnumber;
                                            }
                                        }
                                    }
                                }

                                tr.Commit();
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        public static void AddFieldsToStamp(Table table)
        {
            try
            {
                var dbsi = AcApp.DocumentManager.MdiActiveDocument.Database.SummaryInfo;
                var dbsib = new DatabaseSummaryInfoBuilder(dbsi);

                // База данных штампов
                var doc = XElement.Parse(GetResourceTextFile("Stamps.xml"));
                foreach (var xmlTbl in doc.Elements("Stamp"))
                {
                    if (table.TableStyleName.Equals(xmlTbl.Attribute("tablestylename").Value))
                    {
                        // Даже если имя табличного стиля сошлось - проверяем по количеству ячеек!
                        if (int.Parse(xmlTbl.Attribute("cellcount").Value) != table.Cells.Count())
                        {
                            MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err12"));
                            return;
                        }

                        // Вставка полей, не содержащих фамилии
                        if (xmlTbl.Attribute("hasfields").Value.Equals("true"))
                        {
                            var fieldsNames = xmlTbl.Attribute("fieldsnames").Value.Split(',');
                            var fieldsCoordinates = xmlTbl.Attribute("fieldscoordinates").Value.Split(';');
                            for (var i = 0; i < fieldsNames.Length; i++)
                            {
                                if (dbsib.CustomPropertyTable.Contains(fieldsNames[i]))
                                {
                                    var col = int.Parse(fieldsCoordinates[i].Split(',').GetValue(0).ToString());
                                    var row = int.Parse(fieldsCoordinates[i].Split(',').GetValue(1).ToString());
                                    var cell = new Cell(table, row, col);
                                    table.Cells[row, col].TextString =
                                        cell.TextString =
                                            "%<\\AcVar CustomDP." +
                                            fieldsNames[i] +
                                            ">%";
                                }
                            }
                        }

                        // Вставка полей, содержащих фамилию
                        if (xmlTbl.Attribute("hassurenames").Value.Equals("true"))
                        {
                            var startrow =
                                int.Parse(xmlTbl.Attribute("firstcell").Value.Split(',').GetValue(1).ToString());
                            var startcol =
                                int.Parse(xmlTbl.Attribute("firstcell").Value.Split(',').GetValue(0).ToString());
                            var n = table.Rows.Count - startrow;

                            // Стандартные
                            var surnames = xmlTbl.Attribute("surnames").Value.Split(',').ToList();
                            var surnameskeys = xmlTbl.Attribute("surnameskeys").Value.Split(',').ToList();

                            if (UserConfigFile.ConfigFileXml.Element("Settings") != null)
                            {
                                if (UserConfigFile.ConfigFileXml?.Element("Settings")?.Element("UserSurnames") != null)
                                {
                                    var xElements = UserConfigFile.ConfigFileXml?.Element("Settings")?.Element("UserSurnames")?.Elements("Surname");
                                    if (xElements != null)
                                        foreach (var sn in xElements)
                                        {
                                            surnames.Add(sn.Attribute("Surname").Value);
                                            surnameskeys.Add(sn.Attribute("Id").Value);
                                        }
                                }
                            }

                            for (var i = 0; i < n; i++)
                            {
                                if (!table.Cells[startrow + i, startcol].TextString.Equals(string.Empty))
                                {
                                    for (var j = 0; j < surnames.Count; j++)
                                    {
                                        if (table.Cells[startrow + i, startcol].TextString.Equals(surnames[j]))
                                        {
                                            if (dbsib.CustomPropertyTable.Contains(surnameskeys[j]))
                                            {
                                                var k = 1;
                                                var isMerged = table.Cells[startrow + i, startcol].IsMerged;
                                                if (isMerged != null && isMerged.Value)
                                                    k = 2;
                                                table.Cells[startrow + i, startcol + k].TextString =
                                                    "%<\\AcVar CustomDP." +
                                                    surnameskeys[j] +
                                                    ">%";
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }

        private static string GetResourceTextFile(string filename)
        {
            string result = string.Empty;

            using (Stream stream = Assembly.GetExecutingAssembly().
                GetManifestResourceStream("mpFormats.Resources." + filename))
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    result = sr.ReadToEnd();
                }
            }

            return result;
        }

        #region checkboxes
        private void ChbNumber_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_num.Opacity = 1;
        }

        private void ChbNumber_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_num.Opacity = 0.5;
        }

        private void CbBottomFrame_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(sender is ComboBox cb))
                return;
            var index = cb.SelectedIndex;
            if (index == 0)
            {
                var uriSource = new Uri(@"/mpFormats_" + ModPlusConnector.Instance.AvailProductExternalVersion + ";component/Resources/Preview/F_5.png", UriKind.Relative);
                Image_format.Source = new BitmapImage(uriSource);

                Image_b1.Margin = new Thickness(5, 0, 0, 3);
                Image_b2.Margin = new Thickness(3, 0, 0, 54);
                Image_stamp.Margin = new Thickness(0, 0, 3, 3);
            }
            else
            {
                var uriSource = new Uri(@"/mpFormats_" + ModPlusConnector.Instance.AvailProductExternalVersion + ";component/Resources/Preview/F_10.png", UriKind.Relative);
                Image_format.Source = new BitmapImage(uriSource);

                Image_b1.Margin = new Thickness(5, 0, 0, 5);
                Image_b2.Margin = new Thickness(3, 0, 0, 56);
                Image_stamp.Margin = new Thickness(0, 0, 3, 5);
            }
        }

        private void ChkB1_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_b1.Opacity = 1;
        }

        private void ChkB1_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_b1.Opacity = 0.3;
        }

        private void ChkB2_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_b2.Opacity = 1;
        }

        private void ChkB2_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_b2.Opacity = 0.3;
        }

        private void ChkB3_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_top.Opacity = 1;
        }

        private void ChkB3_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_top.Opacity = 0.3;
        }

        private void ChkStamp_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_stamp.Opacity = 1;
            GridStamp.Visibility = CbDocumentsFor.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Visible;

        }

        private void ChkStamp_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_stamp.Opacity = 0.3;
            GridStamp.Visibility = CbDocumentsFor.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Collapsed;
        }
        #endregion

        private void ShowFormatSize()
        {
            try
            {
                string side, orientation;
                if (RbShort.IsChecked != null && RbShort.IsChecked.Value)
                    side = ModPlusAPI.Language.GetItem(LangItem, "h11");
                else
                    side = ModPlusAPI.Language.GetItem(LangItem, "h12");
                if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                    orientation = ModPlusAPI.Language.GetItem(LangItem, "h8");
                else
                    orientation = ModPlusAPI.Language.GetItem(LangItem, "h9");

                var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                var multiplicity = CbMultiplicity.SelectedItem.ToString();

                MpFormatsAdd.GetFormatSize(format, orientation, side, multiplicity, out var width, out var height);
                TbFormatSize.Text = width.ToString(CultureInfo.InvariantCulture) + " x " +
                                    height.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                TbFormatSize.Text = string.Empty;
            }
        }

        private void CbMultiplicity_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // show format size
            ShowFormatSize();
        }

        private void RadioButton_FormatSettings_OnChecked_OnUnchecked(object sender, RoutedEventArgs e)
        {
            // show format size
            ShowFormatSize();
        }

        private void MainTab_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is TabControl tab)
            {
                if (tab.SelectedIndex != -1)
                    TbFormatSize.Visibility = tab.SelectedIndex == 0 ? Visibility.Visible : Visibility.Collapsed;
                else
                    TbFormatSize.Visibility = Visibility.Visible;
            }
        }

        private void ChkLogoFromBlock_OnChecked(object sender, RoutedEventArgs e)
        {
            ChkLogoFromFile.IsChecked = false;
        }

        private void ChkLogoFromFile_OnChecked(object sender, RoutedEventArgs e)
        {
            ChkLogoFromBlock.IsChecked = false;
        }

        private void ChkLogoFromBlock_OnUnchecked(object sender, RoutedEventArgs e)
        {
            ChkLogoFromFile.IsChecked = true;
        }

        private void ChkLogoFromFile_OnUnchecked(object sender, RoutedEventArgs e)
        {
            ChkLogoFromBlock.IsChecked = true;
        }

        private void BtGetFileForLogo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var ofd = new OpenFileDialog(ModPlusAPI.Language.GetItem(LangItem, "h53"), string.Empty, "dwg", "dialog",
                    OpenFileDialog.OpenFileDialogFlags.SearchPath |
                    OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var sourceDb = new Database(false, true);

                    // read file DB
                    sourceDb.ReadDwgFile(ofd.Filename, FileOpenMode.OpenTryForReadShare, true, string.Empty);

                    // Если файл более поздней версии, то будет ошибка
                    TbLogoFile.Text = ofd.Filename;
                }
                else
                {
                    MessageBox.Show(
                        ModPlusAPI.Language.GetItem(LangItem, "msg7") + Environment.NewLine + ofd.Filename + Environment.NewLine +
                        ModPlusAPI.Language.GetItem(LangItem, "msg8"));
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exception)
            {
                if (exception.ErrorStatus == ErrorStatus.NotImplementedYet)
                {
                    MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "err13"));
                }
                else
                {
                    ExceptionBox.Show(exception);
                }
            }
        }

        // Создание блока из файла (вставка файла в виде блока)
        private string CreateBlockFromFile()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var blockName = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(TbLogoFile.Text))
                {
                    if (File.Exists(TbLogoFile.Text))
                    {
                        using (var sourceDb = new Database(false, true))
                        {
                            sourceDb.ReadDwgFile(TbLogoFile.Text, FileOpenMode.OpenTryForReadShare, true, string.Empty);

                            db.Insert(Path.GetFileNameWithoutExtension(TbLogoFile.Text), sourceDb, true);
                        }

                        blockName = Path.GetFileNameWithoutExtension(TbLogoFile.Text);
                    }
                }

                return blockName;
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
                return blockName;
            }
        }
        
        // Запрет пробела
        private void TboxesNoSpaceBar_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = e.Key == Key.Space;
        }

        // - без минуса
        private void Tb_OnlyNums_NoMinus_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var txt = ((TextBox)sender).Text + e.Text;
            e.Handled = !DoubleCharChecker(txt, false, true, null);
        }

        // Проверка, что число, точка или знак минус
        private static bool DoubleCharChecker(string str, bool checkMinus, bool checkDot, double? max)
        {
            var result = false;
            if (str.Count(c => c.Equals('.')) > 1)
                return false;
            if (str.Count(c => c.Equals('-')) > 1)
                return false;

            // Проверять нужно только последний знак в строке!!!
            var ch = str.Last();
            if (checkMinus)
            {
                if (ch.Equals('-'))
                    result = str.IndexOf(ch) == 0;
            }

            if (checkDot)
            {
                if (ch.Equals('.'))
                    result = true;
            }

            if (char.IsNumber(ch))
                result = true;

            // На "максимальность" проверяем если предыдущие проверки успешны
            if (max != null & result)
            {
                if (double.TryParse(str, out var d))
                {
                    if (Math.Abs(d) > max)
                        result = false;
                }
            }

            return result;
        }
    }
}
