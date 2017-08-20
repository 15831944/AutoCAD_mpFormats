#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using mpMsg;
using mpSettings;
using ModPlus;

using Color = Autodesk.AutoCAD.Colors.Color;
using Exception = System.Exception;
using Visibility = System.Windows.Visibility;

namespace mpFormats
{
    /// <summary>
    /// Логика взаимодействия для MpFormats.xaml
    /// </summary>
    public partial class MpFormats
    {
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
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
            // Настройки видимости для штампа отключаем тут, чтобы видеть в редакторе окна
            GridStamp.Visibility = //DpSurenames.Visibility = //TbLogo.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Collapsed;
        }
        #region window basic
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
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
        private void MpFormats_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;

                _xmlTblsDoc = XElement.Parse(Properties.Resources.Stamps);
                // Заполнение списка масштабов
                var ocm = db.ObjectContextManager;
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                foreach (var ans in occ.Cast<AnnotationScale>())
                {
                    CbScales.Items.Add(ans.Name);
                }
                // Начальное значение масштаба
                var cans = occ.CurrentContext as AnnotationScale;
                if (cans != null) CbScales.SelectedItem = cans.Name;
                // Заполняем список текстовых стилей
                string txtstname;
                using (var acTrans = doc.TransactionManager.StartTransaction())
                {
                    var tst = (TextStyleTable)acTrans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                    foreach (var id in tst)
                    {
                        var tstr = (TextStyleTableRecord)acTrans.GetObject(id, OpenMode.ForRead);
                        CbTextStyle.Items.Add(tstr.Name);
                    }
                    var curtxt = (TextStyleTableRecord)acTrans.GetObject(db.Textstyle, OpenMode.ForRead);
                    txtstname = curtxt.Name;
                    acTrans.Commit();
                }
                CbTextStyle.SelectedItem = txtstname;
                // Логотип
                CbLogo.Items.Clear();
                CbLogo.Items.Add("Нет");
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

                //FillStamps();

                // Загрузка из настроек
                LoadFromSettings();
                // Проверка файла со штампами
                if (!CheckTableFileExist())
                {
                    MpMsgWin.Show("Не найден файл со штампами!" + Environment.NewLine + "Запустите функцию \"Штампы\"");
                    // Видимость
                    GridStamp.Visibility = //DpSurenames.Visibility = //TbLogo.Visibility =
                        CbLogo.Visibility =
                        ChkStamp.Visibility = ChkB1.Visibility = ChkB2.Visibility = ChkB3.Visibility = GridSplitterStamp.Visibility =
                        Image_stamp.Visibility = Image_b1.Visibility = Image_b2.Visibility = Image_top.Visibility =
                        Visibility.Collapsed;
                    ChkStamp.IsChecked = ChkB1.IsChecked = ChkB2.IsChecked = ChkB3.IsChecked = false;
                }
                // Включаем обработчики событий
                //CbFormat.SelectionChanged += CbFormat_SelectionChanged;
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
            catch (Exception exception)
            {
                MpExWin.Show(exception);
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
                var cb = sender as ComboBox;
                var comboBoxItem = cb?.SelectedItem as ComboBoxItem;
                if (cb != null && comboBoxItem != null && cb.SelectedIndex != -1)
                {
                    MpSettings.SetValue("Settings", "mpFormats", "CbDocumentsFor",
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
                MpExWin.Show(ex);
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
            CbDocumentsFor.SelectedIndex =
                    int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbDocumentsFor"), out int index)
                        ? index
                        : 0;
            // format
            int i;
            CbFormat.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbFormat"), out i) ? i : 3;
            CbMultiplicity.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbMultiplicity"), out i) ? i : 0;
            CbBottomFrame.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbBottomFrame"), out i) ? i : 0;
            // Выбранный штамп
            CbTables.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbTables"), out i) ? i : 0;
            // Масштаб
            var scale = MpSettings.GetValue("Settings", "mpFormats", "CbScales");
            CbScales.SelectedIndex = CbScales.Items.Contains(scale)
                ? CbScales.Items.IndexOf(scale)
                : 0;
            bool b;
            ChkB1.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkB1"), out b) && b;
            ChkB2.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkB2"), out b) && b;
            ChkB3.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkB3"), out b) && b;
            ChbCopy.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChbCopy"), out b) && b;
            ChbNumber.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChbNumber"), out b) && b;
            ChkStamp.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkStamp"), out b) && b;
            RbVertical.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "RbVertical"), out b) && b;
            RbLong.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "RbLong"), out b) && b;

            TbFormatHeight.Text = MpSettings.GetValue("Settings", "mpFormats", "TbFormatHeight");
            TbFormatLength.Text = MpSettings.GetValue("Settings", "mpFormats", "TbFormatLength");
            // Текстовый стиль (меняем, если есть в настройках, а иначе оставляем текущий)
            var txtstl = MpSettings.GetValue("Settings", "mpFormats", "CbTextStyle");
            if (CbTextStyle.Items.Contains(txtstl))
                CbTextStyle.SelectedIndex = CbTextStyle.Items.IndexOf(txtstl);
            // Логотип
            var logo = MpSettings.GetValue("Settings", "mpFormats", "CbLogo");
            CbLogo.SelectedIndex = CbLogo.Items.Contains(logo) ? CbLogo.Items.IndexOf(logo) : 0;
            // Поля
            ChbHasFields.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChbHasFields"), out b) && b;
            // Высота текста
            double d;
            TbMainTextHeight.Text = double.TryParse(MpSettings.GetValue("Settings", "mpFormats", "MainTextHeight"),
                out d)
                ? d.ToString(CultureInfo.InvariantCulture)
                : "2.5";
            TbBigTextHeight.Text = double.TryParse(MpSettings.GetValue("Settings", "mpFormats", "BigTextHeight"),
                out d)
                ? d.ToString(CultureInfo.InvariantCulture)
                : "3.5";
            // logo from
            ChkLogoFromBlock.IsChecked = !bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "LogoFromBlock"), out b) || b;
            ChkLogoFromFile.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "LogoFromFile"), out b) && b;
            // logo file
            var ffs = MpSettings.GetValue("Settings", "mpFormats", "LogoFile");
            if (!string.IsNullOrEmpty(ffs))
                if (File.Exists(ffs))
                    TbLogoFile.Text = ffs;
        }
        // Сохранение в настройки
        private void SaveToSettings()
        {
            try
            {
                MpSettings.SetValue("Settings", "mpFormats", "CbFormat", CbFormat.SelectedIndex.ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "CbMultiplicity", CbMultiplicity.SelectedIndex.ToString(),
                    false);
                MpSettings.SetValue("Settings", "mpFormats", "CbBottomFrame", CbBottomFrame.SelectedIndex.ToString(),
                    false);

                MpSettings.SetValue("Settings", "mpFormats", "ChkB1",
                    (ChkB1.IsChecked != null && ChkB1.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "ChkB2",
                    (ChkB2.IsChecked != null && ChkB2.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "ChkB3",
                    (ChkB3.IsChecked != null && ChkB3.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "ChbCopy",
                    (ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "ChbNumber",
                    (ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "ChkStamp",
                    (ChkStamp.IsChecked != null && ChkStamp.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "RbVertical",
                    (RbVertical.IsChecked != null && RbVertical.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "RbLong",
                    (RbLong.IsChecked != null && RbLong.IsChecked.Value).ToString(), false);

                MpSettings.SetValue("Settings", "mpFormats", "TbFormatHeight", TbFormatHeight.Text, false);
                MpSettings.SetValue("Settings", "mpFormats", "TbFormatLength", TbFormatLength.Text, false);
                // Текстовый стиль
                MpSettings.SetValue("Settings", "mpFormats", "CbTextStyle", CbTextStyle.SelectedItem.ToString(), false);
                // Логотип
                MpSettings.SetValue("Settings", "mpFormats", "CbLogo", CbLogo.SelectedItem.ToString(), false);
                // Поля
                MpSettings.SetValue("Settings", "mpFormats", "ChbHasFields",
                    (ChbHasFields.IsChecked != null && ChbHasFields.IsChecked.Value).ToString(), false);
                // Выбранный штамп
                MpSettings.SetValue("Settings", "mpFormats", "CbTables",
                    CbTables.SelectedIndex.ToString(CultureInfo.InvariantCulture), false);
                // Масштаб
                MpSettings.SetValue("Settings", "mpFormats", "CbScales",
                                        CbScales.SelectedItem.ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "MainTextHeight", TbMainTextHeight.Text, false);
                MpSettings.SetValue("Settings", "mpFormats", "BigTextHeight", TbBigTextHeight.Text, false);
                MpSettings.SetValue("Settings", "mpFormats", "LogoFromBlock", (ChkLogoFromBlock.IsChecked != null && ChkLogoFromBlock.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "LogoFromFile", (ChkLogoFromFile.IsChecked != null && ChkLogoFromFile.IsChecked.Value).ToString(), false);
                MpSettings.SetValue("Settings", "mpFormats", "LogoFile",
                    File.Exists(TbLogoFile.Text) ? TbLogoFile.Text : string.Empty, false);

                MpSettings.SaveFile();
            }
            catch (Exception exception)
            {
                MpExWin.Show(exception);
            }
        }
        private static double Scale(string scaleName)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ocm = db.ObjectContextManager;
            var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
            var ansc = occ.GetContext(scaleName) as AnnotationScale;
            Debug.Assert(ansc != null, "ansc != null");
            return (ansc.DrawingUnits / ansc.PaperUnits);
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
            if (MpSettings.XmlMpSettingsFile.Element("Settings") != null)
            {
                var setXml = MpSettings.XmlMpSettingsFile.Element("Settings");
                if (setXml?.Element("UserSurnames") != null)
                {
                    var element = setXml.Element("UserSurnames");
                    if (element != null)
                        foreach (var sn in element.Elements("Surname"))
                        {
                            LbSurnames.Items.Add(sn.Attribute("Surname").Value);
                        }
                }
            }
            // Заполняем значения из файла настроек. Про совпадении в "левом" списке - удаляем
            if (MpSettings.XmlMpSettingsFile.Element("Settings") != null)
            {
                var setXml = MpSettings.XmlMpSettingsFile.Element("Settings");
                if (setXml?.Element("mpStampTblSaves") != null)
                {
                    var element = setXml.Element("mpStampTblSaves");
                    if (element?.Attribute(_currentTblXml.Attribute("tablestylename").Value) != null)
                    {
                        LbStampSurnames.Items.Clear();
                        var xElement = setXml.Element("mpStampTblSaves");
                        if (xElement != null)
                            foreach (var item in xElement.Attribute(
                                _currentTblXml.Attribute("tablestylename").Value).Value.Split('$'))
                            {
                                if (LbSurnames.Items.Contains(item))
                                    LbSurnames.Items.Remove(item);
                                LbStampSurnames.Items.Add(item);
                            }
                    }
                    else FillStampSurnamesEmpty();
                }
                else FillStampSurnamesEmpty();
            }
            else FillStampSurnamesEmpty();
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
                var docFor = ((ComboBoxItem) CbDocumentsFor.SelectedItem).Tag;
                var stamps = new List<Stamp>();
                foreach (var tbl in _xmlTblsDoc.Elements("Stamp"))
                {
                    var docForAttr = tbl.Attribute("DocFor");
                    if (docForAttr != null && tbl.Attribute("noShow") == null & 
                        docForAttr.Value.Equals(docFor))
                        stamps.Add(new Stamp
                        {
                            Name = tbl.Attribute("name")?.Value,
                            TableStyle = tbl.Attribute("tablestylename")?.Value,
                            Description = tbl.Attribute("document")?.Value + " " + tbl.Attribute("description")?.Value
                        });
                }
                CbTables.ItemsSource = stamps;
                // Устанавливаем первое значение
                CbTables.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MpExWin.Show(ex);
            }
        }
        // Выбор таблицы
        private void CbTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (CbTables.Items.Count > 0) // Если не пустой список
                {
                    var stamp = CbTables.SelectedItem as Stamp;
                    // Изменяем значения элементов формы в зависимости от выбранной таблицы
                    foreach (var tbl in _xmlTblsDoc.Elements("Stamp"))
                    {
                        if (stamp != null)
                            if (stamp.Name.Equals(tbl.Attribute("name").Value) &
                                stamp.TableStyle.Equals(tbl.Attribute("tablestylename").Value))
                            {
                                _currentTblXml = tbl;

                                // Изображение
                                try
                                {
                                    var uriSource = new Uri(@"/mpFormats_" + VersionData.FuncVersion + ";component/Resources/Preview/" + tbl.Attribute("img")?.Value + ".png", UriKind.Relative);
                                    Image_stamp.Source = new BitmapImage(uriSource);
                                }
                                catch
                                {
                                    //
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
                            } // Tabel name
                    }
                }
            }
            catch (Exception ex)
            {
                MpExWin.Show(ex);
            }
        }
        // Вызов функции "Поля"
        private void BtFields_Click(object sender, RoutedEventArgs e)
        {
            // Проверка полной версии
            if (!MpCadHelpers.IsFunctionBought("mpStamps", VersionData.FuncVersion))
            {
                MpMsgWin.Show("Доступно при наличии полной версии функции \"Штампы\"");
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
            if (!MpCadHelpers.IsFunctionBought("mpStamps", VersionData.FuncVersion))
            {
                MpMsgWin.Show("Доступно при наличии полной версии функции \"Штампы\"");
            }
            else
            {
                var window = new EditUserSurnames
                {
                    Topmost = true
                };
                window.ShowDialog();
                // Перезаполнем список фамилий
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
                BtRemoveSurname.IsEnabled = false;
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
                            SaveNamesForCurrentTableInConfigfile();
                            break;
                        }
                    }
                }
                // Если кол-во пустых равно 0, значит мы достигли предела
                else
                {
                    MpMsgWin.Show("Нельзя добавлять более " + Namecol.ToString(CultureInfo.InvariantCulture) + " должностей");
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
                    SaveNamesForCurrentTableInConfigfile();
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
                SaveNamesForCurrentTableInConfigfile();
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
                SaveNamesForCurrentTableInConfigfile();
            }
        }
        // Сохранение должностей в штампе для текущего штампа в файл настроек
        private void SaveNamesForCurrentTableInConfigfile()
        {
            if (bool.Parse(_currentTblXml.Attribute("hassurenames").Value))
            {
                var str = LbStampSurnames.Items.Cast<object>()
                    .Aggregate(string.Empty, (current, item) => current + (item.ToString() + "$"));
                if (MpSettings.XmlMpSettingsFile.Element("Settings") != null)
                {
                    var setXml = MpSettings.XmlMpSettingsFile.Element("Settings");
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

                    MpSettings.SaveFile();
                }
            }
        }

        private void CbFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var arr = new ArrayList();
            if (CbFormat.SelectedIndex == 0)// A0
            {
                arr = new ArrayList { "1", "2", "3" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }
            if (CbFormat.SelectedIndex == 1)// A1
            {
                arr = new ArrayList { "1", "3", "4" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }
            if (CbFormat.SelectedIndex == 2)// A2
            {
                arr = new ArrayList { "1", "3", "4", "5" };
                PanelBottomFrame.Visibility = Visibility.Collapsed;
                CbBottomFrame.SelectedIndex = 0;
            }
            if (CbFormat.SelectedIndex == 3)// A3
            {
                arr = new ArrayList { "1", "3", "4", "5", "6", "7" };
                PanelBottomFrame.Visibility = Visibility.Visible;
                int i;
                CbBottomFrame.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbBottomFrame"), out i) ? i : 0;
            }
            if (CbFormat.SelectedIndex == 4)// A4
            {
                arr = new ArrayList { "1", "3", "4", "5", "6", "7", "8", "9" };
                PanelBottomFrame.Visibility = Visibility.Visible;
                int i;
                CbBottomFrame.SelectedIndex = int.TryParse(MpSettings.GetValue("Settings", "mpFormats", "CbBottomFrame"), out i) ? i : 0;
            }
            if (CbFormat.SelectedIndex == 5)// A5
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
                    if (RbShort.IsChecked != null && RbShort.IsChecked.Value) side = "Короткая";
                    else
                        side = "Длинная";
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = "Альбомный";
                    else
                        orientation = "Книжный";
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
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D
                            ))
                            AddStamps(bottomLeftPt, topLeftPt, bottomRightPt, replaceVector3D, Scale(CbScales.SelectedItem.ToString()), blockInsertionPoint3D);// stamps
                    }
                    finally
                    {
                        Show();
                    }
                }
                if (Tabs.SelectedIndex == 1)
                {
                    if (string.IsNullOrEmpty(TbFormatLength.Text))
                    {
                        MpMsgWin.Show("Не введена длина форматки");
                        return;
                    }
                    if (string.IsNullOrEmpty(TbFormatHeight.Text))
                    {
                        MpMsgWin.Show("Не введена высота форматки");
                        return;
                    }
                    if (double.Parse(TbFormatLength.Text) < 30)
                    {
                        MpMsgWin.Show("Длина форматки должна быть не менее 30 мм");
                        return;
                    }
                    if (double.Parse(TbFormatHeight.Text) < 15)
                    {
                        MpMsgWin.Show("Высота форматки должна быть не менее 15 мм");
                        return;
                    }
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        if (MpFormatsAdd.DrawBlockHand(
                                double.Parse(TbFormatLength.Text),
                                double.Parse(TbFormatHeight.Text),
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                false,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString()),
                                null,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D
                            ))
                            AddStamps(bottomLeftPt, topLeftPt, bottomRightPt, replaceVector3D, Scale(CbScales.SelectedItem.ToString()), blockInsertionPoint3D);// stamps
                    }
                    finally
                    {
                        Show();
                    }
                }
            } // try
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
                        side = "Короткая";
                    else
                        side = "Длинная";
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = "Альбомный";
                    else
                        orientation = "Книжный";
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                    var multiplicity = CbMultiplicity.SelectedItem.ToString();

                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        MpFormatsAdd.ReplaceBlock
                            (
                                format,
                                multiplicity,
                                side,
                                orientation,
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbBottomFrame.SelectionBoxItem.ToString(),
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString())
                            );
                    }
                    finally
                    {
                        Show();
                    }
                }
                if (Tabs.SelectedIndex == 1)
                {
                    if (string.IsNullOrEmpty(TbFormatLength.Text))
                    {
                        MpMsgWin.Show("Не введена длина форматки");
                        return;
                    }
                    if (string.IsNullOrEmpty(TbFormatHeight.Text))
                    {
                        MpMsgWin.Show("Не введена высота форматки");
                        return;
                    }
                    if (double.Parse(TbFormatLength.Text) < 30)
                    {
                        MpMsgWin.Show("Длина форматки должна быть не менее 30 мм");
                        return;
                    }
                    if (double.Parse(TbFormatHeight.Text) < 15)
                    {
                        MpMsgWin.Show("Высота форматки должна быть не менее 15 мм");
                        return;
                    }
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    Show();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        MpFormatsAdd.ReplaceBlockHand
                            (
                                double.Parse(TbFormatLength.Text),
                                double.Parse(TbFormatHeight.Text),
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                CbTextStyle.SelectedItem.ToString(),
                                Scale(CbScales.SelectedItem.ToString())
                            );
                    }
                    finally
                    {
                        Show();
                    }
                }
            } // try
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
                    if (RbShort.IsChecked != null && RbShort.IsChecked.Value) side = "Короткая";
                    else side = "Длинная";
                    if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                        orientation = "Альбомный";
                    else orientation = "Книжный";
                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;

                    var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                    var multiplicity = CbMultiplicity.SelectedItem.ToString();
                    // Переменная указывает, следует ли оставить масштаб 1:1
                    bool layoutScaleOneToOne;
                    // Создаем лист
                    if (!CreateLayout(out layoutScaleOneToOne)) return;
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
                                true,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                null,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt,
                                out replaceVector3D,
                                out blockInsertionPoint3D
                            ))
                            // stamps
                            AddStamps(bottomLeftPt, topLeftPt, bottomRightPt, replaceVector3D, layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()), blockInsertionPoint3D);
                    }
                    finally
                    {
                        Show();
                    }
                }
                if (Tabs.SelectedIndex == 1)
                {
                    if (string.IsNullOrEmpty(TbFormatLength.Text))
                    {
                        MpMsgWin.Show("Не введена длина форматки");
                        return;
                    }
                    if (string.IsNullOrEmpty(TbFormatHeight.Text))
                    {
                        MpMsgWin.Show("Не введена высота форматки");
                        return;
                    }
                    if (double.Parse(TbFormatLength.Text) < 30)
                    {
                        MpMsgWin.Show("Длина форматки должна быть не менее 30 мм");
                        return;
                    }
                    if (double.Parse(TbFormatHeight.Text) < 15)
                    {
                        MpMsgWin.Show("Высота форматки должна быть не менее 15 мм");
                        return;
                    }

                    var number = ChbNumber.IsChecked != null && ChbNumber.IsChecked.Value;
                    bool layoutScaleOneToOne;
                    // Создаем лист
                    if (!CreateLayout(out layoutScaleOneToOne)) return;
                    Hide();
                    try
                    {
                        Utils.SetFocusToDwgView();

                        if (MpFormatsAdd.DrawBlockHand(
                                double.Parse(TbFormatLength.Text),
                                double.Parse(TbFormatHeight.Text),
                                number,
                                ChbCopy.IsChecked != null && ChbCopy.IsChecked.Value,
                                true,
                                new Point3d(0.0, 0.0, 0.0),
                                CbTextStyle.SelectedItem.ToString(),
                                layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()),
                                null,
                                out bottomLeftPt,
                                out topLeftPt,
                                out bottomRightPt, out replaceVector3D,
                                out blockInsertionPoint3D
                            ))
                            // stamps
                            AddStamps(bottomLeftPt, topLeftPt, bottomRightPt, replaceVector3D, layoutScaleOneToOne ? 1 : Scale(CbScales.SelectedItem.ToString()), blockInsertionPoint3D);
                    }
                    finally
                    {
                        Show();
                    }
                }
                AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_zoom _all ", false, false, false);
            }
            catch (Exception ex)
            {
                MpExWin.Show(ex);
            }
        }
        private static bool CreateLayout(out bool layoutScaleOneToOne)
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var returned = false;
            layoutScaleOneToOne = true;
            // Создаем лист
            try
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    using (doc.LockDocument())
                    {
                        var lm = LayoutManager.Current;
                        var lname = "Лист" + lm.LayoutCount.ToString(CultureInfo.InvariantCulture);

                        var lnamewin = new LayoutName(lname);
                        // Если оба значения отсутствуют, тогда скрываем вообще ввод
                        if (string.IsNullOrEmpty(_lnamecoord) & string.IsNullOrEmpty(_lnumbercoord))
                            lnamewin.GridAddToStamp.Visibility = Visibility.Collapsed;
                        else lnamewin.GridAddToStamp.Visibility = Visibility.Visible;
                        // Если нет номера
                        if (string.IsNullOrEmpty(_lnumbercoord))
                        {
                            lnamewin.ChkLNumber.Visibility = Visibility.Collapsed;
                            lnamewin.TbLNumber.Visibility = Visibility.Collapsed;
                        }
                        else
                        {
                            lnamewin.ChkLNumber.Visibility = Visibility.Visible;
                            lnamewin.TbLNumber.Visibility = Visibility.Visible;
                            lnamewin.TbLNumber.Text = lm.LayoutCount.ToString(CultureInfo.InvariantCulture);
                        }
                        // Если нет имени
                        lnamewin.ChkAddNameToStamp.Visibility = string.IsNullOrEmpty(_lnamecoord) ? Visibility.Collapsed : Visibility.Visible;
                        bool b;
                        lnamewin.ChkAddNameToStamp.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkAddNameToStamp"), out b) && b;
                        lnamewin.ChkLNumber.IsChecked = bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "ChkLNumber"), out b) && b;
                        // scale
                        lnamewin.ChkLayoutScaleOneToOne.IsChecked = !bool.TryParse(MpSettings.GetValue("Settings", "mpFormats", "LayoutScaleOneToOne"), out b) || b;
                        if (lnamewin.ShowDialog() == true)
                        {
                            MpSettings.SetValue("Settings", "mpFormats", "ChkAddNameToStamp",
                                (lnamewin.ChkAddNameToStamp.IsChecked != null && lnamewin.ChkAddNameToStamp.IsChecked.Value).ToString(), false);
                            MpSettings.SetValue("Settings", "mpFormats", "ChkLNumber",
                                (lnamewin.ChkLNumber.IsChecked != null && lnamewin.ChkLNumber.IsChecked.Value).ToString(), false);
                            MpSettings.SetValue("Settings", "mpFormats", "LayoutScaleOneToOne",
                                (lnamewin.ChkLayoutScaleOneToOne.IsChecked != null && lnamewin.ChkLayoutScaleOneToOne.IsChecked.Value).ToString(), false);
                            if (lnamewin.ChkLayoutScaleOneToOne.IsChecked != null)
                                layoutScaleOneToOne = lnamewin.ChkLayoutScaleOneToOne.IsChecked.Value;

                            lname = lnamewin.TbLayoutName.Text;
                            if (lnamewin.ChkAddNameToStamp.IsChecked != null && lnamewin.ChkAddNameToStamp.IsChecked.Value)
                                _lname = lname;
                            if (lnamewin.ChkLNumber.IsChecked != null && lnamewin.ChkLNumber.IsChecked.Value)
                                _lnumber = lnamewin.TbLNumber.Text;
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
                returned = false;
            }
            return returned;
        }

        private void AddStamps(Point3d bottomLeftPt, Point3d topLeftPt, Point3d bottomRightPt, Vector3d replaceVector3D, double scale, Point3d blockInsertionPoint3D)
        {
            var docFor = ((ComboBoxItem) CbDocumentsFor.SelectedItem).Tag;
            if (ChkB1.IsChecked != null && ChkB1.IsChecked.Value)
            {
                if(docFor.Equals("RU"))
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
                    var stamp = CbTables.SelectedItem as Stamp;
                    if (stamp != null)
                        BtAddTable(stamp.TableStyle, "BottomRight", bottomRightPt, replaceVector3D, scale, blockInsertionPoint3D);
                }
            }
        }
        // Вставка штампа
        private void BtAddTable(string tableStyleName, string pointAligin, Point3d insertPt, Vector3d replaceVector3D, double scale, Point3d blockInsertionPoint3D)
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
                                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("ModPlus");
                                using (key)
                                {
                                    if (key != null)
                                    {
                                        // Директория расположения файла
                                        var dir = Path.Combine(key.GetValue("TopDir").ToString(), "Data", "Dwg");
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
                                                if (ent is Table)
                                                {
                                                    var tblsty = (Table)myT.GetObject(obj, OpenMode.ForWrite);
                                                    if (tblsty.TableStyleName.Equals(xmltbl.Attribute("tablestylename").Value))
                                                    {
                                                        tblIds.Add(tblsty.ObjectId);
                                                        var im = new IdMapping();
                                                        sourceDb.WblockCloneObjects(tblIds, db.CurrentSpaceId, im,
                                                            DuplicateRecordCloning.Ignore, false);
                                                        tbl =
                                                            (Table)
                                                                tr.GetObject(im.Lookup(tblsty.ObjectId).Value,
                                                                    OpenMode.ForWrite);
                                                        break;
                                                    }
                                                }
                                            }
                                            myT.Commit();
                                        }
                                        if (tbl.ObjectId == ObjectId.Null)
                                        {
                                            MpMsgWin.Show("В файле штампов не найден штамп с табличным стилем " +
                                                          tableStyleName + Environment.NewLine +
                                                          "Запустите функцию \"Штампы\" для перезаписи файла штампов");
                                            return;
                                        }

                                        var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                                        // Перемещаем 
                                        var mInsertPt = tbl.Position;
                                        var width = tbl.Width;
                                        var height = tbl.Height;

                                        if (pointAligin.Equals("TopLeft")) mInsertPt = insertPt;
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
                                        // Присваиваем свойства//
                                        /////////////////////////

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
                                                double d;
                                                cell.TextHeight =
                                                    double.TryParse(TbMainTextHeight.Text, out d)
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
                                                double d;
                                                cell.TextHeight =
                                                    double.TryParse(TbBigTextHeight.Text, out d)
                                                        ? d * scale
                                                        : 3.5 * scale;
                                            }
                                        }
                                        // Имя листа и номер
                                        if (xmltbl.Attribute("lnamecoordinates") != null)
                                            if (!string.IsNullOrEmpty(_lnamecoord))
                                                if (!string.IsNullOrEmpty(_lname))
                                                {
                                                    var rown = -1;
                                                    var coln = -1;
                                                    int.TryParse(_lnamecoord.Split(',').GetValue(0).ToString(), out rown);
                                                    int.TryParse(_lnamecoord.Split(',').GetValue(1).ToString(), out coln);
                                                    if (rown != -1 & coln != -1)
                                                        tbl.Cells[rown, coln].TextString = _lname;
                                                }
                                        if (xmltbl.Attribute("lnumbercooridnates") != null)
                                            if (!string.IsNullOrEmpty(_lnumbercoord))
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
                                                            tbl.Cells[rown, coln].TextString =
                                                                tbl.Cells[rown, coln].TextString + " " + _lnumber;
                                                    }
                                                }

                                        tr.Commit();
                                        break; // Тормозим Foreach
                                    }
                                    else
                                    {
                                        MpMsgWin.Show("Не найдена запись в реестре. Запустите конфигуратор!");
                                        return;
                                    }
                                }
                            }// transaction
                        }// if
                    }// foreach
                }
            }// try
            catch (Exception ex)
            {
                MpExWin.Show(ex);
            }
        }
        public static void AddFieldsToStamp(Table table)
        {
            try
            {
                var dbsi = AcApp.DocumentManager.MdiActiveDocument.Database.SummaryInfo;
                var dbsib = new DatabaseSummaryInfoBuilder(dbsi);
                // База данных штампов
                var doc = XElement.Parse(Properties.Resources.Stamps);
                foreach (var xmlTbl in doc.Elements("Stamp"))
                {
                    if (table.TableStyleName.Equals(xmlTbl.Attribute("tablestylename").Value))
                    {
                        // Даже если имя табличного стиля сошлось - проверяем по количеству ячеек!
                        if (int.Parse(xmlTbl.Attribute("cellcount").Value) != table.Cells.Count())
                        {
                            MpMsgWin.Show("Штамп не соответсвует штампу ModPlus!" + "Неверное кол-во ячеек");
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

                            if (MpSettings.XmlMpSettingsFile.Element("Settings") != null)
                                if (MpSettings.XmlMpSettingsFile?.Element("Settings")?.Element("UserSurnames") != null)
                                {
                                    var xElements = MpSettings.XmlMpSettingsFile?.Element("Settings")?.Element("UserSurnames")?.Elements("Surname");
                                    if (xElements != null)
                                        foreach (var sn in xElements)
                                        {
                                            surnames.Add(sn.Attribute("Surname").Value);
                                            surnameskeys.Add(sn.Attribute("Id").Value);
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
            }
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
            var cb = sender as ComboBox;
            if (cb == null) return;
            var index = cb.SelectedIndex;
            if (index == 0)
            {
                var uriSource = new Uri(@"/mpFormats_" + VersionData.FuncVersion + ";component/Resources/Preview/F_5.png", UriKind.Relative);
                Image_format.Source = new BitmapImage(uriSource);

                Image_b1.Margin = new Thickness(5, 0, 0, 3);
                Image_b2.Margin = new Thickness(0, 0, 0, 54);
                Image_stamp.Margin = new Thickness(0, 0, 3, 3);
            }
            else
            {
                var uriSource = new Uri(@"/mpFormats_" + VersionData.FuncVersion + ";component/Resources/Preview/F_10.png", UriKind.Relative);
                Image_format.Source = new BitmapImage(uriSource);

                Image_b1.Margin = new Thickness(5, 0, 0, 5);
                Image_b2.Margin = new Thickness(0, 0, 0, 56);
                Image_stamp.Margin = new Thickness(0, 0, 3, 5);
            }
        }

        private void ChkB1_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_b1.Opacity = 1;
        }

        private void ChkB1_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_b1.Opacity = 0.5;
        }

        private void ChkB2_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_b2.Opacity = 1;
        }

        private void ChkB2_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_b2.Opacity = 0.5;
        }

        private void ChkB3_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_top.Opacity = 1;
        }

        private void ChkB3_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_top.Opacity = 0.5;
        }

        private void ChkStamp_OnChecked(object sender, RoutedEventArgs e)
        {
            Image_stamp.Opacity = 1;
            GridStamp.Visibility = CbDocumentsFor.Visibility =//DpSurenames.Visibility = //TbLogo.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Visible;
            
        }

        private void ChkStamp_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Image_stamp.Opacity = 0.5;
            GridStamp.Visibility = CbDocumentsFor.Visibility =//DpSurenames.Visibility = //TbLogo.Visibility =
            CbLogo.Visibility = GridSplitterStamp.Visibility = Visibility.Collapsed;
        }
        #endregion

        private class Stamp
        {
            public string Name { get; set; }
            public string TableStyle { get; set; }
            public string Description { get; set; }
        }

        private void ShowFormatSize()
        {
            try
            {
                double dlina, visota;
                string side, orientation;
                if (RbShort.IsChecked != null && RbShort.IsChecked.Value) side = "Короткая";
                else
                    side = "Длинная";
                if (RbHorizontal.IsChecked != null && RbHorizontal.IsChecked.Value)
                    orientation = "Альбомный";
                else
                    orientation = "Книжный";

                var format = ((ListBoxItem)CbFormat.SelectedItem).Content.ToString();
                var multiplicity = CbMultiplicity.SelectedItem.ToString();

                MpFormatsAdd.GetFormatSize(format, orientation, side, multiplicity, out dlina, out visota);
                TbFormatSize.Text = dlina.ToString(CultureInfo.InvariantCulture) + " x " +
                                    visota.ToString(CultureInfo.InvariantCulture);
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
            var tab = sender as TabControl;
            if (tab != null)
            {
                if (tab.SelectedIndex != -1)
                    TbFormatSize.Visibility = tab.SelectedIndex == 0 ? Visibility.Visible : Visibility.Collapsed;
                else TbFormatSize.Visibility = Visibility.Visible;
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
                var ofd = new OpenFileDialog("Использовать dwg-файл в качестве логотипа", string.Empty, "dwg", "dialog",
                    OpenFileDialog.OpenFileDialogFlags.SearchPath |
                    OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var sourceDb = new Database(false, true);
                    //read file DB
                    sourceDb.ReadDwgFile(ofd.Filename, FileOpenMode.OpenTryForReadShare, true, string.Empty);
                    // Если файл более поздней версии, то будет ошибка
                    TbLogoFile.Text = ofd.Filename;
                }
                else
                {
                    MpMsgWin.Show("В файле" + Environment.NewLine + ofd.Filename + Environment.NewLine +
                                  "не найдено полей");
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exception)
            {
                if (exception.ErrorStatus == ErrorStatus.NotImplementedYet)
                {
                    MpMsgWin.Show("Файл создан в более поздней версии AutoCad!");
                }
                else MpExWin.Show(exception);
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
                        var sourceDb = new Database(false, true);
                        //read file DB
                        sourceDb.ReadDwgFile(TbLogoFile.Text, FileOpenMode.OpenTryForReadShare, true, string.Empty);

                        db.Insert(Path.GetFileNameWithoutExtension(TbLogoFile.Text), sourceDb, true);
                        blockName = Path.GetFileNameWithoutExtension(TbLogoFile.Text);
                    }
                }
                return blockName;
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
                return blockName;
            }
        }
        // Запрет пробела
        private void TboxesNoSpaceBar_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = (e.Key == Key.Space);
        }
        //  - без минуса
        private void Tb_OnlyNums_NoMinus_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var txt = ((TextBox)sender).Text + e.Text;
            e.Handled = !DoubleCharChecker(txt, false, true, null);
        }
        // Проверка, что число, точка или знак минус
        private static bool DoubleCharChecker(string str, bool checkMinus, bool checkDot, double? max)
        {
            var result = false;
            if (str.Count(c => c.Equals('.')) > 1) return false;
            if (str.Count(c => c.Equals('-')) > 1) return false;
            // Проверять нужно только последний знак в строке!!!
            var ch = str.Last();
            if (checkMinus)
                if (ch.Equals('-'))
                    result = str.IndexOf(ch) == 0;
            if (checkDot)
                if (ch.Equals('.'))
                    result = true;

            if (char.IsNumber(ch))
                result = true;
            // На "максимальность" проверяем если предыдущие провреки успешны
            if (max != null & result)
            {
                double d;
                if (double.TryParse(str, out d))
                    if (Math.Abs(d) > max) result = false;
            }
            return result;
        }

    }
    // Запуск функции и создание блока
    public class MpFormatsAdd
    {
        private MpFormats _mpFormats;

        [CommandMethod("ModPlus", "mpFormats", CommandFlags.Modal)]
        public void StartMpFormats()
        {
            if (_mpFormats == null)
            {
                _mpFormats = new MpFormats();
                _mpFormats.Closed += win_Closed;
            }

            if (_mpFormats.IsLoaded)
                _mpFormats.Activate();
            else
                AcApp.ShowModelessWindow(AcApp.MainWindow.Handle, _mpFormats, false);
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
            var doc = AcApp.DocumentManager.MdiActiveDocument;
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
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 1189;
                    //                visota = 841 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 841;
                    //                visota = 1189 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 841 * int.Parse(multiplicity);
                    //                visota = 1189;
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 1189 * int.Parse(multiplicity);
                    //                visota = 841;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            dlina = 841;
                    //            visota = 1189;
                    //        }
                    //        if (orientation.Equals("Альбомный"))
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
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 841;
                    //                visota = 594 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 594;
                    //                visota = 841 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 594 * int.Parse(multiplicity);
                    //                visota = 841;
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 841 * int.Parse(multiplicity);
                    //                visota = 594;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            dlina = 594;
                    //            visota = 841;
                    //        }
                    //        if (orientation.Equals("Альбомный"))
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
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 594;
                    //                visota = 420 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 420;
                    //                visota = 594 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 420 * int.Parse(multiplicity);
                    //                visota = 594;
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 594 * int.Parse(multiplicity);
                    //                visota = 420;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            dlina = 420;
                    //            visota = 594;
                    //        }
                    //        if (orientation.Equals("Альбомный"))
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
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 420;
                    //                visota = 297 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 297;
                    //                visota = 420 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 297 * int.Parse(multiplicity);
                    //                visota = 420;
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 420 * int.Parse(multiplicity);
                    //                visota = 297;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            dlina = 297;
                    //            visota = 420;
                    //        }
                    //        if (orientation.Equals("Альбомный"))
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
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 297;
                    //                visota = 210 * int.Parse(multiplicity);
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 210;
                    //                visota = 297 * int.Parse(multiplicity);
                    //            }
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            if (side.Equals("Короткая"))
                    //            {
                    //                dlina = 210 * int.Parse(multiplicity);
                    //                visota = 297;
                    //            }
                    //            if (side.Equals("Длинная"))
                    //            {
                    //                dlina = 297 * int.Parse(multiplicity);
                    //                visota = 210;
                    //            }
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (orientation.Equals("Книжный"))
                    //        {
                    //            dlina = 210;
                    //            visota = 297;
                    //        }
                    //        if (orientation.Equals("Альбомный"))
                    //        {
                    //            dlina = 297;
                    //            visota = 210;
                    //        }
                    //    }
                    //}
                    //if (format.Equals("A5"))
                    //{
                    //    if (orientation.Equals("Книжный"))
                    //    {
                    //        dlina = 148;
                    //        visota = 210;
                    //    }
                    //    if (orientation.Equals("Альбомный"))
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
                        if (bottomFrame.Equals("10мм"))
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
                                MessageBox.Show("Неверное имя блока");
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
                            MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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

                                MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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

                                    MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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

                                MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
            var doc = AcApp.DocumentManager.MdiActiveDocument;
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
                                MpMsgWin.Show("Неверное имя блока");
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
                            MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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
                                MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
                                br.XData = new ResultBuffer(
                                    new TypedValue(1001, "MP_FORMAT"),
                                    new TypedValue(1000, "MP_FORMAT"));

                                var entJig = new BlockJig(br);
                                var pr = ed.Drag(entJig);
                                if (pr.Status == PromptStatus.OK)
                                {
                                    replaceVector3D = entJig.ReplaceVector3D;
                                    var ent = entJig.GetEntity();
                                    MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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
                                MpCadHelpers.AddRegAppTableRecord("MP_FORMAT");
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                var peo = new PromptEntityOptions("\n" + "Выберите форматку для замены: ");
                peo.SetRejectMessage("\n" + "Неверный выбор!");
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;
                var peo = new PromptEntityOptions("\n" + "Выберите форматку для замены: ");
                peo.SetRejectMessage("\n" + "Неверный выбор!");
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
            catch (Exception ex)
            {
                MpExWin.Show(ex);
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
                    if (orientation.Equals("Книжный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 1189;
                            visota = 841 * int.Parse(multiplicity);
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 841;
                            visota = 1189 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 841 * int.Parse(multiplicity);
                            visota = 1189;
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 1189 * int.Parse(multiplicity);
                            visota = 841;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals("Книжный"))
                    {
                        dlina = 841;
                        visota = 1189;
                    }
                    if (orientation.Equals("Альбомный"))
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
                    if (orientation.Equals("Книжный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 841;
                            visota = 594 * int.Parse(multiplicity);
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 594;
                            visota = 841 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 594 * int.Parse(multiplicity);
                            visota = 841;
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 841 * int.Parse(multiplicity);
                            visota = 594;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals("Книжный"))
                    {
                        dlina = 594;
                        visota = 841;
                    }
                    if (orientation.Equals("Альбомный"))
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
                    if (orientation.Equals("Книжный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 594;
                            visota = 420 * int.Parse(multiplicity);
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 420;
                            visota = 594 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 420 * int.Parse(multiplicity);
                            visota = 594;
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 594 * int.Parse(multiplicity);
                            visota = 420;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals("Книжный"))
                    {
                        dlina = 420;
                        visota = 594;
                    }
                    if (orientation.Equals("Альбомный"))
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
                    if (orientation.Equals("Книжный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 420;
                            visota = 297 * int.Parse(multiplicity);
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 297;
                            visota = 420 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 297 * int.Parse(multiplicity);
                            visota = 420;
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 420 * int.Parse(multiplicity);
                            visota = 297;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals("Книжный"))
                    {
                        dlina = 297;
                        visota = 420;
                    }
                    if (orientation.Equals("Альбомный"))
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
                    if (orientation.Equals("Книжный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 297;
                            visota = 210 * int.Parse(multiplicity);
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 210;
                            visota = 297 * int.Parse(multiplicity);
                        }
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        if (side.Equals("Короткая"))
                        {
                            dlina = 210 * int.Parse(multiplicity);
                            visota = 297;
                        }
                        if (side.Equals("Длинная"))
                        {
                            dlina = 297 * int.Parse(multiplicity);
                            visota = 210;
                        }
                    }
                }
                else
                {
                    if (orientation.Equals("Книжный"))
                    {
                        dlina = 210;
                        visota = 297;
                    }
                    if (orientation.Equals("Альбомный"))
                    {
                        dlina = 297;
                        visota = 210;
                    }
                }
            }
            if (format.Equals("A5"))
            {
                if (orientation.Equals("Книжный"))
                {
                    dlina = 148;
                    visota = 210;
                }
                if (orientation.Equals("Альбомный"))
                {
                    dlina = 210;
                    visota = 148;
                }
            }
        }
    }

    class BlockJig : EntityJig
    {
        Point3d _mCenterPt, _mActualPoint;

        public BlockJig(BlockReference br)
            : base(br)
        {
            _mCenterPt = br.Position;
        }

        public Vector3d ReplaceVector3D { get; private set; }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var jigOpts = new JigPromptPointOptions
            {
                UserInputControls = (UserInputControls.Accept3dCoordinates
                | UserInputControls.NoZeroResponseAccepted
                | UserInputControls.AcceptOtherInputString
                | UserInputControls.NoNegativeResponseAccepted),
                Message = "\n" + "Точка вставки: "
            };
            var dres = prompts.AcquirePoint(jigOpts);
            if (_mActualPoint == dres.Value)
                return SamplerStatus.NoChange;
            _mActualPoint = dres.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            try
            {
                ((BlockReference)Entity).Position = _mActualPoint;
                ReplaceVector3D = _mCenterPt.GetVectorTo(_mActualPoint);

            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public Entity GetEntity()
        {
            return Entity;
        }

    }
}
