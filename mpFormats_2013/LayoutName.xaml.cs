using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ModPlusAPI.Windows;

namespace mpFormats
{
    public partial class LayoutName
    {
        private const string LangItem = "mpFormats";

        private readonly List<string> _wrongSymbols = new List<string>
        {
            ">","<","/","\\","\"",":",";","?","*","|",",","=","`"
        };
        public LayoutName(string layOutName)
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h42");
            TbLayoutName.Text = layOutName;
        }

        private void BtAccept_OnClick(object sender, RoutedEventArgs e)
        {
            if (_wrongSymbols.Any(wrongSymbol => TbLayoutName.Text.Contains(wrongSymbol)))
            {
                ModPlusAPI.Windows.MessageBox.Show(
                    ModPlusAPI.Language.GetItem(LangItem, "err3") + " " + TbLayoutName.Text + Environment.NewLine +
                    ModPlusAPI.Language.GetItem(LangItem, "err4") + " " + string.Join("", _wrongSymbols.ToArray()), MessageBoxIcon.Alert);
                return;
            }
            DialogResult = true;
        }

        private void ChkLNumber_OnChecked(object sender, RoutedEventArgs e)
        {
            TbLNumber.IsEnabled = true;
        }

        private void ChkLNumber_OnUnchecked(object sender, RoutedEventArgs e)
        {
            TbLNumber.IsEnabled = false;
        }
    }
}
