using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ModPlusAPI.Windows;
using ModPlusAPI.Windows.Helpers;

namespace mpFormats
{
    public partial class LayoutName
    {
        private const string LangItem = "mpFormats";
        readonly List<string> _wrongSymbols = new List<string>
        {
            ">","<","/","\\","\"",":",";","?","*","|",",","=","`"
        };
        public LayoutName(string layOutName)
        {
            InitializeComponent();
            this.OnWindowStartUp();
            TbLayoutName.Text = layOutName;
        }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                DialogResult = false;
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
