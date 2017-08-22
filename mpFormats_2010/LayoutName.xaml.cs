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
                ModPlusAPI.Windows.MessageBox.Show("Недопустимые символы в имени листа: " + TbLayoutName.Text + Environment.NewLine +
                              "Не разрешается использование следующих символов: " + string.Join("",_wrongSymbols.ToArray()), MessageBoxIcon.Alert);
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
