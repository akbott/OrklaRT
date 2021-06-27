using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;

namespace SelectionPane
{
    class ComboBoxWin8: ComboBox
    {
        public ComboBoxWin8()
        {
            Loaded += ComboBoxWin8_Loaded;
        }

        void ComboBoxWin8_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            ControlTemplate ct = Template;
            Border border = ct.FindName("Border", this) as Border;

            if(border != null)
            {
                border.Background = Background;
                BindingExpression be = GetBindingExpression(ComboBoxWin8.BackgroundProperty);
                if(be != null)
                {
                    border.SetBinding(Border.BackgroundProperty, be.ParentBindingBase);
                }
            }
        }
    }
}
