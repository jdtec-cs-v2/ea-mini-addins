using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Input;

namespace MiniAddins.View
{
    /// <summary>
    /// ModelExportView.xaml 的交互逻辑
    /// </summary>
    public partial class ModelExportView : UserControl
    {
        public ModelExportView()
        {
            InitializeComponent();
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsValid(((TextBox)sender).Text + e.Text);
        }

        public static bool IsValid(string str)
        {
            int i;
            return int.TryParse(str, out i) && i >= 1 && i <= 9999;
        }
    }
}
