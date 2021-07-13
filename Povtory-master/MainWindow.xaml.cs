
using Povtory.ViewModels;
using System.Windows;
using System.Windows.Input;

namespace Povtory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new BLViewModel(Comparer.DI.ServiceModule.Instance().ComparerSvc);
        }
    }
}
