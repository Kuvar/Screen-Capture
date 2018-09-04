using System.Windows;
using System.Windows.Interactivity;

namespace ScreenCaptureApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new CaptureScreenViewModel();
        }
    }
}
