using System.Windows.Controls;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class DispatcherBoardView : UserControl
    {
        public DispatcherBoardView()
        {
            InitializeComponent();
            DataContext = new DispatcherBoardViewModel();
        }
    }
}
