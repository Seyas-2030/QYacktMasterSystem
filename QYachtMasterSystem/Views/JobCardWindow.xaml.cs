using System.Windows;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class JobCardWindow : Window
    {
        public JobCardWindow(int customerId)
        {
            InitializeComponent();
            DataContext = new JobCardViewModel(customerId);
        }
    }
}
