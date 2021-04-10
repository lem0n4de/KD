using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace cKeypsDownloader.Views
{
    public class LessonListView : UserControl
    {
        public LessonListView()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }
    }
}