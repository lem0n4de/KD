using System.Collections.Generic;
using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;

namespace cKeypsDownloader.Views {
    public class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
// #if DEBUG
//             this.AttachDevTools();
// #endif
        }


        private void InitializeComponent() {
            AvaloniaXamlLoader.Load(this);
        }
    }
}
