using System;
using System.Collections.Generic;
using System.IO;
using System.Reactive;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using cKeypsDownloader.Views;
using ReactiveUI;

namespace cKeypsDownloader.ViewModels {
    public class MainWindowViewModel : ViewModelBase {
        private string _keypsFileName = "Keyps dosya adı";

        public MainWindowViewModel() {
            LessonListViewModel = new LessonListViewModel();
            KeypsFileDialogButtonOnClick = ReactiveCommand.Create(GetKeypsFileNameAsync);
            ShowSettingsWindow = ReactiveCommand.Create(_showSettingsWindow);
        }
        
        public ReactiveCommand<Unit, Unit> ShowSettingsWindow { get; }

        private void _showSettingsWindow() {
            var setWin = new SettingsWindow();
            setWin.Show();
        }

        public LessonListViewModel LessonListViewModel { get; }

        public string KeypsFileName {
            get => _keypsFileName;
            set => this.RaiseAndSetIfChanged(ref _keypsFileName, value);
        }

        public ReactiveCommand<Unit, Unit> KeypsFileDialogButtonOnClick { get; }

        
        private async void GetKeypsFileNameAsync() {
            var fileDialog = new OpenFileDialog {AllowMultiple = false};
            var fileDialogFilter = new FileDialogFilter() {Name = "Excel", Extensions = new List<string>() {"xlsx"}};
            fileDialog.Filters.Add(fileDialogFilter);

            if (!(Application.Current.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop))
                return;
            var selectedFiles = await fileDialog.ShowAsync(desktop.MainWindow);
            if (selectedFiles.Length <= 0) return;
            KeypsFileName = selectedFiles[0];
            
            LessonListViewModel.GetLessonList(KeypsFileName);
        }
    }
}