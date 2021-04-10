namespace cKeypsDownloader.ViewModels {
    public class SettingsWindowViewModel : ViewModelBase {
        public string DownloadPath { get; set; }
        public string LastKnownKeypsFilePath { get; set; }

        public SettingsWindowViewModel() {
            DownloadPath = Settings.DownloadPath;
            LastKnownKeypsFilePath = Settings.LastKnownKeypsFilePath;
        }
    }
}