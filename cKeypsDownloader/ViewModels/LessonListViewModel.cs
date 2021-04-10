using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using cKeypsDownloader.Services;
using DynamicData;
using keypsDownloaderCore.Models;
using NodaTime;
using ReactiveUI;

namespace cKeypsDownloader.ViewModels {
    public enum DownloadState {
        NotDownloading,
        Downloading,
        DownloadFinished
    }

    public class LessonForUi {
        public Lesson Lesson { get; }
        public bool IsChecked { get; set; }

        public DownloadState DownloadState { get; set; }

        public LessonForUi(Lesson l) {
            Lesson = l;
            IsChecked = false;
        }

        public LessonForUi ShallowCopy() {
            return (LessonForUi) this.MemberwiseClone();
        }
    }

    public class LessonListViewModel : ViewModelBase {
        private readonly ReadOnlyObservableCollection<LessonForUi> _items;
        public ReadOnlyObservableCollection<LessonForUi> Items => _items;

        public ObservableCollection<LessonForUi> Downloads { get; }

        private KeypsCoreCommunicator KeypsCommunicator { get; set; }

        public ReactiveCommand<Unit, Unit> DownloadLessonsButtonClick { get; }
        public ReactiveCommand<Unit, Unit> SingleKeypsLessonDownloadButtonOnClick { get; }


        private async void DownloadLessonsAsync() {
            var selectedLessons = _items.Where(x => x.IsChecked).ToList();

            // Eğer zaten indiriliyorsa tekrar indirme
            foreach (var selected in selectedLessons) {
                if (Downloads.Any(downloading => selected.Lesson.Id == downloading.Lesson.Id)) {
                    selectedLessons.Remove(selected);
                }
                else {
                    selected.DownloadState = DownloadState.Downloading;
                }

                selected.IsChecked = false;
            }

            Downloads.AddRange(selectedLessons);
            // await KeypsCommunicator.DownloadLessonsAsync(selectedLessons.Select(x => x.Lesson), Settings.DownloadPath);
            await KeypsCommunicator.DownloadLessonsAsync(selectedLessons.Select(x => x.Lesson),
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
        }

        private async void DownloadSingleLessonAsync() {
            if (string.IsNullOrEmpty(_singleLessonUrl)) return;
            var lesson = new Lesson(-1, "placeholderName", "", _singleLessonUrl, "https://big.kapitta.com",
                new LocalDateTime());
            await KeypsCommunicator.DownloadSingleLessonAsync(lesson, Settings.DownloadPath);
        }

        private string _singleLessonUrl = "";

        public string SingleLessonUrl {
            get => _singleLessonUrl;
            set => this.RaiseAndSetIfChanged(ref _singleLessonUrl, value);
        }

        private bool _showProgressBar;

        public bool ShowProgressBar {
            get => _showProgressBar;
            set => this.RaiseAndSetIfChanged(ref _showProgressBar, value);
        }

        private bool _showDownloadButton;

        public bool ShowDownloadButton {
            get => _showDownloadButton;
            set => this.RaiseAndSetIfChanged(ref _showDownloadButton, value);
        }


        public LessonListViewModel() {
            KeypsCommunicator = new KeypsCoreCommunicator();
            KeypsCommunicator
                .Connect()
                .ObserveOn(RxApp.MainThreadScheduler)
                .Transform(x => new LessonForUi(x))
                .Bind(out _items)
                .Subscribe();

            KeypsCommunicator.DownloadedLessons.Subscribe(x => {
                var lessonToBeReplaced = Downloads.First(l => l.Lesson.Id == x.Id);
                Downloads.Remove(lessonToBeReplaced);
                lessonToBeReplaced.DownloadState = DownloadState.DownloadFinished;
                Downloads.Add(lessonToBeReplaced);
            });

            Downloads = new ObservableCollection<LessonForUi>();

            DownloadLessonsButtonClick = ReactiveCommand.Create(DownloadLessonsAsync);
            SingleKeypsLessonDownloadButtonOnClick = ReactiveCommand.Create(DownloadSingleLessonAsync);
        }

        public async void GetLessonList(string keypsFileName) {
            ShowProgressBar = true;
            ShowDownloadButton = false;
            await KeypsCommunicator.GetKeypsLessonList(keypsFileName);
            ShowProgressBar = false;
            ShowDownloadButton = true;
        }
    }
}