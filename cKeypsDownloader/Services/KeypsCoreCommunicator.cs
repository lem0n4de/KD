using System;
using System.Collections.Generic;
using System.Reactive.Subjects;
using System.Threading.Tasks;
using DynamicData;
using keypsDownloaderCore;
using keypsDownloaderCore.Models;

namespace cKeypsDownloader.Services {
    public class KeypsCoreCommunicator {
        private readonly SourceList<Lesson> _items = new();

        public KeypsCoreCommunicator() {
            DownloadedLessons = new Subject<Lesson>();
        }
        

        public IObservable<IChangeSet<Lesson>> Connect() => _items.Connect();

        public Subject<Lesson> DownloadedLessons { get; }

        public async Task GetKeypsLessonList(string keypsFileName) {
            KeypsWorksheetParser kwp = new KeypsWorksheetParser(keypsFileName);
            var x = await kwp.ParseAsync();
            _items.AddRange(x);
        }

        public async Task DownloadLessonsAsync(IEnumerable<Lesson> selectedLessons, string downloadPath) {
            var ld = new LessonDownloader(downloadPath);
            ld.DownloadedLessons.Subscribe(x => DownloadedLessons.OnNext(x));
            await ld.DownloadLessonsAsync(selectedLessons);
        }

        public async Task DownloadSingleLessonAsync(Lesson lesson, string downloadPath) {
            var ld = new LessonDownloader(downloadPath);
            await ld.DownloadLesson(lesson);
        }
    }
}