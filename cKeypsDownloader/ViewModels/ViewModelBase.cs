using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Text;
using cKeypsDownloader.Services;

namespace cKeypsDownloader.ViewModels {
    public class ViewModelBase : ReactiveObject {
        protected readonly Settings Settings = Settings.InitSettings();
    }
}
