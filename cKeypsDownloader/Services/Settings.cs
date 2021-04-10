using System;
using Microsoft.Extensions.Configuration;

namespace cKeypsDownloader.Services {
    public class Settings {
        public string DownloadPath { get; set; }
        public string LastKnownKeypsFilePath { get; set; }

        public static Settings InitSettings() {
            var builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", true, true);

            var cfg = builder.Build();
            var newCfg = cfg.Get<Settings>();
            // if (string.IsNullOrEmpty(newCfg.DownloadPath))
            //     newCfg.DownloadPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            return newCfg;
        }
    }
}