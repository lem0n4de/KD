using System;
using System.Collections.Generic;
using System.Text;
using NodaTime;

namespace keypsDownloaderCore.Models {
    public enum Grade {
        First,
        Second,
        Third,
        Fourth,
        Fifth,
        Sixth,
        Undefined
    }

    public class Lesson {
        public int Id { get; set; }
        public string Alan { get; set; }
        public string Name { get; set; }
        public string Url { get; set; }
        public string Teacher { get; set; }
        public Grade Grade { get; set; }
        public LocalDateTime Date { get; set; }

        public string BaseKapittaUrl { get; set; }

        public Lesson(int id, string name, string alan, string url, string baseKapittaUrl, LocalDateTime date,
            Grade grade = Grade.Undefined, string teacher = "") {
            Id = id;
            Name = name;
            Alan = alan;
            Url = url;
            BaseKapittaUrl = baseKapittaUrl;
            Teacher = teacher;
            Grade = grade;
            Date = date;
        }
    }
}