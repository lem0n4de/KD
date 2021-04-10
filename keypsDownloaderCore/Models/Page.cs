namespace keypsDownloaderCore.Models {
    public class Page {
        public int PageNumber { get; set; }
        public string Url { get; set; }
        public string FileName { get; set; }
        
        public Page() {}

        public Page(string url, int number) {
            PageNumber = number;
            Url = url;
        }
    }
}