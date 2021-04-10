namespace keypsDownloaderCore {
    class ColumnInfo {
        public string IdColumn = null;
        public string DateColumn = null;
        public string TimeColumn = null;
        public string GradeColumn = null;
        public string AlanColumn = null;
        public string KonuColumn = null;
        public string TeacherColumn = null;
        public string MeetingIdColumn = null;
        public string KapittaColumn = null;

        internal bool IsCompleted() {
            if (IdColumn != null && DateColumn != null && TimeColumn != null && GradeColumn != null && AlanColumn != null &&
                KonuColumn != null && TeacherColumn != null && MeetingIdColumn != null && KapittaColumn != null) return true;
            return false;
        }
    }
}
