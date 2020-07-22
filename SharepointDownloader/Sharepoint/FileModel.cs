using System;

namespace SharepointDownloader.Sharepoint
{
    public class FileModel
    {
        public string FileName { get; set; }
        public byte[] FileData { get; set; }
        public DateTime DateCreated { get; set; }
        public DateTime DateModified { get; set; }
        public string FileStorageIdentifier { get; set; }
    }
}
