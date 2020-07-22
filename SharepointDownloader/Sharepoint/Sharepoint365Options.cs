namespace SharepointDownloader.Sharepoint
{
    public class Sharepoint365Options
    {
        public string Url { get; set; }
        public string RootFolder { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public int ConnectionTimeout { get; set; }
    }
}
