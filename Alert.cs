namespace SplunkMailProcessor
{
    public class Alert
    {
        public string timestamp { get; set; }
        public string id { get; set; }
        public string description { get; set; }
        public string message { get; set; }
        public string generatedBy { get; set; }
    }
}
