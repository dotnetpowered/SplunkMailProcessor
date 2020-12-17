namespace SplunkMailProcessor
{
    public interface IAlertGenerator
    {
        Alert GetAlertFromMessage(MailItem mailItem);
    }
}
