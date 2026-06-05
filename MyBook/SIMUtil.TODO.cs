namespace MyBook
{
    // TODO: Fetch SMS messages from a SIM modem.
    partial class SIMUtil
    {
        public Task<List<SIMMessage>> FetchSIMMessages()
        {
            throw new NotImplementedException("TODO: Fetch SMS messages from a SIM modem.");
        }
    }

    public sealed record SIMMessage(DateTime Time, string Sender, string Text);
}
