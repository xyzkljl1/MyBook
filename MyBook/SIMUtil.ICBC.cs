namespace MyBook
{
    partial class SIMUtil
    {
        private static bool IsICBCSender(string sender)
        {
            return String.Equals(sender, "95588", StringComparison.OrdinalIgnoreCase);
        }

        private static void ParseICBCSIMMessage(SIMMessage message)
        {
            // TODO: Parse ICBC SMS messages into records.
        }
    }
}
