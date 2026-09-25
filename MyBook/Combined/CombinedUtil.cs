namespace MyBook;

partial class CombinedUtil(DatabaseUtil database, PlaidUtil plaid, MailUtil mail)
{
    private readonly DatabaseUtil database = database;
    private readonly PlaidUtil plaid = plaid;
    private readonly MailUtil mail = mail;
}
