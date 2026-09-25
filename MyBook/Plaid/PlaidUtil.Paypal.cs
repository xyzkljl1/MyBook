using Newtonsoft.Json.Linq;

namespace MyBook;

partial class PlaidUtil
{
    internal const string PayPalInstitutionId = "ins_22";

    internal async Task<TransactionSyncData> ReadPayPalAsync(PlaidItem item, string? cursor, CancellationToken cancellationToken)
    {
        if (item.institutionId != PayPalInstitutionId)
            throw new PlaidRequestException("PayPal: unexpected Item institution.");
        GetLinkedAccount(item, "PAYPAL");
        var data = await GetTransactionUpdatesAsync(item, cursor, cancellationToken).ConfigureAwait(false);
        var accounts = Rows(data.Accounts, "accounts").ToList();
        if (accounts.Count != 1 || Text(accounts[0], "subtype") != "paypal")
            throw new PlaidRequestException("PayPal: Item must return exactly one PayPal account.");
        return data;
    }
}
