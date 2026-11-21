using System.Security.Cryptography;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using SteamKit2;
using SteamKit2.Authentication;

namespace MyBook;

partial class WebUtil
{
    // Supplying credentials explicitly starts a new authorization. Otherwise only the saved token is used.
    public async Task<SteamLoginInfo> LoginSteamAsync(string? loginName = null,
        string? password = null, IAuthenticator? authenticator = null, CancellationToken cancellationToken = default)
        => await WithSteamSessionAsync((readInfo, _, _) => readInfo(), loginName, password, authenticator, cancellationToken).ConfigureAwait(false);

    public Task RefreshSteamSessionAsync() => WithSteamSessionAsync((_, _, _) => Task.FromResult(true));

    private async Task<T> WithSteamSessionAsync<T>(Func<Func<Task<SteamLoginInfo>>, string, CancellationToken, Task<T>> fetch,
        string? loginName = null, string? password = null, IAuthenticator? authenticator = null, CancellationToken cancellationToken = default)
    {
        var authorize = password is not null;
        if (authorize && (String.IsNullOrWhiteSpace(loginName) || String.IsNullOrWhiteSpace(password) || authenticator is null))
            throw new ArgumentException("Initial Steam authorization requires a login name, password and Steam Guard handler.");
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            using var sessionStore = database.OpenLoginSession(LoginProvider.Steam);
            var saved = authorize ? null : JsonSerializer.Deserialize<SteamSessionState>(sessionStore.Read()
                ?? throw new InvalidOperationException("No saved Steam session; complete initial authorization first."));
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var timeoutDuration = TimeSpan.FromMinutes(3);
            var elapsed = System.Diagnostics.Stopwatch.StartNew();
            timeout.CancelAfter(timeoutDuration);
            var operation = SteamSessionOperation.Connect;
            var client = new SteamClient(SteamConfiguration.Create(builder => builder
                .WithProtocolTypes(ProtocolTypes.WebSocket)
                .WithHttpClientFactory(_ => new HttpClient(CreateSteamHttpHandler()))));
            var manager = new CallbackManager(client);
            var user = client.GetHandler<SteamUser>()!;
            var connected = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var loggedOn = new TaskCompletionSource<SteamUser.LoggedOnCallback>(TaskCreationOptions.RunContinuationsAsynchronously);
            var account = new TaskCompletionSource<SteamUser.AccountInfoCallback>(TaskCreationOptions.RunContinuationsAsynchronously);
            var wallet = new TaskCompletionSource<SteamUser.WalletInfoCallback>(TaskCreationOptions.RunContinuationsAsynchronously);
            var disconnected = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var loggedOff = new TaskCompletionSource<EResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            using var connectionSubscription = manager.Subscribe<SteamClient.ConnectedCallback>(_ => connected.TrySetResult());
            using var loginSubscription = manager.Subscribe<SteamUser.LoggedOnCallback>(value => loggedOn.TrySetResult(value));
            using var accountSubscription = manager.Subscribe<SteamUser.AccountInfoCallback>(value => account.TrySetResult(value));
            using var walletSubscription = manager.Subscribe<SteamUser.WalletInfoCallback>(value => wallet.TrySetResult(value));
            using var logoutSubscription = manager.Subscribe<SteamUser.LoggedOffCallback>(value => loggedOff.TrySetResult(value.Result));
            using var disconnectSubscription = manager.Subscribe<SteamClient.DisconnectedCallback>(value =>
            {
                disconnected.TrySetResult(value.UserInitiated);
                timeout.Cancel();
            });
            var pump = Task.Run(async () =>
            {
                try
                {
                    while (!timeout.IsCancellationRequested)
                    {
                        manager.RunCallbacks();
                        await Task.Delay(50, timeout.Token).ConfigureAwait(false);
                    }
                }
                finally { timeout.Cancel(); }
            });
            try
            {
                client.Connect();
                await connected.Task.WaitAsync(timeout.Token).ConfigureAwait(false);
                if (authorize)
                {
                    operation = SteamSessionOperation.Authorize;
                    var auth = await client.Authentication.BeginAuthSessionViaCredentialsAsync(new AuthSessionDetails
                    {
                        Username = loginName!, Password = password!, Authenticator = authenticator!,
                        IsPersistentSession = true
                    }).WaitAsync(timeout.Token).ConfigureAwait(false);
                    var result = await auth.PollingWaitForResultAsync(timeout.Token).ConfigureAwait(false);
                    saved = new SteamSessionState
                    {
                        loginName = result.AccountName, refreshToken = result.RefreshToken
                    };
                }
                operation = SteamSessionOperation.LogOn;
                user.LogOn(new SteamUser.LogOnDetails
                {
                    Username = saved!.loginName, AccessToken = saved.refreshToken, ShouldRememberPassword = true,
                    LoginID = (uint)RandomNumberGenerator.GetInt32(1, Int32.MaxValue), MachineName = "MyBook"
                });
                var login = await loggedOn.Task.WaitAsync(timeout.Token).ConfigureAwait(false);
                if (login.Result != EResult.OK)
                    throw new InvalidOperationException($"Steam ClientLogOn: {login.Result}; extended={login.ExtendedResult}. Reauthorize if the saved session is invalid.");
                // Commit credentials independently of later account queries, so their failure cannot lose a new token.
                if (authorize) sessionStore.Save(JsonSerializer.Serialize(saved), newSession: true);
                operation = SteamSessionOperation.RefreshToken;
                var renewed = await client.Authentication.GenerateAccessTokenForAppAsync(client.SteamID!,
                    saved.refreshToken, allowRenewal: true).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!String.IsNullOrEmpty(renewed.RefreshToken) && renewed.RefreshToken != saved.refreshToken)
                {
                    saved.refreshToken = renewed.RefreshToken;
                    sessionStore.Save(JsonSerializer.Serialize(saved));
                }
                async Task<SteamLoginInfo> ReadInfo()
                {
                    operation = SteamSessionOperation.ReadAccount;
                    var info = await account.Task.WaitAsync(timeout.Token).ConfigureAwait(false);
                    operation = SteamSessionOperation.ReadWallet;
                    var funds = await wallet.Task.WaitAsync(timeout.Token).ConfigureAwait(false);
                    operation = SteamSessionOperation.Fetch;
                    return new SteamLoginInfo(client.SteamID!.ConvertToUInt64(), info.PersonaName,
                        funds.HasWallet, funds.Currency, funds.LongBalance / 100m, funds.LongBalanceDelayed / 100m);
                }
                operation = SteamSessionOperation.Fetch;
                return await fetch(ReadInfo, renewed.AccessToken, timeout.Token).ConfigureAwait(false);
            }
            catch (Exception error)
            {
                // Capture before cleanup cancels the token and deliberately disconnects the client.
                // Only typed protocol results and timing are safe for the local failure log.
                error.Data[typeof(SteamFailureDiagnostic)] = new SteamFailureDiagnostic(
                    operation, elapsed.Elapsed, connected.Task.IsCompletedSuccessfully,
                    timeout.IsCancellationRequested, elapsed.Elapsed >= timeoutDuration,
                    cancellationToken.IsCancellationRequested,
                    disconnected.Task.IsCompletedSuccessfully,
                    disconnected.Task.IsCompletedSuccessfully ? disconnected.Task.Result : null,
                    loggedOn.Task.IsCompletedSuccessfully ? loggedOn.Task.Result.Result : null,
                    loggedOn.Task.IsCompletedSuccessfully ? loggedOn.Task.Result.ExtendedResult : null,
                    loggedOff.Task.IsCompletedSuccessfully ? loggedOff.Task.Result : null,
                    error is AuthenticationException authError ? authError.Result : null);
                throw;
            }
            finally
            {
                timeout.Cancel();
                client.Disconnect();
                try { await pump.ConfigureAwait(false); } catch (OperationCanceledException) { }
            }
        }
        catch (AuthenticationException e)
        {
            var error = new InvalidOperationException($"Steam Authentication: {e.Result}.");
            error.Data[typeof(SteamFailureDiagnostic)] = e.Data[typeof(SteamFailureDiagnostic)];
            throw error;
        }
        catch (OperationCanceledException e) when (!cancellationToken.IsCancellationRequested)
        {
            var error = new InvalidOperationException("Steam login/account query timed out or disconnected.");
            error.Data[typeof(SteamFailureDiagnostic)] = e.Data[typeof(SteamFailureDiagnostic)];
            throw error;
        }
    }

    internal enum SteamSessionOperation { Connect, Authorize, LogOn, RefreshToken, ReadAccount, ReadWallet, Fetch }

    internal sealed record SteamFailureDiagnostic(SteamSessionOperation Operation, TimeSpan Elapsed,
        bool Connected, bool CancellationRequested, bool DeadlineElapsed, bool ExternalCancellation,
        bool Disconnected, bool? DisconnectUserInitiated, EResult? LogOnResult, EResult? LogOnExtendedResult,
        EResult? LogOffResult, EResult? AuthenticationResult);

    private HttpClientHandler CreateSteamHttpHandler()
    {
        var value = config["steam_proxy"];
        Uri? proxy = null;
        if (!String.IsNullOrWhiteSpace(value) && (!Uri.TryCreate(value, UriKind.Absolute, out proxy)
            || proxy.Scheme is not ("http" or "socks5") || proxy.UserInfo.Length > 0))
            throw new InvalidOperationException("Steam: invalid steam_proxy.");
        return new HttpClientHandler { UseProxy = proxy is not null, Proxy = proxy is null ? null : new WebProxy(proxy) };
    }

    private sealed class SteamSessionState
    {
        public required string loginName { get; set; }
        public required string refreshToken { get; set; }
    }

    public sealed record SteamLoginInfo(ulong SteamId, string Nickname, bool HasWallet,
        ECurrencyCode Currency, decimal Balance, decimal PendingBalance);
}
