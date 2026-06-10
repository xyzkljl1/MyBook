using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;

namespace MyBook
{
    // Reads already-stored SMS messages from a USB SIM modem through GSM AT commands.
    // Developed/tested with Quectel EC200A-CNDA (firmware EC200ACNDAR01A19M16).
    partial class SIMUtil
    {
        private static readonly int[] CandidateBaudRates = [115200, 9600, 57600, 38400, 19200];
        private static readonly string[] CandidateMessageStores = ["", "MT", "SM", "ME"];
        private static readonly string[] ProcessingMessageStores = ["ME", "SM", "MT", ""];
        private static readonly bool DeleteProcessedMessages = false; // Temporarily disabled while developing SMS parsers.
        private static readonly TimeSpan DefaultCommandTimeout = TimeSpan.FromSeconds(4);
        private static readonly TimeSpan ListMessagesTimeout = TimeSpan.FromSeconds(20);
        private readonly DatabaseUtil? database;
        private string? cachedPortName;
        private int cachedBaudRate;
        private bool hasFoundSIMModem;

        public SIMUtil()
        {
        }

        public SIMUtil(DatabaseUtil database)
        {
            this.database = database;
        }

        public Task<List<SIMMessage>> FetchSIMMessages(string? portName = null)
        {
            return Task.Run(() => FetchSIMMessagesCore(portName));
        }

        public Task<SIMPollResult> PollConfiguredSIMMessages(string expectedImsi)
        {
            return Task.Run(() => PollConfiguredSIMMessagesCore(expectedImsi));
        }

        private static List<SIMMessage> FetchSIMMessagesCore(string? portName)
        {
            var portNames = GetCandidatePorts(portName);
            if (portNames.Count == 0)
                throw new InvalidOperationException("No serial ports were found for a SIM modem.");

            var failures = new List<string>();
            var foundResponsiveModem = false;
            foreach (var candidatePortName in portNames)
            {
                foreach (var baudRate in CandidateBaudRates)
                {
                    try
                    {
                        using var port = OpenPort(candidatePortName, baudRate);
                        InitializeModem(port);
                        foundResponsiveModem = true;
                        var messages = ReadStoredMessages(port, candidatePortName);
                        var orderedMessages = messages
                            .OrderBy(message => message.Time)
                            .ThenBy(message => message.Index)
                            .ToList();
                        if (orderedMessages.Count > 0 || !String.IsNullOrWhiteSpace(portName))
                            return orderedMessages;

                        break;
                    }
                    catch (Exception exception) when (IsProbeException(exception))
                    {
                        failures.Add($"{candidatePortName}@{baudRate}: {exception.Message}");
                    }
                }
            }

            if (foundResponsiveModem)
                return [];

            throw new InvalidOperationException(
                "No responsive SIM modem was found. Tried: " + String.Join("; ", failures.Take(8)));
        }

        private SIMPollResult PollConfiguredSIMMessagesCore(string expectedImsi)
        {
            var logLines = new List<string>();
            var normalizedExpectedImsi = expectedImsi.Trim();
            if (String.IsNullOrWhiteSpace(normalizedExpectedImsi))
            {
                logLines.Add("SIM SMS poll skipped: sim_imsi is empty.");
                return new SIMPollResult(0, 0, logLines);
            }

            if (!String.IsNullOrWhiteSpace(cachedPortName) && cachedBaudRate > 0)
            {
                try
                {
                    using var cachedPort = OpenPort(cachedPortName, cachedBaudRate);
                    InitializeModem(cachedPort);
                    return ProcessConfiguredSIMPort(cachedPort, cachedPortName, cachedBaudRate, normalizedExpectedImsi, logLines);
                }
                catch (Exception exception) when (IsProbeException(exception))
                {
                    logLines.Add($"SIM SMS poll cached modem {cachedPortName}@{cachedBaudRate} unavailable: {exception.Message}");
                }
            }

            return PollConfiguredSIMMessagesByScanning(normalizedExpectedImsi, logLines);
        }

        private SIMPollResult PollConfiguredSIMMessagesByScanning(string expectedImsi, List<string> logLines)
        {
            var failures = new List<string>();
            var modems = FindResponsiveSIMModems(failures);
            if (modems.Count == 0)
            {
                var message = failures.Count == 0
                    ? "SIM SMS poll found no responsive SIM modems."
                    : "SIM SMS poll found no responsive SIM modems. Tried: " + String.Join("; ", failures.Take(8));
                if (hasFoundSIMModem)
                    throw new InvalidOperationException(message);

                logLines.Add(message);
                return new SIMPollResult(0, 0, logLines);
            }

            var probes = ProbeSIMModemIMSIs(modems, failures);
            var distinctIMSIs = probes
                .Select(probe => probe.IMSI)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (distinctIMSIs.Count > 1)
                throw new InvalidOperationException("SIM SMS poll found multiple SIM cards: " + String.Join(", ", probes.Select(probe => $"{probe.Modem.PortName}@{probe.Modem.BaudRate}")));

            var matchingProbes = probes
                .Where(probe => String.Equals(probe.IMSI, expectedImsi, StringComparison.Ordinal))
                .OrderBy(probe => GetPortSortNumber(probe.Modem.PortName))
                .ThenBy(probe => probe.Modem.PortName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(probe => probe.Modem.BaudRate)
                .ToList();
            if (matchingProbes.Count == 0)
            {
                var message = failures.Count == 0
                    ? "SIM SMS poll found no modem with matching IMSI."
                    : "SIM SMS poll found no modem with matching IMSI. Tried: " + String.Join("; ", failures.Take(8));
                if (hasFoundSIMModem)
                    throw new InvalidOperationException(message);

                logLines.Add(message);
                return new SIMPollResult(0, 0, logLines);
            }

            if (matchingProbes.Count > 1)
            {
                logLines.Add(
                    "SIM SMS poll found multiple responsive ports for the configured SIM; using "
                    + $"{matchingProbes[0].Modem.PortName}@{matchingProbes[0].Modem.BaudRate}.");
            }

            var modem = matchingProbes[0].Modem;
            cachedPortName = modem.PortName;
            cachedBaudRate = modem.BaudRate;
            hasFoundSIMModem = true;

            using var port = OpenPort(modem.PortName, modem.BaudRate);
            InitializeModem(port);
            return ProcessConfiguredSIMPort(port, modem.PortName, modem.BaudRate, expectedImsi, logLines);
        }

        private SIMPollResult ProcessConfiguredSIMPort(
            SerialPort port,
            string portName,
            int baudRate,
            string expectedImsi,
            List<string> logLines)
        {
            string imsi;
            try
            {
                imsi = ReadIMSI(port);
            }
            catch (Exception exception) when (IsProbeException(exception))
            {
                logLines.Add($"SIM SMS poll skipped modem {portName}@{baudRate}: unable to read IMSI: {exception.Message}");
                return new SIMPollResult(0, 0, logLines);
            }

            if (!String.Equals(imsi, expectedImsi, StringComparison.Ordinal))
            {
                logLines.Add($"SIM SMS poll skipped modem {portName}@{baudRate}: IMSI mismatch.");
                return new SIMPollResult(0, 0, logLines);
            }

            logLines.Add($"SIM SMS poll matched modem {portName}@{baudRate}.");
            return ProcessMatchingSIMMessages(port, portName, logLines);
        }

        private SIMPollResult ProcessMatchingSIMMessages(SerialPort port, string portName, List<string> logLines)
        {
            ReadStoredMessagesResult readResult;
            try
            {
                readResult = ReadStoredMessagesForProcessing(port, portName);
            }
            catch (Exception exception) when (IsProbeException(exception) || exception is FormatException)
            {
                logLines.Add($"SIM SMS poll failed to read stored messages: {exception.Message}");
                return new SIMPollResult(0, 0, logLines);
            }

            foreach (var incompleteMessage in readResult.IncompleteConcatMessages)
            {
                logLines.Add(
                    $"SIM SMS poll kept incomplete long SMS from {incompleteMessage.Sender}: "
                    + $"{incompleteMessage.PresentParts}/{incompleteMessage.TotalParts} parts, "
                    + $"stored indexes {String.Join(", ", incompleteMessage.StoredIndexes)}.");
            }

            var messages = readResult.Messages;
            if (messages.Count == 0)
            {
                logLines.Add(readResult.IncompleteConcatMessages.Count == 0
                    ? "SIM SMS poll found no received messages."
                    : "SIM SMS poll found no complete received messages.");
                return new SIMPollResult(0, 0, logLines);
            }

            var parsedCount = 0;
            var deletedCount = 0;
            foreach (var message in messages.OrderBy(message => message.Time).ThenBy(message => message.Index))
            {
                SIMMessageProcessResult result;
                try
                {
                    result = ProcessSIMMessage(message, logLines);
                    parsedCount++;
                }
                catch (Exception exception)
                {
                    logLines.Add($"SIM SMS poll failed to process message index {message.Index} from {message.Sender}: {exception.Message}");
                    continue;
                }

                if (!result.CanDelete)
                {
                    logLines.Add($"SIM SMS poll retained message index {message.Index} from {message.Sender}: {result.Description}.");
                    continue;
                }

                if (!DeleteProcessedMessages)
                {
                    logLines.Add($"SIM SMS poll retained message index {message.Index} from {message.Sender}: deletion is temporarily disabled.");
                    continue;
                }

                foreach (var index in message.StoredIndexes)
                {
                    try
                    {
                        DeleteStoredMessage(port, message.Storage, index);
                        deletedCount++;
                        logLines.Add($"SIM SMS poll deleted message index {index} from {message.Sender}.");
                    }
                    catch (Exception exception) when (IsProbeException(exception))
                    {
                        logLines.Add($"SIM SMS poll failed to delete message index {index} from {message.Sender}: {exception.Message}");
                    }
                }
            }

            logLines.Add($"SIM SMS poll processed {parsedCount} message(s), deleted {deletedCount} stored item(s).");
            return new SIMPollResult(parsedCount, deletedCount, logLines);
        }

        private SIMMessageProcessResult ProcessSIMMessage(SIMMessage message, List<string> logLines)
        {
            if (IsICBCSender(message.Sender))
            {
                var result = ParseICBCSIMMessage(message);
                logLines.Add($"SIM SMS poll processed ICBC message index {message.Index} from {message.Sender}: {result.Description}.");
                return result;
            }

            logLines.Add(DeleteProcessedMessages
                ? $"SIM SMS poll has no parser for sender {message.Sender}; message index {message.Index} will be deleted."
                : $"SIM SMS poll has no parser for sender {message.Sender}; message index {message.Index} is retained because deletion is temporarily disabled.");
            return new SIMMessageProcessResult("unsupported sender", true);
        }

        private static string ReadIMSI(SerialPort port)
        {
            var response = SendCommand(port, "AT+CIMI");
            var imsi = GetResponseLines(response).FirstOrDefault(line => Regex.IsMatch(line, @"\A\d{10,20}\z"));
            if (String.IsNullOrWhiteSpace(imsi))
                throw new InvalidOperationException("AT+CIMI returned no IMSI.");

            return imsi;
        }

        private static List<string> GetCandidatePorts(string? portName)
        {
            if (!String.IsNullOrWhiteSpace(portName))
                return [portName.Trim()];

            return SerialPort.GetPortNames()
                .OrderBy(GetPortSortNumber)
                .ThenBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<SIMModemInfo> FindResponsiveSIMModems(List<string> failures)
        {
            var modems = new List<SIMModemInfo>();
            foreach (var candidatePortName in GetCandidatePorts(null))
            {
                foreach (var baudRate in CandidateBaudRates)
                {
                    try
                    {
                        using var port = OpenPort(candidatePortName, baudRate);
                        InitializeModem(port);
                        if (!TrySendCommand(port, "AT+CSMS?", out var smsServiceResponse))
                        {
                            failures.Add($"{candidatePortName}@{baudRate}: SMS service unavailable: {smsServiceResponse}");
                            continue;
                        }

                        modems.Add(new SIMModemInfo(candidatePortName, baudRate));
                        break;
                    }
                    catch (Exception exception) when (IsProbeException(exception))
                    {
                        failures.Add($"{candidatePortName}@{baudRate}: {exception.Message}");
                    }
                }
            }

            return modems;
        }

        private static List<SIMModemProbe> ProbeSIMModemIMSIs(List<SIMModemInfo> modems, List<string> failures)
        {
            var probes = new List<SIMModemProbe>();
            foreach (var modem in modems)
            {
                try
                {
                    using var port = OpenPort(modem.PortName, modem.BaudRate);
                    InitializeModem(port);
                    probes.Add(new SIMModemProbe(modem, ReadIMSI(port)));
                }
                catch (Exception exception) when (IsProbeException(exception))
                {
                    failures.Add($"{modem.PortName}@{modem.BaudRate}: unable to read IMSI: {exception.Message}");
                }
            }

            return probes;
        }

        private static int GetPortSortNumber(string portName)
        {
            var match = Regex.Match(portName, @"\d+");
            return match.Success
                ? Int32.Parse(match.Value, CultureInfo.InvariantCulture)
                : Int32.MaxValue;
        }

        private static SerialPort OpenPort(string portName, int baudRate)
        {
            var port = new SerialPort(portName, baudRate)
            {
                DtrEnable = true,
                RtsEnable = true,
                ReadTimeout = 250,
                WriteTimeout = (int)DefaultCommandTimeout.TotalMilliseconds,
                NewLine = "\r\n",
                Encoding = Encoding.ASCII
            };
            port.Open();
            port.DiscardInBuffer();
            port.DiscardOutBuffer();
            return port;
        }

        private static void InitializeModem(SerialPort port)
        {
            SendCommand(port, "AT");
            SendCommand(port, "ATE0");
            TrySendCommand(port, "AT+CMEE=2", out _);
        }

        private static List<SIMMessage> ReadStoredMessages(SerialPort port, string portName)
        {
            return ReadStoredMessages(port, portName, CandidateMessageStores);
        }

        private static List<SIMMessage> ReadStoredMessages(SerialPort port, string portName, IReadOnlyList<string> messageStores)
        {
            return ReadStoredMessagesCore(port, portName, messageStores, keepIncompleteConcatParts: true).Messages;
        }

        private static ReadStoredMessagesResult ReadStoredMessagesForProcessing(SerialPort port, string portName)
        {
            return ReadStoredMessagesCore(port, portName, ProcessingMessageStores, keepIncompleteConcatParts: false);
        }

        private static ReadStoredMessagesResult ReadStoredMessagesCore(
            SerialPort port,
            string portName,
            IReadOnlyList<string> messageStores,
            bool keepIncompleteConcatParts)
        {
            var parsedMessages = new List<ParsedSIMMessage>();
            var pduModeWorked = false;
            foreach (var store in messageStores)
            {
                if (!String.IsNullOrWhiteSpace(store) && !TrySendCommand(port, $"AT+CPMS=\"{store}\"", out _))
                    continue;

                if (TryReadPduMessages(port, portName, store, out var pduMessages))
                {
                    pduModeWorked = true;
                    parsedMessages.AddRange(pduMessages);
                }
            }

            if (!pduModeWorked)
            {
                foreach (var store in messageStores)
                {
                    if (!String.IsNullOrWhiteSpace(store) && !TrySendCommand(port, $"AT+CPMS=\"{store}\"", out _))
                        continue;

                    if (TryReadTextMessages(port, portName, store, out var textMessages))
                        parsedMessages.AddRange(textMessages);
                }
            }

            return CombineAndDeduplicate(parsedMessages, keepIncompleteConcatParts);
        }

        private static void DeleteStoredMessage(SerialPort port, string store, int index)
        {
            if (!String.IsNullOrWhiteSpace(store))
                SendCommand(port, $"AT+CPMS=\"{store}\"");

            SendCommand(port, $"AT+CMGD={index}");
        }

        private static bool TryReadPduMessages(
            SerialPort port,
            string portName,
            string store,
            out List<ParsedSIMMessage> messages)
        {
            messages = [];
            try
            {
                SendCommand(port, "AT+CMGF=0");
                var response = SendCommand(port, "AT+CMGL=4", ListMessagesTimeout);
                messages = ParsePduMessageList(response, portName, store);
                return true;
            }
            catch (Exception exception) when (IsProbeException(exception) || exception is FormatException)
            {
                return false;
            }
        }

        private static bool TryReadTextMessages(
            SerialPort port,
            string portName,
            string store,
            out List<ParsedSIMMessage> messages)
        {
            messages = [];
            try
            {
                SendCommand(port, "AT+CMGF=1");
                TrySendCommand(port, "AT+CSCS=\"GSM\"", out _);
                var response = SendCommand(port, "AT+CMGL=\"ALL\"", ListMessagesTimeout);
                messages = ParseTextMessageList(response, portName, store);
                return true;
            }
            catch (Exception exception) when (IsProbeException(exception) || exception is FormatException)
            {
                return false;
            }
        }

        private static bool TrySendCommand(SerialPort port, string command, out string response)
        {
            try
            {
                response = SendCommand(port, command);
                return true;
            }
            catch (Exception exception) when (IsProbeException(exception))
            {
                response = exception.Message;
                return false;
            }
        }

        private static string SendCommand(SerialPort port, string command, TimeSpan? timeout = null)
        {
            port.DiscardInBuffer();
            port.Write(command + "\r");
            var response = ReadUntilFinalResponse(port, timeout ?? DefaultCommandTimeout);
            if (IsErrorResponse(response))
                throw new InvalidOperationException($"{command} failed: {GetFinalResponseLine(response)}");

            return response;
        }

        private static string ReadUntilFinalResponse(SerialPort port, TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow + timeout;
            var builder = new StringBuilder();
            while (DateTime.UtcNow < deadline)
            {
                var chunk = port.ReadExisting();
                if (!String.IsNullOrEmpty(chunk))
                {
                    builder.Append(chunk);
                    var response = builder.ToString();
                    if (HasFinalResponse(response))
                        return response;
                }

                Thread.Sleep(50);
            }

            throw new TimeoutException("Timed out waiting for modem response.");
        }

        private static bool HasFinalResponse(string response)
        {
            return GetResponseLines(response).Any(line =>
                line.Equals("OK", StringComparison.OrdinalIgnoreCase)
                || line.Equals("ERROR", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CME ERROR:", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CMS ERROR:", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsErrorResponse(string response)
        {
            return GetResponseLines(response).Any(line =>
                line.Equals("ERROR", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CME ERROR:", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CMS ERROR:", StringComparison.OrdinalIgnoreCase));
        }

        private static string GetFinalResponseLine(string response)
        {
            return GetResponseLines(response).LastOrDefault(line =>
                line.Equals("OK", StringComparison.OrdinalIgnoreCase)
                || line.Equals("ERROR", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CME ERROR:", StringComparison.OrdinalIgnoreCase)
                || line.StartsWith("+CMS ERROR:", StringComparison.OrdinalIgnoreCase)) ?? "unknown modem error";
        }

        private static List<string> GetResponseLines(string response)
        {
            return response
                .Replace("\r", "\n")
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(line =>
                    !line.Equals("AT", StringComparison.OrdinalIgnoreCase)
                    && !line.StartsWith("AT+", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private static List<ParsedSIMMessage> ParsePduMessageList(string response, string portName, string store)
        {
            var lines = GetResponseLines(response);
            var messages = new List<ParsedSIMMessage>();
            for (var i = 0; i < lines.Count; i++)
            {
                if (!lines[i].StartsWith("+CMGL:", StringComparison.OrdinalIgnoreCase))
                    continue;

                var header = ParseCmglHeader(lines[i]);
                if (!IsReceivedStatus(header.Status))
                {
                    i++;
                    continue;
                }

                if (i + 1 >= lines.Count)
                    throw new FormatException("PDU CMGL response has no PDU line.");

                var pdu = lines[++i].Trim();
                if (pdu.Equals("OK", StringComparison.OrdinalIgnoreCase))
                    throw new FormatException("PDU CMGL response ended before PDU line.");

                var parsed = ParseSmsDeliverPdu(pdu);
                messages.Add(new ParsedSIMMessage(
                    parsed.Time,
                    parsed.Sender,
                    parsed.Text,
                    header.Index,
                    header.Status,
                    portName,
                    store,
                    parsed.Concat));
            }

            return messages;
        }

        private static List<ParsedSIMMessage> ParseTextMessageList(string response, string portName, string store)
        {
            var lines = GetResponseLines(response);
            var messages = new List<ParsedSIMMessage>();
            for (var i = 0; i < lines.Count; i++)
            {
                if (!lines[i].StartsWith("+CMGL:", StringComparison.OrdinalIgnoreCase))
                    continue;

                var header = ParseCmglHeader(lines[i]);
                if (!IsReceivedStatus(header.Status))
                    continue;

                var body = new List<string>();
                while (i + 1 < lines.Count
                    && !lines[i + 1].StartsWith("+CMGL:", StringComparison.OrdinalIgnoreCase)
                    && !lines[i + 1].Equals("OK", StringComparison.OrdinalIgnoreCase))
                {
                    body.Add(lines[++i]);
                }

                messages.Add(new ParsedSIMMessage(
                    ParseTextModeTimestamp(header.Timestamp),
                    DecodeMaybeUcs2Hex(header.Sender),
                    DecodeMaybeUcs2Hex(String.Join("\n", body)),
                    header.Index,
                    header.Status,
                    portName,
                    store,
                    null));
            }

            return messages;
        }

        private static CmglHeader ParseCmglHeader(string line)
        {
            var colon = line.IndexOf(':');
            if (colon < 0)
                throw new FormatException($"Invalid CMGL header: {line}");

            var fields = SplitAtFields(line[(colon + 1)..]);
            if (fields.Count < 2 || !Int32.TryParse(fields[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
                throw new FormatException($"Invalid CMGL header: {line}");

            return new CmglHeader(
                index,
                Unquote(fields[1]),
                fields.Count > 2 ? DecodeMaybeUcs2Hex(fields[2]) : "",
                fields.Count > 4 ? DecodeMaybeUcs2Hex(fields[4]) : "");
        }

        private static List<string> SplitAtFields(string text)
        {
            var fields = new List<string>();
            var builder = new StringBuilder();
            var inQuotes = false;
            foreach (var character in text)
            {
                if (character == '"')
                {
                    inQuotes = !inQuotes;
                    builder.Append(character);
                    continue;
                }

                if (character == ',' && !inQuotes)
                {
                    fields.Add(builder.ToString().Trim());
                    builder.Clear();
                    continue;
                }

                builder.Append(character);
            }

            fields.Add(builder.ToString().Trim());
            return fields;
        }

        private static bool IsReceivedStatus(string status)
        {
            return status.Equals("0", StringComparison.OrdinalIgnoreCase)
                || status.Equals("1", StringComparison.OrdinalIgnoreCase)
                || status.Equals("REC UNREAD", StringComparison.OrdinalIgnoreCase)
                || status.Equals("REC READ", StringComparison.OrdinalIgnoreCase);
        }

        private static ParsedPduMessage ParseSmsDeliverPdu(string pdu)
        {
            var bytes = ParseHexBytes(pdu);
            var offset = 0;
            var smscLength = ReadOctet(bytes, ref offset);
            offset += smscLength;

            var firstOctet = ReadOctet(bytes, ref offset);
            if ((firstOctet & 0x03) != 0)
                throw new FormatException("Only SMS-DELIVER PDUs are supported.");

            var hasUserDataHeader = (firstOctet & 0x40) != 0;
            var addressLength = ReadOctet(bytes, ref offset);
            var addressType = ReadOctet(bytes, ref offset);
            var addressByteCount = (addressLength + 1) / 2;
            var sender = DecodeAddress(bytes.Skip(offset).Take(addressByteCount).ToArray(), addressLength, addressType);
            offset += addressByteCount;

            _ = ReadOctet(bytes, ref offset);
            var dataCodingScheme = ReadOctet(bytes, ref offset);
            var timestamp = DecodeTimestamp(bytes, ref offset);
            var userDataLength = ReadOctet(bytes, ref offset);
            var userData = bytes.Skip(offset).ToArray();
            var decoded = DecodeUserData(dataCodingScheme, userDataLength, hasUserDataHeader, userData);
            return new ParsedPduMessage(timestamp, sender, decoded.Text, decoded.Concat);
        }

        private static DecodedUserData DecodeUserData(
            int dataCodingScheme,
            int userDataLength,
            bool hasUserDataHeader,
            byte[] userData)
        {
            var userDataHeaderLength = 0;
            SmsConcatInfo? concat = null;
            if (hasUserDataHeader && userData.Length > 0)
            {
                userDataHeaderLength = Math.Min(userData.Length, userData[0] + 1);
                concat = ParseConcatInfo(userData.Take(userDataHeaderLength).ToArray());
            }

            if ((dataCodingScheme & 0x0C) == 0x08)
            {
                var start = userDataHeaderLength;
                var count = Math.Max(0, Math.Min(userData.Length - start, userDataLength - userDataHeaderLength));
                return new DecodedUserData(Encoding.BigEndianUnicode.GetString(userData, start, count).TrimEnd('\0'), concat);
            }

            if ((dataCodingScheme & 0x0C) == 0x04)
            {
                var start = userDataHeaderLength;
                var count = Math.Max(0, Math.Min(userData.Length - start, userDataLength - userDataHeaderLength));
                return new DecodedUserData(Encoding.Latin1.GetString(userData, start, count).TrimEnd('\0'), concat);
            }

            var skippedBits = userDataHeaderLength * 8;
            var headerSeptets = hasUserDataHeader ? (skippedBits + 6) / 7 : 0;
            var septetCount = Math.Max(0, userDataLength - headerSeptets);
            return new DecodedUserData(DecodeGsm7(userData, septetCount, skippedBits), concat);
        }

        private static SmsConcatInfo? ParseConcatInfo(byte[] userDataHeader)
        {
            var offset = 1;
            while (offset + 1 < userDataHeader.Length)
            {
                var identifier = userDataHeader[offset++];
                var length = userDataHeader[offset++];
                if (offset + length > userDataHeader.Length)
                    return null;

                if (identifier == 0x00 && length == 3)
                    return new SmsConcatInfo(userDataHeader[offset].ToString(CultureInfo.InvariantCulture), userDataHeader[offset + 2], userDataHeader[offset + 1]);

                if (identifier == 0x08 && length == 4)
                {
                    var reference = (userDataHeader[offset] << 8) | userDataHeader[offset + 1];
                    return new SmsConcatInfo(reference.ToString(CultureInfo.InvariantCulture), userDataHeader[offset + 3], userDataHeader[offset + 2]);
                }

                offset += length;
            }

            return null;
        }

        private static string DecodeAddress(byte[] bytes, int length, int type)
        {
            var typeOfNumber = type & 0x70;
            if (typeOfNumber == 0x50)
                return DecodeGsm7(bytes, length);

            var builder = new StringBuilder();
            foreach (var item in bytes)
            {
                AppendSemiOctet(builder, item & 0x0F);
                AppendSemiOctet(builder, (item >> 4) & 0x0F);
            }

            if (builder.Length > length)
                builder.Length = length;

            var number = builder.ToString();
            return (type & 0x90) == 0x90 ? "+" + number : number;
        }

        private static void AppendSemiOctet(StringBuilder builder, int value)
        {
            if (value <= 9)
                builder.Append((char)('0' + value));
        }

        private static DateTime DecodeTimestamp(byte[] bytes, ref int offset)
        {
            var year = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            var month = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            var day = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            var hour = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            var minute = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            var second = DecodeSemiOctetDecimal(ReadOctet(bytes, ref offset));
            _ = ReadOctet(bytes, ref offset);
            return new DateTime(year >= 90 ? 1900 + year : 2000 + year, month, day, hour, minute, second);
        }

        private static int DecodeSemiOctetDecimal(int value)
        {
            return (value & 0x0F) * 10 + ((value >> 4) & 0x0F);
        }

        private static int ReadOctet(byte[] bytes, ref int offset)
        {
            if (offset >= bytes.Length)
                throw new FormatException("PDU ended unexpectedly.");

            return bytes[offset++];
        }

        private static byte[] ParseHexBytes(string text)
        {
            var normalized = Regex.Replace(text, @"\s+", "");
            if (normalized.Length % 2 != 0 || !Regex.IsMatch(normalized, @"\A[0-9A-Fa-f]*\z"))
                throw new FormatException("Invalid hexadecimal PDU text.");

            var bytes = new byte[normalized.Length / 2];
            for (var i = 0; i < bytes.Length; i++)
                bytes[i] = Byte.Parse(normalized.Substring(i * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            return bytes;
        }

        private static string DecodeGsm7(byte[] bytes, int septetCount, int skipBits = 0)
        {
            var builder = new StringBuilder();
            for (var i = 0; i < septetCount; i++)
            {
                var value = 0;
                var bitOffset = skipBits + i * 7;
                for (var bit = 0; bit < 7; bit++)
                {
                    var sourceBit = bitOffset + bit;
                    var byteIndex = sourceBit / 8;
                    if (byteIndex >= bytes.Length)
                        break;

                    if ((bytes[byteIndex] & (1 << (sourceBit % 8))) != 0)
                        value |= 1 << bit;
                }

                if (value == 0x1B && i + 1 < septetCount)
                {
                    var extension = ReadGsm7Septet(bytes, skipBits + (++i) * 7);
                    builder.Append(DecodeGsm7Extension(extension));
                }
                else if (value != 0x1B)
                {
                    builder.Append(DecodeGsm7Basic(value));
                }
            }

            return builder.ToString().TrimEnd('\0');
        }

        private static int ReadGsm7Septet(byte[] bytes, int bitOffset)
        {
            var value = 0;
            for (var bit = 0; bit < 7; bit++)
            {
                var sourceBit = bitOffset + bit;
                var byteIndex = sourceBit / 8;
                if (byteIndex >= bytes.Length)
                    break;

                if ((bytes[byteIndex] & (1 << (sourceBit % 8))) != 0)
                    value |= 1 << bit;
            }

            return value;
        }

        private static char DecodeGsm7Basic(int value)
        {
            return value switch
            {
                0x00 => '@',
                0x01 => '\u00A3',
                0x02 => '$',
                0x03 => '\u00A5',
                0x04 => '\u00E8',
                0x05 => '\u00E9',
                0x06 => '\u00F9',
                0x07 => '\u00EC',
                0x08 => '\u00F2',
                0x09 => '\u00C7',
                0x0A => '\n',
                0x0B => '\u00D8',
                0x0C => '\u00F8',
                0x0D => '\r',
                0x0E => '\u00C5',
                0x0F => '\u00E5',
                0x10 => '\u0394',
                0x11 => '_',
                0x12 => '\u03A6',
                0x13 => '\u0393',
                0x14 => '\u039B',
                0x15 => '\u03A9',
                0x16 => '\u03A0',
                0x17 => '\u03A8',
                0x18 => '\u03A3',
                0x19 => '\u0398',
                0x1A => '\u039E',
                0x1C => '\u00C6',
                0x1D => '\u00E6',
                0x1E => '\u00DF',
                0x1F => '\u00C9',
                >= 0x20 and <= 0x5A => (char)value,
                0x5B => '\u00C4',
                0x5C => '\u00D6',
                0x5D => '\u00D1',
                0x5E => '\u00DC',
                0x5F => '\u00A7',
                0x60 => '\u00BF',
                >= 0x61 and <= 0x7A => (char)value,
                0x7B => '\u00E4',
                0x7C => '\u00F6',
                0x7D => '\u00F1',
                0x7E => '\u00FC',
                0x7F => '\u00E0',
                _ => ' '
            };
        }

        private static char DecodeGsm7Extension(int value)
        {
            return value switch
            {
                0x0A => '\f',
                0x14 => '^',
                0x28 => '{',
                0x29 => '}',
                0x2F => '\\',
                0x3C => '[',
                0x3D => '~',
                0x3E => ']',
                0x40 => '|',
                0x65 => '\u20AC',
                _ => ' '
            };
        }

        private static string DecodeMaybeUcs2Hex(string text)
        {
            var value = Unquote(text).Trim();
            if (value.Length < 4 || value.Length % 4 != 0 || !Regex.IsMatch(value, @"\A[0-9A-Fa-f]*\z"))
                return value;

            try
            {
                return Encoding.BigEndianUnicode.GetString(ParseHexBytes(value)).TrimEnd('\0');
            }
            catch (FormatException)
            {
                return value;
            }
        }

        private static DateTime ParseTextModeTimestamp(string text)
        {
            var value = DecodeMaybeUcs2Hex(text);
            var match = Regex.Match(value, @"(?<year>\d{2})/(?<month>\d{2})/(?<day>\d{2}),(?<hour>\d{2}):(?<minute>\d{2}):(?<second>\d{2})");
            if (!match.Success)
                return DateTime.MinValue;

            var year = Int32.Parse(match.Groups["year"].Value, CultureInfo.InvariantCulture);
            return new DateTime(
                year >= 90 ? 1900 + year : 2000 + year,
                Int32.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(match.Groups["hour"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(match.Groups["minute"].Value, CultureInfo.InvariantCulture),
                Int32.Parse(match.Groups["second"].Value, CultureInfo.InvariantCulture));
        }

        private static string Unquote(string text)
        {
            var value = text.Trim();
            return value.Length >= 2 && value[0] == '"' && value[^1] == '"'
                ? value[1..^1]
                : value;
        }

        private static ReadStoredMessagesResult CombineAndDeduplicate(List<ParsedSIMMessage> messages, bool keepIncompleteConcatParts)
        {
            var combined = new List<SIMMessage>();
            var incompleteMessages = new List<IncompleteConcatMessage>();
            var singleMessages = messages.Where(message => message.Concat is null);
            combined.AddRange(singleMessages.Select(ToSIMMessage));

            foreach (var group in messages
                .Where(message => message.Concat is not null)
                .GroupBy(message => $"{message.PortName}|{message.Storage}|{message.Sender}|{message.Concat!.Reference}|{message.Concat.Total}"))
            {
                var ordered = group.OrderBy(message => message.Concat!.Sequence).ToList();
                var expectedTotal = ordered[0].Concat!.Total;
                var presentSequences = ordered.Select(message => message.Concat!.Sequence).Distinct().ToList();
                if (ordered.Count == expectedTotal && presentSequences.Count == expectedTotal)
                {
                    combined.Add(new SIMMessage(
                        ordered.Min(message => message.Time),
                        ordered[0].Sender,
                        String.Concat(ordered.Select(message => message.Text)),
                        ordered.Min(message => message.Index),
                        ordered[0].Status,
                        ordered[0].PortName,
                        ordered[0].Storage,
                        ordered.Select(message => message.Index).Distinct().OrderBy(index => index).ToArray()));
                }
                else
                {
                    incompleteMessages.Add(new IncompleteConcatMessage(
                        ordered[0].Sender,
                        expectedTotal,
                        presentSequences.Count,
                        ordered.Select(message => message.Index).Distinct().OrderBy(index => index).ToArray()));

                    if (keepIncompleteConcatParts)
                        combined.AddRange(ordered.Select(ToSIMMessage));
                }
            }

            var deduplicatedMessages = combined
                .GroupBy(message => $"{message.Time:O}|{message.Sender}|{message.Text}")
                .Select(group => group.First())
                .ToList();

            return new ReadStoredMessagesResult(deduplicatedMessages, incompleteMessages);
        }

        private static SIMMessage ToSIMMessage(ParsedSIMMessage message)
        {
            return new SIMMessage(
                message.Time,
                message.Sender,
                message.Text,
                message.Index,
                message.Status,
                message.PortName,
                message.Storage,
                [message.Index]);
        }

        private static bool IsProbeException(Exception exception)
        {
            return exception is IOException
                or InvalidOperationException
                or TimeoutException
                or UnauthorizedAccessException;
        }

        private sealed record CmglHeader(int Index, string Status, string Sender, string Timestamp);
        private sealed record ParsedPduMessage(DateTime Time, string Sender, string Text, SmsConcatInfo? Concat);
        private sealed record DecodedUserData(string Text, SmsConcatInfo? Concat);
        private sealed record SmsConcatInfo(string Reference, int Sequence, int Total);
        private sealed record ParsedSIMMessage(DateTime Time, string Sender, string Text, int Index, string Status, string PortName, string Storage, SmsConcatInfo? Concat);
        private sealed record IncompleteConcatMessage(string Sender, int TotalParts, int PresentParts, IReadOnlyList<int> StoredIndexes);
        private sealed record ReadStoredMessagesResult(List<SIMMessage> Messages, List<IncompleteConcatMessage> IncompleteConcatMessages);
        private sealed record SIMMessageProcessResult(string Description, bool CanDelete);
        private sealed record SIMModemInfo(string PortName, int BaudRate);
        private sealed record SIMModemProbe(SIMModemInfo Modem, string IMSI);
    }

    public sealed record SIMMessage(
        DateTime Time,
        string Sender,
        string Text,
        int Index = 0,
        string Status = "",
        string PortName = "",
        string Storage = "",
        IReadOnlyList<int>? DeleteIndexes = null)
    {
        public IReadOnlyList<int> StoredIndexes => DeleteIndexes ?? new[] { Index };
    }

    public sealed record SIMPollResult(
        int ParsedMessages,
        int DeletedStoredItems,
        IReadOnlyList<string> LogLines);
}
