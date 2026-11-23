using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Mpc.Core
{
    /// <summary>One result in an Authentication-Results header, such as "dkim=pass header.d=green.mpc".</summary>
    public sealed class AuthResult
    {
        public AuthResult(string method, string result, IDictionary<string, string> properties)
        {
            Method = (method ?? string.Empty).ToLowerInvariant();
            Result = (result ?? string.Empty).ToLowerInvariant();
            Properties = new Dictionary<string, string>(properties ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>"spf", "dkim", "dmarc", ... (lower case, without a version).</summary>
        public string Method { get; private set; }

        /// <summary>"pass", "fail", "none", ... (lower case).</summary>
        public string Result { get; private set; }

        /// <summary>Properties such as header.d, header.from or smtp.mailfrom, and reason.</summary>
        public IReadOnlyDictionary<string, string> Properties { get; private set; }

        public string GetProperty(string name)
        {
            string value;
            return Properties.TryGetValue(name, out value) ? value : null;
        }

        public override string ToString()
        {
            return Method + "=" + Result;
        }
    }

    /// <summary>
    /// A parsed Authentication-Results header (RFC 8601): the name of the server that ran the
    /// checks (authserv-id) and its results. Comments in parentheses are ignored.
    /// </summary>
    public sealed class AuthenticationResults
    {
        private AuthenticationResults(string authservId, IReadOnlyList<AuthResult> results)
        {
            AuthservId = authservId;
            Results = results;
        }

        /// <summary>The server that added the header, e.g. "relay.provider-a.test".</summary>
        public string AuthservId { get; private set; }

        public IReadOnlyList<AuthResult> Results { get; private set; }

        /// <summary>Parses one header value. Returns false when it has no authserv-id.</summary>
        public static bool TryParse(string value, out AuthenticationResults parsed)
        {
            parsed = null;
            if (string.IsNullOrWhiteSpace(value)) return false;

            List<string> parts = SplitOutsideQuotes(RemoveComments(value), ';');
            List<string> idTokens = Tokens(parts[0]);
            if (idTokens.Count == 0) return false;

            var results = new List<AuthResult>();
            for (int i = 1; i < parts.Count; i++)
            {
                List<string> tokens = Tokens(parts[i]);
                if (tokens.Count == 0) continue;

                // "method[/version]=result"; a lone "none" means no checks were run.
                int eq = tokens[0].IndexOf('=');
                if (eq <= 0) continue;
                string method = tokens[0].Substring(0, eq);
                int slash = method.IndexOf('/');
                if (slash >= 0) method = method.Substring(0, slash);
                string result = Unquote(tokens[0].Substring(eq + 1));

                var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int t = 1; t < tokens.Count; t++)
                {
                    int peq = tokens[t].IndexOf('=');
                    if (peq <= 0) continue;
                    properties[tokens[t].Substring(0, peq)] = Unquote(tokens[t].Substring(peq + 1));
                }
                results.Add(new AuthResult(method.Trim(), result.Trim(), properties));
            }

            parsed = new AuthenticationResults(idTokens[0], results);
            return true;
        }

        /// <summary>Every Authentication-Results header in a raw header block that can be parsed.</summary>
        public static IReadOnlyList<AuthenticationResults> FromHeaders(string rawHeaders)
        {
            var all = new List<AuthenticationResults>();
            foreach (string value in InternetHeaders.GetAll(rawHeaders, MpcConstants.HeaderAuthenticationResults))
            {
                AuthenticationResults parsed;
                if (TryParse(value, out parsed)) all.Add(parsed);
            }
            return all;
        }

        private static string RemoveComments(string s)
        {
            var sb = new StringBuilder(s.Length);
            int depth = 0;
            bool quoted = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\\' && i + 1 < s.Length)
                {
                    if (depth == 0) sb.Append(c).Append(s[i + 1]);
                    i++;
                    continue;
                }
                if (depth == 0 && c == '"') quoted = !quoted;
                if (!quoted && c == '(') { depth++; continue; }
                if (!quoted && c == ')' && depth > 0) { depth--; sb.Append(' '); continue; }
                if (depth == 0) sb.Append(c);
            }
            return sb.ToString();
        }

        private static List<string> SplitOutsideQuotes(string s, char separator)
        {
            var parts = new List<string>();
            var current = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\\' && i + 1 < s.Length) { current.Append(c).Append(s[++i]); continue; }
                if (c == '"') quoted = !quoted;
                if (c == separator && !quoted) { parts.Add(current.ToString()); current.Clear(); continue; }
                current.Append(c);
            }
            parts.Add(current.ToString());
            return parts;
        }

        private static List<string> Tokens(string s)
        {
            var tokens = new List<string>();
            var current = new StringBuilder();
            bool quoted = false;
            foreach (char c in s)
            {
                if (c == '"') quoted = !quoted;
                if (!quoted && char.IsWhiteSpace(c))
                {
                    if (current.Length > 0) { tokens.Add(current.ToString()); current.Clear(); }
                    continue;
                }
                current.Append(c);
            }
            if (current.Length > 0) tokens.Add(current.ToString());
            return tokens;
        }

        private static string Unquote(string s)
        {
            s = s.Trim();
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"') s = s.Substring(1, s.Length - 2).Replace("\\\"", "\"").Replace("\\\\", "\\");
            return s;
        }
    }

    /// <summary>Whether a message's chunks passed the sender checks, and if not, why.</summary>
    public sealed class SenderCheckResult
    {
        internal SenderCheckResult(IReadOnlyList<string> problems)
        {
            Problems = problems;
        }

        public bool Passed
        {
            get { return Problems.Count == 0; }
        }

        /// <summary>Why the message failed, one line per problem; empty when it passed.</summary>
        public IReadOnlyList<string> Problems { get; private set; }
    }

    /// <summary>
    /// Stage 1 of malware detection (sender and connection checks), on the receiving side.
    /// <para>
    /// Each chunk is checked where it enters the recipient's MPC providers: the relay named in its
    /// X-MPC-Via runs SPF, DKIM and DMARC and records the results in an Authentication-Results
    /// header named after itself, and the recipient's mail server removes any it didn't get from
    /// those relays. A message is rejoined only when, for every chunk, the results from its
    /// relay say DMARC passed and a DKIM signature from the From: domain verified, and all chunks
    /// have the same From: address. A chunk without results from its relay fails: unchecked is
    /// not trusted.
    /// </para>
    /// </summary>
    public static class SenderChecks
    {
        public static SenderCheckResult Evaluate<T>(ChunkGroup<T> group)
        {
            if (group == null) throw new ArgumentNullException("group");

            var problems = new List<string>();
            var fromAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (ReceivedChunk<T> chunk in group.Chunks)
            {
                string label = "chunk " + chunk.Chunk.Index.ToString(CultureInfo.InvariantCulture);
                string fromAddress = EmailAddress.GetAddress(chunk.From);
                string fromDomain = EmailAddress.GetDomain(chunk.From);
                if (fromAddress == null || fromDomain == null)
                {
                    problems.Add(label + ": no From: address");
                    continue;
                }
                fromAddresses.Add(fromAddress);

                string relay = NormalizeHost(chunk.Via);
                List<AuthResult> results = chunk.AuthenticationResults
                    .Where(ar => string.Equals(NormalizeHost(ar.AuthservId), relay, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(ar => ar.Results)
                    .ToList();
                if (results.Count == 0)
                {
                    problems.Add(label + ": no sender-check results from " + chunk.Via);
                    continue;
                }

                // DMARC: SPF or DKIM passed for the From: domain.
                List<AuthResult> dmarc = results.Where(r => r.Method == "dmarc").ToList();
                bool dmarcPass = dmarc.Any(r => r.Result == "pass" && SameDomain(r.GetProperty("header.from"), fromDomain, true));
                if (!dmarcPass)
                {
                    problems.Add(label + ": DMARC " + (dmarc.Count == 0 ? "was not checked" : dmarc[0].Result) + " for " + fromDomain + " (at " + chunk.Via + ")");
                }

                // DKIM: a valid signature from the From: domain itself, so the chunk came from that
                // domain's mail server unchanged (passing on SPF alone isn't enough).
                bool dkimPass = results.Any(r => r.Method == "dkim" && r.Result == "pass" && SameDomain(r.GetProperty("header.d"), fromDomain, false));
                if (!dkimPass)
                {
                    problems.Add(label + ": no valid DKIM signature from " + fromDomain + " (at " + chunk.Via + ")");
                }
            }

            if (fromAddresses.Count > 1)
            {
                problems.Add("the chunks have different From: addresses (" + string.Join(", ", fromAddresses.OrderBy(a => a, StringComparer.OrdinalIgnoreCase)) + ")");
            }

            return new SenderCheckResult(problems);
        }

        private static string NormalizeHost(string host)
        {
            return (host ?? string.Empty).Trim().TrimEnd('.');
        }

        private static bool SameDomain(string value, string domain, bool missingIsOk)
        {
            if (string.IsNullOrWhiteSpace(value)) return missingIsOk;
            return string.Equals(NormalizeHost(value), domain, StringComparison.OrdinalIgnoreCase);
        }
    }
}
