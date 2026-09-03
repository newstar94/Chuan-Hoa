using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ChuanHoa.Client.Core.Safety
{
    public sealed class PersistentAuthorizationReplayStore : IAuthorizationReplayStore
    {
        private const int MaximumJtiLength = 512;
        private const int CompactAfterRecordCount = 1024;
        private readonly string _filePath;
        private readonly IUtcClock _clock;

        public PersistentAuthorizationReplayStore(string filePath, IUtcClock clock)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("A replay-store file path is required.", nameof(filePath));
            }

            _filePath = Path.GetFullPath(filePath);
            _clock = clock ?? throw new ArgumentNullException(nameof(clock));
        }

        public bool TryConsume(string jti, DateTimeOffset expiresAtUtc)
        {
            if (string.IsNullOrWhiteSpace(jti) ||
                jti.Length > MaximumJtiLength ||
                expiresAtUtc <= _clock.UtcNow)
            {
                return false;
            }

            try
            {
                var directory = Path.GetDirectoryName(_filePath);
                if (string.IsNullOrWhiteSpace(directory))
                {
                    return false;
                }

                Directory.CreateDirectory(directory);
                using (var stream = new FileStream(
                    _filePath,
                    FileMode.OpenOrCreate,
                    FileAccess.ReadWrite,
                    FileShare.None,
                    4096,
                    FileOptions.WriteThrough))
                {
                    var records = ReadRecords(stream);
                    var now = _clock.UtcNow;
                    ReplayRecord existing;
                    if (records.TryGetValue(jti, out existing) && existing.ExpiresAtUtc > now)
                    {
                        return false;
                    }

                    records[jti] = new ReplayRecord(jti, expiresAtUtc);
                    WriteRecords(stream, records, now);
                    stream.Flush(true);
                    return true;
                }
            }
            catch (IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        private static Dictionary<string, ReplayRecord> ReadRecords(Stream stream)
        {
            stream.Position = 0;
            var records = new Dictionary<string, ReplayRecord>(StringComparer.Ordinal);
            using (var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    var separator = line.IndexOf('\t');
                    if (separator <= 0 || separator == line.Length - 1)
                    {
                        throw new IOException("The replay store is malformed.");
                    }

                    var jti = Encoding.UTF8.GetString(Convert.FromBase64String(line.Substring(0, separator)));
                    long ticks;
                    if (string.IsNullOrWhiteSpace(jti) ||
                        jti.Length > MaximumJtiLength ||
                        !long.TryParse(line.Substring(separator + 1), NumberStyles.None, CultureInfo.InvariantCulture, out ticks))
                    {
                        throw new IOException("The replay store is malformed.");
                    }

                    records[jti] = new ReplayRecord(jti, new DateTimeOffset(ticks, TimeSpan.Zero));
                }
            }

            return records;
        }

        private static void WriteRecords(Stream stream, Dictionary<string, ReplayRecord> records, DateTimeOffset now)
        {
            stream.Position = 0;
            stream.SetLength(0);
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false), 4096, true))
            {
                var written = 0;
                foreach (var record in records.Values)
                {
                    if (record.ExpiresAtUtc <= now)
                    {
                        continue;
                    }

                    writer.Write(Convert.ToBase64String(Encoding.UTF8.GetBytes(record.Jti)));
                    writer.Write('\t');
                    writer.Write(record.ExpiresAtUtc.UtcDateTime.Ticks.ToString(CultureInfo.InvariantCulture));
                    writer.Write('\n');
                    written++;
                    if (written > CompactAfterRecordCount * 16)
                    {
                        throw new IOException("The replay store exceeded its fail-closed safety bound.");
                    }
                }

                writer.Flush();
            }
        }

        private sealed class ReplayRecord
        {
            public ReplayRecord(string jti, DateTimeOffset expiresAtUtc)
            {
                Jti = jti;
                ExpiresAtUtc = expiresAtUtc;
            }

            public string Jti { get; }

            public DateTimeOffset ExpiresAtUtc { get; }
        }
    }

    public sealed class SystemUtcClock : IUtcClock
    {
        public DateTimeOffset UtcNow => DateTimeOffset.UtcNow;
    }
}
