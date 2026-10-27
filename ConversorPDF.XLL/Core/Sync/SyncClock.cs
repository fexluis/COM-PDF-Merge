using System;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// Representa el reloj lógico de sincronización de CR-SQLite.
    /// En CR-SQLite (v0.14+), el tiempo avanza con el db_version (que incrementa por transacción).
    /// El seq es un índice intra-transaccional que se reinicia a 0 con cada nuevo db_version.
    /// </summary>
    internal struct SyncClock : IComparable<SyncClock>
    {
        public long DbVersion { get; }
        public long Seq { get; }

        public SyncClock(long dbVersion, long seq)
        {
            DbVersion = dbVersion;
            Seq = seq;
        }

        public int CompareTo(SyncClock other)
        {
            int dbVersionComparison = DbVersion.CompareTo(other.DbVersion);
            if (dbVersionComparison != 0)
            {
                return dbVersionComparison;
            }
            return Seq.CompareTo(other.Seq);
        }

        public static bool operator >(SyncClock left, SyncClock right)
        {
            return left.CompareTo(right) > 0;
        }

        public static bool operator <(SyncClock left, SyncClock right)
        {
            return left.CompareTo(right) < 0;
        }

        public static bool operator >=(SyncClock left, SyncClock right)
        {
            return left.CompareTo(right) >= 0;
        }

        public static bool operator <=(SyncClock left, SyncClock right)
        {
            return left.CompareTo(right) <= 0;
        }

        public override string ToString()
        {
            return $"v{DbVersion}.{Seq}";
        }
    }
}
