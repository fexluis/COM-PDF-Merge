using Xunit;

// Deshabilitar el paralelismo para evitar colisiones con el Mutex global de EncryptedSQLiteContext
[assembly: CollectionBehavior(DisableTestParallelization = true)]
