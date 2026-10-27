using System;
using System.Data;
using System.Data.Common;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context
{
    /// <summary>
    /// Wrapper DbConnection que impide que el llamador cierre o desheche la conexión subyacente.
    /// El ciclo de vida real pertenece a AcademicDbSession.
    /// </summary>
    internal sealed class NonDisposableDbConnection : DbConnection
    {
        private readonly DbConnection _inner;

        public NonDisposableDbConnection(DbConnection inner) => _inner = inner;

        /// <summary>
        /// Expone la conexión interna para operaciones que requieren el tipo concreto.
        /// </summary>
        public DbConnection Unwrap() => _inner;

        public override string ConnectionString
        {
            get => _inner.ConnectionString;
            set => _inner.ConnectionString = value;
        }

        public override string Database => _inner.Database;
        public override string DataSource => _inner.DataSource;
        public override string ServerVersion => _inner.ServerVersion;
        
        public override ConnectionState State
        {
            get
            {
                try { return _inner.State; }
                catch (ObjectDisposedException) { return ConnectionState.Closed; }
            }
        }

        public override void ChangeDatabase(string databaseName) => _inner.ChangeDatabase(databaseName);

        protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel) => _inner.BeginTransaction(isolationLevel);
        
        protected override DbCommand CreateDbCommand() => new NonDisposableDbCommand(_inner.CreateCommand(), this);

        public override void Open()
        {
            if (State != ConnectionState.Open)
                _inner.Open();
        }

        public override void Close()
        {
            // No-op: la sesión dueña controla el cierre
        }

        protected override void Dispose(bool disposing)
        {
            // No-op: la sesión dueña controla el dispose
        }
    }

    /// <summary>
    /// Wrapper DbCommand para interceptar la asignación de la conexión y evitar InvalidCastException
    /// cuando herramientas como Dapper intentan asignar el wrapper a la propiedad Connection del comando concreto.
    /// </summary>
    internal sealed class NonDisposableDbCommand : DbCommand
    {
        private readonly DbCommand _inner;

        public NonDisposableDbCommand(DbCommand inner, DbConnection connection)
        {
            _inner = inner;
            if (connection != null) DbConnection = connection;
        }

        public override string CommandText { get => _inner.CommandText; set => _inner.CommandText = value; }
        public override int CommandTimeout { get => _inner.CommandTimeout; set => _inner.CommandTimeout = value; }
        public override CommandType CommandType { get => _inner.CommandType; set => _inner.CommandType = value; }
        public override bool DesignTimeVisible { get => _inner.DesignTimeVisible; set => _inner.DesignTimeVisible = value; }
        public override UpdateRowSource UpdatedRowSource { get => _inner.UpdatedRowSource; set => _inner.UpdatedRowSource = value; }

        protected override DbConnection DbConnection
        {
            get => _inner.Connection;
            set => _inner.Connection = value is NonDisposableDbConnection ndc ? ndc.Unwrap() : value;
        }

        protected override DbTransaction DbTransaction
        {
            get => _inner.Transaction;
            set => _inner.Transaction = value;
        }

        public override void Prepare() => _inner.Prepare();
        public override void Cancel() => _inner.Cancel();

        protected override DbParameterCollection DbParameterCollection => _inner.Parameters;

        protected override DbParameter CreateDbParameter() => _inner.CreateParameter();
        protected override DbDataReader ExecuteDbDataReader(CommandBehavior behavior) => _inner.ExecuteReader(behavior);
        public override int ExecuteNonQuery() => _inner.ExecuteNonQuery();
        public override object ExecuteScalar() => _inner.ExecuteScalar();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
        }
    }
}