using System;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Representa el resultado de una operación de base de datos.
    /// </summary>
    /// <typeparam name="T">Tipo del valor de retorno.</typeparam>
    public class SqliteResult<T>
    {
        public bool IsSuccess { get; }
        public bool IsFailure => !IsSuccess;
        public T Value { get; }
        public string Error { get; }

        private SqliteResult(bool isSuccess, T value, string error)
        {
            IsSuccess = isSuccess;
            Value = value;
            Error = error;
        }

        public static SqliteResult<T> Success(T value) => new SqliteResult<T>(true, value, null);
        public static SqliteResult<T> Failure(string error) => new SqliteResult<T>(false, default, error);
    }

    /// <summary>
    /// Representa el resultado de una operación de base de datos sin valor de retorno.
    /// </summary>
    public class SqliteResult
    {
        public bool IsSuccess { get; }
        public bool IsFailure => !IsSuccess;
        public string Error { get; }

        private SqliteResult(bool isSuccess, string error)
        {
            IsSuccess = isSuccess;
            Error = error;
        }

        public static SqliteResult Success() => new SqliteResult(true, null);
        public static SqliteResult Failure(string error) => new SqliteResult(false, error);
    }
}
