using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace OpenExcel.OleDb
{
    public class OleDbExcelReader : IDisposable
    {
        private static string _provider = "Microsoft.ACE.OLEDB.12.0";
        private static string _connStrIMEX = @"Provider=" + _provider + @";Data Source={0};Extended Properties=""Excel 8.0;HDR=No;ReadOnly=True;IMEX=1""";
        private static string _connStrNoIMEX = @"Provider=" + _provider + @";Data Source={0};Extended Properties=""Excel 8.0;HDR=No;ReadOnly=True;""";
        private OleDbConnection _conn;
        private string _path;

        public OleDbExcelWorksheets Worksheets { get; protected set; }
        public ReaderOptions Options { get; protected set; }

        public OleDbExcelReader(string path)
        {
            this.Worksheets = new OleDbExcelWorksheets(this);
            this._path = path;
            this.Options = new ReaderOptions();
        }

        public OleDbExcelReader(string path, ReaderOptions options)
        {
            this.Worksheets = new OleDbExcelWorksheets(this);
            this._path = path;
            this.Options = options;
        }

        internal OleDbConnection OpenConnection(bool useImex)
        {
            if (useImex)
                _conn = new OleDbConnection(string.Format(_connStrIMEX, _path));
            else
                _conn = new OleDbConnection(string.Format(_connStrNoIMEX, _path));
            _conn.Open();
            return _conn;
        }

        #region Dispose/Finalize
        private bool _disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Clean managed resources
                }
                // Clean unmanaged resources
                if (_conn != null)
                    _conn.Dispose();

                _disposed = true;
            }
        }

        ~OleDbExcelReader()
        {
            Dispose(false);
        }
        #endregion
    }
}
