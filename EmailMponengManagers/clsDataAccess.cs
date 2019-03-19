//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.ComponentModel;
using System.Diagnostics;
using System.Configuration;

namespace MWDataManager
{

    public enum ExecutionType
    {
        StoreProcedure,
        GeneralSQLStatement,
        DirectTableAccess
    }

    public enum ReturnType
    {
        SQLDataReader,
        DataTable,
        longNumber,
        DataSet
    }


    class clsDataAccess
    {
        #region class constructors
        public clsDataAccess()
        {
        }
        #endregion class constructors

        #region class properties and globals

        private string _connectionString;
        private string _sqlStatement;
        private SqlParameter[] _paramCollection;
        private string _resultsTableName;
        private DataTable _resultsDataTable;
        private DataSet _resultsDataSet;
        private SqlDataReader _resultsDataReader;
        private long _longAffectedRecords;

        public ExecutionType queryExecutionType;
        public ReturnType queryReturnType;
        public DataTable ResultsDataTable { get { return _resultsDataTable; } }
        public SqlDataReader ResultsDataReader { get { return _resultsDataReader; } }
        public DataSet ResultsDataSet { get { return _resultsDataSet; } }
        public long LongAffectedRecords { get { return _longAffectedRecords; } }
        public SqlParameter[] ParamCollection { get { return _paramCollection; } set { _paramCollection = value; } }
        public string SqlStatement { get { return _sqlStatement; } set { _sqlStatement = value; } }
        public string ResultsTableName { get { return _resultsTableName; } set { _resultsTableName = value; } }
        public string ConnectionString { get { return _connectionString; } set { _connectionString = value; } }


        #endregion class properties and globals

        #region class methods

        public bool ExecuteInstruction()
        {

            SqlConnection _connection = null;
            SqlCommand _command = null;
            SqlDataAdapter _daDataAdapter = null;
            DataTable _dtTable = null;
            DataSet _dsDataSet = null;
            bool _executionResult = false;
           
            try
            {
                _connection = new SqlConnection(this._connectionString);
                //_connection.ConnectionTimeout = (99999);
                //_command.CommandTimeout = 0;
                 _command = new SqlCommand(this._sqlStatement, _connection);
                switch (this.queryExecutionType)
                {
                    case ExecutionType.StoreProcedure:
                        {
                            if (this._paramCollection != null)
                            {
                                foreach (SqlParameter _currentParameter in this._paramCollection)
                                    _command.Parameters.Add(_currentParameter);
                            }
                            _command.CommandType = CommandType.StoredProcedure;
                            break;
                        }
                    case ExecutionType.GeneralSQLStatement:
                        {
                          
                            _command.CommandType = CommandType.Text;
                            break;
                        }
                    case ExecutionType.DirectTableAccess:
                        {
                            _command.CommandType = CommandType.TableDirect;
                            break;
                        }
                    default:
                        break;
                }
                _connection.Open();
                switch (this.queryReturnType)
                {
                    case ReturnType.DataTable:
                        _daDataAdapter = new SqlDataAdapter(_command);
                        _dtTable = new DataTable(this._resultsTableName);
                        _daDataAdapter.Fill(_dtTable);
                        this._resultsDataTable = _dtTable;
                        break;
                    case ReturnType.DataSet:
                        _daDataAdapter = new SqlDataAdapter(_command);
                        _dsDataSet = new DataSet();
                        _daDataAdapter.Fill(_dsDataSet);
                        this._resultsDataSet = _dsDataSet;
                        break;
                    case ReturnType.SQLDataReader:
                        this._resultsDataReader = _command.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
                        break;
                    case ReturnType.longNumber:
                        this._longAffectedRecords = _command.ExecuteNonQuery();
                        break;
                    default:
                        break;
                }
                _executionResult = true;

            }
            catch (Exception _exception)
            {
                throw new ApplicationException(_exception.Message, _exception);
            }
            finally
            {
                if (this.queryReturnType != ReturnType.SQLDataReader)
                    _connection.Close();
                _daDataAdapter = null;
                _connection = null;
                _command = null;
            }
            return _executionResult;
        }

        public SqlParameter CreateParameter(string parameterName, SqlDbType dataType, int size)
        {
            SqlParameter _dbParameter = null;
            try
            {
                _dbParameter = new SqlParameter(parameterName, dataType, size, ParameterDirection.Output, false, 10, 0, "", DataRowVersion.Proposed, "");
                _dbParameter.Value = "";
                return _dbParameter;
            }
            catch (Exception _exception)
            {
                throw new ApplicationException(_exception.Message, _exception);
            }
            finally
            {
                _dbParameter = null;
            }
        }

        public SqlParameter CreateParameter(string parameterName, SqlDbType dataType, int size, object paramValue)
        {
            SqlParameter _dbParameter = null;
            try
            {
                _dbParameter = new SqlParameter(parameterName, dataType, size);
                _dbParameter.Value = paramValue;
                return _dbParameter;
            }
            catch (Exception _exception)
            {
                throw new ApplicationException(_exception.Message, _exception);
            }
            finally
            {
                _dbParameter = null;
            }
        }
        #endregion class methods
    }

}
