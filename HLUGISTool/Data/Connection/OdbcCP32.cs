// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Runtime.InteropServices;

namespace HLU.Data.Connection
{
    /// <summary>
    /// Wrapper class for 32-bit ODBC manager.
    /// </summary>
    class OdbcCP32
    {
        /// <summary>
        /// The driver to use for the datasource.
        /// </summary>
        private const string MS_ACCESS_DRIVER = "Microsoft Access Driver (*.mdb)";

        /// <summary>
        /// A handle to a window that will never be displayed.
        /// </summary>
        private const uint NULL_HWND = 0;

        private enum RequestFlags : int
        {
            /// <summary>
            /// Add a new user data source.
            /// </summary>
            ODBC_ADD_DSN = 1,
            /// <summary>
            /// Configure (modify) an existing user data source.
            /// </summary>
            ODBC_CONFIG_DSN = 2,
            /// <summary>
            /// Remove an existing user data source.
            /// </summary>
            ODBC_REMOVE_DSN = 3,
            /// <summary>
            /// Add a new system data source.
            /// </summary>
            ODBC_ADD_SYS_DSN = 4,
            /// <summary>
            /// Modify an existing system data source.
            /// </summary>
            ODBC_CONFIG_SYS_DSN = 5,
            /// <summary>
            /// Remove an existing system data source.
            /// </summary>
            ODBC_REMOVE_SYS_DSN = 6,
            /// <summary>
            /// Remove the default data source specification section from the system information.
            /// </summary>
            ODBC_REMOVE_DEFAULT_DSN = 7
        }
        
        public OdbcCP32() { }

		#region Win32 API Imports

        [DllImport("ODBCCP32.dll")]
		private static extern bool SQLManageDataSources(IntPtr hwnd);
		
		[DllImport("ODBCCP32.dll")]
		private static extern bool SQLCreateDataSource(IntPtr hwnd, string lpszDS);

        /// <summary>
        /// A method to dynamically add DSN-names to the system. This method also
        /// aids with the creation, and subsequent manipulation, of Microsoft
        /// Access database files.
        /// <see cref="http://msdn.microsoft.com/library/default.asp?url=/library/en-us/odbc/htm/odbcsqlconfigdatasource.asp"/>
        /// </summary>
        /// </summary>
        /// <param name="hwndParent">Parent window handle. The function will not display
        /// any dialog boxes if the handle is null.</param>
        /// <param name="fRequest">One of the OdbcConstant enum values to specify the
        /// type of the request (RequestFlags.ODBC_ADD_DSN to create an MDB).</param>
        /// <param name="lpszDriver">Driver description (usually the name of the
        /// associated DBMS) presented to users instead of the physical driver name.</param>
        /// <param name="lpszAttributes">List of attributes in the form of keyword-value
        /// pairs. For more information, see
        /// <see cref="http://msdn.microsoft.com/library/en-us/odbc/htm/odbcconfigdsn.asp">ConfigDSN</see>
        /// in Chapter 22: Setup DLL Function Reference.</param>
        /// <returns>The function returns TRUE if it is successful, FALSE if it fails.
        /// If no entry exists in the system information when this function is called,
        /// the function returns FALSE.</returns>
        [DllImport("ODBCCP32.DLL", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SQLConfigDataSourceW(UInt32 hwndParent, RequestFlags fRequest, string lpszDriver, string lpszAttributes);
        
        #endregion

		#region ODBC Manager

        public bool ManageDatasources(IntPtr hwnd)
		{
            return SQLManageDataSources(hwnd);
		}

		public bool CreateDatasource(IntPtr hwnd, string szDsn)
		{
			return SQLCreateDataSource(hwnd, szDsn);
        }

        #endregion

        #region MS Jet

        /// <summary>
        /// Compacts an MS Access database.
        /// </summary>
        /// <param name="DatabasePath">The path of the database to be compacted.</param>
        /// <returns>A boolean value indicating success.</returns>
        public bool CompactDatabase(string DatabasePath)
        {
            string attributes = String.Format("COMPACT_DB=\"{0}\" \"{0}\" General\0", DatabasePath);
            return SQLConfigDataSourceW(NULL_HWND, RequestFlags.ODBC_ADD_DSN, MS_ACCESS_DRIVER, attributes);
        }

        /// <summary>
        /// Creates an MS Access database.
        /// </summary>
        /// <param name="DatabasePath">The path of the database to be created.</param>
        /// <returns>A boolean value indicating success.</returns>
        public bool CreateDatabase(string DatabasePath)
        {
            string attributes = String.Format("CREATE_DB=\"{0}\" General\0", DatabasePath);
            return SQLConfigDataSourceW(NULL_HWND, RequestFlags.ODBC_ADD_DSN, MS_ACCESS_DRIVER, attributes);
        }

        /// <summary>
        /// Repairs an MS Access Database.
        /// </summary>
        /// <param name="DatabasePath">The path of the database to be repaired.</param>
        /// <returns>A boolean value indicating success.</returns>
        public bool RepairDatabase(string DatabasePath)
        {
            string attributes = String.Format("REPAIR_DB=\"{0}\" General\0", DatabasePath);
            return SQLConfigDataSourceW(NULL_HWND, RequestFlags.ODBC_ADD_DSN, MS_ACCESS_DRIVER, attributes);
        }

        #endregion
    }
}
