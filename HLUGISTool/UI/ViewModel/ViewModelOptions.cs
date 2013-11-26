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
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.UserControls;

namespace HLU.UI.ViewModel
{
    class ViewModelOptions : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private ICommand _browseMapCommand;
        private string _displayName = "HLU Options";
        private SelectionList<string> _historyColumns;
        private HluDataSet.incid_mm_polygonsDataTable _incidMMPolygonsTable = 
            new HluDataSet.incid_mm_polygonsDataTable();
        private List<int> _gisIDColumnOrdinals;

        private int _dbConnectionTimeout = Settings.Default.DbConnectionTimeout;
        private int _incidTablePageSize = Settings.Default.IncidTablePageSize;
        private int _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;
        private bool _bulkUpdateBlankRowMeansDelete = Settings.Default.BulkUpdateBlankRowMeansDelete;

        private string _mapPath = Settings.Default.MapPath;
        private int _preferredGis = Settings.Default.PreferredGis;
        private bool _warnOnGISSelect = Settings.Default.WarnOnGISSelect;

        private string _seasonSpring = Settings.Default.SeasonNames[0];
        private string _seasonSummer = Settings.Default.SeasonNames[1];
        private string _seasonAutumn = Settings.Default.SeasonNames[2];
        private string _seasonWinter = Settings.Default.SeasonNames[3];
        private string _vagueDateDelimiter = Settings.Default.VagueDateDelimiter;

        private string _bakMapPath;

        #endregion

        #region Constructor

        public ViewModelOptions()
        {
            _gisIDColumnOrdinals = Settings.Default.GisIDColumnOrdinals.Cast<string>()
                .Select(s => Int32.Parse(s)).ToList();
            
            _historyColumns = new SelectionList<string>(_incidMMPolygonsTable.Columns.Cast<DataColumn>()
                .Where(c => !_gisIDColumnOrdinals.Contains(c.Ordinal) && !c.ColumnName.StartsWith("shape_"))
                .Select(c => EscapeAccessKey(c.ColumnName)).ToArray());

            List<int> historyColumnOrdinals = Settings.Default.HistoryColumnOrdinals.Cast<string>()
                .Select(s => Int32.Parse(s)).Where(i => !_gisIDColumnOrdinals.Contains(i) &&
                  !_incidMMPolygonsTable.Columns[i].ColumnName.StartsWith("shape_")).ToList();

            foreach (SelectionItem<string> si in _historyColumns)
                si.IsSelected = historyColumnOrdinals.Contains(
                    _incidMMPolygonsTable.Columns[UnescapeAccessKey(si.Item)].Ordinal);
        }

        private string EscapeAccessKey(string s)
        {
            return s.Replace("_", "__");
        }

        private string UnescapeAccessKey(string s)
        {
            return s.Replace("__", "_");
        }

        #endregion

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public override string WindowTitle
        {
            get { return _displayName; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(bool saveSettings);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Ok Command

        /// <summary>
        /// Create Ok button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand OkCommand
        {
            get
            {
                if (_okCommand == null)
                {
                    Action<object> okAction = new Action<object>(this.OkCommandClick);
                    _okCommand = new RelayCommand(okAction, param => this.CanOk);
                }

                return _okCommand;
            }
        }

        /// <summary>
        /// Handles event when Ok button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void OkCommandClick(object param)
        {
            Settings.Default.DbConnectionTimeout = _dbConnectionTimeout;
            Settings.Default.IncidTablePageSize = _incidTablePageSize;
            Settings.Default.HistoryDisplayLastN = _historyDisplayLastN;
            Settings.Default.BulkUpdateBlankRowMeansDelete = _bulkUpdateBlankRowMeansDelete;

            Settings.Default.MapPath = _mapPath;
            Settings.Default.PreferredGis = _preferredGis;
            Settings.Default.WarnOnGISSelect = _warnOnGISSelect;
            
            Settings.Default.HistoryColumnOrdinals = new StringCollection();
            Settings.Default.HistoryColumnOrdinals.AddRange(_historyColumns.Where(c => c.IsSelected)
                .Select(c => _incidMMPolygonsTable.Columns[UnescapeAccessKey(c.Item)].Ordinal.ToString()).ToArray());

            Settings.Default.SeasonNames[0] = _seasonSpring;
            Settings.Default.SeasonNames[1] = _seasonSummer;
            Settings.Default.SeasonNames[2] = _seasonAutumn;
            Settings.Default.SeasonNames[3] = _seasonWinter;
            Settings.Default.VagueDateDelimiter = _vagueDateDelimiter;

            Settings.Default.Save();

            this.RequestClose(true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return String.IsNullOrEmpty(Error); } }

        #endregion

        #region Cancel Command

        /// <summary>
        /// Create Cancel button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand CancelCommand
        {
            get
            {
                if (_cancelCommand == null)
                {
                    Action<object> cancelAction = new Action<object>(this.CancelCommandClick);
                    _cancelCommand = new RelayCommand(cancelAction);
                }

                return _cancelCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CancelCommandClick(object param)
        {
            Settings.Default.MapPath = _bakMapPath;
            this.RequestClose(false);
        }

        #endregion

        #region Database

        public int DbConnectionTimeout
        {
            get { return _dbConnectionTimeout; }
            set { _dbConnectionTimeout = value; }
        }

        public int IncidTablePageSize
        {
            get { return _incidTablePageSize; }
            set { _incidTablePageSize = value; }
        }

        public int HistoryDisplayLastN
        {
            get { return _historyDisplayLastN; }
            set { _historyDisplayLastN = value; }
        }

        public bool BulkUpdateBlankRowMeansDelete
        {
            get { return _bulkUpdateBlankRowMeansDelete; }
            set { _bulkUpdateBlankRowMeansDelete = value; }
        }

        #endregion

        #region GIS

        public ICommand BrowseMapCommand
        {
            get
            {
                if (_browseMapCommand == null)
                {
                    Action<object> browseMapAction = new Action<object>(this.BrowseMapClicked);
                    _browseMapCommand = new RelayCommand(browseMapAction, param => this.CanBrowseMap);
                }

                return _browseMapCommand;
            }
        }

        public bool CanBrowseMap
        {
            get { return _preferredGis != (int)GISApplications.None; }
        }

        private void BrowseMapClicked(object param)
        {
            _bakMapPath = _mapPath;
            Settings.Default.MapPath = String.Empty;
            MapDocument = String.Empty;
            MapDocument = GISAppFactory.GetMapPath((GISApplications)_preferredGis);
            if (String.IsNullOrEmpty(MapDocument))
            {
                Settings.Default.MapPath = _bakMapPath;
                MapDocument = _bakMapPath;
            }
            OnPropertyChanged("MapDocument");
        }

        public string MapDocument
        {
            get { return Path.GetFileName(_mapPath); }
            set { _mapPath = value; }
        }

        public string MapPath
        {
            get { return _mapPath; }
            set { _mapPath = value; }
        }

        public GISApplications[] GisApps
        {
            get 
            {
                return Enum.GetValues(typeof(GISApplications)).Cast<GISApplications>()
                    .Where(t => t != GISApplications.None).ToArray();
            }
            set { }
        }

        public bool GisAppsEnabled
        {
            get { return GISAppFactory.ArcGisInstalled && GISAppFactory.MapInfoInstalled; }
        }

        public GISApplications PreferredGis
        {
            get { return (GISApplications)_preferredGis; }
            set
            {
                _preferredGis = (int)value;
                OnPropertyChanged("CanBrowseMap");
            }
        }

        public string HistoryColumnsLabel
        {
            get { return "History\nColumns"; }
            set { }
        }

        public SelectionList<string> HistoryColumns
        {
            get { return _historyColumns; }
            set { _historyColumns = value; }
        }

        public bool WarnOnGISSelect
        {
            get { return _warnOnGISSelect; }
            set { _warnOnGISSelect = value; }
        }

        #endregion

        #region Date

        public string SeasonSpring
        {
            get { return _seasonSpring; }
            set { _seasonSpring = value; }
        }

        public string SeasonSummer
        {
            get { return _seasonSummer; }
            set { _seasonSummer = value; }
        }

        public string SeasonAutumn
        {
            get { return _seasonAutumn; }
            set { _seasonAutumn = value; }
        }

        public string SeasonWinter
        {
            get { return _seasonWinter; }
            set { _seasonWinter = value; }
        }

        public string VagueDateDelimiter
        {
            get { return _vagueDateDelimiter; }
            set
            {
                if (value != CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator)
                    _vagueDateDelimiter = value;
            }
        }

        #endregion

        #region IDataErrorInfo Members

        private bool ValidateMapPath(out string message)
        {
            message = null;
            switch (PreferredGis)
            {
                case GISApplications.ArcGIS:
                    if (!Path.GetExtension(_mapPath).Equals(".mxd", StringComparison.CurrentCultureIgnoreCase))
                    {
                        message = String.Format("'{0}' does not appear to be an ArcGIS map document.", MapPath);
                        return false;
                    }
                    break;
                case GISApplications.MapInfo:
                    if (!Path.GetExtension(_mapPath).Equals(".wor", StringComparison.CurrentCultureIgnoreCase))
                    {
                        message = String.Format("'{0}' does not appear to be a MapInfo workspace.", MapPath);
                        return false;
                    }
                    break;
            }
            return true;
        }

        public string Error
        {
            get 
            {
                StringBuilder error = new StringBuilder();

                if (DbConnectionTimeout <= 0)
                    error.Append("\n" + "Please enter a database timeout greater than 0 seconds.");
                if (IncidTablePageSize <= 0)
                    error.Append("\n" + "Please enter a database page size greater than 0 rows.");
                if (HistoryDisplayLastN <= 0)
                    error.Append("\n" + "Number of history rows to be displayed must be greater than 0.");
                if (GisAppsEnabled && (PreferredGis == GISApplications.None))
                    error.Append("\n" + "Please select your preferred GIS application.");
                if (String.IsNullOrEmpty(_mapPath))
                {
                    error.Append("\n" + "Please enter a path to a GIS workspace.");
                }
                else
                {
                    string msg;
                    if (!ValidateMapPath(out msg)) error.Append(msg);
                }
                //---------------------------------------------------------------------
                // FIX: Allow the user to not display any of the history columns
                // Remove the validation enforcing the user to display at least one
                // of the history columns because they are all updated when creating
                // history regardless of whether the user wants to display any of them.
                //
                //if (_historyColumns.Count(h => h.IsSelected) == 0)
                //    error.Append("\n" + "Please select columns to be recorded in history trail.");
                //---------------------------------------------------------------------
                if (String.IsNullOrEmpty(SeasonSpring))
                    error.Append("\n" + "Please enter a season name for spring.");
                if (String.IsNullOrEmpty(SeasonSummer))
                    error.Append("\n" + "Please enter a season name for summer.");
                if (String.IsNullOrEmpty(SeasonAutumn))
                    error.Append("\n" + "Please enter a season name for autumn.");
                if (String.IsNullOrEmpty(SeasonWinter))
                    error.Append("\n" + "Please enter a season name for winter.");
                if (String.IsNullOrEmpty(VagueDateDelimiter))
                    error.Append("\n" + "Please enter a vague date delimiter character.");
                else if (VagueDateDelimiter.Length > 1)
                    error.Append("\n" + "Vague date delimiter must be a single character.");

                if (error.Length > 0)
                    return error.ToString();
                else
                    return null;
            }
        }

        public string this[string columnName]
        {
            get 
            {
                string error = null;

                switch (columnName)
                {
                    case "DbConnectionTimeout":
                        if (DbConnectionTimeout <= 0)
                            error = "Please enter a database timeout greater than 0 seconds.";
                        break;
                    case "IncidTablePageSize":
                        if (IncidTablePageSize <= 0)
                            error = "Please enter a database page size greater than 0 rows.";
                        break;
                    case "HistoryDisplayLastN":
                        if (HistoryDisplayLastN <= 0)
                            error = "Number of history rows to be displayed must be greater than 0.";
                        break;
                    case "PreferredGis":
                        if (GisAppsEnabled && (PreferredGis == GISApplications.None))
                            error = "Please select your preferred GIS application.";
                        break;
                    case "MapPath":
                        if (String.IsNullOrEmpty(MapDocument))
                        {
                            error = "Please enter a path to a GIS workspace.";
                        }
                        else
                        {
                            string msg;
                            if (!ValidateMapPath(out msg)) error = msg;
                        }
                        break;
                    case "HistoryColumns":
                        if (_historyColumns.Count(h => h.IsSelected) == 0)
                            error = "Please select columns to be recorded in history trail.";
                        break;
                    case "SeasonSpring":
                        if (String.IsNullOrEmpty(SeasonSpring))
                            error = "Please enter a season name for spring.";
                        break;
                    case "SeasonSummer":
                        if (String.IsNullOrEmpty(SeasonSummer))
                            error = "Please enter a season name for summer.";
                        break;
                    case "SeasonAutumn":
                        if (String.IsNullOrEmpty(SeasonAutumn))
                            error = "Please enter a season name for autumn.";
                        break;
                    case "SeasonWinter":
                        if (String.IsNullOrEmpty(SeasonWinter))
                            error = "Please enter a season name for winter.";
                        break;
                    case "VagueDateDelimiter":
                        if (String.IsNullOrEmpty(VagueDateDelimiter))
                            error = "Please enter a vague date delimiter character.";
                        else if (VagueDateDelimiter.Length > 1)
                            error = "Vague date delimiter must be a single character.";
                        break;
                }

                CommandManager.InvalidateRequerySuggested();
                
                return error;
            }
        }

        #endregion
    }
}
