// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013, 2016 Thames Valley Environmental Records Centre
// Copyright © 2014 Sussex Biodiversity Record Centre
// Copyright © 2019 London & South East Record Centres (LaSER)
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
using System.Windows.Forms;
using HLU.Data.Model;
using HLU.UI.ViewModel;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.UserControls;

namespace HLU.UI.ViewModel
{
    class ViewModelOptions : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        private ICommand _saveCommand;
        private ICommand _cancelCommand;
        private ICommand _browseMapPathCommand;
        private ICommand _browseExportPathCommand;
        private ICommand _browseSqlPathCommand;
        private string _displayName = "HLU Options";
        private SelectionList<string> _historyColumns;
        private HluDataSet.incid_mm_polygonsDataTable _incidMMPolygonsTable = 
            new HluDataSet.incid_mm_polygonsDataTable();
        private List<int> _gisIDColumnOrdinals;
        private List<string> _showOSMMUpdatesOptions;

        private int? _dbConnectionTimeout = Settings.Default.DbConnectionTimeout;
        private int? _incidTablePageSize = Settings.Default.IncidTablePageSize;
        private int? _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;
        private bool _bulkUpdateBlankRowMeansDelete = Settings.Default.BulkUpdateBlankRowMeansDelete;

        private int _preferredGis = Settings.Default.PreferredGis;
        private string _mapPath = Settings.Default.MapPath;
        private string _exportPath = Settings.Default.ExportPath;
        private int? _minAutoZoom = Settings.Default.MinimumAutoZoom;

        private int? _subsetUpdateAction = Settings.Default.SubsetUpdateAction;
        private string _preferredHabitatClass = Settings.Default.PreferredHabitatClass;
        private bool _showNVCCodes = Settings.Default.ShowNVCCodes;
        private bool _showGroupHeaders = Settings.Default.ShowGroupHeaders;
        private bool _notifyOnSplitMerge = Settings.Default.NotifyOnSplitMerge;
        private string _showOSMMUpdatesOption = Settings.Default.ShowOSMMUpdatesOption;
        private bool _resetOSMMUpdatesStatus = Settings.Default.ResetOSMMUpdatesStatus;

        private int? _warnBeforeGISSelect = Settings.Default.WarnBeforeGISSelect;
        private bool _useAdvancedSQL = Settings.Default.UseAdvancedSQL;
        private int? _getValueRows = Settings.Default.GetValueRows;
        private string _sqlPath = Settings.Default.SqlPath;

        private string _seasonSpring = Settings.Default.SeasonNames[0];
        private string _seasonSummer = Settings.Default.SeasonNames[1];
        private string _seasonAutumn = Settings.Default.SeasonNames[2];
        private string _seasonWinter = Settings.Default.SeasonNames[3];
        private string _vagueDateDelimiter = Settings.Default.VagueDateDelimiter;

        private string _bakMapPath;
        private string _bakExportPath;
        private string _bakSqlPath;

        #endregion

        #region Constructor

        /// <summary>
        /// Get the default values from settings.
        /// </summary>
        /// <remarks></remarks>
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

            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // A new option to enable the user to determine whether to show
            // the OSMM update attributes for the current incid.
            _showOSMMUpdatesOptions = Settings.Default.ShowOSMMUpdatesOptions.Cast<string>().ToList();
            //---------------------------------------------------------------------

            //---------------------------------------------------------------------
            // FIX: 010 Don't clear the map path when cancelling option updates
            // Store the map path so that it can be reset if the user 
            // cancels updates to the options.
            _bakMapPath = _mapPath;
            //---------------------------------------------------------------------
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

        #region Save Command

        /// <summary>
        /// Create Save button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand SaveCommand
        {
            get
            {
                if (_saveCommand == null)
                {
                    Action<object> saveAction = new Action<object>(this.SaveCommandClick);
                    _saveCommand = new RelayCommand(saveAction, param => this.CanSave);
                }

                return _saveCommand;
            }
        }

        /// <summary>
        /// Handles event when Save button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void SaveCommandClick(object param)
        {
            // Database options
            Settings.Default.DbConnectionTimeout = (int)_dbConnectionTimeout;
            Settings.Default.IncidTablePageSize = (int)_incidTablePageSize;
            Settings.Default.HistoryDisplayLastN = (int)_historyDisplayLastN;
            Settings.Default.BulkUpdateBlankRowMeansDelete = _bulkUpdateBlankRowMeansDelete;

            // GIS options
            Settings.Default.HistoryColumnOrdinals = new StringCollection();
            Settings.Default.HistoryColumnOrdinals.AddRange(_historyColumns.Where(c => c.IsSelected)
                .Select(c => _incidMMPolygonsTable.Columns[UnescapeAccessKey(c.Item)].Ordinal.ToString()).ToArray());
            Settings.Default.PreferredGis = _preferredGis;
            Settings.Default.MapPath = _mapPath;
            Settings.Default.ExportPath = _exportPath;
            Settings.Default.MinimumAutoZoom = (int)_minAutoZoom;

            // Interface options
            Settings.Default.SubsetUpdateAction = (int)_subsetUpdateAction;
            Settings.Default.PreferredHabitatClass = _preferredHabitatClass;
            Settings.Default.ShowNVCCodes = _showNVCCodes;
            Settings.Default.ShowGroupHeaders = _showGroupHeaders;
            Settings.Default.NotifyOnSplitMerge = _notifyOnSplitMerge;
            Settings.Default.ShowOSMMUpdatesOption = _showOSMMUpdatesOption;
            Settings.Default.ResetOSMMUpdatesStatus = _resetOSMMUpdatesStatus;

            // SQL Query options
            Settings.Default.WarnBeforeGISSelect = (int)_warnBeforeGISSelect;
            Settings.Default.UseAdvancedSQL = _useAdvancedSQL;
            Settings.Default.SqlPath = _sqlPath;
            Settings.Default.GetValueRows = (int)_getValueRows;

            // Vague Dates options
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
        private bool CanSave { get { return String.IsNullOrEmpty(Error); } }

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
            // Reset the map path incase the user has selected a new
            // path and is now cancelling the changes.
            Settings.Default.MapPath = _bakMapPath;
            this.RequestClose(false);
        }

        #endregion

        #region Database

        public int? DbConnectionTimeout
        {
            get { return _dbConnectionTimeout; }
            set { _dbConnectionTimeout = value; }
        }

        public int? IncidTablePageSize
        {
            get { return _incidTablePageSize; }
            set { _incidTablePageSize = value; }
        }

        public int? HistoryDisplayLastN
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

        public ICommand BrowseMapPathCommand
        {
            get
            {
                if (_browseMapPathCommand == null)
                {
                    Action<object> browseMapPathAction = new Action<object>(this.BrowseMapPathClicked);
                    _browseMapPathCommand = new RelayCommand(browseMapPathAction, param => this.CanBrowseMapPath);
                }

                return _browseMapPathCommand;
            }
        }

        public bool CanBrowseMapPath
        {
            get { return _preferredGis != (int)GISApplications.None; }
        }

        private void BrowseMapPathClicked(object param)
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
            get { return GISAppFactory.ArcGisInstalled &&  GISAppFactory.MapInfoInstalled; }
        }

        public GISApplications PreferredGis
        {
            get { return (GISApplications)_preferredGis; }
            set
            {
                _preferredGis = (int)value;
                OnPropertyChanged("CanBrowseMapPath");
                OnPropertyChanged("MapDocument");
                OnPropertyChanged("CanBrowseExportPath");
            }
        }

        public string HistoryColumnsLabel
        {
            get { return "History\nDisplay\nColumns"; }
            set { }
        }

        public SelectionList<string> HistoryColumns
        {
            get { return _historyColumns; }
            set { _historyColumns = value; }
        }

        #endregion

        #region Interface

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // A new option to enable the user to determine what to do
        // if they try to update a subset of features for the current
        // incid.
        // 
        /// <summary>
        /// Gets or sets the list of available subset update actions from
        /// the enum.
        /// </summary>
        /// <value>
        /// The list of subset update actions.
        /// </value>
        public SubsetUpdateActions[] SubsetUpdateActions
        {
            get
            {
                return Enum.GetValues(typeof(SubsetUpdateActions)).Cast<SubsetUpdateActions>()
                    .ToArray();
            }
            set { }
        }

        /// <summary>
        /// Gets or sets the preferred subset update action.
        /// </summary>
        /// <value>
        /// The preferred subset update action.
        /// </value>
        public SubsetUpdateActions? SubsetUpdateAction
        {
            get { return (SubsetUpdateActions)_subsetUpdateAction; }
            set
            {
                _subsetUpdateAction = (int)value;
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR29 (Habitat classification and conversion to IHS)
        // Add an option for the user to select their preferred
        // habitat class which will be automatically selected when
        // the tool first starts.
        //
        /// <summary>
        /// Gets or sets the list of possible habitat class codes.
        /// </summary>
        /// <value>
        /// The list of possible habitat class codes.
        /// </value>
        public HluDataSet.lut_habitat_classRow[] HabitatClassCodes
        {
            get { return ViewModelWindowMain.HabitatClasses; }
            set { }
        }

        public string PreferredHabitatClass
        {
            get
            {
                var q = HabitatClassCodes.Where(h => h.code == _preferredHabitatClass);
                if (q.Count() > 0)
                    return _preferredHabitatClass;
                else
                    return null;
            }
            set
            {
                _preferredHabitatClass = value;
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // A new option to enable the user to select how many rows
        // to retrieve when getting values for a data column when
        // building a sql query.
        // 
        /// <summary>
        /// Gets or sets the maximum number of value rows to retrieve.
        /// </summary>
        /// <value>
        /// The maximum get value rows.
        /// </value>
        public int? GetValueRows
        {
            get { return _getValueRows; }
            set { _getValueRows = value; }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIX: 056 A new option to enable NVC Codes to be shown or hidden.
        // 
        /// <summary>
        /// Gets or sets the preferred option to show or hide NVC Codes.
        /// </summary>
        /// <value>
        /// The preferred option for showing or hidding NVC Codes.
        /// </value>
        public bool ShowNVCCodes
        {
            get { return _showNVCCodes; }
            set { _showNVCCodes = value; }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIX: 076 A new option to hide group headers to reduce window height.
        // 
        /// <summary>
        /// Gets or sets the preferred option to show or hide group headers.
        /// </summary>
        /// <value>
        /// The preferred option for showing or hidding group headers.
        /// </value>
        public bool ShowGroupHeaders
        {
            get { return _showGroupHeaders; }
            set { _showGroupHeaders = value; }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // A new option to enable the user to determine whether to show
        // the OSMM update attributes for the current incid.
        // 
        /// <summary>
        /// Gets or sets the list of available show OSMM Update options from
        /// the class.
        /// </summary>
        /// <value>
        /// The list of subset update actions.
        /// </value>
        public string[] ShowOSMMUpdatesOptions
        {
            get
            {
                return _showOSMMUpdatesOptions.Cast<string>().ToArray();
            }
            set { }
        }

        /// <summary>
        /// Gets or sets the preferred show OSMM Update option.
        /// </summary>
        /// <value>
        /// The preferred show OSMM Update option.
        /// </value>
        public string ShowOSMMUpdatesOption
        {
            get { return _showOSMMUpdatesOption; }
            set
            {
                _showOSMMUpdatesOption = value;
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // A new option to enable the user to determine whether to reset
        // the OSMM update process flag when manually updating the current
        // incid.
        // 
        /// <summary>
        /// Gets or sets the preferred option to reset the OSMM Update
        /// process flag when applying manual updates.
        /// </summary>
        /// <value>
        /// The preferred option to reset the OSMM Update process flag
        /// when applying manual updates.
        /// </value>
        public bool ResetOSMMUpdatesStatus
        {
            get { return _resetOSMMUpdatesStatus; }
            set { _resetOSMMUpdatesStatus = value; }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // A new option to enable the user to determine when to warn
        // the user before performing a GIS selection.
        // 
        /// <summary>
        /// Gets or sets the list of available warn before GIS selection
        /// options from the enum.
        /// </summary>
        /// <value>
        /// The list of options for warning before any GIS selection.
        /// </value>
        public WarnBeforeGISSelect[] WarnBeforeGISSelectOptions
        {
            get
            {
                return Enum.GetValues(typeof(WarnBeforeGISSelect)).Cast<WarnBeforeGISSelect>()
                    .ToArray();
            }
            set { }
        }

        /// <summary>
        /// Gets or sets the preferred warning before any GIS selection option.
        /// </summary>
        /// <value>
        /// The preferred option for warning before any GIS selection.
        /// </value>
        public WarnBeforeGISSelect? WarnBeforeGISSelect
        {
            get { return (WarnBeforeGISSelect)_warnBeforeGISSelect; }
            set
            {
                _warnBeforeGISSelect = (int)value;
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR39 (Split and merge complete messages)
        // A new option to enable the user to determine if they
        // want to be notified following the completion of a
        // split or merge.
        //
        /// <summary>
        /// Gets or sets the choice of whether the user will
        /// be notified when a split or merge has completed.
        /// </summary>
        /// <value>
        /// If the user will be notified after a split or merge.
        /// </value>
        public bool NotifyOnSplitMerge
        {
            get { return _notifyOnSplitMerge; }
            set { _notifyOnSplitMerge = value; }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIX: 040 Enable MapInfo users to set a default export
        // folder path.
        // 
        /// <summary>
        /// Get the browse Export Path command.
        /// </summary>
        /// <value>
        /// The browse Export path command.
        /// </value>
        public ICommand BrowseExportPathCommand
        {
            get
            {
                if (_browseExportPathCommand == null)
                {
                    Action<object> browseExportPathAction = new Action<object>(this.BrowseExportPathClicked);
                    _browseExportPathCommand = new RelayCommand(browseExportPathAction, param => this.CanBrowseExportPath);
                }

                return _browseExportPathCommand;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the user can browse for
        /// the default Export path.
        /// </summary>
        /// <value>
        /// <c>true</c> if the user can browse for the default Export path; otherwise, <c>false</c>.
        /// </value>
        public bool CanBrowseExportPath
        {
            get { return (PreferredGis == GISApplications.MapInfo); }
    }

        /// <summary>
        /// Action when the browse SQL button is clicked.
        /// </summary>
        /// <param name="param"></param>
        private void BrowseExportPathClicked(object param)
        {
            _bakExportPath = _exportPath;
            ExportPath = String.Empty;
            ExportPath = GetExportPath();

            if (String.IsNullOrEmpty(ExportPath))
            {
                ExportPath = _bakExportPath;
            }
            OnPropertyChanged("ExportPath");
        }

        /// <summary>
        /// Gets or sets the default Export path.
        /// </summary>
        /// <value>
        /// The Export path.
        /// </value>
        public string ExportPath
        {
            get { return _exportPath; }
            set { _exportPath = value; }
        }

        /// <summary>
        /// Prompt the user to set the default Export path.
        /// </summary>
        /// <returns></returns>
        public static string GetExportPath()
        {
            try
            {
                string exportPath = Settings.Default.ExportPath;

                FolderBrowserDialog openFolderDlg = new FolderBrowserDialog();
                openFolderDlg.Description = "Select Export Default Directory";
                openFolderDlg.SelectedPath = exportPath;
                //openFolderDlg.RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFolderDlg.ShowNewFolderButton = true;
                if (openFolderDlg.ShowDialog() == DialogResult.OK)
                {
                    if (Directory.Exists(openFolderDlg.SelectedPath))
                        return openFolderDlg.SelectedPath;
                }
            }
            catch { }

            return null;
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIX: 071 Add minimum auto zoom scale to options.
        // Validate the minimum auto zoom scale.
        //
        /// <summary>
        /// Gets the default minimum auto zoom scale text.
        /// </summary>
        /// <value>
        /// The Minimum auto zoom scale text.
        /// </value>
        public string MinAutoZoomText
        {
            get
            {
                string distUnits = Settings.Default.MapDistanceUnits;
                return string.Format("Minimum auto zoom size ({0}):", distUnits);
            }
        }

        /// <summary>
        /// Gets or sets the default minimum auto zoom scale.
        /// </summary>
        /// <value>
        /// The Minimum auto zoom scale.
        /// </value>
        public int? MinAutoZoom
        {
            get { return _minAutoZoom; }
            set { _minAutoZoom = value; }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Sql Query

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by Attributes Interface)
        // A new option to enable the user to use the new
        // (advanced) 'Select by Attributes' form or the
        // old (original) form.
        //
        /// <summary>
        /// Gets or sets the choice of whether the user will
        /// use the advanced 'Select by Attributes' form.
        /// </summary>
        /// <value>
        /// If the user will use the advanced 'Select by Attributes'
        /// form.
        /// </value>
        public bool UseAdvancedSQL
        {
            get { return _useAdvancedSQL; }
            set
            {
                _useAdvancedSQL = value;
                OnPropertyChanged("CanBrowseSqlPath");
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // A new option to enable the user to set the default
        // folder for saving and loading SQL queries.
        // 
        /// <summary>
        /// Get the browse SQL path command.
        /// </summary>
        /// <value>
        /// The browse SQL path command.
        /// </value>
        public ICommand BrowseSqlPathCommand
        {
            get
            {
                if (_browseSqlPathCommand == null)
                {
                    Action<object> browseSqlPathAction = new Action<object>(this.BrowseSqlPathClicked);
                    _browseSqlPathCommand = new RelayCommand(browseSqlPathAction, param => this.CanBrowseSqlPath);
                }

                return _browseSqlPathCommand;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the user can browse for
        /// the default SQL path.
        /// </summary>
        /// <value>
        /// <c>true</c> if the user can browse for the default SQL path; otherwise, <c>false</c>.
        /// </value>
        public bool CanBrowseSqlPath
        {
            get { return _useAdvancedSQL == true; }
        }

        /// <summary>
        /// Action when the browse SQL path button is clicked.
        /// </summary>
        /// <param name="param"></param>
        private void BrowseSqlPathClicked(object param)
        {
            _bakSqlPath = _sqlPath;
            SqlPath = String.Empty;
            SqlPath = GetSqlPath();

            if (String.IsNullOrEmpty(SqlPath))
            {
                SqlPath = _bakSqlPath;
            }
            OnPropertyChanged("SqlPath");
        }

        /// <summary>
        /// Gets or sets the default SQL path.
        /// </summary>
        /// <value>
        /// The SQL path.
        /// </value>
        public string SqlPath
        {
            get { return _sqlPath; }
            set { _sqlPath = value; }
        }

        /// <summary>
        /// Prompt the user to set the default SQL path.
        /// </summary>
        /// <returns></returns>
        public static string GetSqlPath()
        {
            try
            {
                string sqlPath = Settings.Default.SqlPath;

                FolderBrowserDialog openFolderDlg = new FolderBrowserDialog();
                openFolderDlg.Description = "Select Sql Query Default Directory";
                openFolderDlg.SelectedPath = sqlPath;
                //openFolderDlg.RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFolderDlg.ShowNewFolderButton = true;
                if (openFolderDlg.ShowDialog() == DialogResult.OK)
                {
                    if (Directory.Exists(openFolderDlg.SelectedPath))
                        return openFolderDlg.SelectedPath;
                }
            }
            catch { }

            return null;
        }
        //---------------------------------------------------------------------

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

                //---------------------------------------------------------------------
                // FIX: 017 Validate that mandatory values are not blank.
                // Validate that the database timeout period, database page size and
                // history rows to display are not null.
                if (Convert.ToInt32(DbConnectionTimeout) <= 0 || DbConnectionTimeout == null)
                    error.Append("\n" + "Please enter a database timeout greater than 0 seconds.");
                if (Convert.ToInt32(IncidTablePageSize) <= 0 || IncidTablePageSize == null)
                    error.Append("\n" + "Please enter a database page size greater than 0 rows.");
                if (Convert.ToInt32(HistoryDisplayLastN) <= 0 || HistoryDisplayLastN == null)
                    error.Append("\n" + "Number of history rows to be displayed must be greater than 0.");
                //---------------------------------------------------------------------
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
                // FIX: 071 Add minimum auto zoom scale to options.
                // Validate the minimum auto zoom scale.
                if (Convert.ToInt32(MinAutoZoom) <= 0 || MinAutoZoom == null)
                    error.Append("\n" + "Minimum auto zoom scale must be greater than 0.");
                if (Convert.ToInt32(MinAutoZoom) > Settings.Default.MaximumAutoZoom)
                    error.Append("\n" + String.Format("Minimum auto zoom scale must not be greater than {0}.", Settings.Default.MaximumAutoZoom));
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // CHANGED: CR10 (Attribute updates for incid subsets)
                // Validate the users preferred action when updating a
                // subset of features for an incid.
                if (SubsetUpdateAction == null)
                    error.Append("Please select the action to take when updating an incid subset.");
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // CHANGED: CR29 (Habitat classification and conversion to IHS)
                // Validate the users preferred habitat class.
                if (PreferredHabitatClass == null)
                    error.Append("Please select your preferred habitat class.");
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // A new option to enable the user to determine whether to show
                // the OSMM update attributes for the current incid.
                if (ShowOSMMUpdatesOption == null)
                    error.Append("Please select the option of when to display any OSMM Updates.");
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // CHANGED: CR5 (Select by attributes interface)
                // Validate the maximum number of rows to be retrieved
                // when getting values for a data column when building
                // a sql query.
                if (Convert.ToInt32(GetValueRows) <= 0 || GetValueRows == null)
                    error.Append("\n" + "Number of value rows to be retrieved must be greater than 0.");
                if (Convert.ToInt32(GetValueRows) > Settings.Default.MaxGetValueRows)
                    error.Append("\n" + String.Format("Number of value rows to be retrieved must not be greater than {0}.", Settings.Default.MaxGetValueRows));
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // FIX: 006 Allow the user to not display any of the history columns
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
                    //---------------------------------------------------------------------
                    // FIX: 017 Validate that mandatory values are not blank
                    // Validate that the database timeout period, database page size and
                    // history rows to display are not null.
                    //---------------------------------------------------------------------
                    case "DbConnectionTimeout":
                        if (Convert.ToInt32(DbConnectionTimeout) <= 0 || DbConnectionTimeout == null)
                            error = "Error: Enter a database timeout greater than 0 seconds.";
                        break;
                    case "IncidTablePageSize":
                        if (Convert.ToInt32(IncidTablePageSize) <= 0 || IncidTablePageSize == null)
                            error = "Error: Enter a database page size greater than 0 rows.";
                        break;
                    case "HistoryDisplayLastN":
                        if (Convert.ToInt32(HistoryDisplayLastN) <= 0 || HistoryDisplayLastN == null)
                            error = "Error: Number of history rows to be displayed must be greater than 0.";
                        break;
                    //---------------------------------------------------------------------
                    case "PreferredGis":
                        if (GisAppsEnabled && (PreferredGis == GISApplications.None))
                            error = "Error: Select your preferred GIS application.";
                        break;
                    case "MapDocument":
                        if (String.IsNullOrEmpty(MapDocument))
                        {
                            error = "Error: Enter a path to a GIS workspace.";
                        }
                        else
                        {
                            string msg;
                            if (!ValidateMapPath(out msg)) error = msg;
                        }
                        break;
                    case "HistoryColumns":
                        if (_historyColumns.Count(h => h.IsSelected) == 0)
                            error = "Error: Select columns to be recorded in history trail.";
                        break;
                    //---------------------------------------------------------------------
                    // FIX: 071 Add minimum auto zoom scale to options.
                    // Validate the minimum auto zoom scale.
                    case "MinAutoZoom":
                        if (Convert.ToInt32(MinAutoZoom) <= 0 || MinAutoZoom == null)
                            error = "Error: Minimum auto zoom scale must be greater than 0.";
                        if (Convert.ToInt32(MinAutoZoom) > Settings.Default.MaximumAutoZoom)
                            error = String.Format("Error: Minimum auto zoom scale must not be greater than {0}.", Settings.Default.MaximumAutoZoom);
                        break;
                    //---------------------------------------------------------------------
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Validate the users preferred action when updating a
                    // subset of features for an incid.
                    case "SubsetUpdateAction":
                        if (SubsetUpdateAction == null)
                            error = "Error: Select the action to take when updating an incid subset.";
                        break;
                    //---------------------------------------------------------------------
                    //---------------------------------------------------------------------
                    // CHANGED: CR29 (Habitat classification and conversion to IHS)
                    // Validate the users preferred habitat class.
                    case "PreferredHabitatClass":
                        if (PreferredHabitatClass == null)
                            error = "Error: Select your preferred habitat class.";
                        break;
                    //---------------------------------------------------------------------
                    //---------------------------------------------------------------------
                    // CHANGED: CR49 Process proposed OSMM Updates
                    // A new option to enable the user to determine whether to show
                    // the OSMM update attributes for the current incid.
                    case "ShowOSMMUpdatesOption":
                        if (ShowOSMMUpdatesOption == null)
                            error = "Error: Select option of when to display any OSMM Updates.";
                        break;
                    //---------------------------------------------------------------------
                    //---------------------------------------------------------------------
                    // CHANGED: CR5 (Select by attributes interface)
                    // Validate the maximum number of rows to be retrieved
                    // when getting values for a data column when building
                    // a sql query.
                    case "GetValueRows":
                        if (Convert.ToInt32(GetValueRows) <= 0 || GetValueRows == null)
                            error = "Error: Number of value rows to be retrieved must be greater than 0.";
                        if (Convert.ToInt32(GetValueRows) > Settings.Default.MaxGetValueRows)
                            error = String.Format("Error: Number of value rows to be retrieved must not be greater than {0}.", Settings.Default.MaxGetValueRows);
                        break;
                    //---------------------------------------------------------------------
                    case "SeasonSpring":
                        if (String.IsNullOrEmpty(SeasonSpring))
                            error = "Error: You must enter a season name for spring.";
                        break;
                    case "SeasonSummer":
                        if (String.IsNullOrEmpty(SeasonSummer))
                            error = "Error: You must enter a season name for summer.";
                        break;
                    case "SeasonAutumn":
                        if (String.IsNullOrEmpty(SeasonAutumn))
                            error = "Error: You must enter a season name for autumn.";
                        break;
                    case "SeasonWinter":
                        if (String.IsNullOrEmpty(SeasonWinter))
                            error = "Error: You must enter a season name for winter.";
                        break;
                    case "VagueDateDelimiter":
                        if (String.IsNullOrEmpty(VagueDateDelimiter))
                            error = "Error: You must enter a vague date delimiter character.";
                        else if (VagueDateDelimiter.Length > 1)
                            error = "Error: Vague date delimiter must be a single character.";
                        break;
                }

                CommandManager.InvalidateRequerySuggested();
                
                return error;
            }
        }

        #endregion
    }
}
