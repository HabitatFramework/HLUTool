// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
// Copyright © 2014 Sussex Biodiversity Record Centre
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
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using HLU.Data;
using HLU.Data.Connection;
using HLU.Data.Model;
using HLU.Data.Model.HluDataSetTableAdapters;
using HLU.Date;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.View;

namespace HLU.UI.ViewModel
{
    /// <summary>
    /// An enumeration of the different options for what to do when
    /// attempting to update a subset of features for an incid.
    /// </summary>
    public enum SubsetUpdateActions
    {
        Prompt,
        Split,
        All
    };

    public class ViewModelWindowMain : ViewModelBase, IDataErrorInfo
    {
        #region Enums

        public enum GeometryTypes { Point, Line, Polygon, Unknown };

        public enum Operations { PhysicalMerge, PhysicalSplit, LogicalMerge, LogicalSplit, AttributeUpdate, BulkUpdate };

        #endregion

        #region Fields

        private ICommand _navigateFirstCommand;
        private ICommand _navigatePreviousCommand;
        private ICommand _navigateNextCommand;
        private ICommand _navigateLastCommand;
        private ICommand _selectByAttributesCommand;
        private ICommand _clearFilterCommand;
        private ICommand _selectOnMapCommand;
        private ICommand _readMapSelectionCommand;
        private ICommand _selectByIncidCommand;
        private ICommand _switchGisLayerCommand;
        private ICommand _splitCommand;
        private ICommand _mergeCommand;
        private ICommand _updateCommand;
        private ICommand _bulkUpdateCommand;
        private ICommand _bulkUpdateCommandMenu;
        private ICommand _cancelBulkUpdateCommand;
        private ICommand _exportCommand;
        private ICommand _closeCommand;
        private ICommand _copyCommand;
        private ICommand _pasteCommand;
        private ICommand _zoomSelectionCommand;
        private ICommand _gisWinMaximiseCommand;
        private ICommand _gisWinSideBySideCommand;
        private ICommand _gisWinSwitchToCommand;
        private ICommand _optionsCommand;
        private ICommand _aboutCommand;

        private WindowMainCopySwitches _copySwitches = new WindowMainCopySwitches();
        private WindowOptions _windowOptions;
        private WindowQueryBuilder _qryBuilderWindow;
        private WindowWarnOnGISSelect _windowWarnGISSelect;
        private WindowWarnOnSubsetUpdate _windowWarnSubsetUpdate;
        private WindowSwitchGISLayer _windowSwitchGISLayer;
        private ViewModelOptions _viewModelOptions;
        private ViewModelQueryBuilder _qryBuilderViewModel;
        private ViewModelWindowWarnOnGISSelect _viewModelWinWarnGISSelect;
        private ViewModelWindowWarnOnSubsetUpdate _viewModelWinWarnSubsetUpdate;
        private ViewModelWindowMainBulkUpdate _viewModelBulkUpdate;
        private ViewModelWindowMainUpdate _viewModelUpd;
        private ViewModelWindowSwitchGISLayer _viewModelSwitchGISLayer;
        private WindowCompletePhysicalSplit _windowCompSplit;
        private ViewModelCompletePhysicalSplit _vmCompSplit;

        private string _displayName = "HLU GIS Tool";
        private string _logoPath = String.Empty;
        private DbBase _db;
        private GISApp _gisApp;
        private GeometryTypes _gisLayerType = GeometryTypes.Polygon;
        private HluDataSet _hluDS;
        private TableAdapterManager _hluTableAdapterMgr;
        private IEnumerable<DataRelation> _hluDataRelations;
        private RecordIds _recIDs;
        private int _incidCurrentRowIndex;
        private DataTable _incidSelection;
        private DataTable _gisSelection;

        private HluDataSet.incidRow _incidCurrentRow;
        private HluDataSet.incidRow _incidCurrentRowClone;
        private HluDataSet.incid_ihs_matrixRow[] _incidIhsMatrixRows;
        private HluDataSet.incid_ihs_formationRow[] _incidIhsFormationRows;
        private HluDataSet.incid_ihs_managementRow[] _incidIhsManagementRows;
        private HluDataSet.incid_ihs_complexRow[] _incidIhsComplexRows;
        private HluDataSet.incid_bapRow[] _incidBapRows;
        private HluDataSet.incid_sourcesRow[] _incidSourcesRows;
        private HluDataSet.historyRow[] _incidHistoryRows;
        private HluDataSet.lut_reasonRow[] _reasonCodes;
        private HluDataSet.lut_processRow[] _processCodes;
        private HluDataSet.lut_ihs_categoryRow[] _ihsCategoryCodes;
        private HluDataSet.lut_ihs_matrixRow[] _lutIhsMatrixCodes;
        private HluDataSet.lut_ihs_formationRow[] _lutIhsFormationCodes;
        private HluDataSet.lut_ihs_managementRow[] _lutIhsManagementCodes;
        private HluDataSet.lut_ihs_complexRow[] _lutIhsComplexCodes;
        private HluDataSet.lut_bap_quality_determinationRow[] _bapDeterminationQualityCodes;
        private HluDataSet.lut_bap_quality_interpretationRow[] _bapInterpretationQualityCodes;
        private HluDataSet.lut_habitat_classRow[] _sourceHabitatClassCodes;
        private HluDataSet.lut_importanceRow[] _sourceImportanceCodes;
        private HluDataSet.lut_habitat_typeRow[] _bapHabitatCodes;
        private ObservableCollection<BapEnvironment> _incidBapRowsAuto;
        private ObservableCollection<BapEnvironment> _incidBapRowsUser;
        private IEnumerable<HluDataSet.lut_ihs_habitatRow> _ihsHabitatCodes;
        private HistoryRowEqualityComparer _histRowEqComp = new HistoryRowEqualityComparer();

        private string[] _nvcCategoryCodes;
        private double _incidArea;
        private double _incidLength;
        private string _process;
        private string _reason;
        private string _nvcCategory;
        private string _ihsCategory;
        private bool _reasonProcessEnabled = true;
        private bool _tabControlDataEnabled = true;
        private bool _tabItemIhsEnabled = true;
        private bool _tabItemDetailsEnabled = true;
        private bool _tabItemSourcesEnabled = true;
        private bool _tabItemHistoryEnabled = true;
        private bool _windowEnabled = true;
        private bool _editMode;
        private bool _pasting = false;
        private bool _changed = false;
        private bool _saving = false;
        private bool _autoSplit = true;
        private bool _splitting = false;
        private bool _comingFromIncidIhsMatrix2 = false;
        private bool _comingFromIncidIhsMatrix3 = false;
        private bool _comingFromIncidIhsFormation2 = false;
        private bool _comingFromIncidIhsManagement2 = false;
        private bool _comingFromIncidIhsComplex2 = false;
        private Cursor _windowCursor = Cursors.Arrow;
        private DataColumn[] _gisIDColumns;
        private int[] _gisIDColumnOrdinals;
        private IEnumerable<string> _incidsSelectedMap;
        private IEnumerable<string> _toidsSelectedMap;
        private IEnumerable<string> _fragsSelectedMap;
        private int _incidsSelectedMapCount = -1;
        private int _toidsSelectedMapCount = -1;
        private int _fragsSelectedMapCount = -1;
        private int _toidsIncidGisCount = -1;
        private int _fragsIncidGisCount = -1;
        private int _toidsIncidDbCount = -1;
        private int _fragsIncidDbCount = -1;
        private int _origIncidIhsMatrixCount = 0;
        private int _origIncidIhsFormationCount = 0;
        private int _origIncidIhsManagementCount = 0;
        private int _origIncidIhsComplexCount = 0;
        private SqlFilterCondition _incidMMPolygonsIncidFilter;
        private DataColumn[] _historyColumns;
        private int _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;
        private int _incidRowCount;
        private int _incidPageRowNoMin = 0;
        private int _incidPageRowNoMax = 0;
        private string _incidIhsHabitat;
        private string _incidLastModifiedUser;
        private DateTime _incidLastModifiedDate;
        private Dictionary<Type, List<SqlFilterCondition>> _childRowFilterDict;
        private Dictionary<Type, string> _childRowOrderByDict;
        private List<List<SqlFilterCondition>> _incidSelectionWhereClause;
        private List<string> _exportMdbs = new List<string>();
        private string _userName;
        private Nullable<bool> _isAuthorisedUser;
        private Nullable<bool> _canBulkUpdate;
        private Nullable<bool> _bulkUpdateMode = false;
        private bool _bulkUpdateCreateHistory;
        private bool _bulkUpdatePrimaryBap = Settings.Default.BulkUpdatePotentialBap;
        private VagueDateInstance _incidSource1DateEntered;
        private VagueDateInstance _incidSource2DateEntered;
        private VagueDateInstance _incidSource3DateEntered;
        private string _codeDeleteRow = Settings.Default.CodeDeleteRow;
        private string _processingMsg = "Processing ...";
        private bool _saved = false;
        private bool _savingAttempted;
        private List<string> _ihsErrors = new List<string>();
        private List<string> _detailsErrors = new List<string>();
        private List<string[]> _source1Errors;
        private List<string[]> _source2Errors;
        private List<string[]> _source3Errors;
        private bool _updateCancelled = true;
        private bool _updateAllFeatures = true;

        public static string HistoryGeometry1ColumnName = Settings.Default.HistoryGeometry1ColumnName;
        public static string HistoryGeometry2ColumnName = Settings.Default.HistoryGeometry2ColumnName;
        internal static string LutDescriptionFieldName = Settings.Default.LutDescriptionFieldName;
        internal static int LutDescriptionFieldOrdinal = Settings.Default.LutDescriptionFieldOrdinal;
        internal static int IncidPageSize = Settings.Default.IncidTablePageSize;

        #endregion

        #region Constructor

        public ViewModelWindowMain()
        {
        }

        internal bool Initialize()
        {
            try
            {
                ProcessStartupArguments();

                bool haveSplashWin = App.SplashViewModel != null;

                // open database connection and test whether it points to a valid HLU database
                while (true)
                {
                    if (haveSplashWin)
                    {
                        DispatcherHelper.DoEvents();
                        App.SplashViewModel.ProgressText = "Opening database connection...";
                    }
                    if ((_db = DbFactory.CreateConnection()) == null)
                        throw new Exception("No database connection.");

                    _hluDS = new HluDataSet();

                    if (haveSplashWin)
                    {
                        App.SplashViewModel.ProgressText = "Inspecting database schema...";
                        DispatcherHelper.DoEvents();
                    }
                    string errorMessage;
                    if (!_db.ContainsDataSet(_hluDS, out errorMessage))
                    {
                        DbFactory.ClearSettings();

                        if (String.IsNullOrEmpty(errorMessage))
                        {
                            errorMessage = String.Empty;
                        }
                        else if (errorMessage.Length > 200)
                        {
                            if (MessageBox.Show("There were errors loading data from the database." +
                                "\n\nWould like to see a list of those errors?", "HLU Dataset",
                                MessageBoxButton.YesNo, MessageBoxImage.Exclamation) == MessageBoxResult.Yes)
                                ShowMessageWindow.ShowMessage(errorMessage, "HLU Dataset");
                            errorMessage = String.Empty;
                        }
                        if (MessageBox.Show("There were errors loading data from the database." +
                            errorMessage + "\n\nWould like to connect to another database?", "HLU Dataset",
                            MessageBoxButton.YesNo, MessageBoxImage.Exclamation) == MessageBoxResult.No)
                            throw new Exception("cancelled");
                    }
                    else
                    {
                        break;
                    }
                }

                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = "Populating dataset...";
                    DispatcherHelper.DoEvents();
                }
                else
                {
                    ChangeCursor(Cursors.Wait, "Processing ...");
                }

                // create table adapter manager for the dataset and connection
                _hluTableAdapterMgr = new TableAdapterManager(_db, TableAdapterManager.Scope.AllButMMPolygonsHistory);

                // fill a dictionary of parent-child tables and relations between them
                _hluDataRelations = HluDataset.Relations.Cast<DataRelation>();

                // translate DataRelation objects into database condtions and build order by clauses
                _childRowFilterDict = BuildChildRowFilters();
                _childRowOrderByDict = BuildChildRowOrderByClauses();

                // fill lookup tables (at least lut_site_id must be filled at this point)
                _hluTableAdapterMgr.Fill(_hluDS, TableAdapterManager.Scope.Lookup, false);

                // create RecordIds object for the db
                _recIDs = new RecordIds(_db, _hluDS, _hluTableAdapterMgr, GisLayerType);

                // wire up event handler for copy switches
                _copySwitches.PropertyChanged += new PropertyChangedEventHandler(_copySwitches_PropertyChanged);

                int result;
                // columns that identify map polygons and are returned by GIS
                _gisIDColumnOrdinals = (from s in Settings.Default.GisIDColumnOrdinals.Cast<string>()
                                        where Int32.TryParse(s, out result) && (result >= 0) &&
                                        (result < _hluDS.incid_mm_polygons.Columns.Count)
                                        select Int32.Parse(s)).ToArray();
                _gisIDColumns = _gisIDColumnOrdinals.Select(i => _hluDS.incid_mm_polygons.Columns[i]).ToArray();
                
                // columns to be saved to history when altering GIS layer (always includes _gisIDColumns)
                _historyColumns = InitializeHistoryColumns(_historyColumns);

                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = "Creating scratch database...";
                    DispatcherHelper.DoEvents();
                }
                ScratchDb.CreateScratchMdb(_hluDS.incid, _hluDS.incid_mm_polygons);

                // count rows of incid table
                IncidRowCount(true);

                // move to first row
                IncidCurrentRowIndex = 1;

                // open the GIS application
                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = "Opening GIS application...";
                    DispatcherHelper.DoEvents();
                }

                // ?
                _gisApp = GISAppFactory.CreateGisApp();
                if (_gisApp == null)
                {
                    GISAppFactory.ClearSettings();
                    return false;
                }
                else if (!_gisApp.Start(ProcessWindowStyle.Maximized))
                {
                    string gisAppName = "GIS application";
                    if (_gisApp != null)
                    {
                        gisAppName = _gisApp.ApplicationType.ToString();
                        _gisApp.Close();
                    }
                    MessageBox.Show(String.Format("{0} failed to open a valid HLU workspace.\nShutting down.", 
                        gisAppName), "HLU", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return false;
                }

                _viewModelUpd = new ViewModelWindowMainUpdate(this);

                // Read the selected features from the map
                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = "Reading map selection...";
                    DispatcherHelper.DoEvents();
                }
                ReadMapSelection(false);

                // Clear the splash window status bar (or reset the cursor to an arrow)
                if (haveSplashWin)
                    App.SplashViewModel.ProgressText = String.Empty;
                else
                    ChangeCursor(Cursors.Arrow, null);

                return true;
            }
            catch (Exception ex)
            {
                if (ex.Message != "cancelled")
                    MessageBox.Show(ex.Message + "\n\nShutting down.", "HLU", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                App.Current.Shutdown();
                return false;
            }
        }

        /// <summary>
        /// Get columns to be saved to history when altering GIS layer.
        /// </summary>
        /// <param name="historyColumns">Old value to be restored if method fails.</param>
        /// <returns>Array of columns to be saved to history when altering GIS layer.
        /// Always includes _gisIDColumns.</returns>
        private DataColumn[] InitializeHistoryColumns(DataColumn[] historyColumns)
        {
            try
            {
                //---------------------------------------------------------------------
                // FIX: Always save all (both) of the history columns
                // Make sure that all the available history columns are updated when
                // creating history even if the user only wants to display some of them.
                return _gisIDColumns.Concat(_hluDS.incid_mm_polygons.Columns.Cast<DataColumn>()
                    .Where(c => !_gisIDColumnOrdinals.Contains(c.Ordinal) && !c.ColumnName.StartsWith("shape_"))).ToArray();
                //int result;
                //return _gisIDColumns.Concat((from s in Settings.Default.HistoryColumnOrdinals.Cast<string>()
                //                                              where Int32.TryParse(s, out result) && (result >= 0) &&
                //                                                   (result < _hluDS.incid_mm_polygons.Columns.Count) &&
                //                                                   !_gisIDColumnOrdinals.Contains(result)
                //                                              select _hluDS.incid_mm_polygons.Columns[Int32.Parse(s)])).ToArray();
                //---------------------------------------------------------------------
            }
            catch { return historyColumns; }
        }

        private void ProcessStartupArguments()
        {
            foreach (string s in App.StartupArguments)
            {
                switch (s.Replace(" ", String.Empty).ToLower())
                {
                    //---------------------------------------------------------------------
                    // FIX: Allow the user to reset only the database settings "/d"
                    // or only the GIS settings "/g" instead of always both "/c"
                    case "/d":
                        DbFactory.ClearSettings();
                        break;
                    case "/g":
                        GISAppFactory.ClearSettings();
                        break;
                    //---------------------------------------------------------------------
                    case "/c":
                        DbFactory.ClearSettings();
                        GISAppFactory.ClearSettings();
                        break;
                    case "/arcgis": // start with ArcGIS
                        Settings.Default.PreferredGis = (int)GISApplications.ArcGIS;
                        if (Settings.Default.MapPath.ToLower().EndsWith(".wor"))
                            Settings.Default.MapPath = String.Empty;
                        break;
                    case "/mapinfo": // start with MapInfo
                        Settings.Default.PreferredGis = (int)GISApplications.MapInfo;
                        if (Settings.Default.MapPath.ToLower().EndsWith(".mxd"))
                            Settings.Default.MapPath = String.Empty;
                        break;
                }
            }
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
            get
            {
                //---------------------------------------------------------------------
                // CHANGED: CR19 (Feature layer position in GIS)
                // Also display the current HLU layer in the window title.
                //
                // If no HLU layer has been identified yet (GIS is still loading) then
                // don't return the layer name
                if (_gisApp.CurrentHluLayer == null)
                    return String.Format("{0}{1}", DisplayName, _editMode ? String.Empty : " [READONLY]");
                else
                {
                    // Include the layer name and active layer/map window number in the window title.
                    return String.Format("{0} - {1} [{2}]{3}", DisplayName, _gisApp.CurrentHluLayer.LayerName, _gisApp.CurrentHluLayer.MapNum, _editMode ? String.Empty : " [READONLY]");
                }
                //---------------------------------------------------------------------
            }
        }

        #endregion

        #region Internal properties

        internal GISApp GISApplication
        {
            get { return _gisApp; }
        }

        internal DbBase DataBase
        {
            get { return _db; }
        }

        internal HluDataSet HluDataset
        {
            get
            {
                if (_hluDS == null) Initialize();
                return _hluDS;
            }
        }

        internal IEnumerable<DataRelation> HluDataRelations
        {
            get { return _hluDataRelations; }
        }

        internal TableAdapterManager HluTableAdapterManager
        {
            get { return _hluTableAdapterMgr; }
        }

        internal int[] GisIDColumnOrdinals
        {
            get { return _gisIDColumnOrdinals; }
        }

        internal DataColumn[] GisIDColumns
        {
            get { return _gisIDColumns; }
        }

        internal DataColumn[] HistoryColumns
        {
            get { return _historyColumns; }
        }

        internal DataTable GisSelection
        {
            get { return _gisSelection; }
            set { _gisSelection = value; }
        }

        internal DataTable IncidSelection
        {
            get { return _incidSelection; }
            set { _incidSelection = value; }
        }

        internal List<List<SqlFilterCondition>> IncidSelectionWhereClause
        {
            get { return _incidSelectionWhereClause; }
            set { _incidSelectionWhereClause = value; }
        }

        internal List<string> ExportMdbs
        {
            get { return _exportMdbs; }
            set { _exportMdbs = value; }
        }

        internal DateTime IncidLastModifiedDateVal
        {
            get { return _incidLastModifiedDate; }
            set { _incidLastModifiedDate = value; }
        }

        internal string IncidLastModifiedUserId
        {
            get { return _incidLastModifiedUser; }
            set { if ((IncidCurrentRow != null) && (value != null)) _incidLastModifiedUser = value; }
        }

        internal HluDataSet.lut_ihs_matrixRow[] IhsMatrixCodes
        {
            get { return _lutIhsMatrixCodes; }
        }

        internal HluDataSet.lut_ihs_formationRow[] IhsFormationCodes
        {
            get { return _lutIhsFormationCodes; }
        }

        internal HluDataSet.lut_ihs_managementRow[] IhsManagementCodes
        {
            get { return _lutIhsManagementCodes; }
        }

        internal HluDataSet.lut_ihs_complexRow[] IhsComplexCodes
        {
            get { return _lutIhsComplexCodes; }
        }

        internal HluDataSet.incid_ihs_matrixRow[] IncidIhsMatrixRows
        {
            get { return _incidIhsMatrixRows; }
            set { _incidIhsMatrixRows = value; }
        }

        internal HluDataSet.incid_ihs_formationRow[] IncidIhsFormationRows
        {
            get { return _incidIhsFormationRows; }
            set { _incidIhsFormationRows = value; }
        }

        internal HluDataSet.incid_ihs_managementRow[] IncidIhsManagementRows
        {
            get { return _incidIhsManagementRows; }
            set { _incidIhsManagementRows = value; }
        }

        internal HluDataSet.incid_ihs_complexRow[] IncidIhsComplexRows
        {
            get { return _incidIhsComplexRows; }
            set { _incidIhsComplexRows = value; }
        }

        internal HluDataSet.incid_bapRow[] IncidBapRows
        {
            get { return _incidBapRows; }
            set { _incidBapRows = value; }
        }

        internal HluDataSet.incid_sourcesRow[] IncidSourcesRows
        {
            get { return _incidSourcesRows; }
            set { _incidSourcesRows = value; }
        }

        internal ObservableCollection<BapEnvironment> IncidBapRowsAuto
        {
            get { return _incidBapRowsAuto; }
            set { _incidBapRowsAuto = value; }
        }

        internal ObservableCollection<BapEnvironment> IncidBapRowsUser
        {
            get { return _incidBapRowsUser; }
            set { _incidBapRowsUser = value; }
        }

        internal RecordIds RecIDs
        {
            get { return _recIDs; }
            set { _recIDs = value; }
        }

        internal bool Saved
        {
            get { return _saved; }
            set { _saved = value; }
        }

        internal bool Pasting
        {
            get { return _pasting; }
            set { _pasting = value; }
        }

        internal bool Changed
        {
            get { return _changed; }
            set
            {
                // If this is another change by the user but the data is no longer
                // dirty (i.e. the user has reversed out their changes) then
                // reset the changed flag.
                if (value == true && !IsDirty)
                    _changed = false;
                else
                    _changed = value; 
            }
        }

        internal bool Saving
        {
            get { return _saving; }
            set { _saving = value; }
        }

        internal bool SavingAttempted
        {
            get { return _savingAttempted; }
            set { _savingAttempted = value; }
        }

        internal IEnumerable<string> IncidsSelectedMap
        {
            get { return _incidsSelectedMap; }
        }

        internal IEnumerable<string> ToidsSelectedMap
        {
            get { return _toidsSelectedMap; }
        }

        internal IEnumerable<string> FragsSelectedMap
        {
            get { return _fragsSelectedMap; }
        }

        internal int IncidsSelectedMapCount
        {
            get { return _incidsSelectedMapCount; }
        }

        internal int ToidsSelectedMapCount
        {
            get { return _toidsSelectedMapCount; }
        }

        internal int FragsSelectedMapCount
        {
            get { return _fragsSelectedMapCount; }
        }

        internal ViewModelWindowMainUpdate ViewModelUpdate
        {
            get { return _viewModelUpd; }
        }

        internal List<string[]> Source1Errors
        {
            get { return _source1Errors; }
            set { _source1Errors = value; }
        }

        internal List<string[]> Source2Errors
        {
            get { return _source2Errors; }
            set { _source2Errors = value; }
        }

        internal List<string[]> Source3Errors
        {
            get { return _source3Errors; }
            set { _source3Errors = value; }
        }

        #endregion

        #region Logo

        public string LogoPath
        {
            get { return Settings.Default.LogoPath; }
            set { }
        }

        public int LogoHeight { get { return Settings.Default.LogoHeight; } }

        public object LogoSource
        {
            get
            {
                BitmapImage image = new BitmapImage();

                try
                {
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                    image.DecodePixelHeight = 100;
                    image.UriSource = new Uri(Settings.Default.LogoPath, UriKind.Absolute);
                    image.EndInit();
                }
                catch
                {
                    return DependencyProperty.UnsetValue;
                }

                return image;
            }
        }

        #endregion
 
        #region Cursor

        public Cursor WindowCursor { get { return _windowCursor; } }

        public bool WindowEnabled { get { return _windowEnabled; } }

        public void ChangeCursor(Cursor cursorType, string processingMessage)
        {
            _windowCursor = cursorType;
            _windowEnabled = cursorType != Cursors.Wait;
            OnPropertyChanged("WindowCursor");
            OnPropertyChanged("WindowEnabled");
            OnPropertyChanged("TabControlDataEnabled");
            if (cursorType == Cursors.Arrow)
                _processingMsg = "Processing ...";
            else
                _processingMsg = processingMessage;
            OnPropertyChanged("StatusTop");
            if (cursorType == Cursors.Wait)
                DispatcherHelper.DoEvents();
        }

        #endregion

        #region User ID

        public string UserID
        {
            get 
            {
                if (!String.IsNullOrEmpty(Environment.UserDomainName))
                    return Environment.UserDomainName + @"\" + Environment.UserName;
                else
                    return Environment.UserName;
            }
        }

        public string UserName
        {
            get { return _userName; }
        }

        public bool IsAuthorisedUser
        {
            get
            {
                if (_isAuthorisedUser == null) GetUserInfo();
                return _isAuthorisedUser == true;
            }
        }

        /// <summary>
        /// Checks the current userid is found in the lut_table, determines
        /// if the user has bulk update authority and retrieves the user's
        /// name.
        /// </summary>
        private void GetUserInfo()
        {
            try
            {
                object result = _db.ExecuteScalar(String.Format("SELECT {0} FROM {1} WHERE {2} = {3}",
                    _db.QuoteIdentifier(_hluDS.lut_user.bulk_updateColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.lut_user.TableName),
                    _db.QuoteIdentifier(_hluDS.lut_user.user_idColumn.ColumnName),
                    _db.QuoteValue(this.UserID)), _db.Connection.ConnectionTimeout, CommandType.Text);
                if (result != null)
                {
                    _isAuthorisedUser = true;
                    _canBulkUpdate = (bool)result;
                }
                else
                {
                    _isAuthorisedUser = false;
                    _canBulkUpdate = false;
                }

                //---------------------------------------------------------------------
                // CHANGED: CR9 (Current userid)
                // Get the current user's username from the lut_table to display with
                // the userid in the 'About' box.
                //
                result = _db.ExecuteScalar(String.Format("SELECT {0} FROM {1} WHERE {2} = {3}",
                    _db.QuoteIdentifier(_hluDS.lut_user.user_nameColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.lut_user.TableName),
                    _db.QuoteIdentifier(_hluDS.lut_user.user_idColumn.ColumnName),
                    _db.QuoteValue(this.UserID)), _db.Connection.ConnectionTimeout, CommandType.Text);
                if (result != null)
                {
                    _userName = (string)result;
                }
                else
                {
                    _userName = "(guest)";
                }
                //---------------------------------------------------------------------
            }
            catch
            {
                _isAuthorisedUser = null;
                _canBulkUpdate = null;
            }
        }

        #endregion

        #region Close Command

        /// <summary>
        /// Returns the command that, when invoked, attempts
        /// to remove this workspace from the user interface.
        /// </summary>
        public ICommand CloseCommand
        {
            get
            {
                if (_closeCommand == null)
                    _closeCommand = new RelayCommand(param => this.OnRequestClose());

                return _closeCommand;
            }
        }

        /// <summary>
        /// Raised when main window should be closed.
        /// </summary>
        public event EventHandler RequestClose;

        public void OnRequestClose()
        {
            EventHandler handler = this.RequestClose;
            if (handler != null)
            {
                MessageBoxResult userResponse = CheckDirty();

                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                       if (!_viewModelUpd.Update()) return;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        return;
                }

                if (HaveGisApp && MessageBox.Show(String.Format("Close {0} as well?",
                    _gisApp.ApplicationType), "HLU: Exit", MessageBoxButton.YesNo,
                    MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    _gisApp.Close();

                    ScratchDb.CleanUp();

                    if (_exportMdbs != null)
                    {
                        foreach (string path in _exportMdbs)
                        {
                            try { File.Delete(path); }
                            catch { }
                        }
                    }
                }

                handler(this, EventArgs.Empty);
            }
        }

        #endregion

        #region Copy and Paste

        public WindowMainCopySwitches CopySwitches
        {
            get { return _copySwitches; }
            set { _copySwitches = value; }
        }

        void _copySwitches_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName.StartsWith("Copy"))
                OnPropertyChanged("CanCopy");
            else
                OnPropertyChanged("CanPaste");
        }

        /// <summary>
        /// Copy command.
        /// </summary>
        public ICommand CopyCommand
        {
            get
            {
                if (_copyCommand == null)
                {
                    Action<object> copyAction = new Action<object>(this.CopyClicked);
                    _copyCommand = new RelayCommand(copyAction, param => this.CanCopy);
                }
                return _copyCommand;
            }
        }

        private void CopyClicked(object param)
        {
            _copySwitches.CopyValues(this);
        }

        public bool CanCopy
        {
            get
            {
                return IncidCurrentRow != null && _copySwitches != null &&
                    typeof(WindowMainCopySwitches).GetProperties().Where(p => p.Name.StartsWith("Copy"))
                    .Count(p => (bool)typeof(WindowMainCopySwitches).GetProperty(p.Name)
                        .GetValue(_copySwitches, null)) > 0;
            }
        }

        /// <summary>
        /// Paste command.
        /// </summary>
        public ICommand PasteCommand
        {
            get
            {
                if (_pasteCommand == null)
                {
                    Action<object> pasteAction = new Action<object>(this.PasteClicked);
                    _pasteCommand = new RelayCommand(pasteAction, param => this.CanPaste);
                }
                return _pasteCommand;
            }
        }

        private void PasteClicked(object param)
        {
            _copySwitches.PasteValues(this);
        }

        public bool CanPaste
        {
            get
            {
                return IncidCurrentRow != null && _copySwitches != null &&
                    typeof(WindowMainCopySwitches).GetProperties().Where(p => !p.Name.StartsWith("Copy"))
                    .Count(p => typeof(WindowMainCopySwitches).GetProperty(p.Name)
                        .GetValue(_copySwitches, null) != null) > 0;
            }
        }

        #endregion

        #region Navigation Commands

        /// <summary>
        /// Navigate to first record command.
        /// </summary>
        public ICommand NavigateFirstCommand
        {
            get
            {
                if (_navigateFirstCommand == null)
                {
                    Action<object> nagigateFirstAction = new Action<object>(this.NavigateFirstClicked);
                    _navigateFirstCommand = new RelayCommand(nagigateFirstAction, param => this.CanNavigateBackward);
                }
                return _navigateFirstCommand;
            }
        }

        private void NavigateFirstClicked(object param)
        {
            if (_bulkUpdateMode == true) return;

            //---------------------------------------------------------------------
            // CHANGED: CR22 (Record selectors)
            // Show the wait cursor and processing message in the status area
            // whilst moving to the new Incid.
            ChangeCursor(Cursors.Wait, "Processing ...");

            IncidCurrentRowIndex = 1;

            ChangeCursor(Cursors.Arrow, null);
            //---------------------------------------------------------------------
        
        }

        /// <summary>
        /// Navigate to previous record command.
        /// </summary>
        public ICommand NavigatePreviousCommand
        {
            get
            {
                if (_navigatePreviousCommand == null)
                {
                    Action<object> nagigatePreviousAction = new Action<object>(this.NavigatePreviousClicked);
                    _navigatePreviousCommand = new RelayCommand(nagigatePreviousAction, param => this.CanNavigateBackward);
                }
                return _navigatePreviousCommand;
            }
        }

        private void NavigatePreviousClicked(object param)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR22 (Record selectors)
            // Show the wait cursor and processing message in the status area
            // whilst moving to the new Incid.
            ChangeCursor(Cursors.Wait, "Processing ...");

            IncidCurrentRowIndex -= 1;

            ChangeCursor(Cursors.Arrow, null);
            //---------------------------------------------------------------------
        }

        private bool CanNavigateBackward
        {
            get { return _bulkUpdateMode == false && IncidCurrentRowIndex > 1; }
        }

        /// <summary>
        /// Navigate to next record command.
        /// </summary>
        public ICommand NavigateNextCommand
        {
            get
            {
                if (_navigateNextCommand == null)
                {
                    Action<object> nagigateNextAction = new Action<object>(this.NavigateNextClicked);
                    _navigateNextCommand = new RelayCommand(nagigateNextAction, param => this.CanNavigateForward);
                }
                return _navigateNextCommand;
            }
        }

        private void NavigateNextClicked(object param)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR22 (Record selectors)
            // Show the wait cursor and processing message in the status area
            // whilst moving to the new Incid.
            ChangeCursor(Cursors.Wait, "Processing ...");

            IncidCurrentRowIndex += 1;

            ChangeCursor(Cursors.Arrow, null);
            //---------------------------------------------------------------------
        }

        private bool CanNavigateForward
        {
            get
            {
                return ((_bulkUpdateMode == false) && (IsFiltered && (IncidCurrentRowIndex < _incidSelection.Rows.Count)) ||
                    (!IsFiltered && (IncidCurrentRowIndex < _incidRowCount)));
            }
        }

        /// <summary>
        /// Navigate to last record command.
        /// </summary>
        public ICommand NavigateLastCommand
        {
            get
            {
                if (_navigateLastCommand == null)
                {
                    Action<object> nagigateLastAction = new Action<object>(this.NavigateLastClicked);
                    _navigateLastCommand = new RelayCommand(nagigateLastAction, param => this.CanNavigateForward);
                }
                return _navigateLastCommand;
            }
        }

        private void NavigateLastClicked(object param)
        {
            if (_bulkUpdateMode == true) return;

            //---------------------------------------------------------------------
            // CHANGED: CR22 (Record selectors)
            // Show the wait cursor and processing message in the status area
            // whilst moving to the new Incid.
            ChangeCursor(Cursors.Wait, "Processing ...");

            IncidCurrentRowIndex = IsFiltered ? _incidSelection.Rows.Count : _incidRowCount;

            ChangeCursor(Cursors.Arrow, null);
            //---------------------------------------------------------------------
        }

        #endregion

        #region Split

        /// <summary>
        /// Split command.
        /// </summary>
        public ICommand SplitCommand
        {
            get
            {
                if (_splitCommand == null)
                {
                    Action<object> splitAction = new Action<object>(this.SplitClicked);
                    _splitCommand = new RelayCommand(splitAction, param => this.CanSplit);
                }
                return _splitCommand;
            }
        }

        /// <summary>
        /// SplitCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void SplitClicked(object param)
        {
            _autoSplit = false;
            ReadMapSelection(false);
            _autoSplit = true;

            ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
            vmSplit.Split();
        }

        private bool CanSplit
        {
            get
            {
                return _bulkUpdateMode == false && HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) && 
                    (_gisSelection != null) && (_gisSelection.Rows.Count > 0) && (_incidsSelectedMapCount == 1) && 
                    (((_toidsSelectedMapCount == 1) && (_fragsSelectedMapCount == 1)) || 
                    ((_toidsSelectedMapCount > 1) || (_fragsSelectedMapCount > 1)) || (_fragsSelectedMapCount == 1));
            }
        }

        #endregion

        #region Merge

        /// <summary>
        /// Merge command.
        /// </summary>
        public ICommand MergeCommand
        {
            get
            {
                if (_mergeCommand == null)
                {
                    Action<object> mergeAction = new Action<object>(this.MergeClicked);
                    _mergeCommand = new RelayCommand(mergeAction, param => this.CanMerge);
                }
                return _mergeCommand;
            }
        }

        private void MergeClicked(object param)
        {
            ReadMapSelection(false);

            ViewModelWindowMainMerge vmMerge = new ViewModelWindowMainMerge(this);
            vmMerge.Merge();
        }

        private bool CanMerge
        {
            get
            {
                return (_bulkUpdateMode == false) && HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) && 
                    _gisSelection != null && _gisSelection.Rows.Count > 1 && (_fragsSelectedMapCount > 1) && 
                    ((_toidsSelectedMapCount == 1) || (_incidsSelectedMapCount > 1));
            }
        }

        #endregion

        #region Update

        /// <summary>
        /// Update command.
        /// </summary>
        public ICommand UpdateCommand
        {
            get
            {
                if (_updateCommand == null)
                {
                    Action<object> updateAction = new Action<object>(this.UpdateClicked);
                    _updateCommand = new RelayCommand(updateAction, param => this.CanUpdate);
                }
                return _updateCommand;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // Enable users to change the attributes for an incid and then
        // when saving them changes check if only a subset of features for
        // the current incid are selected in GIS. If only a subset is
        // selected confirm how to proceed - either logically split the
        // subset of features first before updating, update all of the
        // features for the incid, or cancel the update.
        //
        /// <summary>
        /// UpdateCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void UpdateClicked(object param)
        {
            // If there are no features selected in the GIS (because there is no
            // active filter) then re-select the current incid features in GIS.
            if (_incidsSelectedMapCount <= 0)
                SelectOnMap();

            // If there are any features selected in the GIS ...
            if (_incidsSelectedMapCount > 0)
            {
                // If in bulk update mode then perform the bulk update.
                if (_bulkUpdateMode == true)
                    BulkUpdateClicked(param);
                // If there is no saving already in progress ...
                else if (!_savingAttempted)
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Check if the record has changed and if it hasn't ask the user
                    // if they still want to update the record (to create new history).
                    MessageBoxResult userResponse = CheckClean();

                    switch (userResponse)
                    {
                        case MessageBoxResult.Yes:
                            _viewModelUpd.Update();
                            break;
                        case MessageBoxResult.No:
                            Changed = false;
                            break;
                        case MessageBoxResult.Cancel:
                            return;
                    }
                    //---------------------------------------------------------------------

                    // If all of the features for the current incid have been
                    // selected in GIS then update them all.
                    if (_fragsIncidGisCount == _fragsIncidDbCount)
                    {
                        _viewModelUpd.Update();
                    }
                    // Otherwise, check if/how the subset of features for the
                    // incid should be updated.
                    else if (ConfirmSubsetUpdate())
                    {
                        // The user does not want to update all the features for the incid
                        // then logically split the subset of features first
                        if (_updateAllFeatures == false)
                        {
                            // If a split can be performed then go ahead.
                            if (CanSplit)
                            {

            // If there are still no features selected in the GIS this suggests
            // that the feature layer contains only a subset of the database
            // features so this incid cannot be updated.
            if (_incidsSelectedMapCount <= 0)
                return;

            // If in bulk update mode then perform the bulk update and exit.
            if (_bulkUpdateMode == true)
            {
                BulkUpdateClicked(param);
                return;
            }

            // If there is no filter active (and hence all the features for the
            // current incid are to be updated) or all of the features for the
            // current incid have been selected in GIS then update them all and exit.
            if ((!IsFiltered) || (_fragsIncidGisCount == _fragsIncidDbCount))
            {
                // Update the current incid.
                _saving = true;
                _savingAttempted = false;
                _viewModelUpd.Update();
                return;
            }

            ChangeCursor(Cursors.Wait, "Processing ...");

            DispatcherHelper.DoEvents();

            _gisSelection = NewGisSelectionTable();
            _gisApp.ReadMapSelection(ref _gisSelection);

            // Recheck the selected features in GIS to make sure they
            // all belong to the current incid.
            //_incidSelectionWhereClause = null;
            //AnalyzeGisSelectionSet();
            RefreshStatus();
            CountToidFrags();

            ChangeCursor(Cursors.Arrow, null);

            // If there are no features for the current incid
            // selected in GIS then cancel the update.
            if (_fragsIncidGisCount < 1)
            {
                MessageBox.Show("No map features for the current incid are selected in the map.",
                    "HLU: Save Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            // If all of the features for the current incid have been
            // selected in GIS then update them all.
            else if (_fragsIncidGisCount == _fragsIncidDbCount)
            {
                _saving = true;
                _savingAttempted = false;
                _viewModelUpd.Update();
            }
            else
            {
                // Check if/how the subset of features for the incid should be updated.
                _updateCancelled = false;
                if (ConfirmSubsetUpdate())
                {
                    // The user does not want to update all the features for the incid
                    // then logically split the subset of features first
                    if (_updateAllFeatures == false)
                    {
                        // Set the status to processing and the cursor to wait.
                        ChangeCursor(Cursors.Wait, "Processing ...");

                        // Logically split the features for the current incid into a new incid.
                        ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
                        _splitting = true;
                        if (!vmSplit.Split())
                        {
                            //MessageBox.Show("Could not complete logical split - update cancelled.\nPlease invoke the Split command before applying any updates.",
                            //    "HLU: Save Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            _updateCancelled = true;
                        }
                        _splitting = false;


                        // If the update failed then restore any active filter exactly as
                        // it was.
                        if (_updateCancelled == true)
                        {
                            // Reset the status message and the cursor.
                            ChangeCursor(Cursors.Arrow, null);
                            return;
                        }

                    }
                    // Apply the updates on the current incid.
                    _saving = true;
                    _savingAttempted = false;
                    _viewModelUpd.Update();
                }
                else
                {
                    MessageBox.Show("The changes have not been applied - the update was cancelled.",
                        "HLU: Save Cancelled", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Update is disabled if not currently in edit mode, if no changes have been made by
        /// the user, or if the current record is in error.
        /// </summary>
        private bool CanUpdate
        {
            get { return EditMode && (Changed == true) && String.IsNullOrEmpty(this.Error); }
        }

        /// <summary>
        /// Edit mode is enabled if the user is authorised (i.e. is in the lut_user table),
        /// if there is a GIS application known to be running and if the HLU Layer is currently
        /// being editing in GIS.
        /// </summary>
        public bool EditMode
        {
            get
            {
                bool editMode = IsAuthorisedUser && HaveGisApp && _gisApp.IsEditing;
                if (_editMode != editMode)
                {
                    _editMode = editMode;
                    OnPropertyChanged("WindowTitle");
                    OnPropertyChanged("CanBulkUpdate");
                }
                return _editMode;
            }
            set { }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // Check if the user still wants to go ahead because only a subset
        // of all the features in an incid have been selected. Also checks
        // if the user wants to logically split the subset of features first
        // or updates all the incid features.
        //
        /// <summary>
        /// Confirms with the user if the update is to go ahead
        /// </summary>
        /// <returns>
        /// True if the update is to go ahead, or false if it is cancelled.
        /// </returns>
        private bool ConfirmSubsetUpdate()
        {
            // The user settings indicate that only the selected features
            // should be updated (by logically splitting them first).
            if (Settings.Default.SubsetUpdateAction == 1)
            {
                _updateAllFeatures = false;
                return true;
            }
            // The user settings indicate that all the features for the incid
            // should be updated.
            else if (Settings.Default.SubsetUpdateAction == 2)
            {
                _updateAllFeatures = true;
                return true;
            }
            // If the user settings do not indicate that all the features for the
            // incid should be updated, or that only the selected features should
            // be updated, then prompt the user for their choice.
            else
            {
                _updateCancelled = true;

                _windowWarnSubsetUpdate = new WindowWarnOnSubsetUpdate();
                if ((_windowWarnSubsetUpdate.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowWarnSubsetUpdate.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _viewModelWinWarnSubsetUpdate = new ViewModelWindowWarnOnSubsetUpdate(
                    _fragsIncidGisCount, _toidsIncidGisCount, _fragsIncidDbCount, _toidsIncidDbCount, _gisLayerType);

                // when ViewModel asks to be closed, close window
                _viewModelWinWarnSubsetUpdate.RequestClose +=
                    new ViewModelWindowWarnOnSubsetUpdate.RequestCloseEventHandler(_viewModelWinWarnSubsetUpdate_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowWarnSubsetUpdate.DataContext = _viewModelWinWarnSubsetUpdate;

                // show window
                _windowWarnSubsetUpdate.ShowDialog();

                return (!_updateCancelled);
            }
        }

        void _viewModelWinWarnSubsetUpdate_RequestClose(bool proceed, bool split)
        {
            _viewModelWinWarnSubsetUpdate.RequestClose -= _viewModelWinWarnSubsetUpdate_RequestClose;
            _windowWarnSubsetUpdate.Close();

            // If the user wants to proceed with the update then set whether they
            // want to update all the features or perform a logically split first.
            if (proceed)
            {
                _updateCancelled = false;
                if (split)
                    _updateAllFeatures = false;
                else
                    _updateAllFeatures = true;
            }
            else
            {
                _updateCancelled = true;
            }
            ChangeCursor(Cursors.Arrow, null);
        }
        //---------------------------------------------------------------------

        #endregion

        #region Bulk Update

        /// <summary>
        /// Bulk Update command.
        /// </summary>
        public ICommand BulkUpdateCommand
        {
            get
            {
                if (_bulkUpdateCommand == null)
                {
                    Action<object> bulkUpdateAction = new Action<object>(this.BulkUpdateClicked);
                    _bulkUpdateCommand = new RelayCommand(bulkUpdateAction, param => this.CanBulkUpdate);
                }
                return _bulkUpdateCommand;
            }
        }

        private void BulkUpdateClicked(object param)
        {
            _saving = false;
            if (_viewModelBulkUpdate == null)
                _viewModelBulkUpdate = new ViewModelWindowMainBulkUpdate(this);

            if (_bulkUpdateMode == true)
            {
                _viewModelBulkUpdate.BulkUpdate();
            }
            else
            {
                MessageBoxResult userResponse = CheckDirty();

                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                        if (!_viewModelUpd.Update()) return;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        return;
                }

                _viewModelBulkUpdate.StartBulkUpdate();
            }
        }

        public bool CanBulkUpdate
        {
            get
            {
                if (_canBulkUpdate == null) GetUserInfo();
                return EditMode && _canBulkUpdate == true;
            }
        }

        public ICommand CancelBulkUpdateCommand
        {
            get
            {
                if (_cancelBulkUpdateCommand == null)
                {
                    Action<object> cancelBulkUpdateAction = new Action<object>(this.CancelBulkUpdateClicked);
                    _cancelBulkUpdateCommand = new RelayCommand(cancelBulkUpdateAction, param => this.CanCancelBulkUpdate);
                }
                return _cancelBulkUpdateCommand;
            }
        }

        private void CancelBulkUpdateClicked(object param)
        {
            if (_viewModelBulkUpdate != null)
            {
                _viewModelBulkUpdate.CancelBulkUpdate();
                _viewModelBulkUpdate = null;
            }
        }

        public bool CanCancelBulkUpdate { get { return _bulkUpdateMode == true; } }

        internal Nullable<bool> BulkUpdateMode
        {
            get { return _bulkUpdateMode; }
            set { _bulkUpdateMode = value; }
        }

        public bool HideInBulkUpdateMode 
        { 
            get { return _bulkUpdateMode == false; }
            set { }
        }

        public string BulkUpdateCommandHeader
        {
            get { return _bulkUpdateMode == true ? "Cancel _Bulk Update" : "_Bulk Update"; }
        }

        public ICommand BulkUpdateCommandMenu
        {
            get
            {
                if (_bulkUpdateCommandMenu == null)
                {
                    Action<object> bulkUpdateMenuAction = new Action<object>(this.BulkUpdateCommandMenuClicked);
                    _bulkUpdateCommandMenu = new RelayCommand(bulkUpdateMenuAction);
                }
                return _bulkUpdateCommandMenu;
            }
        }

        private void BulkUpdateCommandMenuClicked(object param)
        {
            if (_bulkUpdateMode == true)
                CancelBulkUpdateClicked(param);
            else
                BulkUpdateClicked(param);
        }

        public bool ShowInBulkUpdateMode
        {
            get { return _bulkUpdateMode == true; }
            set { }
        }

        public bool BulkUpdateCreateHistory
        {
            get { return _bulkUpdateCreateHistory; }
            set { _bulkUpdateCreateHistory = value; }
        }

        public string TopControlsGroupHeader
        {
            get { return (_bulkUpdateMode == false) ? "INCID" : "Bulk Update"; }
        }

        #endregion

        #region View

        public ICommand ZoomSelectionCommand
        {
            get
            {
                if (_zoomSelectionCommand == null)
                {
                    Action<object> zoomSelectionAction = new Action<object>(this.ZoomSelectionClicked);
                    _zoomSelectionCommand = new RelayCommand(zoomSelectionAction, param => this.CanZoomSelection);
                }
                return _zoomSelectionCommand;
            }
        }

        private void ZoomSelectionClicked(object param)
        {
            _gisApp.ZoomSelected();
        }

        public bool CanZoomSelection { get { return HaveGisApp && _gisSelection != null; } }

        public ICommand GisWinMaximiseCommand
        {
            get
            {
                if (_gisWinMaximiseCommand == null)
                {
                    Action<object> gisWinMaximiseAction = new Action<object>(this.GisWinMaximiseClicked);
                    _gisWinMaximiseCommand = new RelayCommand(gisWinMaximiseAction, param => this.CanGisWinMaximise);
                }
                return _gisWinMaximiseCommand;
            }
        }

        private void GisWinMaximiseClicked(object param)
        {
            _gisApp.Window(ProcessWindowStyle.Maximized, IntPtr.Zero);
        }

        public bool CanGisWinMaximise { get { return HaveGisApp; } }
        
        public ICommand GisWinSideBySideCommand
        {
            get
            {
                if (_gisWinSideBySideCommand == null)
                {
                    Action<object> gisWinSideBySideAction = new Action<object>(this.GisWinSideBySideClicked);
                    _gisWinSideBySideCommand = new RelayCommand(gisWinSideBySideAction, param => this.CanGisWinSideBySide);
                }
                return _gisWinSideBySideCommand;
            }
        }

        private void GisWinSideBySideClicked(object param)
        {
            _gisApp.Window(ProcessWindowStyle.Normal, 
                new WindowInteropHelper(App.Current.MainWindow).Handle);
        }

        public bool CanGisWinSideBySide { get { return HaveGisApp; } }

        public ICommand GisWinSwitchToCommand
        {
            get
            {
                if (_gisWinSwitchToCommand == null)
                {
                    Action<object> gisWinSwitchToAction = new Action<object>(this.GisWinSwitchToClicked);
                    _gisWinSwitchToCommand = new RelayCommand(gisWinSwitchToAction, param => this.CanGisWinSwitchTo);
                }
                return _gisWinSwitchToCommand;
            }
        }

        private void GisWinSwitchToClicked(object param)
        {
            _gisApp.Activate();
        }

        public bool CanGisWinSwitchTo { get { return HaveGisApp; } }

        #endregion

        #region Options

        public ICommand OptionsCommand
        {
            get
            {
                if (_optionsCommand == null)
                {
                    Action<object> optionsAction = new Action<object>(this.OptionsClicked);
                    _optionsCommand = new RelayCommand(optionsAction);
                }
                return _optionsCommand;
            }
        }

        private void OptionsClicked(object param)
        {
            _windowOptions = new WindowOptions();
            _windowOptions.Owner = App.Current.MainWindow;
            _windowOptions.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            _viewModelOptions = new ViewModelOptions();
            _viewModelOptions.RequestClose += 
                new ViewModelOptions.RequestCloseEventHandler(_viewModelOptions_RequestClose);

            _windowOptions.DataContext = _viewModelOptions;

            _windowOptions.ShowDialog();
        }

        void _viewModelOptions_RequestClose(bool saveSettings)
        {
            _viewModelOptions.RequestClose -= _viewModelOptions_RequestClose;
            _windowOptions.Close();

            // re-set static variables (IncidPageSize might be dangerous to change on the fly)
            if (saveSettings)
            {
                _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;
                _historyColumns = InitializeHistoryColumns(_historyColumns);
                VagueDate.Delimiter = Settings.Default.VagueDateDelimiter;
                VagueDate.SeasonNames = Settings.Default.SeasonNames.Cast<string>().ToArray();
            }
        }

        #endregion

        #region About

        public ICommand AboutCommand
        {
            get
            {
                if (_aboutCommand == null)
                {
                    Action<object> aboutAction = new Action<object>(this.AboutClicked);
                    _aboutCommand = new RelayCommand(aboutAction);
                }
                return _aboutCommand;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR9 (Current userid)
        // Retrieve the copyright notice for the assembly to display with the
        // current userid and name in the 'About' box.
        //
        /// <summary>
        /// Gets the assembly copyright notice.
        /// </summary>
        /// <value>The assembly copyright.</value>
        public string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                // If there aren't any Copyright attributes, return an empty string
                if (attributes.Length == 0)
                    return null;
                // If there is a Copyright attribute, return its value
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }
        //---------------------------------------------------------------------

        private void AboutClicked(object param)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR9 (Current userid)
            // Show the current userid and username together with the version
            // and copyright notice in the 'About' box.
            //
            MessageBox.Show(App.Current.MainWindow, String.Format("   {0} Version:  {1}\n\n   Userid:   {2}\n   Username: {3}\n\n   {4}",
                _displayName, Assembly.GetExecutingAssembly().GetName().Version.ToString(), UserID, UserName, AssemblyCopyright), 
                String.Format("About {0}", _displayName), MessageBoxButton.OK, MessageBoxImage.Information);
            //---------------------------------------------------------------------
        }

        #endregion

        #region Export

        /// <summary>
        /// Export command.
        /// </summary>
        public ICommand ExportCommand
        {
            get
            {
                if (_exportCommand == null)
                {
                    Action<object> exportAction = new Action<object>(this.ExportClicked);
                    _exportCommand = new RelayCommand(exportAction, param => this.CanExport);
                }
                return _exportCommand;
            }
        }

        private void ExportClicked(object param)
        {
            ViewModelWindowMainExport viewModelExport = new ViewModelWindowMainExport(this);
            viewModelExport.InitiateExport();
        }

        public bool CanExport { get { return _bulkUpdateMode == false && _hluDS != null && HaveGisApp; } }

        #endregion

        #region Select

        #region Select by Attributes Command

        /// <summary>
        /// SelectByAttributes command.
        /// </summary>
        public ICommand SelectByAttributesCommand
        {
            get
            {
                if (_selectByAttributesCommand == null)
                {
                    Action<object> selectByAttributesAction = new Action<object>(this.SelectByAttributesClicked);
                    _selectByAttributesCommand = new RelayCommand(selectByAttributesAction, param => this.CanSelectByAttributes);
                }
                return _selectByAttributesCommand;
            }
        }

        private void SelectByAttributesClicked(object param)
        {
            OpenQueryBuilder();
        }

        private bool CanSelectByAttributes
        {
            get { return _bulkUpdateMode == false && IncidCurrentRow != null; }
        }

        private void OpenQueryBuilder()
        {
            try
            {
                _qryBuilderWindow = new WindowQueryBuilder();
                if ((_qryBuilderWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _qryBuilderWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _qryBuilderViewModel = new ViewModelQueryBuilder(HluDataset);
                _qryBuilderViewModel.DisplayName = "HLU Query Builder";

                // when ViewModel asks to be closed, close window
                _qryBuilderViewModel.RequestClose +=
                    new ViewModelQueryBuilder.RequestCloseEventHandler(_qryBuilderViewModel_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _qryBuilderWindow.DataContext = _qryBuilderViewModel;

                // show window
                _qryBuilderWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        protected void _qryBuilderViewModel_RequestClose(ObservableCollection<QueryItem> queryItems)
        {
            _qryBuilderViewModel.RequestClose -= _qryBuilderViewModel_RequestClose;
            _qryBuilderWindow.Close();

            if (queryItems != null)
            {
                try
                {
                    ChangeCursor(Cursors.Wait, "Processing ...");

                    // extract where conditions from query items
                    _incidSelectionWhereClause = new List<List<SqlFilterCondition>>();
                    int i = 0;
                    while (i < queryItems.Count)
                    {
                        List<SqlFilterCondition> whereBlock = new List<SqlFilterCondition>();
                        int openParenths = 0;
                        int closeParenths = 0;
                        int j = i;
                        while (j < queryItems.Count)
                        {
                            QueryItem qi = queryItems.ElementAt(j);
                            if (!String.IsNullOrEmpty(qi.OpenParentheses))
                                openParenths += qi.OpenParentheses.Trim().Length;
                            if (!String.IsNullOrEmpty(qi.CloseParentheses))
                                closeParenths += qi.CloseParentheses.Trim().Length;
                            if (qi.IsComplete)
                            {
                                if ((whereBlock.Count >= IncidPageSize) &&
                                    (qi.WhereCondition.BooleanOperator.ToUpper() == "OR") &&
                                    (closeParenths == openParenths))
                                {
                                    _incidSelectionWhereClause.Add(whereBlock);
                                    break;
                                }
                                else
                                {
                                    whereBlock.Add(qi.WhereCondition);
                                    if (j == queryItems.Count - 1)
                                        _incidSelectionWhereClause.Add(whereBlock);
                                    j++;
                                }
                            }
                        }
                        i = j;
                    }

                    // create a selection DataTable of PK values of IncidTable
                    if (_incidSelectionWhereClause.Count > 0)
                    {
                        _incidSelection = _db.SqlSelect(true, IncidTable.PrimaryKey, _incidSelectionWhereClause[0]);
                        for (i = 1; i < _incidSelectionWhereClause.Count; i++)
                        {
                            DataTable tSel = _db.SqlSelect(true, IncidTable.PrimaryKey, _incidSelectionWhereClause[i]);
                            foreach (DataRow r in tSel.Rows)
                                _incidSelection.ImportRow(r);
                        }
                    }
                    else
                    {
                        _incidSelectionWhereClause = null;
                        _incidSelection = null;
                    }

                    if (IsFiltered)
                    {
                        if (ConfirmGISSelect()) PerformGisSelection(true);
                    }
                    else
                    {
                        MessageBox.Show(App.Current.MainWindow, "No records found.", "HLU Query",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                        _incidSelection = null;
                        ChangeCursor(Cursors.Arrow, null);
                    }
                }
                catch (Exception ex)
                {
                    _incidSelection = null;
                    ChangeCursor(Cursors.Arrow, null);
                    MessageBox.Show(App.Current.MainWindow, ex.Message, "HLU", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally { RefreshStatus(); }
            }
        }

        private bool ConfirmGISSelect()
        {
            if (!Settings.Default.WarnOnGISSelect)
            {
                return true;
            }
            else
            {
                int expectedNumFeatures = ExpectedSelectionFeatures();

                _windowWarnGISSelect = new WindowWarnOnGISSelect();
                if ((_windowWarnGISSelect.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowWarnGISSelect.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _viewModelWinWarnGISSelect = new ViewModelWindowWarnOnGISSelect(
                    expectedNumFeatures, expectedNumFeatures > -1 ? _gisLayerType : GeometryTypes.Unknown);

                // when ViewModel asks to be closed, close window
                _viewModelWinWarnGISSelect.RequestClose += 
                    new ViewModelWindowWarnOnGISSelect.RequestCloseEventHandler(_viewModelWinWarnGISSelect_RequestClose);
                    
                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowWarnGISSelect.DataContext = _viewModelWinWarnGISSelect;

                // show window
                _windowWarnGISSelect.ShowDialog();

                return IsFiltered;
            }
        }

        void _viewModelWinWarnGISSelect_RequestClose(bool proceed)
        {
            _viewModelWinWarnGISSelect.RequestClose -= _viewModelWinWarnGISSelect_RequestClose;
            _windowWarnGISSelect.Close();

            if (!proceed)
            {
                _incidSelectionWhereClause = null;
                _incidSelection = null;
                ChangeCursor(Cursors.Arrow, null);
            }
        }

        #endregion

        #region Select On Map Command

        /// <summary>
        /// ShowOnMap command.
        /// </summary>
        public ICommand SelectOnMapCommand
        {
            get
            {
                if (_selectOnMapCommand == null)
                {
                    Action<object> selectOnMapAction = new Action<object>(this.SelectOnMapClicked);
                    _selectOnMapCommand = new RelayCommand(selectOnMapAction, param => this.CanSelectOnMap);
                }
                return _selectOnMapCommand;
            }
        }

        private void SelectOnMapClicked(object param)
        {
            SelectOnMap();
        }

        private bool CanSelectOnMap
        {
            get { return _bulkUpdateMode == false && HaveGisApp && IncidCurrentRow != null; }
        }

        /// <summary>
        /// Select current DB record on map when button pressed.
        /// </summary>
        private void SelectOnMap()
        {
            if (IncidCurrentRow == null) return;

            try
            {
                // Save the current table of selected incids.
                DataTable prevIncidSelection = NewIncidSelectionTable();
                prevIncidSelection = _incidSelection;

                // Save the current table of selected GIS features.
                DataTable prevGISSelection = NewGisSelectionTable();
                prevGISSelection = _gisSelection;

                // Reset the table of selected incids.
                _incidSelection = NewIncidSelectionTable();

                // Set the table of selected incids to the current incid.
                DataRow selRow = _incidSelection.NewRow();
                foreach (DataColumn c in _incidSelection.Columns)
                    selRow[c] = IncidCurrentRow[c.ColumnName];
                _incidSelection.Rows.Add(selRow);

                // Select all the features for the current incid in GIS
                // and don't set the filter (i.e. don't go to the first
                // record).
                PerformGisSelection(false);

                //Restore the current table of selected incids.
                _incidSelection = prevIncidSelection;

                //Restore the current table of selected GIS features.
                _gisSelection = prevGISSelection;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Read Map Selection Command

        /// <summary>
        /// ReadMapSelection command.
        /// </summary>
        public ICommand ReadMapSelectionCommand
        {
            get
            {
                if (_readMapSelectionCommand == null)
                {
                    Action<object> readMapSelectionAction = new Action<object>(this.ReadMapSelectionClicked);
                    _readMapSelectionCommand = new RelayCommand(readMapSelectionAction, param => this.CanReadMapSelection);
                }
                return _readMapSelectionCommand;
            }
        }

        private void ReadMapSelectionClicked(object param)
        {
            ReadMapSelection(true);
        }

        private bool CanReadMapSelection
        {
            get { return _bulkUpdateMode == false && HaveGisApp; }
        }

        internal void ReadMapSelection(bool showMessage)
        {
            try
            {
                ChangeCursor(Cursors.Wait, "Processing ...");

                DispatcherHelper.DoEvents();

                _gisSelection = NewGisSelectionTable();
                _gisApp.ReadMapSelection(ref _gisSelection);
                _incidSelectionWhereClause = null;
                AnalyzeGisSelectionSet();
                if (_gisSelection.Rows.Count > 0)
                {
                    SetFilter();

                    if (_autoSplit && (_gisSelection != null) && (_gisSelection.Rows.Count > 1) && (_incidsSelectedMapCount == 1) &&
                        (_toidsSelectedMapCount == 1) && (_fragsSelectedMapCount == 1))
                    {
                        if (IsAuthorisedUser)
                        {
                            if (!CanSplit)
                            {
                                _windowCompSplit = new WindowCompletePhysicalSplit();
                                _windowCompSplit.Owner = App.Current.MainWindow;
                                _windowCompSplit.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                                _vmCompSplit = new ViewModelCompletePhysicalSplit(_reason, _process, _reasonCodes, _processCodes);
                                _vmCompSplit.RequestClose += new ViewModelCompletePhysicalSplit.RequestCloseEventHandler(vmCompSplit_RequestClose);
                                _windowCompSplit.DataContext = _vmCompSplit;
                                _windowCompSplit.ShowDialog();
                            }
                            if (CanSplit)
                            {
                                ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
                                if (vmSplit.Split()) MessageBox.Show("Physical split completed.",
                                    "HLU: Physical Split", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            else
                            {
                                MessageBox.Show("Could not complete physical split.\nPlease invoke the Split command before altering the map selection.",
                                    "Physical Split", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Could not complete physical split because you are not an authorized user.\n" + 
                                "Please undo your map changes to prevent map and database going out of sync.",
                                "Physical Split", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        }
                    }
                }
                else
                {
                    if (showMessage) MessageBox.Show("No map features selected.", "HLU Selection", 
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    // Reset the incid and map selections and move
                    // to the first incid in the database.
                    ClearFilter(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ChangeCursor(Cursors.Arrow, null);
            }
        }

        void vmCompSplit_RequestClose(string reason, string process)
        {
            _vmCompSplit.RequestClose -= vmCompSplit_RequestClose;
            _windowCompSplit.Close();
            if (!String.IsNullOrEmpty(reason))
            {
                _reason = reason;
                OnPropertyChanged("Reason");
            }
            if (!String.IsNullOrEmpty(process))
            {
                _process = process;
                OnPropertyChanged("Process");
            }
        }

        #endregion

        #region Select By Incid Command

        /// <summary>
        /// SelectByIncid command.
        /// </summary>
        public ICommand SelectByIncidCommand
        {
            get
            {
                if (_selectByIncidCommand == null)
                {
                    Action<object> selectByIncidAction = new Action<object>(this.SelectByIncidClicked);
                    _selectByIncidCommand = new RelayCommand(selectByIncidAction, param => this.CanSelectByIncid);
                }
                return _selectByIncidCommand;
            }
        }

        private void SelectByIncidClicked(object param)
        {
            SelectByIncid();
        }

        private bool CanSelectByIncid
        {
            get { return _bulkUpdateMode == false && HaveGisApp && _gisSelection != null && _incidsSelectedMapCount == 1; }
        }

        /// <summary>
        /// Select current DB record on map when button pressed.
        /// </summary>
        private void SelectByIncid()
        {
            if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0) && (_incidsSelectedMapCount == 1))
            {
                _incidSelection = NewIncidSelectionTable();
                DataRow selRow = _incidSelection.NewRow();
                foreach (DataColumn c in _incidSelection.Columns)
                {
                    if (_gisSelection.Columns.Contains(c.ColumnName))
                        selRow[c] = _gisSelection.Rows[0][c.ColumnName];
                }
                _incidSelection.Rows.Add(selRow);

                PerformGisSelection(false);
            }
        }

        #endregion

        #region Clear Filter Command

        /// <summary>
        /// ClearFilter command.
        /// </summary>
        public ICommand ClearFilterCommand
        {
            get
            {
                if (_clearFilterCommand == null)
                {
                    Action<object> qryBuilderAction = new Action<object>(this.ClearFilterClicked);
                    _clearFilterCommand = new RelayCommand(qryBuilderAction, param => this.IsFiltered);
                }
                return _clearFilterCommand;
            }
        }

        private void ClearFilterClicked(object param)
        {
            // Reset the incid and map selections and move
            // to the first incid in the database.
            ClearFilter(true);
        }

        /// <summary>
        /// Clears any active incid filter and optionally moves to the first incid in the index.
        /// </summary>
        /// <param name="resetRowIndex">If set to <c>true</c> the first incid in the index is loaded.</param>
        internal void ClearFilter(bool resetRowIndex)
        {
            if (IsFiltered)
            {
                _incidSelection = null;
                _incidSelectionWhereClause = null;
                _gisSelection = null;
                _toidsSelectedMapCount = -1;
                _fragsSelectedMapCount = -1;
                _incidPageRowNoMax = -1;

                //---------------------------------------------------------------------
                // CHANGED: CR10 (Attribute updates for incid subsets)
                // Only move to the first incid in the index if required, to save
                // changing the index here and then again immediately after from
                // the calling method.
                if (resetRowIndex)
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR22 (Record selectors)
                    // Show the wait cursor and processing message in the status area
                    // whilst moving to the new Incid.
                    ChangeCursor(Cursors.Wait, "Processing ...");

                    IncidCurrentRowIndex = 1;

                    ChangeCursor(Cursors.Arrow, null);
                    //---------------------------------------------------------------------

                    OnPropertyChanged("IsFiltered");
                    OnPropertyChanged("CanBulkUpdate");
                    OnPropertyChanged("CanZoomSelection");
                    RefreshStatus();
                }
                //---------------------------------------------------------------------
            }
        }

        #endregion

        #region Select Helpers

        /// <summary>
        /// Count how many incids, toids and fragments are selected in GIS.
        /// </summary>
        private void AnalyzeGisSelectionSet()
        {
            _incidsSelectedMapCount = -1;
            _toidsSelectedMapCount = -1;
            _fragsSelectedMapCount = -1;
            if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0))
            {
                switch (_gisSelection.Columns.Count)
                {
                    case 3:
                        // Count the number of fragments selected in GIS.
                        _fragsSelectedMap = from r in _gisSelection.AsEnumerable()
                                            group r by new
                                            {
                                                incid = r.Field<string>(0),
                                                toid = r.Field<string>(1),
                                                fragment = r.Field<string>(2)
                                            }
                                                into g
                                                select g.Key.fragment;
                        _fragsSelectedMapCount = _fragsSelectedMap.Count();
                        goto case 2;
                    case 2:
                        // Count the number of toids selected in GIS.
                        _toidsSelectedMap = _gisSelection.AsEnumerable()
                            .GroupBy(r => r.Field<string>(_gisIDColumns[1].ColumnName)).Select(g => g.Key);
                        _toidsSelectedMapCount = _toidsSelectedMap.Count();
                        goto case 1;
                    case 1:
                        // Count the number of incids selected in GIS.
                        _incidsSelectedMap = _gisSelection.AsEnumerable()
                            .GroupBy(r => r.Field<string>(_gisIDColumns[0].ColumnName)).Select(g => g.Key);
                        _incidsSelectedMapCount = _incidsSelectedMap.Count();
                        break;
                }

                if (_incidSelectionWhereClause == null)
                    _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                        IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);

                if (_incidsSelectedMapCount > 0) GisToDbSelection();
            }
            else
            {
                _incidSelectionWhereClause = null;
            }
            RefreshStatus();
        }

        private void GisToDbSelection()
        {
            _incidSelection = NewIncidSelectionTable();
            foreach (string s in _incidsSelectedMap)
                _incidSelection.Rows.Add(new object[] { s });
        }

        private DataTable NewIncidSelectionTable()
        {
            DataTable outTable = new DataTable();
            outTable.Columns.Add(new DataColumn(IncidTable.incidColumn.ColumnName, IncidTable.incidColumn.DataType));
            outTable.DefaultView.Sort = IncidTable.incidColumn.ColumnName;
            return outTable;
        }

        private DataTable NewGisSelectionTable()
        {
            DataTable outTable = new DataTable();
            foreach (DataColumn c in _gisIDColumns)
                outTable.Columns.Add(new DataColumn(c.ColumnName, c.DataType));
            return outTable;
        }

        private int ExpectedSelectionFeatures()
        {
            if (HaveGisApp && (_incidSelection != null) && (_incidSelection.Rows.Count > 0) &&
                (_incidSelectionWhereClause != null) && (_incidSelectionWhereClause.Count > 0))
            {
                try
                {
                    HluDataSet.incid_mm_polygonsDataTable t = new HluDataSet.incid_mm_polygonsDataTable();

                    IEnumerable<DataTable> queryTables = _incidSelectionWhereClause.SelectMany(cond => cond.Select(c => c.Table)).Distinct();
                    DataTable[] selTables = new DataTable[] { t }.Union(queryTables).ToArray();
                    DataRelation rel;
                    IEnumerable<SqlFilterCondition> joinCond = queryTables.Select(st =>
                        st.GetType() == typeof(HluDataSet.incidDataTable) ? 
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", st.Columns[_hluDS.incid.incidColumn.Ordinal]) :
                        (rel = GetRelation(_hluDS.incid, st)) != null ?
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", rel.ChildColumns[0]) : null).Where(c => c != null);

                    int numFeatures = 0;
                    for (int i = 0; i < _incidSelectionWhereClause.Count; i++)
                        numFeatures += _db.SqlCount(selTables, joinCond.Concat(_incidSelectionWhereClause[i]).ToList());

                    return numFeatures;
                }
                catch { }
            }
            return -1;
        }

        private void PerformGisSelection(bool setFilter)
        {
            ChangeCursor(Cursors.Wait, "Processing ...");

            if (_gisApp != null)
            {
                if (_incidSelection.Rows.Count > 1)
                {
                    ScratchDb.WriteSelectionScratchTable(_gisIDColumns, _incidSelection);
                    DispatcherHelper.DoEvents();
                    _gisSelection = _gisApp.SqlSelect(ScratchDb.ScratchMdbPath, 
                        ScratchDb.ScratchSelectionTable, _gisIDColumns);
                }
                else
                {
                    DispatcherHelper.DoEvents();
                    _gisSelection = _gisApp.SqlSelect(true, false, _gisIDColumns, 
                        ScratchDb.GisWhereClause(_incidSelection, _gisApp));
                }
            }

            AnalyzeGisSelectionSet();

            if (setFilter) SetFilter();

            if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                MessageBox.Show(App.Current.MainWindow, "No map features selected.", "HLU Selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);

            ChangeCursor(Cursors.Arrow, null);
        }

        private void SetFilter()
        {
            try
            {
                if (IsFiltered && ((_incidsSelectedMapCount > 0) || (_gisApp == null)))
                    // If currently splitting a feature then go to the last incid
                    // in the filter (which will be the new incid).
                    if (_splitting)
                        IncidCurrentRowIndex = IsFiltered ? _incidSelection.Rows.Count : _incidRowCount;
                    else
                        IncidCurrentRowIndex = 1;
            }
            finally
            {
                RefreshStatus();
            }
        }

        #endregion

        #region Switch GIS Layer
        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        public ICommand SwitchGisLayerCommand
        {
            get
            {
                if (_switchGisLayerCommand == null)
                {
                    Action<object> SwitchGISLayerAction = new Action<object>(this.SwitchGISLayerClicked);
                    _switchGisLayerCommand = new RelayCommand(SwitchGISLayerAction);
                }
                return _switchGisLayerCommand;
            }
        }

        private void SwitchGISLayerClicked(object param)
        {
            if (_gisApp.ListHluLayers() > 0)
            {
                _windowSwitchGISLayer = new WindowSwitchGISLayer();
                _windowSwitchGISLayer.Owner = App.Current.MainWindow;
                _windowSwitchGISLayer.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                _viewModelSwitchGISLayer = new ViewModelWindowSwitchGISLayer(_gisApp.ValidHluLayers, _gisApp.CurrentHluLayer);
                _viewModelSwitchGISLayer.RequestClose +=
                    new ViewModelWindowSwitchGISLayer.RequestCloseEventHandler(_viewModelSwitchGISLayer_RequestClose);

                _windowSwitchGISLayer.DataContext = _viewModelSwitchGISLayer;

                _windowSwitchGISLayer.ShowDialog();
            }
        }

        private bool CanSwitchGISLayer
        {
            get { return _gisApp.ValidHluLayers.Count() > 0; }
        }

        void _viewModelSwitchGISLayer_RequestClose(bool switchGISLayer, GISLayer selectedHLULayer)
        {
            _viewModelSwitchGISLayer.RequestClose -= _viewModelSwitchGISLayer_RequestClose;
            _windowSwitchGISLayer.Close();

            // Switch the GIS layer
            if ((switchGISLayer) && (selectedHLULayer != _gisApp.CurrentHluLayer))
            {
                // Check if there are unsaved edits
                MessageBoxResult userResponse = CheckDirty();

                // Check the user's response to saving unsaved edits
                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                        // Save the unsaved edits
                        if (_viewModelUpd.Update())
                            goto case MessageBoxResult.No;
                        break;
                    case MessageBoxResult.No:
                        // Switch the GIS layer
                        if (_gisApp.IsHluLayer(selectedHLULayer))
                        {
                            // Inform the user that the switch worked
                            if (selectedHLULayer.MapName == null)
                                MessageBox.Show(string.Format("GIS Layer switched to {0} [{1}].", selectedHLULayer.LayerName, selectedHLULayer.MapNum),
                                    "Switch GIS Layer",MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            else
                                MessageBox.Show(string.Format("GIS Layer switched to {0} in {1} [{2}]", selectedHLULayer.LayerName, selectedHLULayer.MapName, selectedHLULayer.MapNum),
                                    "Switch GIS Layer", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                            // Refresh the window title
                            OnPropertyChanged("WindowTitle");

                            // Get the new GIS layer selection
                            ReadMapSelection(false);
                        }
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #endregion

        #region Data Tables

        public HluDataSet.incidDataTable IncidTable
        {
            get
            {
                if (HluDataset.incid.IsInitialized && (HluDataset.incid.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incidTableAdapter == null)
                        _hluTableAdapterMgr.incidTableAdapter =
                            new HluTableAdapter<HluDataSet.incidDataTable, HluDataSet.incidRow>(_db);
                    IncidCurrentRowIndex = 1;
                }
                return _hluDS.incid;
            }
        }

        public HluDataSet.incid_ihs_matrixDataTable IncidIhsMatrixTable
        {
            get
            {
                if (HluDataset.incid_ihs_matrix.IsInitialized && (HluDataset.incid_ihs_matrix.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_ihs_matrixTableAdapter == null)
                        _hluTableAdapterMgr.incid_ihs_matrixTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_ihs_matrixDataTable, HluDataSet.incid_ihs_matrixRow>(_db);
                }
                return _hluDS.incid_ihs_matrix;
            }
        }

        public HluDataSet.incid_ihs_formationDataTable IncidIhsFormationTable
        {
            get
            {
                if (HluDataset.incid_ihs_formation.IsInitialized && (HluDataset.incid_ihs_formation.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_ihs_formationTableAdapter == null)
                        _hluTableAdapterMgr.incid_ihs_formationTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_ihs_formationDataTable, HluDataSet.incid_ihs_formationRow>(_db);
                }
                return _hluDS.incid_ihs_formation;
            }
        }

        public HluDataSet.incid_ihs_managementDataTable IncidIhsManagementTable
        {
            get
            {
                if (HluDataset.incid_ihs_management.IsInitialized && (HluDataset.incid_ihs_management.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_ihs_managementTableAdapter == null)
                        _hluTableAdapterMgr.incid_ihs_managementTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_ihs_managementDataTable, HluDataSet.incid_ihs_managementRow>(_db);
                }
                return _hluDS.incid_ihs_management;
            }
        }

        public HluDataSet.incid_ihs_complexDataTable IncidIhsComplexTable
        {
            get
            {
                if (HluDataset.incid_ihs_complex.IsInitialized && (HluDataset.incid_ihs_complex.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_ihs_complexTableAdapter == null)
                        _hluTableAdapterMgr.incid_ihs_complexTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_ihs_complexDataTable, HluDataSet.incid_ihs_complexRow>(_db);
                }
                return _hluDS.incid_ihs_complex;
            }
        }

        public HluDataSet.incid_bapDataTable IncidBapTable
        {
            get
            {
                if (HluDataset.incid_bap.IsInitialized && (HluDataset.incid_bap.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_bapTableAdapter == null)
                        _hluTableAdapterMgr.incid_bapTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_bapDataTable, HluDataSet.incid_bapRow>(_db);
                }
                return _hluDS.incid_bap;
            }
        }

        public HluDataSet.incid_sourcesDataTable IncidSourcesTable
        {
            get
            {
                if (HluDataset.incid_sources.IsInitialized && (HluDataset.incid_sources.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.incid_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_sourcesDataTable, HluDataSet.incid_sourcesRow>(_db);
                }
                return _hluDS.incid_sources;
            }
        }

        #endregion

        #region Data Rows

        public bool IsFiltered
        {
            get
            {
                return _bulkUpdateMode != true && _incidSelection != null && _incidSelection.Rows.Count > 0;
            }
        }

        public int IncidRowCount(bool recount)
        {
            if (recount || (_incidRowCount <= 0))
            {
                try
                {
                    _incidRowCount = (int)_db.ExecuteScalar(String.Format(
                        "SELECT COUNT(*) FROM {0}", _db.QualifyTableName(_hluDS.incid.TableName)),
                        _db.Connection.ConnectionTimeout, CommandType.Text);
                    RefreshStatus();
                }
                catch { return -1; }
            }
            return _incidRowCount;
        }

        public string StatusIncid
        {
            get
            {
                if (IsFiltered)
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR22 (Record selectors)
                    // Include the total toid and fragment counts for the current Incid
                    // in the status area in addition to the currently select toid and
                    // fragment counts.
                    //
                    //return String.Format("of {0} (filtered){1}{2}", _incidSelection.Rows.Count,
                    //    String.Format(" [T:{0}]", numToids.ToString()), String.Format(" [F:{0}]", numFrags.ToString()));
                    return String.Format("of {0} (filtered) {1}{2} of {3}{4}", _incidSelection.Rows.Count,
                        String.Format("[{0}:", _toidsIncidGisCount.ToString()),
                        String.Format("{0}", _fragsIncidGisCount.ToString()),
                        String.Format("{0}:", _toidsIncidDbCount.ToString()),
                        String.Format("{0}]", _fragsIncidDbCount.ToString()));
                    //---------------------------------------------------------------------
                }
                else if ((_bulkUpdateMode == true) && (_incidSelection != null) && (_incidSelection.Rows.Count > 0))
                {
                    return String.Format("[I:{0}]{1}{2}", _incidsSelectedMapCount,
                        String.Format(" [T:{0}]", _toidsSelectedMapCount.ToString()),
                        String.Format(" [F:{0}]", _fragsSelectedMapCount.ToString()));
                }
                else
                {
                    return String.Format("of {0}", _incidRowCount);
                }
            }
        }

        public string StatusIncidToolTip { get { return IsFiltered ? "Double click to clear filter" : null; } }

        public string StatusTop
        {
            get { return _windowCursor == Cursors.Wait ? _processingMsg : String.Empty; }
        }

        public int NumIncidSelected
        {
            get { return _incidSelection.Rows.Count; }
            set { }
        }

        public int NumIncidSelectedMap
        {
            get { return _incidsSelectedMapCount; }
            set { }
        }

        public int NumToidSelectedMap
        {
            get { return _toidsSelectedMapCount; }
            set { }
        }

        public int NumFragmentsSelectedMap
        {
            get { return _fragsSelectedMapCount; }
            set { }
        }

        public HluDataSet.incidRow IncidCurrentRow
        {
            get { return _incidCurrentRow; }
            set { _incidCurrentRow = value; }
        }

        public int IncidCurrentRowIndex
        {
            get { return _incidCurrentRowIndex; }
            set
            {
                MessageBoxResult userResponse = CheckDirty();

                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                        if (_viewModelUpd.Update())
                            goto case MessageBoxResult.No;
                        break;
                    case MessageBoxResult.No:
                        if (_bulkUpdateMode != false || ((value > 0) &&
                            (IsFiltered && ((_incidSelection == null) || (value <= _incidSelection.Rows.Count))) ||
                            (!IsFiltered && ((_incidSelection == null) || (value <= _incidRowCount)))))
                            _incidCurrentRowIndex = value;
                        NewIncidCurrentRow();
                        break;
                    case MessageBoxResult.Cancel:
                        break;
                }
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR2 (Apply button)
        // Check if the record has changed and if it hasn't ask the user
        // if they still want to update the record (to create new history).
        //
        /// <summary>
        /// If no changes have been made reset the changed flag and then
        /// check if the user still wants to save the record.
        /// </summary>
        /// <returns>The user's response to save or not save the record.</returns>
        private MessageBoxResult CheckClean()
        {
            MessageBoxResult userResponse = MessageBoxResult.No;

            if (EditMode && (_bulkUpdateMode == false))
            {
                userResponse = MessageBoxResult.Yes;
                if (!IsDirty)
                {
                    userResponse = _saving ? MessageBox.Show("The current record has not been changed.\n" +
                        "Would you still like to save the record?", "HLU: Save Changes",
                        MessageBoxButton.YesNoCancel, MessageBoxImage.Question) : MessageBoxResult.Yes;
                }
            }

            return userResponse;
        }
        //---------------------------------------------------------------------

        private MessageBoxResult CheckDirty()
        {
            MessageBoxResult userResponse = MessageBoxResult.No;

            if (EditMode && (_bulkUpdateMode == false) && (_splitting == false) && IsDirty)
            {
                if (CanUpdate)
                {
                    userResponse = _saving ? MessageBoxResult.Yes :
                        MessageBox.Show("The current record has been changed.\n" +
                        "Would you like to save your changes?", "HLU: Save Changes",
                        MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                }
                else if (!HaveGisApp)
                {
                    MessageBox.Show("There is no GIS application known to be running.\n" +
                        "Without a GIS application the GIS layer cannot be synchronized.\n" +
                        "Therefore updates are disabled.", "HLU: Updates Disabled",
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    userResponse = MessageBoxResult.Yes;
                }
                else
                {
                    userResponse = MessageBox.Show("The current record has been changed, " +
                        "but it cannot be saved at this time because it is in error." +
                        "\n\nWould you like to leave this record discarding your changes?",
                        "HLU: Cannot Save", MessageBoxButton.YesNo, MessageBoxImage.Exclamation,
                        MessageBoxResult.No);
                    if (userResponse == MessageBoxResult.Yes)
                        userResponse = MessageBoxResult.No;
                    else
                        userResponse = MessageBoxResult.Cancel;
                }

                if (userResponse == MessageBoxResult.No) RestoreIncidCurrentRow();
            }

            return userResponse;
        }

        public bool IsDirty
        {
            get
            {
                if (_saved)
                {
                    _saved = false;
                    return false;
                }

                return IsDirtyIncid() || IsDirtyIncidIhsMatrix() || IsDirtyIncidIhsFormation() || 
                    IsDirtyIncidIhsManagement() || IsDirtyIncidIhsComplex() || IsDirtyIncidBap() || 
                    IsDirtyIncidSources();
            }
        }

        #endregion

        #region Data Row Helpers

        /// <summary>
        /// Initiates all the necessary actions when moving to another incid row.
        /// </summary>
        private void NewIncidCurrentRow()
        {
            bool canMove = false;
            if (!IsFiltered)
            {
                int newRowIndex = SeekIncid(_incidCurrentRowIndex);
                if ((canMove = newRowIndex != -1))
                    _incidCurrentRow = _hluDS.incid[newRowIndex];
            }
            else
            {
                if ((canMove = (_incidCurrentRowIndex != -1) &&
                    (_incidCurrentRowIndex <= _incidSelection.Rows.Count)))
                    _incidCurrentRow = SeekIncidFiltered(_incidCurrentRowIndex);
            }

            if (canMove)
            {
                CloneIncidCurrentRow();

                _incidArea = -1;
                _incidLength = -1;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has not been changed yes so that the
                // apply button does not appear.
                Changed = false;
                //---------------------------------------------------------------------

                // without this IncidIhsHabitat becomes null, called from IhsHabitatCodes, when coming 
                // from a previous row with valid IHS habitat code 
                // (seemingly alternating rows when browsing, i.e. 1 ok, 2 wrong, 3 ok, ...)
                _incidIhsHabitat = null;
                OnPropertyChanged("IncidIhsHabitat");

                IncidCurrentRowDerivedValuesRetrieve();
                GetIncidChildRows(IncidCurrentRow);

                _ihsHabitatCodes = null;
                if (String.IsNullOrEmpty(_incidIhsHabitat))
                {
                    IhsCategory = null;
                    NvcCategory = null;
                }
                else
                {
                    string cat = FindIhsCategory(_incidIhsHabitat);
                    if (!String.IsNullOrEmpty(cat))
                    {
                        IhsCategory = cat;
                        OnPropertyChanged("IhsCategory");
                    }
                    else
                    {
                        NvcCategory = FindNvcCategory(_incidIhsHabitat);
                        OnPropertyChanged("NvcCategory");
                    }
                }

                // Count the number of toids and fragments for the current incid
                // selected in the GIS and in the database.
                CountToidFrags();

                OnPropertyChanged("IncidCurrentRowIndex");
                OnPropertyChanged("IncidCurrentRow");
                IhsMultiplexCodes(_incidIhsHabitat);
                RefreshStatus();
                RefreshHeader();
                RefreshIhsTab();
                RefreshDetailsTab();
                RefreshSource1();
                RefreshSource2();
                RefreshSource3();
                RefreshHistory();
            }
            CheckEditingControlState();
        }

        /// <summary>
        /// Count the number of toids and fragments for the current incid
        /// selected in the GIS and in the database.
        /// </summary>
        private void CountToidFrags()
        {
            //---------------------------------------------------------------------
            // CHANGED: CR10 (Attribute updates for incid subsets)
            // Count the number of toids and fragments for this incid selected
            // in the GIS. They are counted here, once when the incid changes,
            // instead of in StatusIncid() which is constantly being called.
            _toidsIncidGisCount = -1;
            _fragsIncidGisCount = -1;
            if (_gisSelection != null)
            {
                DataRow[] gisRows = _gisSelection.AsEnumerable()
                    .Where(r => r[HluDataset.incid_mm_polygons.incidColumn.ColumnName].Equals(_incidCurrentRow.incid)).ToArray();
                _toidsIncidGisCount = gisRows.GroupBy(r => r[HluDataset.incid_mm_polygons.toidColumn.ColumnName]).Count();
                _fragsIncidGisCount = gisRows.Length;
            }
            //---------------------------------------------------------------------

            //---------------------------------------------------------------------
            // CHANGED: CR22 (Record selectors)
            // Count the total number of toids and fragments in the database
            // for this incid so that they can be included in the status area.
            _fragsIncidDbCount = -1;
            _toidsIncidDbCount = -1;
            if (IsFiltered)
            {
                // Count the total number of fragments in the database for
                // this incid.
                _fragsIncidDbCount = (int)_db.ExecuteScalar(String.Format(
                    "SELECT COUNT(*) FROM {0} WHERE {1} = {2}",
                    _db.QualifyTableName(_hluDS.incid_mm_polygons.TableName),
                    _db.QuoteIdentifier(_hluDS.incid_mm_polygons.incidColumn.ColumnName),
                    _db.QuoteValue(_incidCurrentRow.incid)),
                    _db.Connection.ConnectionTimeout, CommandType.Text);

                // Count the total number of toids in the database for
                // this incid.
                _toidsIncidDbCount = (int)_db.ExecuteScalar(String.Format(
                    "SELECT COUNT(*) FROM (SELECT DISTINCT {0} FROM {1} WHERE {2} = {3}) AS T",
                    _db.QuoteIdentifier(_hluDS.incid_mm_polygons.toidColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.incid_mm_polygons.TableName),
                    _db.QuoteIdentifier(_hluDS.incid_mm_polygons.incidColumn.ColumnName),
                    _db.QuoteValue(_incidCurrentRow.incid)),
                    _db.Connection.ConnectionTimeout, CommandType.Text);
            }
            //---------------------------------------------------------------------

        }

        private void IncidCurrentRowDerivedValuesRetrieve()
        {
            _incidLastModifiedUser = _incidCurrentRow.last_modified_user_id;
            _incidLastModifiedDate = _incidCurrentRow.last_modified_date;
            _incidIhsHabitat = _incidCurrentRow.Isihs_habitatNull() ? null : _incidCurrentRow.ihs_habitat;
        }

        private void CloneIncidCurrentRow()
        {
            _incidCurrentRowClone = _hluDS.incid.NewincidRow(); // IncidTable.NewincidRow();
            for (int i = 0; i < IncidTable.Columns.Count; i++)
                _incidCurrentRowClone[i] = _incidCurrentRow[i];
        }

        private void RestoreIncidCurrentRow()
        {
            if (_incidCurrentRowClone != null)
            {
                for (int i = 0; i < _hluDS.incid.Columns.Count; i++) // IncidTable.Columns.Count; i++)
                    _incidCurrentRow[i] = _incidCurrentRowClone[i];
            }
        }

        private bool CompareIncidCurrentRowClone()
        {
            if (_incidCurrentRowClone != null)
            {
                for (int i = 0; i < _hluDS.incid.Columns.Count; i++) // IncidTable.Columns.Count; i++)
                {
                    if ((_incidCurrentRow.IsNull(i) != _incidCurrentRowClone.IsNull(i)) ||
                        !_incidCurrentRow[i].Equals(_incidCurrentRowClone[i])) return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Finds the DB row corresponding to the incid passed in and loads the next _incidPageSize rows
        /// starting at that row. If the method succeeds, the row corresponding to incid will be row 0 of
        /// the HluDataset.incid DataTable.
        /// </summary>
        /// <param name="incid">The incid whose row is to be made current.</param>
        /// <returns>The row number in HluDataset.incid corresponding to the incid passed in, 
        /// or -1 if the search fails.</returns>
        private int GoToIncid(string incid)
        {
            int incidPageRowNoMinBak = _incidPageRowNoMin;
            int incidPageRowNoMaxBak = _incidPageRowNoMax;
            
            try
            {
                StringBuilder whereClause = new StringBuilder(String.Format("{0} >= {1}", 
                    _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName), _db.QuoteValue(incid)));

                int seekRowNumber = (int)_db.ExecuteScalar(
                    String.Format("SELECT COUNT(*) FROM (SELECT {0} FROM {1} ORDER BY {0} ASC) WHERE {2}",
                    _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.incid.TableName), whereClause),
                    _db.Connection.ConnectionTimeout, CommandType.Text);

                whereClause.Append(String.Format(" AND {0} < {1} ORDER BY {0} ASC",
                    _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName),
                    _db.QuoteValue(_recIDs.IncidString(RecordIds.IncidNumber(incid) + IncidPageSize))));

                _hluTableAdapterMgr.Fill(_hluDS, typeof(HluDataSet.incidDataTable), whereClause.ToString(), true);

                _incidPageRowNoMin = seekRowNumber;
                _incidPageRowNoMax = _incidPageRowNoMin + _hluDS.incid.Count;

                return 0;
            }
            catch
            {
                _incidPageRowNoMin = incidPageRowNoMinBak;
                _incidPageRowNoMax = incidPageRowNoMaxBak;
                return -1;
            }
        }

        /// <summary>
        /// Translates a row number in the incid remote DB table, ordered by incid, into a row number in 
        /// the in-memory incid DataTable, which only contains the current page of the entire DB table. 
        /// If necessary, a new page is loaded from the database. 
        /// </summary>
        /// <param name="seekRowNumber">Row number in the remote DB incid table, ordered by incid, whose 
        /// corresponding row number in in-memory DataTable HluDataset.incid is sought.</param>
        /// <returns>The row number in in-memory DataTable HluDataset.incid that corresponds to
        /// row number seekRowNumber in the remote DB incid table, ordered by incid. 
        /// If loading of a new page fails, -1 is returned and _incidPageRowNoMin and _incidPageRowNoMax 
        /// are reset to their values before the attempted move.</returns>
        private int SeekIncid(int seekRowNumber)
        {
            if ((seekRowNumber >= _incidPageRowNoMin) && (seekRowNumber <= _incidPageRowNoMax))
            {
                return seekRowNumber - _incidPageRowNoMin;
            }
            else
            {
                int incidPageRowNoMinBak = _incidPageRowNoMin;
                int incidPageRowNoMaxBak = _incidPageRowNoMax;

                try
                {
                    int seekIncidNumber = seekRowNumber;

                    if (seekRowNumber < 2)
                    {
                        seekIncidNumber = RecordIds.IncidNumber(_db.ExecuteScalar(
                            String.Format("SELECT {0} FROM {1} ORDER BY {0} ASC",
                            _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName),
                            _db.QualifyTableName(_hluDS.incid.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text).ToString());
                    }
                    else if (seekRowNumber >= _incidRowCount)
                    {
                        seekIncidNumber = RecordIds.IncidNumber(_db.ExecuteScalar(
                            String.Format("SELECT {0} FROM {1} ORDER BY {0} DESC",
                            _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName),
                            _db.QualifyTableName(_hluDS.incid.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text).ToString());
                    }
                    else
                    {
                        _incidRowCount = (int)_db.ExecuteScalar(String.Format(
                            "SELECT COUNT(*) FROM {0}", _db.QualifyTableName(_hluDS.incid.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text);

                        string countSql = String.Format("SELECT COUNT(*) FROM {0} WHERE {1} <= {2}",
                            _db.QualifyTableName(_hluDS.incid.TableName),
                            _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName), "{0}");

                        int count = 0;

                        do
                        {
                            count = (int)_db.ExecuteScalar(String.Format(countSql,
                                _db.QuoteValue(_recIDs.IncidString(seekIncidNumber))),
                                _db.Connection.ConnectionTimeout, CommandType.Text);
                            seekIncidNumber += seekRowNumber - count;
                        } while ((count < seekRowNumber) && (count < _incidRowCount));
                    }

                    string loadWhereClauseTemplate = String.Format("{0} >= {1} AND {0} < {2} ORDER BY {0}",
                        _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName), "{0}", "{1}");

                    _hluTableAdapterMgr.Fill(_hluDS, typeof(HluDataSet.incidDataTable), String.Format(
                        loadWhereClauseTemplate, _db.QuoteValue(_recIDs.IncidString(seekIncidNumber)),
                        _db.QuoteValue(_recIDs.IncidString(seekIncidNumber + IncidPageSize))), true);

                    _incidPageRowNoMin = seekRowNumber;
                    _incidPageRowNoMax = _incidPageRowNoMin + _hluDS.incid.Count;

                    return 0;
                }
                catch
                {
                    _incidPageRowNoMin = incidPageRowNoMinBak;
                    _incidPageRowNoMax = incidPageRowNoMaxBak;
                    return -1;
                }
            }
        }

        /// <summary>
        /// Retrieves the row of the in-memory incid DataTable that corresponds to the incid of the row
        /// of the _incidSelection DataTable whose row number is passed in as parameter seekRowNumber.
        /// If necessary, a new page of selected incid rows is loaded from the database. 
        /// </summary>
        /// <param name="seekRowNumber">Row number in the _incidSelection DataTable whose 
        /// corresponding row in in-memory DataTable HluDataset.incid is sought.</param>
        /// <returns>The row of in-memory DataTable HluDataset.incid that corresponds to
        /// row number seekRowNumber in the _incidSelection DataTable. 
        /// If loading of a new page fails, null is returned.</returns>
        private HluDataSet.incidRow SeekIncidFiltered(int seekRowNumber)
        {
            seekRowNumber--;

            if (seekRowNumber < 0)
                seekRowNumber = 0;
            else if (seekRowNumber > _incidSelection.Rows.Count - 1)
                seekRowNumber = _incidSelection.Rows.Count - 1;

            string seekIncid = (string)_incidSelection.DefaultView[seekRowNumber][0];
            HluDataSet.incidRow returnRow = _hluDS.incid.FindByincid(seekIncid);

            if (returnRow != null)
            {
                return returnRow;
            }
            else
            {
                int seekIncidNumber = RecordIds.IncidNumber(seekIncid);
                int incidNumberPageMin;
                int incidNumberPageMax;
                if (_hluDS.incid.Rows.Count == 0)
                {
                    incidNumberPageMin = seekIncidNumber;
                    incidNumberPageMax = incidNumberPageMin + IncidPageSize;
                }
                else
                {
                    incidNumberPageMin = RecordIds.IncidNumber(_hluDS.incid[0].incid);
                    incidNumberPageMax = RecordIds.IncidNumber(_hluDS.incid[_hluDS.incid.Count - 1].incid);
                }

                int start = _incidCurrentRowIndex > 0 ? _incidCurrentRowIndex - 1 : 0;
                int stop = start;
                bool moveForward = true;

                if (seekIncidNumber < incidNumberPageMin) // moving backward
                {
                    start = seekRowNumber - IncidPageSize > 0 ? seekRowNumber - IncidPageSize : 0;
                    stop = seekRowNumber > start ? seekRowNumber : IncidPageSize < _incidSelection.Rows.Count ? 
                    IncidPageSize : _incidSelection.Rows.Count - 1;
                    moveForward = false;
                }
                else if (seekIncidNumber > incidNumberPageMax) // moving forward
                {
                    start = seekRowNumber;
                    stop = seekRowNumber + IncidPageSize < _incidSelection.Rows.Count ?
                        seekRowNumber + IncidPageSize : _incidSelection.Rows.Count - 1;
                }

                try
                {
                    string[] incids = new string[start == stop ? 1 : stop - start + 1];

                    for (int i = 0; i < incids.Length; i++)
                        incids[i] = _db.QuoteValue(_incidSelection.DefaultView[start + i][0]);

                    _hluTableAdapterMgr.incidTableAdapter.Fill(_hluDS.incid, String.Format("{0} IN ({1}) ORDER BY {0}",
                        _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName), String.Join(",", incids)));

                    if (_hluDS.incid.Count == 0)
                    {
                        MessageBox.Show("No database record retrieved.", "HLU",
                            MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        // Reset the incid and map selections and move
                        // to the first incid in the database.
                        ClearFilter(true);
                        return _incidCurrentRow;
                    }
                    else
                    {
                        return moveForward ? _hluDS.incid[0] : _hluDS.incid[_hluDS.incid.Count - 1];
                    }
                }
                catch { return null; }
            }
        }

        private Dictionary<Type, string> BuildChildRowOrderByClauses()
        {
            Dictionary<Type, string> childRowOrberByDict = new Dictionary<Type, string>();

            childRowOrberByDict.Add(typeof(HluDataSet.incid_ihs_matrixDataTable), _hluDS.incid_ihs_matrix.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_ihs_formationDataTable), _hluDS.incid_ihs_formation.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_ihs_managementDataTable), _hluDS.incid_ihs_management.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_ihs_complexDataTable), _hluDS.incid_ihs_complex.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_bapDataTable), _hluDS.incid_bap.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_sourcesDataTable), _hluDS.incid_sources.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            return childRowOrberByDict;
        }

        private Dictionary<Type, List<SqlFilterCondition>> BuildChildRowFilters()
        {
            Dictionary<Type, List<SqlFilterCondition>> childRowFilterDict =
                new Dictionary<Type, List<SqlFilterCondition>>();

            childRowFilterDict.Add(typeof(HluDataSet.incid_ihs_matrixDataTable), 
                ChildRowFilter(_hluDS.incid, _hluDS.incid_ihs_matrix));

            childRowFilterDict.Add(typeof(HluDataSet.incid_ihs_formationDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_ihs_formation));

            childRowFilterDict.Add(typeof(HluDataSet.incid_ihs_managementDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_ihs_management));

            childRowFilterDict.Add(typeof(HluDataSet.incid_ihs_complexDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_ihs_complex));

            childRowFilterDict.Add(typeof(HluDataSet.incid_bapDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_bap));

            childRowFilterDict.Add(typeof(HluDataSet.incid_sourcesDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_sources));

            childRowFilterDict.Add(typeof(HluDataSet.historyDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.history));

            _incidMMPolygonsIncidFilter = new SqlFilterCondition();
            _incidMMPolygonsIncidFilter.BooleanOperator = "OR";
            _incidMMPolygonsIncidFilter.OpenParentheses = "(";
            _incidMMPolygonsIncidFilter.Column = _hluDS.incid_mm_polygons.incidColumn;
            _incidMMPolygonsIncidFilter.Table = _hluDS.incid_mm_polygons;
            _incidMMPolygonsIncidFilter.Value = String.Empty;
            _incidMMPolygonsIncidFilter.CloseParentheses = ")";

            return childRowFilterDict;
        }

        internal SqlFilterCondition ChildRowFilter<T>(T table, DataColumn incidColumn)
            where T : DataTable
        {
            SqlFilterCondition cond = new SqlFilterCondition();
            cond.BooleanOperator = "OR";
            cond.OpenParentheses = "(";
            cond.Column = incidColumn;
            cond.Table = table;
            cond.Value = String.Empty;
            cond.CloseParentheses = ")";
            return cond;
        }

        internal List<SqlFilterCondition> ChildRowFilter<P, C>(P parentTable, C childTable)
            where P : DataTable
            where C : DataTable
        {
            DataRelation rel = GetRelation<P, C>(parentTable, childTable);
            List<SqlFilterCondition> condList = new List<SqlFilterCondition>();

            for (int i = 0; i < rel.ChildColumns.Length; i++)
            {
                DataColumn c = rel.ChildColumns[i];
                SqlFilterCondition cond = new SqlFilterCondition();
                if (i == 0)
                {
                    cond.BooleanOperator = "OR";
                    cond.OpenParentheses = "(";
                    cond.CloseParentheses = String.Empty;
                }
                else
                {
                    cond.BooleanOperator = "AND";
                    cond.OpenParentheses = String.Empty;
                }
                cond.Column = c;
                cond.Table = childTable;
                cond.ColumnSystemType = c.DataType;
                cond.Operator = "=";
                cond.Value = String.Empty;
                if (i == rel.ChildColumns.Length - 1)
                    cond.CloseParentheses = ")";
                else
                    cond.CloseParentheses = String.Empty;
                condList.Add(cond);
            }

            return condList;
        }

        private void GetIncidChildRows(HluDataSet.incidRow incidRow)
        {
            if (incidRow == null) return;

            string[] relValues = new string[] { incidRow.incid };

            HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable = _hluDS.incid_ihs_matrix;
            _incidIhsMatrixRows = GetIncidChildRowsDb(relValues,
               _hluTableAdapterMgr.incid_ihs_matrixTableAdapter, ref ihsMatrixTable);
            _origIncidIhsMatrixCount = _incidIhsMatrixRows.Length;

            HluDataSet.incid_ihs_formationDataTable ihsFormationTable = _hluDS.incid_ihs_formation;
            _incidIhsFormationRows = GetIncidChildRowsDb(relValues, 
                _hluTableAdapterMgr.incid_ihs_formationTableAdapter, ref ihsFormationTable);
            _origIncidIhsFormationCount = _incidIhsFormationRows.Length;

            HluDataSet.incid_ihs_managementDataTable ihsManagementTable = _hluDS.incid_ihs_management;
            _incidIhsManagementRows = GetIncidChildRowsDb(relValues, 
                _hluTableAdapterMgr.incid_ihs_managementTableAdapter, ref ihsManagementTable);
            _origIncidIhsManagementCount = _incidIhsManagementRows.Length;

            HluDataSet.incid_ihs_complexDataTable ihsComplexTable = _hluDS.incid_ihs_complex;
            _incidIhsComplexRows = GetIncidChildRowsDb(relValues, 
                _hluTableAdapterMgr.incid_ihs_complexTableAdapter, ref ihsComplexTable);
            _origIncidIhsComplexCount = _incidIhsComplexRows.Length;

            HluDataSet.incid_bapDataTable incidBapTable = _hluDS.incid_bap;
            _incidBapRows = GetIncidChildRowsDb(relValues, 
                _hluTableAdapterMgr.incid_bapTableAdapter, ref incidBapTable);
            GetBapEnvironments();

            HluDataSet.incid_sourcesDataTable incidSourcesTable = _hluDS.incid_sources;
            _incidSourcesRows = GetIncidChildRowsDb(relValues, 
                _hluTableAdapterMgr.incid_sourcesTableAdapter, ref incidSourcesTable);

            HluDataSet.historyDataTable historyTable = _hluDS.history;
            _incidHistoryRows = GetIncidChildRowsDb(relValues,
                _hluTableAdapterMgr.historyTableAdapter, ref historyTable);
        }

        private DataRelation GetRelation<P, C>(P parentTable, C childTable)
            where P : DataTable
            where C : DataTable
        {
            try
            {
                return _hluDataRelations.Single(r => r.ParentTable == parentTable && r.ChildTable == childTable);
            }
            catch { return null; }
        }

        private R[] GetIncidChildRows<P, C, R>(P parentTable, C childTable)
            where P : DataTable
            where C : DataTable
            where R : DataRow
        {
            if (IncidCurrentRow == null) return null;

            DataRelation rel = GetRelation(parentTable, childTable);
            if (rel != null)
                return (R[])IncidCurrentRow.GetChildRows(rel, DataRowVersion.Default);
            else
                return new R[0];
        }

        internal R[] GetIncidChildRowsDb<C, R>(object[] relValues, HluTableAdapter<C, R> adapter, ref C childTable)
            where C : DataTable, new()
            where R : DataRow
        {
            List<SqlFilterCondition> childConds;

            if (_childRowFilterDict.TryGetValue(typeof(C), out childConds))
            {
                for (int i = 0; i < childConds.Count; i++)
                {
                    SqlFilterCondition cond = childConds[i];
                    cond.Value = relValues[i];
                    childConds[i] = cond;
                }

                string orderByClause;
                if (_childRowOrderByDict.TryGetValue(typeof(C), out orderByClause))
                    adapter.Fill(childTable, String.Format("{0} ORDER BY {1}",
                        _db.WhereClause(false, true, true, childConds), orderByClause));
                else
                    adapter.Fill(childTable, childConds);
                
                return (R[])childTable.Select();
            }
            else
            {
                return new R[0];
            }
        }

        internal void GetIncidMMPolygonRows(List<List<SqlFilterCondition>> whereClause,
            ref HluDataSet.incid_mm_polygonsDataTable table)
        {
            if ((whereClause != null) && (whereClause.Count > 0))
            {
                if (_hluTableAdapterMgr.incid_mm_polygonsTableAdapter == null)
                    _hluTableAdapterMgr.incid_mm_polygonsTableAdapter =
                        new HluTableAdapter<HluDataSet.incid_mm_polygonsDataTable, 
                            HluDataSet.incid_mm_polygonsRow>(_db);

                _hluTableAdapterMgr.incid_mm_polygonsTableAdapter.Fill(table, whereClause); 
            }
        }

        internal bool IsDirtyIncid()
        {
            return ((_incidCurrentRow != null) && (_incidCurrentRow.RowState != DataRowState.Detached) && ((_incidCurrentRow.Isihs_habitatNull() && 
                !String.IsNullOrEmpty(_incidIhsHabitat)) || _incidIhsHabitat != _incidCurrentRow.ihs_habitat || 
                !CompareIncidCurrentRowClone()));
        }

        internal bool IsDirtyIncidIhsMatrix()
        {
            if (_incidIhsMatrixRows != null)
            {
                if (_incidIhsMatrixRows.Count(r => r != null) != _origIncidIhsMatrixCount) return true;

                foreach (DataRow r in _incidIhsMatrixRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidIhsMatrixCount != 0;
        }

        internal bool IsDirtyIncidIhsFormation()
        {
            if (_incidIhsFormationRows != null)
            {
                if (_incidIhsFormationRows.Count(r => r != null) != _origIncidIhsFormationCount) return true;

                foreach (DataRow r in _incidIhsFormationRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidIhsFormationCount != 0;
        }

        internal bool IsDirtyIncidIhsManagement()
        {
            if (_incidIhsManagementRows != null)
            {
                if (_incidIhsManagementRows.Count(r => r != null) != _origIncidIhsManagementCount) return true;

                foreach (DataRow r in _incidIhsManagementRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidIhsManagementCount != 0;
        }

        internal bool IsDirtyIncidIhsComplex()
        {
            if (_incidIhsComplexRows != null)
            {
                if (_incidIhsComplexRows.Count(r => r != null) != _origIncidIhsComplexCount) return true;

                foreach (DataRow r in _incidIhsComplexRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidIhsComplexCount != 0;
        }

        internal bool IsDirtyIncidBap()
        {
            if (_incidBapRows.Count(r => r.RowState == DataRowState.Deleted) > 0) return true;
            int incidBapRowsAutoNum = 0;
            if (_incidBapRowsAuto != null)
            {
                incidBapRowsAutoNum = _incidBapRowsAuto.Count;
                if (_incidBapRowsAuto.Count(be => IncidBapRowDirty(be)) > 0) return true;
            }
            int incidBapRowsAutoUserNum = 0;
            if (_incidBapRowsUser != null)
            {
                incidBapRowsAutoUserNum = _incidBapRowsUser.Count;
                if (_incidBapRowsUser.Count(be => IncidBapRowDirty(be)) > 0) return true;
            }

            if ((_incidBapRows != null) && (incidBapRowsAutoNum + incidBapRowsAutoUserNum !=
                _incidBapRows.Count())) return true;

            if (_incidBapRows != null)
            {
                foreach (DataRow r in _incidBapRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;
            }
            return false;
        }

        private bool IsDirtyIncidSources()
        {
            if (_incidSourcesRows != null)
            {
                foreach (DataRow r in _incidSourcesRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;
            }
            return false;
        }

        private bool IncidBapRowDirty(BapEnvironment be)
        {
            // deleted user BAP row
            var q = _incidBapRows.Where(r => r.RowState != DataRowState.Deleted && r.bap_id == be.bap_id);
            switch (q.Count())
            {
                case 0:
                    return true; // new row;
                case 1:
                    //---------------------------------------------------------------------
                    // FIX: Don't flag existing invalid incid_bap rows as dirty
                    // Only flag an incid_bap row that is invalid as dirty if it has
                    // been added by the user. This allows existing records to be
                    // viewed in the user interface without warning the user that
                    // the data has changed.
                    if (!be.IsValid() && be.IsAdded) return true;
                    //---------------------------------------------------------------------
                    HluDataSet.incid_bapRow oldRow = q.ElementAt(0);
                    object[] itemArray = be.ToItemArray();
                    for (int i = 0; i < itemArray.Length; i++)
                    {
                        if (oldRow.IsNull(i))
                        {
                            if (itemArray[i] != null) return true;
                        }
                        else if (!oldRow[i].Equals(itemArray[i]))
                        {
                            return true;
                        }
                    }
                    return false;
                default:
                    return true; // duplicate row must be new or altered
            }
        }

        #endregion

        #region Refresh

        internal void RefreshAll()
        {
            OnPropertyChanged("CanCopy");
            OnPropertyChanged("CanPaste");
            RefreshBulkUpdateControls();
            RefreshStatus();
            RefreshHeader();
            RefreshIhsTab();
            RefreshDetailsTab();
            RefreshSources();
            RefreshHistory();
            CheckEditingControlState();
        }

        private void RefreshBulkUpdateControls()
        {
            OnPropertyChanged("ShowInBulkUpdateMode");
            OnPropertyChanged("HideInBulkUpdateMode");
            OnPropertyChanged("BulkUpdateCommandHeader");
            OnPropertyChanged("TopControlsGroupHeader");
            OnPropertyChanged("TabItemHistoryEnabled");
            OnPropertyChanged("NumIncidSelected");
            OnPropertyChanged("NumIncidSelectedMap");
            OnPropertyChanged("NumToidSelectedMap");
            OnPropertyChanged("NumFragmentsSelectedMap");
            OnPropertyChanged("BapHabitatsAutoEnabled");
            OnPropertyChanged("BapHabitatsUserEnabled");
        }

        private void RefreshStatus()
        {
            OnPropertyChanged("EditMode");
            OnPropertyChanged("IncidCurrentRowIndex");
            OnPropertyChanged("StatusIncid");
            OnPropertyChanged("StatusIncidToolTip");
            OnPropertyChanged("StatusTop");
            OnPropertyChanged("CanZoomSelection");
            OnPropertyChanged("CanBulkUpdate");
            OnPropertyChanged("IsFiltered");
        }

        private void RefreshHeader()
        {
            OnPropertyChanged("ReasonCodes");
            OnPropertyChanged("Reason");
            OnPropertyChanged("ProcessCodes");
            OnPropertyChanged("Process");
            OnPropertyChanged("Incid");
            OnPropertyChanged("IncidArea");
            OnPropertyChanged("IncidLength");
            OnPropertyChanged("IncidCreatedDate");
            OnPropertyChanged("IncidLastModifiedDate");
            OnPropertyChanged("IncidCreatedUser");
            OnPropertyChanged("IncidLastModifiedUser");
        }

        private void RefreshIhsTab()
        {
            OnPropertyChanged("IhsCategoryCodes");
            OnPropertyChanged("NvcCategoryCodes");
            OnPropertyChanged("IhsCategory");
            OnPropertyChanged("IncidIhsHabitat");
            OnPropertyChanged("NvcCodes");
            RefreshIhsMulitplexValues();
            RefreshIhsMulitplexCodes();
        }

        private void RefreshIhsMulitplexValues()
        {
            OnPropertyChanged("IncidIhsMatrix1");
            OnPropertyChanged("IncidIhsMatrix2");
            OnPropertyChanged("IncidIhsMatrix3");
            OnPropertyChanged("IncidIhsFormation1");
            OnPropertyChanged("IncidIhsFormation2");
            OnPropertyChanged("IncidIhsManagement1");
            OnPropertyChanged("IncidIhsManagement2");
            OnPropertyChanged("IncidIhsComplex1");
            OnPropertyChanged("IncidIhsComplex2");
            OnPropertyChanged("IncidIhsSummary");
            OnPropertyChanged("IncidLegacyHabitat");
        }

        private void RefreshIhsMulitplexCodes()
        {
            OnPropertyChanged("IhsMatrix1Codes");
            OnPropertyChanged("IhsMatrix2Codes");
            OnPropertyChanged("IhsMatrix3Codes");
            OnPropertyChanged("IhsMatrix1Enabled");
            OnPropertyChanged("IhsMatrix2Enabled");
            OnPropertyChanged("IhsMatrix3Enabled");
            OnPropertyChanged("IhsFormation1Codes");
            OnPropertyChanged("IhsFormation2Codes");
            OnPropertyChanged("IhsFormation1Enabled");
            OnPropertyChanged("IhsFormation2Enabled");
            OnPropertyChanged("IhsManagement1Codes");
            OnPropertyChanged("IhsManagement2Codes");
            OnPropertyChanged("IhsManagement1Enabled");
            OnPropertyChanged("IhsManagement2Enabled");
            OnPropertyChanged("IhsComplex1Codes");
            OnPropertyChanged("IhsComplex2Codes");
            OnPropertyChanged("IhsComplex1Enabled");
            OnPropertyChanged("IhsComplex2Enabled");
            RefreshIhsMulitplexValues();
        }

        private void RefreshDetailsTab()
        {
            OnPropertyChanged("IncidBapHabitatsAuto");
            OnPropertyChanged("IncidBapHabitatsUser");
            OnPropertyChanged("BapHabitatsUserEnabled");
            OnPropertyChanged("IncidGeneralComments");
            OnPropertyChanged("IncidBoundaryBaseMap");
            OnPropertyChanged("IncidDigitisationBaseMap");
            OnPropertyChanged("BapHabitatsAutoEnabled");
            OnPropertyChanged("BapHabitatsUserEnabled");
            OnPropertyChanged("IncidSiteName");
        }

        private void RefreshSources()
        {
            RefreshSource1();
            RefreshSource2();
            RefreshSource3();
        }

        private void RefreshSource1()
        {
            OnPropertyChanged("IncidSource1Id");
            OnPropertyChanged("Source1Names");
            OnPropertyChanged("IncidSource1Date");
            OnPropertyChanged("IncidSource1HabitatClass");
            OnPropertyChanged("IncidSource1HabitatType");
            OnPropertyChanged("Source1HabitatTypeCodes");
            //---------------------------------------------------------------------
            // FIXED: KI103 (Record selectors)
            // FIXED: KI109 (Source habitat types)
            // Both issues seem to relate to the source habitat type value
            // not being displayed after the list of source habitat types
            // has been refreshed following a change of habitat class.
            // So the 'OnPropertyChanged' method is called again to
            // trigger a display refresh.
            OnPropertyChanged("IncidSource1HabitatType");
            //---------------------------------------------------------------------
            OnPropertyChanged("IncidSource1BoundaryImportance");
            OnPropertyChanged("IncidSource1HabitatImportance");
            OnPropertyChanged("IncidSource1Enabled");
        }

        private void RefreshSource2()
        {
            OnPropertyChanged("IncidSource2Id");
            OnPropertyChanged("Source2Names");
            OnPropertyChanged("IncidSource2Date");
            OnPropertyChanged("IncidSource2HabitatClass");
            OnPropertyChanged("IncidSource2HabitatType");
            OnPropertyChanged("Source2HabitatTypeCodes");
            //---------------------------------------------------------------------
            // FIXED: KI103 (Record selectors)
            // FIXED: KI109 (Source habitat types)
            // Both issues seem to relate to the source habitat type value
            // not being displayed after the list of source habitat types
            // has been refreshed following a change of habitat class.
            // So the 'OnPropertyChanged' method is called again to
            // trigger a display refresh.
            OnPropertyChanged("IncidSource2HabitatType");
            //---------------------------------------------------------------------
            OnPropertyChanged("IncidSource2BoundaryImportance");
            OnPropertyChanged("IncidSource2HabitatImportance");
            OnPropertyChanged("IncidSource2Enabled");
        }

        private void RefreshSource3()
        {
            OnPropertyChanged("IncidSource3Id");
            OnPropertyChanged("Source3Names");
            OnPropertyChanged("IncidSource3Date");
            OnPropertyChanged("IncidSource3HabitatClass");
            OnPropertyChanged("IncidSource3HabitatType");
            OnPropertyChanged("Source3HabitatTypeCodes");
            //---------------------------------------------------------------------
            // FIXED: KI103 (Record selectors)
            // FIXED: KI109 (Source habitat types)
            // Both issues seem to relate to the source habitat type value
            // not being displayed after the list of source habitat types
            // has been refreshed following a change of habitat class.
            // So the 'OnPropertyChanged' method is called again to
            // trigger a display refresh.
            OnPropertyChanged("IncidSource3HabitatType");
            //---------------------------------------------------------------------
            OnPropertyChanged("IncidSource3BoundaryImportance");
            OnPropertyChanged("IncidSource3HabitatImportance");
            OnPropertyChanged("IncidSource3Enabled");
        }

        private void RefreshHistory()
        {
            OnPropertyChanged("IncidHistory");
        }

        #endregion

        #region Lock Editing Controls

        private void CheckEditingControlState()
        {
            OnPropertyChanged("ReasonProcessEnabled");
            OnPropertyChanged("TabControlDataEnabled");
        }

        public void ChangeEditingControlState(bool enable)
        {
            _reasonProcessEnabled = enable;
            _tabControlDataEnabled = enable;
            CheckEditingControlState();
        }

        public bool ReasonProcessEnabled
        {
            get
            {
                if ((_bulkUpdateMode == false) && IncidCurrentRow == null) _reasonProcessEnabled = false;
                return _reasonProcessEnabled;
            }
            set { _reasonProcessEnabled = value; }
        }

        public bool TabControlDataEnabled
        {
            get
            {
                if ((_bulkUpdateMode == false) && IncidCurrentRow == null) _tabControlDataEnabled = false;
                return _windowEnabled && _tabControlDataEnabled;
            }
            set { _tabControlDataEnabled = value; }
        }

        public bool TabItemIhsEnabled
        {
            get { return _tabItemIhsEnabled; }
            set { _tabItemIhsEnabled = value; }
        }

        public bool TabItemDetailsEnabled
        {
            get { return _tabItemDetailsEnabled; }
            set { _tabItemDetailsEnabled = value; }
        }

        public bool TabItemSourcesEnabled
        {
            get { return _tabItemSourcesEnabled; }
            set { _tabItemSourcesEnabled = value; }
        }

        public bool TabItemHistoryEnabled
        {
            get { return _tabItemHistoryEnabled; }
            set { _tabItemHistoryEnabled = value; }
        }

        #endregion

        #region Header Fields

        public string Incid
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.incidColumn))
                    return IncidCurrentRow.incid;
                else
                    return null;
            }
            set { if ((IncidCurrentRow != null) && (value != null)) IncidCurrentRow.incid = value; }
        }

        public string IncidArea
        {
            get
            {
                if (_bulkUpdateMode == false)
                {
                    GetIncidMeasures();
                    return _incidArea.ToString();
                }
                else
                {
                    return null;
                }
            }
            set { }
        }

        public string IncidLength
        {
            get
            {
                if (_bulkUpdateMode == false)
                {
                    GetIncidMeasures();
                    return _incidLength.ToString();
                }
                else
                {
                    return null;
                }
            }
        }

        private void GetIncidMeasures()
        {
            if (((_incidArea != -1) && (_incidLength != -1)) || (IncidCurrentRow == null)) return;

            _incidMMPolygonsIncidFilter.Value = Incid;
            HluDataSet.incid_mm_polygonsDataTable table = HluDataset.incid_mm_polygons;

            List<SqlFilterCondition> incidCond = 
                new List<SqlFilterCondition>(new SqlFilterCondition[] { _incidMMPolygonsIncidFilter });
            List<List<SqlFilterCondition>> incidCondList = new List<List<SqlFilterCondition>>();
            incidCondList.Add(incidCond);
            GetIncidMMPolygonRows(incidCondList, ref table);

            _incidArea = 0;
            _incidLength = 0;
            foreach (HluDataSet.incid_mm_polygonsRow r in table)
            {
                _incidArea += r.shape_area;
                _incidLength += r.shape_length;
            }

            _incidArea = Math.Round(_incidArea / 10000, 4);
            _incidLength = Math.Round(_incidLength / 1000, 3);
        }

        public string IncidCreatedDate
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.created_dateColumn))
                    return IncidCurrentRow.created_date.ToShortDateString();
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    DateTime newDate;
                    if (DateTime.TryParse(value, out newDate))
                        IncidCurrentRow.created_date = newDate;
                }
            }
        }

        public string IncidLastModifiedDate
        {
            get { return _incidLastModifiedDate.ToShortDateString(); }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    DateTime newDate;
                    if (DateTime.TryParse(value, out newDate))
                        _incidLastModifiedDate = newDate;
                }
            }
        }

        public string IncidCreatedUser
        {
            get
            {
                //---------------------------------------------------------------------
                // FIX: Display user's name instead of their user_id
                // Display the created user's name from the lut_user table
                // (if found) instead of the user_id
                //
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.created_user_idColumn))
                    return String.IsNullOrEmpty(IncidCurrentRow.lut_userRowByfk_incid_user_created.user_name)
                        ? IncidCurrentRow.created_user_id : IncidCurrentRow.lut_userRowByfk_incid_user_created.user_name;
                //---------------------------------------------------------------------
                else
                    return null;
            }
            set { if ((IncidCurrentRow != null) && (value != null)) IncidCurrentRow.created_user_id = value; }
        }

        public string IncidLastModifiedUser
        {
            get
            {
                //---------------------------------------------------------------------
                // FIX: Display user's name instead of their user_id
                // Display the last modified user's name from the lut_user table
                // (if found) instead of the user_id
                //
                return String.IsNullOrEmpty(IncidCurrentRow.lut_userRowByfk_incid_user_modified.user_name)
                        ? IncidCurrentRow.last_modified_user_id : IncidCurrentRow.lut_userRowByfk_incid_user_modified.user_name;
                //---------------------------------------------------------------------
            }
            set { if ((IncidCurrentRow != null) && (value != null)) _incidLastModifiedUser = value; }
        }

        #region Reason and Process

        public HluDataSet.lut_reasonRow[] ReasonCodes
        {
            get
            {
                if (_reasonCodes == null)
                {
                    if (HluDataset.lut_reason.IsInitialized && (HluDataset.lut_reason.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_reasonTableAdapter == null)
                            _hluTableAdapterMgr.lut_reasonTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_reasonDataTable, HluDataSet.lut_reasonRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_reasonDataTable) }, false);
                    }
                    _reasonCodes = HluDataset.lut_reason.OrderBy(r => r.sort_order).ToArray();
                }
                return _reasonCodes;
            }
            set { }
        }

        public string Reason
        {
            get { return _reason; }
            set { _reason = value; }
        }

        public HluDataSet.lut_processRow[] ProcessCodes
        {
            get
            {
                if (_processCodes == null)
                {
                    if (HluDataset.lut_process.IsInitialized && (HluDataset.lut_process.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_processTableAdapter == null)
                            _hluTableAdapterMgr.lut_processTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_processDataTable, HluDataSet.lut_processRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_processDataTable) }, false);
                    }
                    _processCodes = HluDataset.lut_process.OrderBy(r => r.sort_order).ToArray();
                }
                return _processCodes;
            }
            set { }
        }

        public string Process
        {
            get { return _process; }
            set { _process = value; }
        }

        #endregion

        #endregion

        #region IHS Tab

        //---------------------------------------------------------------------
        // FIX: Show field errors on tab labels.
        // Set the Ihs tab label from here so that validation can be done.
        // This will enable tooltips to be shown so that validation errors
        // in any fields in the tab can be highlighted by flagging the tab
        // label as in error.
        public string IhsTabLabel
        {
            get { return "IHS"; }
        }
        //---------------------------------------------------------------------

        #region IHS Category

        public HluDataSet.lut_ihs_categoryRow[] IhsCategoryCodes
        {
            get
            {
                if (_ihsCategoryCodes == null)
                {
                    if (HluDataset.lut_ihs_category.IsInitialized && (HluDataset.lut_ihs_category.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_categoryTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_categoryTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_categoryDataTable, HluDataSet.lut_ihs_categoryRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_categoryDataTable) }, false);
                    }
                    _ihsCategoryCodes = HluDataset.lut_ihs_category.OrderBy(r => r.sort_order).ToArray();
                }
                return _ihsCategoryCodes;
            }
            set { }
        }

        public string IhsCategory
        {
            get { return _ihsCategory; }
            set
            {
                _ihsCategory = value;

                if (!String.IsNullOrEmpty(_ihsCategory))
                {
                    _nvcCategory = null;
                    switch (_gisLayerType)
                    {
                        case GeometryTypes.Point:
                            _ihsHabitatCodes = from h in HluDataset.lut_ihs_habitat
                                                    where h.is_local && h.point && h.category == IhsCategory
                                                    select h;
                            break;
                        case GeometryTypes.Line:
                            _ihsHabitatCodes = from h in HluDataset.lut_ihs_habitat
                                                    where h.is_local && h.line && h.category == IhsCategory
                                                    select h;
                            break;
                        case GeometryTypes.Polygon:
                            _ihsHabitatCodes = from h in HluDataset.lut_ihs_habitat
                                                    where h.is_local && h.polygon && h.category == IhsCategory
                                                    select h;
                            break;
                    }
                }
                else
                {
                    _ihsHabitatCodes = null;
                }

                OnPropertyChanged("IhsHabitatCodes");
                if ((_ihsHabitatCodes != null) && (_ihsHabitatCodes.Count() == 1))
                    OnPropertyChanged("IncidIhsHabitat");
                OnPropertyChanged("IhsHabitatListEnabled");
                OnPropertyChanged("NvcCategory");
                OnPropertyChanged("NvcCodes");
            }
        }

        public string[] NvcCategoryCodes
        {
            get
            {
                if (_nvcCategoryCodes == null)
                {
                    if (HluDataset.lut_ihs_habitat_ihs_nvc.IsInitialized &&
                        HluDataset.lut_ihs_habitat_ihs_nvc.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_ihs_habitat_ihs_nvcTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_habitat_ihs_nvcTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_nvcDataTable,
                                    HluDataSet.lut_ihs_habitat_ihs_nvcRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_ihs_habitat_ihs_nvcDataTable) }, false);
                    }
                    _nvcCategoryCodes =
                        HluDataset.lut_ihs_habitat_ihs_nvc.Select(r => r.code_nvc).Distinct().OrderBy(s => s).ToArray();
                }
                return _nvcCategoryCodes;
            }
            set { }
        }

        public string NvcCategory
        {
            get { return _nvcCategory; }
            set
            {
                _nvcCategory = value;
                if (!String.IsNullOrEmpty(_nvcCategory))
                {
                    _ihsCategory = null;
                    _ihsHabitatCodes = from h in HluDataset.lut_ihs_habitat
                                            join n in HluDataset.lut_ihs_habitat_ihs_nvc on h.code equals n.code_habitat
                                            where n.code_nvc == NvcCategory
                                            select h;
                }
                else
                {
                    _ihsHabitatCodes = null;
                }
                OnPropertyChanged("IhsHabitatCodes");
                OnPropertyChanged("IhsHabitatListEnabled");
                OnPropertyChanged("IhsCategory");
                OnPropertyChanged("NvcCodes");
            }
        }

        public string NvcCodes
        {
            get
            {
                if (String.IsNullOrEmpty(_incidIhsHabitat)) return null;

                //---------------------------------------------------------------------
                // FIXED: KI102 (NVC Codes)
                // Select NVC codes based on current IHS habitat
                var q = _ihsHabitatCodes.Where(h => h.code == _incidIhsHabitat);
                //var q = _ihsHabitatCodes.Where(h => !h.IsNull(HluDataset.lut_ihs_habitat.nvc_codesColumn));
                //---------------------------------------------------------------------
                if (q.Count() > 0)
                    return q.ElementAt(0).nvc_codes;
                else
                    return null;
            }
        }

        #endregion

        #region IHS Multiplex Codes

        #region IHS Habitat

        public HluDataSet.lut_ihs_habitatRow[] IhsHabitatCodes
        {
            get
            {
                if (HluDataset.lut_ihs_habitat.IsInitialized && HluDataset.lut_ihs_habitat.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_ihs_habitatTableAdapter == null)
                        _hluTableAdapterMgr.lut_ihs_habitatTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_ihs_habitatDataTable, HluDataSet.lut_ihs_habitatRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_categoryDataTable) }, false);
                }

                if (_ihsHabitatCodes != null)
                {
                    if ((IncidCurrentRow != null) && (_ihsHabitatCodes.Count() == 1) &&
                        (IncidIhsHabitat != _ihsHabitatCodes.First().code))
                        IncidIhsHabitat = _ihsHabitatCodes.ElementAt(0).code;

                    OnPropertyChanged("IncidIhsHabitat");
                    return _ihsHabitatCodes.ToArray();
                }
                else if (!String.IsNullOrEmpty(IncidIhsHabitat))
                {
                    return HluDataset.lut_ihs_habitat.Where(r => r.code == IncidIhsHabitat).ToArray();
                }
                else
                {
                    return null;
                }
            }
            set { }
        }

        public bool IhsHabitatListEnabled
        {
            get { return IncidIhsHabitat != null || (IhsHabitatCodes != null && IhsHabitatCodes.Count() > 0); }
        }

        public string IncidIhsHabitat
        {
            get { return _incidIhsHabitat; }
            set
            {
                if (IncidCurrentRow != null)
                {
                    if (_pasting && (_ihsHabitatCodes.Count(r => r.code == value) == 0))
                    {
                        _pasting = false;
                        var q = HluDataset.lut_ihs_habitat.Where(r => r.code == value);
                        if (q.Count() > 0) IhsCategory = q.First().category;
                    }

                    _incidIhsHabitat = value;
                    _incidLastModifiedUser = UserID;
                    _incidLastModifiedDate = DateTime.Now;
                    OnPropertyChanged("IncidLastModifiedUser");
                    OnPropertyChanged("IncidLastModifiedDate");

                    if (value != null)
                    {
                        // set default IncidIhsComplex1 value, if any
                        if (_ihsHabitatCodes != null)
                        {
                            var q = _ihsHabitatCodes.Where(h => h.code == value && !h.IsNull(HluDataset.lut_ihs_habitat.default_complex_code));
                            if ((q.Count() > 0) && (IncidIhsComplex1 == null))
                            {
                                IncidIhsComplex1 = q.ElementAt(0).default_complex_code;
                                OnPropertyChanged("IncidIhsComplex1");
                            }
                        }

                        OnPropertyChanged("NvcCodes");
                        GetBapEnvironments();
                        OnPropertyChanged("IncidBapHabitatsAuto");
                        OnPropertyChanged("IncidBapHabitatsUser");
                        OnPropertyChanged("BapHabitatsAutoEnabled");
                        OnPropertyChanged("BapHabitatsUserEnabled");
                    }
                    IhsMultiplexCodes(_incidIhsHabitat);
                    RefreshIhsMulitplexCodes();
                }
            }
        }

        private string FindIhsCategory(string ihsHabitatCode)
        {
            if (!String.IsNullOrEmpty(ihsHabitatCode))
            {
                if (_ihsHabitatCodes != null)
                {
                    IEnumerable<HluDataSet.lut_ihs_habitatRow> hCat = _ihsHabitatCodes.Where(h => h.code == ihsHabitatCode);
                    if (hCat.Count() > 0) return hCat.First().category;
                }
                else if ((HluDataset != null) && (HluDataset.lut_ihs_habitat != null))
                {
                    IEnumerable<string> q = null;

                    switch (_gisLayerType)
                    {
                        case GeometryTypes.Point:
                            q = from h in HluDataset.lut_ihs_habitat
                                where h.is_local && h.point && h.code == ihsHabitatCode
                                select h.category;
                                break;
                        case GeometryTypes.Line:
                            q = from h in HluDataset.lut_ihs_habitat
                                where h.is_local && h.line && h.code == ihsHabitatCode
                                select h.category;
                            break;
                        case GeometryTypes.Polygon:
                            q = from h in HluDataset.lut_ihs_habitat
                                where h.is_local && h.polygon && h.code == ihsHabitatCode
                                select h.category;
                            break;
                    }
                    if ((q != null) && (q.Count() > 0)) return q.First();
                }
            }
            return null;
        }

        private string FindNvcCategory(string ihsHabitatCode)
        {
            if (!String.IsNullOrEmpty(ihsHabitatCode))
            {
                if (_ihsHabitatCodes != null)
                {
                    IEnumerable<HluDataSet.lut_ihs_habitatRow> hCat = _ihsHabitatCodes.Where(h => h.code == ihsHabitatCode);
                    if (hCat.Count() > 0) return hCat.First().category;
                }
                else if ((HluDataset != null) && (HluDataset.lut_ihs_habitat != null))
                {
                    IEnumerable<string> q = from h in HluDataset.lut_ihs_habitat
                                            join n in HluDataset.lut_ihs_habitat_ihs_nvc on h.code equals n.code_habitat
                                            where h.code == ihsHabitatCode
                                            select n.code_nvc;
                    if ((q != null) && (q.Count() > 0)) return q.First();
                }
            }
            return null;
        }

        #endregion

        #region IHS Matrix

        public HluDataSet.lut_ihs_matrixRow[] IhsMatrix1Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_matrixRow> q = null;
                HluDataSet.lut_ihs_matrixRow[] ihsMatrix1Codes;

                if ((_lutIhsMatrixCodes == null) || (_lutIhsMatrixCodes.Length == 0))
                {
                    ihsMatrix1Codes = new HluDataSet.lut_ihs_matrixRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first matrix code if there
                // is only one in the list.
                //
                //else if (_lutIhsMatrixCodes.Length == 1)
                //{
                //    ihsMatrix1Codes = _lutIhsMatrixCodes;
                //    IncidIhsMatrix1 = _lutIhsMatrixCodes[0].code;
                //    OnPropertyChanged("IncidIhsMatrix1");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsMatrix1))
                    {
                        if (_lutIhsMatrixCodes != null)
                            q = _lutIhsMatrixCodes.Where(r => r.code != IncidIhsMatrix2 && r.code != IncidIhsMatrix3);
                        else
                            q = _hluDS.lut_ihs_matrix.Where(r => r.code == IncidIhsMatrix1);

                        if (q == null)
                            ihsMatrix1Codes = new HluDataSet.lut_ihs_matrixRow[0];
                        else
                            ihsMatrix1Codes = ClearRowIhsMatrix(-3).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        ihsMatrix1Codes = _lutIhsMatrixCodes;
                    }
                }
                return ihsMatrix1Codes;
            }
            set { }
        }

        public HluDataSet.lut_ihs_matrixRow[] IhsMatrix2Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_matrixRow> q = null;
                HluDataSet.lut_ihs_matrixRow[] ihsMatrix2Codes;

                if ((_lutIhsMatrixCodes == null) || (_lutIhsMatrixCodes.Length == 0))
                {
                    ihsMatrix2Codes = new HluDataSet.lut_ihs_matrixRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first matrix code if there
                // is only one in the list.
                //
                //else if ((_lutIhsMatrixCodes.Length == 1) && (_lutIhsMatrixCodes[0].code != IncidIhsMatrix1))
                //{
                //    ihsMatrix2Codes = _lutIhsMatrixCodes;
                //    IncidIhsMatrix2 = _lutIhsMatrixCodes[0].code;
                //    OnPropertyChanged("IncidIhsMatrix2");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsMatrix2))
                    {
                        if (_lutIhsMatrixCodes != null)
                            q = _lutIhsMatrixCodes.Where(r => r.code != IncidIhsMatrix1 && r.code != IncidIhsMatrix3);
                        else
                            q = _hluDS.lut_ihs_matrix.Where(r => r.code == IncidIhsMatrix2);

                        if (q != null)
                            ihsMatrix2Codes = ClearRowIhsMatrix(-2).Concat(q).OrderBy(r => r.sort_order).ToArray();
                        else
                            ihsMatrix2Codes = new HluDataSet.lut_ihs_matrixRow[0];
                    }
                    else
                    {
                        if (_lutIhsMatrixCodes != null)
                            ihsMatrix2Codes = _lutIhsMatrixCodes.Where(r => r.code != IncidIhsMatrix1 && r.code != IncidIhsMatrix3).ToArray();
                        else
                            ihsMatrix2Codes = _lutIhsMatrixCodes;
                    }
                }
                return ihsMatrix2Codes;
            }
            set { }
        }

        public HluDataSet.lut_ihs_matrixRow[] IhsMatrix3Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_matrixRow> q = null;
                HluDataSet.lut_ihs_matrixRow[] ihsMatrix3Codes;

                if ((_lutIhsMatrixCodes == null) || (_lutIhsMatrixCodes.Length == 0))
                {
                    ihsMatrix3Codes = new HluDataSet.lut_ihs_matrixRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first matrix code if there
                // is only one in the list.
                //
                //else if ((_lutIhsMatrixCodes.Length == 1) && (_lutIhsMatrixCodes[0].code != IncidIhsMatrix1) && 
                //    (_lutIhsMatrixCodes[0].code != IncidIhsMatrix2))
                //{
                //    ihsMatrix3Codes = _lutIhsMatrixCodes;
                //    IncidIhsMatrix2 = _lutIhsMatrixCodes[0].code;
                //    OnPropertyChanged("IncidIhsMatrix2");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsMatrix3))
                    {
                        if (_lutIhsMatrixCodes != null)
                            q = _lutIhsMatrixCodes.Where(r => r.code != IncidIhsMatrix1 && r.code != IncidIhsMatrix2);
                        else
                            q = _hluDS.lut_ihs_matrix.Where(r => r.code == IncidIhsMatrix3);

                        if (q != null)
                            ihsMatrix3Codes = ClearRowIhsMatrix(-1).Concat(q).OrderBy(r => r.sort_order).ToArray();
                        else
                            ihsMatrix3Codes = new HluDataSet.lut_ihs_matrixRow[0];
                    }
                    else
                    {
                        if (_lutIhsMatrixCodes != null)
                            ihsMatrix3Codes = _lutIhsMatrixCodes.Where(r => r.code != IncidIhsMatrix1 && r.code != IncidIhsMatrix2).ToArray();
                        else
                            ihsMatrix3Codes = _lutIhsMatrixCodes;
                    }
                }
                return ihsMatrix3Codes;
            }
            set { }
        }

        public bool IhsMatrix1Enabled
        {
            get { return IhsMatrix1Codes.Count() > 0; }
        }

        public bool IhsMatrix2Enabled
        {
            get { return !String.IsNullOrEmpty(IncidIhsMatrix1) && IhsMatrix2Codes.Count() > 0; }
        }

        public bool IhsMatrix3Enabled
        {
            get
            {
                return !String.IsNullOrEmpty(IncidIhsMatrix1) && 
                    !String.IsNullOrEmpty(IncidIhsMatrix2) && IhsMatrix3Codes.Count() > 0;
            }
        }

        public string IncidIhsMatrix1
        {
            get
            {
                if (!CheckIhsMatrix()) return null;
                if ((_incidIhsMatrixRows.Length > 0) && (_incidIhsMatrixRows[0] != null) &&
                    !_incidIhsMatrixRows[0].IsNull(_hluDS.incid_ihs_matrix.matrixColumn.Ordinal))
                    return _incidIhsMatrixRows[0].matrix;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsMatrixRows.Length > 0) && (_incidIhsMatrixRows[0] != null))
                        RemoveIncidIhsMatrixRow(0);
                }
                else
                {
                    if ((_incidIhsMatrixRows.Length > 0) && (_incidIhsMatrixRows[0] != null))
                    {
                        _incidIhsMatrixRows[0].matrix = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsMatrixRow(value, 0);
                        OnPropertyChanged("IhsMatrix1Codes");
                    }
                }
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsMatrix1Codes");
                    OnPropertyChanged("IhsMatrix1Enabled");
                }
                if (!_comingFromIncidIhsMatrix2)
                {
                    OnPropertyChanged("IhsMatrix2Codes");
                    OnPropertyChanged("IhsMatrix2Enabled");
                }
                else
                {
                    _comingFromIncidIhsMatrix2 = false;
                }
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidIhsMatrix2
        {
            get
            {
                if (!CheckIhsMatrix()) return null;
                if ((_incidIhsMatrixRows.Length > 1) && (_incidIhsMatrixRows[1] != null) &&
                    !_incidIhsMatrixRows[1].IsNull(_hluDS.incid_ihs_matrix.matrixColumn.Ordinal))
                    return _incidIhsMatrixRows[1].matrix;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsMatrixRows.Length > 1) && (_incidIhsMatrixRows[1] != null))
                        RemoveIncidIhsMatrixRow(1);
                }
                else if ((_incidIhsMatrixRows[0] != null) && (_incidIhsMatrixRows[0].matrix == value))
                {
                    return;
                }
                else
                {
                    if ((_incidIhsMatrixRows.Length > 1) && (_incidIhsMatrixRows[1] != null))
                    {
                        _incidIhsMatrixRows[1].matrix = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsMatrixRow(value, 1);
                        OnPropertyChanged("IhsMatrix2Codes");
                    }
                }

                _comingFromIncidIhsMatrix2 = true;
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsMatrix2Codes");
                    OnPropertyChanged("IhsMatrix2Enabled");
                }
                OnPropertyChanged("IhsMatrix1Codes");
                if (!_comingFromIncidIhsMatrix3)
                {
                    OnPropertyChanged("IhsMatrix3Codes");
                    OnPropertyChanged("IhsMatrix3Enabled");
                }
                else
                {
                    _comingFromIncidIhsMatrix3 = false;
                }
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidIhsMatrix3
        {
            get
            {
                if (!CheckIhsMatrix()) return null;
                if ((_incidIhsMatrixRows.Length > 2) && (_incidIhsMatrixRows[2] != null) &&
                    !_incidIhsMatrixRows[2].IsNull(_hluDS.incid_ihs_matrix.matrixColumn.Ordinal))
                    return _incidIhsMatrixRows[2].matrix;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsMatrixRows.Length > 2) && (_incidIhsMatrixRows[2] != null))
                        RemoveIncidIhsMatrixRow(2);
                }
                else if (((_incidIhsMatrixRows[1] != null) && (_incidIhsMatrixRows[1].matrix == value)) ||
                    (_incidIhsMatrixRows[0] != null) && (_incidIhsMatrixRows[0].matrix == value))
                {
                    return;
                }
                else
                {
                    if ((_incidIhsMatrixRows.Length > 2) && (_incidIhsMatrixRows[2] != null))
                    {
                        _incidIhsMatrixRows[2].matrix = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsMatrixRow(value, 2);
                        OnPropertyChanged("IhsMatrix3Codes");
                    }
                }
                _comingFromIncidIhsMatrix3 = true;
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsMatrix3Codes");
                    OnPropertyChanged("IhsMatrix3Enabled");
                }
                OnPropertyChanged("IhsMatrix1Codes");
                OnPropertyChanged("IhsMatrix2Codes");
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        private bool CheckIhsMatrix()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidIhsMatrixRows == null)
            {
                HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable = _hluDS.incid_ihs_matrix;
                GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_ihs_matrixTableAdapter, ref ihsMatrixTable);
            }

            return _incidIhsMatrixRows != null;
        }

        private HluDataSet.lut_ihs_matrixRow[] ClearRowIhsMatrix(int sortOrder)
        {
            HluDataSet.lut_ihs_matrixRow clearRow = _hluDS.lut_ihs_matrix.Newlut_ihs_matrixRow();
            clearRow.code = _codeDeleteRow;
            clearRow.sort_order = sortOrder;
            return new HluDataSet.lut_ihs_matrixRow[] { clearRow };
        }

        private void AddIncidIhsMatrixRow(string newMatrixCode, int ix)
        {
            if (newMatrixCode != null)
            {
                HluDataSet.incid_ihs_matrixRow newRow = IncidIhsMatrixTable.Newincid_ihs_matrixRow();
                if ((_bulkUpdateMode == false))
                {
                    newRow.matrix_id = NextIncidIhsMatrixId;
                    newRow.incid = IncidCurrentRow.incid;
                    newRow.matrix = newMatrixCode;
                    IncidIhsMatrixTable.Addincid_ihs_matrixRow(newRow);
                    _incidIhsMatrixRows = GetIncidChildRows<HluDataSet.incidDataTable, 
                        HluDataSet.incid_ihs_matrixDataTable, HluDataSet.incid_ihs_matrixRow>(
                        IncidTable, IncidIhsMatrixTable);
                }
                else
                {
                    newRow.matrix_id = ix;
                    newRow.matrix = newMatrixCode;
                    _incidIhsMatrixRows[ix] = newRow;
                }
                OnPropertyChanged("IncidIhsMatrix1");
                OnPropertyChanged("IncidIhsMatrix2");
                OnPropertyChanged("IncidIhsMatrix3");
                OnPropertyChanged("IncidIhsSummary");
            }
        }

        private void RemoveIncidIhsMatrixRow(int rowNumber)
        {
            if (_incidIhsMatrixRows[rowNumber].RowState != DataRowState.Detached)
                _incidIhsMatrixRows[rowNumber].Delete();
            _incidIhsMatrixRows[rowNumber] = null;
        }

        #endregion

        #region IHS Formation

        public HluDataSet.lut_ihs_formationRow[] IhsFormation1Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_formationRow> q = null;
                HluDataSet.lut_ihs_formationRow[] ihsFormation1Codes;

                if ((_lutIhsFormationCodes == null) || (_lutIhsFormationCodes.Length == 0))
                {
                    ihsFormation1Codes = new HluDataSet.lut_ihs_formationRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first formation code if there
                // is only one in the list.
                //
                //else if (_lutIhsFormationCodes.Length == 1)
                //{
                //    ihsFormation1Codes = _lutIhsFormationCodes;
                //    IncidIhsFormation1 = _lutIhsFormationCodes[0].code;
                //    OnPropertyChanged("IncidIhsFormation1");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsFormation1))
                    {
                        if (_lutIhsFormationCodes != null)
                            q = _lutIhsFormationCodes.Where(r => r.code != IncidIhsFormation2);
                        else
                            q = _hluDS.lut_ihs_formation.Where(r => r.code == IncidIhsFormation1);

                        if (q == null)
                            ihsFormation1Codes = new HluDataSet.lut_ihs_formationRow[0];
                        else
                            ihsFormation1Codes = ClearRowIhsFormation(-2).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        ihsFormation1Codes = _lutIhsFormationCodes;
                    }
                }
                return ihsFormation1Codes;
            }
            set { }
        }

        public HluDataSet.lut_ihs_formationRow[] IhsFormation2Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_formationRow> q = null;
                HluDataSet.lut_ihs_formationRow[] ihsFormation2Codes;

                if ((_lutIhsFormationCodes == null) || (_lutIhsFormationCodes.Length == 0))
                {
                    ihsFormation2Codes = new HluDataSet.lut_ihs_formationRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first formation code if there
                // is only one in the list.
                //
                //else if ((_lutIhsFormationCodes.Length == 1) && (_lutIhsFormationCodes[0].code != IncidIhsFormation1))
                //{
                //    ihsFormation2Codes = _lutIhsFormationCodes;
                //    IncidIhsFormation2 = _lutIhsFormationCodes[0].code;
                //    OnPropertyChanged("IncidIhsFormation2");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsFormation2))
                    {
                        if (_lutIhsFormationCodes != null)
                            q = _lutIhsFormationCodes.Where(r => r.code != IncidIhsFormation1);
                        else
                            q = _hluDS.lut_ihs_formation.Where(r => r.code == IncidIhsFormation2);

                        if (q == null)
                            ihsFormation2Codes = new HluDataSet.lut_ihs_formationRow[0];
                        else
                            ihsFormation2Codes = ClearRowIhsFormation(-1).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        if (_lutIhsFormationCodes != null)
                            ihsFormation2Codes = _lutIhsFormationCodes.Where(r => r.code != IncidIhsFormation1).ToArray();
                        else
                            ihsFormation2Codes = _lutIhsFormationCodes;
                    }
                }
                return ihsFormation2Codes;
            }
            set { }
        }

        public bool IhsFormation1Enabled
        {
            get { return IhsFormation1Codes.Count() > 0; }
        }

        public bool IhsFormation2Enabled
        {
            get { return !String.IsNullOrEmpty(IncidIhsFormation1) && IhsFormation2Codes.Count() > 0; }
        }

        public string IncidIhsFormation1
        {
            get
            {
                if (!CheckIhsFormation()) return null;
                if ((_incidIhsFormationRows.Length > 0) && (_incidIhsFormationRows[0] != null) &&
                    !_incidIhsFormationRows[0].IsNull(_hluDS.incid_ihs_formation.formationColumn.Ordinal))
                    return _incidIhsFormationRows[0].formation;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsFormationRows.Length > 0) && (_incidIhsFormationRows[0] != null))
                        RemoveIncidIhsFormationRow(0);
                }
                else
                {
                    if ((_incidIhsFormationRows.Length > 0) && (_incidIhsFormationRows[0] != null))
                    {
                        _incidIhsFormationRows[0].formation = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsFormationRow(value, 0);
                        OnPropertyChanged("IhsFormation1Codes");
                    }
                }
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsFormation1Codes");
                    OnPropertyChanged("IhsFormation1Enabled");
                }
                if (!_comingFromIncidIhsFormation2)
                {
                    OnPropertyChanged("IhsFormation2Codes");
                    OnPropertyChanged("IhsFormation2Enabled");
                }
                else
                {
                    _comingFromIncidIhsFormation2 = false;
                }
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidIhsFormation2
        {
            get
            {
                if (!CheckIhsFormation()) return null;
                if ((_incidIhsFormationRows.Length > 1) && (_incidIhsFormationRows[1] != null) &&
                    !_incidIhsFormationRows[1].IsNull(_hluDS.incid_ihs_formation.formationColumn.Ordinal))
                    return _incidIhsFormationRows[1].formation;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsFormationRows.Length > 1) && (_incidIhsFormationRows[1] != null))
                        RemoveIncidIhsFormationRow(1);
                }
                else if ((_incidIhsFormationRows[0] != null) && (_incidIhsFormationRows[0].formation == value))
                {
                    return;
                }
                else
                {
                    if ((_incidIhsFormationRows.Length > 1) && (_incidIhsFormationRows[1] != null))
                    {
                        _incidIhsFormationRows[1].formation = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsFormationRow(value, 1);
                        OnPropertyChanged("IhsFormation2Codes");
                    }
                }

                _comingFromIncidIhsFormation2 = true;
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsFormation2Codes");
                    OnPropertyChanged("IhsFormation2Enabled");
                }
                OnPropertyChanged("IhsFormation1Codes");
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        private bool CheckIhsFormation()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidIhsFormationRows == null)
            {
                HluDataSet.incid_ihs_formationDataTable ihsFormationTable = _hluDS.incid_ihs_formation;
                GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_ihs_formationTableAdapter, ref ihsFormationTable);
            }
            return _incidIhsFormationRows != null;
        }

        private HluDataSet.lut_ihs_formationRow[] ClearRowIhsFormation(int sortOrder)
        {
            HluDataSet.lut_ihs_formationRow clearRow = _hluDS.lut_ihs_formation.Newlut_ihs_formationRow();
            clearRow.code = _codeDeleteRow;
            clearRow.sort_order = sortOrder;
            return new HluDataSet.lut_ihs_formationRow[] { clearRow };
        }
        
        private void AddIncidIhsFormationRow(string newFormationCode, int ix)
        {
            if (newFormationCode != null)
            {
                HluDataSet.incid_ihs_formationRow newRow = IncidIhsFormationTable.Newincid_ihs_formationRow();
                if ((_bulkUpdateMode == false))
                {
                    newRow.formation_id = NextIncidIhsFormationId;
                    newRow.incid = IncidCurrentRow.incid;
                    newRow.formation = newFormationCode;
                    IncidIhsFormationTable.Addincid_ihs_formationRow(newRow);
                    _incidIhsFormationRows = GetIncidChildRows<HluDataSet.incidDataTable, 
                        HluDataSet.incid_ihs_formationDataTable, HluDataSet.incid_ihs_formationRow>(
                        IncidTable, IncidIhsFormationTable);
                }
                else
                {
                    newRow.formation_id = ix;
                    newRow.formation = newFormationCode;
                    _incidIhsFormationRows[ix] = newRow;
                }
                
                OnPropertyChanged("IncidIhsFormation1");
                OnPropertyChanged("IncidIhsFormation2");
                OnPropertyChanged("IncidIhsSummary");
            }
        }

        private void RemoveIncidIhsFormationRow(int rowNumber)
        {
            if (_incidIhsFormationRows[rowNumber].RowState != DataRowState.Detached)
                _incidIhsFormationRows[rowNumber].Delete();
            _incidIhsFormationRows[rowNumber] = null;
        }

        #endregion

        #region IHS Management

        public HluDataSet.lut_ihs_managementRow[] IhsManagement1Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_managementRow> q = null;
                HluDataSet.lut_ihs_managementRow[] ihsManagement1Codes;

                if ((_lutIhsManagementCodes == null) || (_lutIhsManagementCodes.Length == 0))
                {
                    ihsManagement1Codes = new HluDataSet.lut_ihs_managementRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first management code if there
                // is only one in the list.
                //
                //else if (_lutIhsManagementCodes.Length == 1)
                //{
                //    ihsManagement1Codes = _lutIhsManagementCodes;
                //    IncidIhsManagement1 = _lutIhsManagementCodes[0].code;
                //    OnPropertyChanged("IncidIhsManagement1");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsManagement1))
                    {
                        if (_lutIhsManagementCodes != null)
                            q = _lutIhsManagementCodes.Where(r => r.code != IncidIhsManagement2);
                        else
                            q = _hluDS.lut_ihs_management.Where(r => r.code == IncidIhsManagement1);

                        if (q == null)
                            ihsManagement1Codes = new HluDataSet.lut_ihs_managementRow[0];
                        else
                            ihsManagement1Codes = ClearRowIhsManagement(-2).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        ihsManagement1Codes = _lutIhsManagementCodes;
                    }
                }
                return ihsManagement1Codes;
            }
            set { }
        }

        public HluDataSet.lut_ihs_managementRow[] IhsManagement2Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_managementRow> q = null;
                HluDataSet.lut_ihs_managementRow[] ihsManagement2Codes;

                if ((_lutIhsManagementCodes == null) || (_lutIhsManagementCodes.Length == 0))
                {
                    ihsManagement2Codes = new HluDataSet.lut_ihs_managementRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first management code if there
                // is only one in the list.
                //
                //else if ((_lutIhsManagementCodes.Length == 1) && (_lutIhsManagementCodes[0].code != IncidIhsManagement1))
                //{
                //    ihsManagement2Codes = _lutIhsManagementCodes;
                //    IncidIhsManagement2 = _lutIhsManagementCodes[0].code;
                //    OnPropertyChanged("IncidIhsManagement2");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsManagement2))
                    {
                        if (_lutIhsManagementCodes != null)
                            q = _lutIhsManagementCodes.Where(r => r.code != IncidIhsManagement1);
                        else
                            q = _hluDS.lut_ihs_management.Where(r => r.code == IncidIhsManagement2);

                        if (q == null)
                            ihsManagement2Codes = new HluDataSet.lut_ihs_managementRow[0];
                        else
                            ihsManagement2Codes = ClearRowIhsManagement(-1).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        if (_lutIhsManagementCodes != null)
                            ihsManagement2Codes = _lutIhsManagementCodes.Where(r => r.code != IncidIhsManagement1).ToArray();
                        else
                            ihsManagement2Codes = _lutIhsManagementCodes;
                    }
                }
                return ihsManagement2Codes;
            }
            set { }
        }

        public bool IhsManagement1Enabled
        {
            get { return IhsManagement1Codes.Count() > 0; }
        }

        public bool IhsManagement2Enabled
        {
            get { return !String.IsNullOrEmpty(IncidIhsManagement1) && IhsManagement2Codes.Count() > 0; }
        }

        public string IncidIhsManagement1
        {
            get
            {
                if (!CheckIhsManagement()) return null;
                if ((_incidIhsManagementRows.Length > 0) && (_incidIhsManagementRows[0] != null) && 
                    !_incidIhsManagementRows[0].IsNull(_hluDS.incid_ihs_management.managementColumn.Ordinal))
                    return _incidIhsManagementRows[0].management;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsManagementRows.Length > 0) && (_incidIhsManagementRows[0] != null))
                        RemoveIncidIhsManagementRow(0);
                }
                else
                {
                    if ((_incidIhsManagementRows.Length > 0) && (_incidIhsManagementRows[0] != null))
                    {
                        _incidIhsManagementRows[0].management = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsManagementRow(value, 0);
                        OnPropertyChanged("IhsManagement1Codes");
                    }
                }

                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsManagement1Codes");
                    OnPropertyChanged("IhsManagement1Enabled");
                }
                if (!_comingFromIncidIhsManagement2)
                {
                    OnPropertyChanged("IhsManagement2Codes");
                    OnPropertyChanged("IhsManagement2Enabled");
                }
                else
                {
                    _comingFromIncidIhsManagement2 = false;
                }
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidIhsManagement2
        {
            get
            {
                if (!CheckIhsManagement()) return null;
                if ((_incidIhsManagementRows.Length > 1) && (_incidIhsManagementRows[1] != null) &&
                    !_incidIhsManagementRows[1].IsNull(_hluDS.incid_ihs_management.managementColumn.Ordinal))
                    return _incidIhsManagementRows[1].management;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsManagementRows.Length > 1) && (_incidIhsManagementRows[1] != null))
                        RemoveIncidIhsManagementRow(1);
                }
                else if ((_incidIhsManagementRows[0] != null) && (_incidIhsManagementRows[0].management == value))
                {
                    return;
                }
                else
                {
                    if ((_incidIhsManagementRows.Length > 1) && (_incidIhsManagementRows[1] != null))
                    {
                        _incidIhsManagementRows[1].management = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsManagementRow(value, 1);
                        OnPropertyChanged("IhsManagement2Codes");
                    }
                }

                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsManagement2Codes");
                    OnPropertyChanged("IhsManagement2Enabled");
                }
                _comingFromIncidIhsManagement2 = true;
                OnPropertyChanged("IhsManagement1Codes");
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        private bool CheckIhsManagement()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidIhsManagementRows == null)
            {
                HluDataSet.incid_ihs_managementDataTable ihsManagementTable = _hluDS.incid_ihs_management;
                GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_ihs_managementTableAdapter, ref ihsManagementTable);
            }

            return _incidIhsManagementRows != null;
        }
        
        private HluDataSet.lut_ihs_managementRow[] ClearRowIhsManagement(int sortOrder)
        {
            HluDataSet.lut_ihs_managementRow clearRow = _hluDS.lut_ihs_management.Newlut_ihs_managementRow();
            clearRow.code = _codeDeleteRow;
            clearRow.sort_order = sortOrder;
            return new HluDataSet.lut_ihs_managementRow[] { clearRow };
        }

        private void AddIncidIhsManagementRow(string newManagementCode, int ix)
        {
            if (newManagementCode != null)
            {
                HluDataSet.incid_ihs_managementRow newRow = IncidIhsManagementTable.Newincid_ihs_managementRow();
                if ((_bulkUpdateMode == false))
                {
                    newRow.management_id = NextIncidIhsManagementId;
                    newRow.incid = IncidCurrentRow.incid;
                    newRow.management = newManagementCode;
                    IncidIhsManagementTable.Addincid_ihs_managementRow(newRow);
                    _incidIhsManagementRows = GetIncidChildRows<HluDataSet.incidDataTable, 
                        HluDataSet.incid_ihs_managementDataTable, HluDataSet.incid_ihs_managementRow>(
                        IncidTable, IncidIhsManagementTable);
                }
                else
                {
                    newRow.management_id = ix;
                    newRow.management = newManagementCode;
                    _incidIhsManagementRows[ix] = newRow;
                }

                OnPropertyChanged("IncidIhsManagement1");
                OnPropertyChanged("IncidIhsManagement2");
                OnPropertyChanged("IncidIhsSummary");
            }
        }

        private void RemoveIncidIhsManagementRow(int rowNumber)
        {
            if (_incidIhsManagementRows[rowNumber].RowState != DataRowState.Detached)
                _incidIhsManagementRows[rowNumber].Delete();
            _incidIhsManagementRows[rowNumber] = null;
        }

        #endregion

        #region IHS Complex

        public HluDataSet.lut_ihs_complexRow[] IhsComplex1Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_complexRow> q = null;
                HluDataSet.lut_ihs_complexRow[] ihsComplex1Codes;

                if ((_lutIhsComplexCodes == null) || (_lutIhsComplexCodes.Length == 0))
                {
                    ihsComplex1Codes = new HluDataSet.lut_ihs_complexRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first complex code if there
                // is only one in the list.
                //
                //else if (_lutIhsComplexCodes.Length == 1)
                //{
                //    ihsComplex1Codes = _lutIhsComplexCodes;
                //    IncidIhsComplex1 = _lutIhsComplexCodes[0].code;
                //    OnPropertyChanged("IncidIhsComplex1");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsComplex1))
                    {
                        if (_lutIhsComplexCodes != null)
                            q = _lutIhsComplexCodes.Where(r => r.code != IncidIhsComplex2);
                        else
                            q = _hluDS.lut_ihs_complex.Where(r => r.code == IncidIhsComplex1);

                        if (q == null)
                            ihsComplex1Codes = new HluDataSet.lut_ihs_complexRow[0];
                        else
                            ihsComplex1Codes = ClearRowIhsComplex(-2).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        ihsComplex1Codes = _lutIhsComplexCodes;
                    }
                }
                return ihsComplex1Codes;
            }
            set { }
        }

        public HluDataSet.lut_ihs_complexRow[] IhsComplex2Codes
        {
            get
            {
                IEnumerable<HluDataSet.lut_ihs_complexRow> q = null;
                HluDataSet.lut_ihs_complexRow[] ihsComplex2Codes;

                if ((_lutIhsComplexCodes == null) || (_lutIhsComplexCodes.Length == 0))
                {
                    ihsComplex2Codes = new HluDataSet.lut_ihs_complexRow[0];
                }
                //---------------------------------------------------------------------
                // FIXED: KI105 (Mandatory multiplex codes)
                // Don't automatically select the first complex code if there
                // is only one in the list.
                //
                //else if ((_lutIhsComplexCodes.Length == 1) && (_lutIhsComplexCodes[0].code != IncidIhsComplex1))
                //{
                //    ihsComplex2Codes = _lutIhsComplexCodes;
                //    IncidIhsComplex2 = _lutIhsComplexCodes[0].code;
                //    OnPropertyChanged("IncidIhsComplex2");
                //}
                //---------------------------------------------------------------------
                else
                {
                    if (!String.IsNullOrEmpty(IncidIhsComplex2))
                    {
                        if (_lutIhsComplexCodes != null)
                            q = _lutIhsComplexCodes.Where(r => r.code != IncidIhsComplex1);
                        else
                            q = _hluDS.lut_ihs_complex.Where(r => r.code == IncidIhsComplex2);

                        if (q == null)
                            ihsComplex2Codes = new HluDataSet.lut_ihs_complexRow[0];
                        else
                            ihsComplex2Codes = ClearRowIhsComplex(-1).Concat(q).OrderBy(r => r.sort_order).ToArray();
                    }
                    else
                    {
                        if (_lutIhsComplexCodes != null)
                            ihsComplex2Codes = _lutIhsComplexCodes.Where(r => r.code != IncidIhsComplex1).ToArray();
                        else
                            ihsComplex2Codes = _lutIhsComplexCodes;
                    }
                }
                return ihsComplex2Codes;
            }
            set { }
        }

        public bool IhsComplex1Enabled
        {
            get { return IhsComplex1Codes.Count() > 0; }
        }

        public bool IhsComplex2Enabled
        {
            get { return !String.IsNullOrEmpty(IncidIhsComplex1) && IhsComplex2Codes.Count() > 0; }
        }

        public string IncidIhsComplex1
        {
            get
            {
                if (!CheckIhsComplex()) return null;
                if ((_incidIhsComplexRows.Length > 0) && (_incidIhsComplexRows[0] != null) && 
                    !_incidIhsComplexRows[0].IsNull(_hluDS.incid_ihs_complex.complexColumn.Ordinal))
                    return _incidIhsComplexRows[0].complex;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsComplexRows.Length > 0) && (_incidIhsComplexRows[0] != null))
                        RemoveIncidIhsComplexRow(0);
                }
                else
                {
                    if ((_incidIhsComplexRows.Length > 0) && (_incidIhsComplexRows[0] != null))
                    {
                        _incidIhsComplexRows[0].complex = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsComplexRow(value, 0);
                        OnPropertyChanged("IhsComplex1Codes");
                    }
                }

                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsComplex1Codes");
                    OnPropertyChanged("IhsComplex1Enabled");
                }

                if (!_comingFromIncidIhsComplex2)
                {
                    OnPropertyChanged("IhsComplex2Codes");
                    OnPropertyChanged("IhsComplex2Enabled");
                }
                else
                {
                    _comingFromIncidIhsComplex2 = false;
                }
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidIhsComplex2
        {
            get
            {
                if (!CheckIhsComplex()) return null;
                if ((_incidIhsComplexRows.Length > 1) && (_incidIhsComplexRows[1] != null) &&
                    !_incidIhsComplexRows[1].IsNull(_hluDS.incid_ihs_complex.complexColumn.Ordinal))
                    return _incidIhsComplexRows[1].complex;
                else
                    return null;
            }
            set
            {
                bool removeDeleteCode = value == _codeDeleteRow;
                if (removeDeleteCode) value = null;

                if (value == null)
                {
                    if ((_incidIhsComplexRows.Length > 1) && (_incidIhsComplexRows[1] != null))
                        RemoveIncidIhsComplexRow(1);
                }
                else if ((_incidIhsComplexRows[0] != null) && (_incidIhsComplexRows[0].complex == value))
                {
                    return;
                }
                else
                {
                    if ((_incidIhsComplexRows.Length > 1) && (_incidIhsComplexRows[1] != null))
                    {
                        _incidIhsComplexRows[1].complex = value;
                        OnPropertyChanged("IncidIhsSummary");
                    }
                    else
                    {
                        AddIncidIhsComplexRow(value, 1);
                        OnPropertyChanged("IhsComplex2Codes");
                    }
                }

                _comingFromIncidIhsComplex2 = true;
                if (removeDeleteCode)
                {
                    OnPropertyChanged("IhsComplex2Codes");
                    OnPropertyChanged("IhsComplex2Enabled");
                }
                OnPropertyChanged("IhsComplex1Codes");
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                //---------------------------------------------------------------------
                // FIXED: KI96 (BAP Habitats)
                // Enable editing of bap habitats when they are only associated
                // with matrix, formation, management or complex codes (rather
                // than habitat codes.
                OnPropertyChanged("BapHabitatsAutoEnabled");
                //---------------------------------------------------------------------
                OnPropertyChanged("BapHabitatsUserEnabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        private bool CheckIhsComplex()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidIhsComplexRows == null)
            {
                HluDataSet.incid_ihs_complexDataTable ihsComplexTable = _hluDS.incid_ihs_complex;
                GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_ihs_complexTableAdapter, ref ihsComplexTable);
            }
            return _incidIhsComplexRows != null;
        }

        private HluDataSet.lut_ihs_complexRow[] ClearRowIhsComplex(int sortOrder)
        {
            HluDataSet.lut_ihs_complexRow clearRow = _hluDS.lut_ihs_complex.Newlut_ihs_complexRow();
            clearRow.code = _codeDeleteRow;
            clearRow.sort_order = sortOrder;
            return new HluDataSet.lut_ihs_complexRow[] { clearRow };
        }

        private void AddIncidIhsComplexRow(string newComplexCode, int ix)
        {
            if (newComplexCode != null)
            {
                HluDataSet.incid_ihs_complexRow newRow = IncidIhsComplexTable.Newincid_ihs_complexRow();
                if ((_bulkUpdateMode == false))
                {
                    newRow.complex_id = NextIncidIhsComplexId;
                    newRow.incid = IncidCurrentRow.incid;
                    newRow.complex = newComplexCode;
                    IncidIhsComplexTable.Addincid_ihs_complexRow(newRow);
                    _incidIhsComplexRows = GetIncidChildRows<HluDataSet.incidDataTable, 
                        HluDataSet.incid_ihs_complexDataTable, HluDataSet.incid_ihs_complexRow>(
                        IncidTable, IncidIhsComplexTable);
                }
                else
                {
                    newRow.complex_id = ix;
                    newRow.complex = newComplexCode;
                    _incidIhsComplexRows[ix] = newRow;
                }

                OnPropertyChanged("IncidIhsComplex1");
                OnPropertyChanged("IncidIhsComplex2");
                OnPropertyChanged("IncidIhsSummary");
            }
        }

        private void RemoveIncidIhsComplexRow(int rowNumber)
        {
            if (_incidIhsComplexRows[rowNumber].RowState != DataRowState.Detached)
                _incidIhsComplexRows[rowNumber].Delete();
            _incidIhsComplexRows[rowNumber] = null;
        }

        #endregion

        private void IhsMultiplexCodes(string incidIhsHabitat)
        {
            if (!String.IsNullOrEmpty(incidIhsHabitat))
            {
                if (HluDataset.lut_ihs_matrix.IsInitialized && HluDataset.lut_ihs_matrix.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_ihs_matrixTableAdapter == null)
                        _hluTableAdapterMgr.lut_ihs_matrixTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_matrixDataTable) }, false);
                }

                _lutIhsMatrixCodes = (from m in HluDataset.lut_ihs_matrix
                                      join h in HluDataset.lut_ihs_habitat_ihs_matrix on m.code equals h.code_matrix
                                      where h.code_habitat == IncidIhsHabitat
                                      select m).OrderBy(m => m.sort_order).ToArray();

                if (HluDataset.lut_ihs_formation.IsInitialized && HluDataset.lut_ihs_formation.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_ihs_formationTableAdapter == null)
                        _hluTableAdapterMgr.lut_ihs_formationTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_formationDataTable) }, false);
                }

                _lutIhsFormationCodes = (from f in HluDataset.lut_ihs_formation
                                         join h in HluDataset.lut_ihs_habitat_ihs_formation on f.code equals h.code_formation
                                         where h.code_habitat == IncidIhsHabitat
                                         select f).OrderBy(f => f.sort_order).ToArray();

                if (HluDataset.lut_ihs_management.IsInitialized && HluDataset.lut_ihs_management.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_ihs_managementTableAdapter == null)
                        _hluTableAdapterMgr.lut_ihs_managementTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_managementDataTable) }, false);
                }

                _lutIhsManagementCodes = (from m in HluDataset.lut_ihs_management
                                          join h in HluDataset.lut_ihs_habitat_ihs_management on m.code equals h.code_management
                                          where h.code_habitat == IncidIhsHabitat
                                          select m).OrderBy(m => m.sort_order).ToArray();

                if (HluDataset.lut_ihs_complex.IsInitialized && HluDataset.lut_ihs_complex.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_ihs_complexTableAdapter == null)
                        _hluTableAdapterMgr.lut_ihs_complexTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_complexDataTable) }, false);
                }

                _lutIhsComplexCodes = (from c in HluDataset.lut_ihs_complex
                                       join h in HluDataset.lut_ihs_habitat_ihs_complex on c.code equals h.code_complex
                                       where h.code_habitat == incidIhsHabitat
                                       select c).OrderBy(c => c.sort_order).ToArray();
            }
            else
            {
                _lutIhsMatrixCodes = null;
                _lutIhsFormationCodes = null;
                _lutIhsManagementCodes = null;
                _lutIhsComplexCodes = null;
            }
        }

        #endregion

        #region Summary

        public string IncidIhsSummary
        {
            get
            {
                return ViewModelWindowMainHelpers.IhsSummary(new string[] { 
                    IncidIhsHabitat, IncidIhsMatrix1, IncidIhsMatrix2, IncidIhsMatrix3,
                    IncidIhsFormation1, IncidIhsFormation2, IncidIhsManagement1, IncidIhsManagement2,
                    IncidIhsComplex1, IncidIhsComplex2 });
            }
        }

        public string IncidLegacyHabitat
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.legacy_habitatColumn))
                    return IncidCurrentRow.legacy_habitat;
                else
                    return null;
            }
            set
            { 
                if ((IncidCurrentRow != null) && (value != null)) IncidCurrentRow.legacy_habitat = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        #endregion

        #endregion

        #region Details Tab

        //---------------------------------------------------------------------
        // FIX: Show field errors on tab labels.
        // Set the Details tab label from here so that validation can be done.
        // This will enable tooltips to be shown so that validation errors
        // in any fields in the tab can be highlighted by flagging the tab
        // label as in error.
        public string DetailsTabLabel
        {
            get { return "Details"; }
        }
        //---------------------------------------------------------------------

        #region BAP Habitat

        public HluDataSet.lut_habitat_typeRow[] BapHabitatCodes
        {
            get
            {
                if (_bapHabitatCodes == null)
                {
                    if (HluDataset.lut_habitat_type.IsInitialized && (HluDataset.lut_habitat_type.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_categoryTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    _bapHabitatCodes = (from r in HluDataset.lut_habitat_type
                                        where r.habitat_class_code == "PHAP"
                                        select r).ToArray();
                }

                return _bapHabitatCodes;
            }
        }

        public HluDataSet.lut_bap_quality_determinationRow[] BapDeterminationQualityCodesAuto
        {
            get
            {
                if (BapDeterminationQualityCodes != null)
                    return BapDeterminationQualityCodes.Where(r => r.code != BapEnvironment.BAPDetQltyUserAdded).OrderBy(r => r.sort_order).ToArray();
                else
                    return null;
            }
        }

        public HluDataSet.lut_bap_quality_determinationRow[] BapDeterminationQualityCodesUser
        {
            get
            {
                //---------------------------------------------------------------------
                // FIX: Allow 'None' habitats to be managed
                // Show all determination quality values in the drop-down list (instead
                // of just 'Not present but close to definition') but validate the
                // selected value later.
                if (BapDeterminationQualityCodes != null)
                    //return BapDeterminationQualityCodes.Where(r => r.code == BapEnvironment.BAPDetQltyUserAdded).OrderBy(r => r.sort_order).ToArray();
                    return BapDeterminationQualityCodes.OrderBy(r => r.sort_order).ToArray();
                else
                    return null;
                //---------------------------------------------------------------------
            }
        }

        public HluDataSet.lut_bap_quality_determinationRow[] BapDeterminationQualityCodes
        {
            get
            {
                if (_bapDeterminationQualityCodes == null)
                {
                    if (HluDataset.lut_bap_quality_determination.IsInitialized &&
                        HluDataset.lut_bap_quality_determination.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_bap_quality_determinationTableAdapter == null)
                            _hluTableAdapterMgr.lut_bap_quality_determinationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_bap_quality_determinationDataTable,
                                    HluDataSet.lut_bap_quality_determinationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_bap_quality_determinationDataTable) }, false);
                    }
                    _bapDeterminationQualityCodes =
                        HluDataset.lut_bap_quality_determination.OrderBy(r => r.sort_order).ToArray();
                }
                return _bapDeterminationQualityCodes;
            }
        }

        public HluDataSet.lut_bap_quality_interpretationRow[] BapInterpretationQualityCodes
        {
            get
            {
                if (_bapInterpretationQualityCodes == null)
                {
                    if (HluDataset.lut_bap_quality_interpretation.IsInitialized &&
                        HluDataset.lut_bap_quality_interpretation.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_bap_quality_interpretationTableAdapter == null)
                            _hluTableAdapterMgr.lut_bap_quality_interpretationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_bap_quality_interpretationDataTable,
                                    HluDataSet.lut_bap_quality_interpretationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_bap_quality_interpretationDataTable) }, false);
                    }
                    _bapInterpretationQualityCodes =
                        HluDataset.lut_bap_quality_interpretation.OrderBy(r => r.sort_order).ToArray();
                }
                return _bapInterpretationQualityCodes;
            }
        }

        public ObservableCollection<BapEnvironment> IncidBapHabitatsAuto
        {
            get { return _incidBapRowsAuto; }
            set
            {
                _incidBapRowsAuto = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// The bap_id of existing secondary priority habitats is multiplied by -1 (and same again when 
        /// saving back to DB) to distinguish them from primary priority habitats in UI validation methods.
        /// </summary>
        public ObservableCollection<BapEnvironment> IncidBapHabitatsUser
        {
            get { return _incidBapRowsUser; }
            set 
            { 
                _incidBapRowsUser = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public bool BapHabitatsAutoEnabled
        {
            get { return IncidBapHabitatsAuto != null && IncidBapHabitatsAuto.Count > 0; } // return _bulkUpdateMode == false; }
        }

        public bool BapHabitatsUserEnabled
        {
            get
            {
                return true;
                //return _bulkUpdateMode == true || (IncidBapHabitatsAuto != null && 
                //    IncidBapHabitatsAuto.Count > 0) || (IncidBapHabitatsUser.Count > 0);
            }
        }

        /// <summary>
        /// Gets the bap environments.
        /// </summary>
        public void GetBapEnvironments()
        {
            IEnumerable<string> primaryBap = PrimaryBapEnvironments(IncidIhsHabitat, IncidIhsMatrix1,
                IncidIhsMatrix2, IncidIhsMatrix3, IncidIhsFormation1, IncidIhsFormation2, IncidIhsManagement1,
                IncidIhsManagement2, IncidIhsComplex1, IncidIhsComplex2);

            IEnumerable<HluDataSet.incid_bapRow> incidBapRowsUndel = 
                _incidBapRows.Where(r => r.RowState != DataRowState.Deleted);

            if ((incidBapRowsUndel != null) && (primaryBap != null))
            {
                // primary BAP environments from DB (real bap_id) and new (bap_id = -1)
                IEnumerable<BapEnvironment> prevBapRowsAuto = null;
                IEnumerable<BapEnvironment> newBapRowsAuto = null;
                if (incidBapRowsUndel == null)
                {
                    prevBapRowsAuto = new BapEnvironment[0].AsEnumerable();
                    newBapRowsAuto = new BapEnvironment[0].AsEnumerable();
                }
                else
                {
                    prevBapRowsAuto = from r in incidBapRowsUndel
                                      join pot in primaryBap on r.bap_habitat equals pot
                                      where _incidCurrentRow.incid != null && r.incid == _incidCurrentRow.incid
                                      select new BapEnvironment(false, false, r);

                    newBapRowsAuto = from r in incidBapRowsUndel
                                     join pot in primaryBap on r.bap_habitat equals pot
                                     where prevBapRowsAuto.Count(p => p.bap_habitat == r.bap_habitat) == 0
                                     select new BapEnvironment(false, false, r);
                }

                var potBap = from p in primaryBap
                             where prevBapRowsAuto.Count(a => a.bap_habitat == p) == 0
                             where incidBapRowsUndel.Count(row => row.bap_habitat == p) == 0
                             select new BapEnvironment(false, false, -1, Incid, p, null, null, null);

                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>(
                    prevBapRowsAuto.Concat(newBapRowsAuto).Concat(potBap));
            }
            else if (incidBapRowsUndel != null)
            {
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                // no primary BAP, but user-added Db rows
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            }
            else if ((primaryBap != null) && (primaryBap.Count() > 0))
            {
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                // primary BAP, none in DB
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>(
                    primaryBap.Select(p => new BapEnvironment(false, false, -1, Incid, p, null, null, null)));
            }
            else
            {
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            }

            _incidBapRowsAuto.CollectionChanged += _incidBapRowsAuto_CollectionChanged;

            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Track when the BAP data has been changed so that the apply button
            // will appear.
            foreach (BapEnvironment be in _incidBapRowsAuto)
            {
                be.DataChanged += _incidBapRowsAuto_DataChanged;
            };
            //---------------------------------------------------------------------

            //---------------------------------------------------------------------
            // FIX: Show field errors on tab labels.
            // Check if there are any errors in the primary BAP records to see
            // if the Details tab label should be flagged as also in error.
            if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
            {
                int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    UpdateErrorList(ref _detailsErrors, "BapAuto", true);
                else
                    UpdateErrorList(ref _detailsErrors, "BapAuto", false);
            }
            else
                UpdateErrorList(ref _detailsErrors, "BapAuto", false);

            OnPropertyChanged("DetailsTabLabel");
            //---------------------------------------------------------------------

            OnPropertyChanged("IncidBapHabitatsAuto");

            if ((incidBapRowsUndel != null) && (_incidBapRowsAuto != null))
            {
                List<BapEnvironment> prevBapRowsUser = null;
                if (_incidBapRowsUser == null)
                {
                    prevBapRowsUser = new List<BapEnvironment>();
                }
                else
                {
                    prevBapRowsUser = (from r in _incidBapRowsUser
                                       where _incidCurrentRow.incid != null && r.incid == _incidCurrentRow.incid
                                       where _incidBapRowsAuto.Count(row => row.bap_habitat == r.bap_habitat) == 0
                                       select r).ToList();
                    prevBapRowsUser.ForEach(delegate(BapEnvironment be)
                    {
                        //---------------------------------------------------------------------
                        // FIX: Allow 'None' habitats to be managed
                        // Don't overwrite the determination quality value loaded from the
                        // database with 'Not present but close to definition' as other
                        // values may be valid and will be validated later.
                        //
                        //be.quality_determination = BapEnvironment.BAPDetQltyUserAdded;
                        //---------------------------------------------------------------------
                        be.BulkUpdateMode = _bulkUpdateMode == true;
                    });
                }

                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                _incidBapRowsUser = new ObservableCollection<BapEnvironment>(prevBapRowsUser.Concat(
                    from r in incidBapRowsUndel
                    where _incidBapRowsAuto.Count(a => a.bap_habitat == r.bap_habitat) == 0
                    where prevBapRowsUser.Count(p => p.bap_habitat == r.bap_habitat) == 0
                    select new BapEnvironment(_bulkUpdateMode == true, true, r)));
            }
            else if (incidBapRowsUndel != null)
            {
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                _incidBapRowsUser = new ObservableCollection<BapEnvironment>(
                   incidBapRowsUndel.Select(r => new BapEnvironment(_bulkUpdateMode == true, true, r)));
            }
            else
            {
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }
                //---------------------------------------------------------------------

                _incidBapRowsUser = new ObservableCollection<BapEnvironment>();
            }
            _incidBapRowsUser.CollectionChanged += _incidBapRowsUser_CollectionChanged;

            BapEnvironment.BapEnvironmentList = _incidBapRowsAuto.Concat(_incidBapRowsUser);

            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Track when the BAP data has been changed so that the apply button
            // will appear.
            foreach (BapEnvironment be in _incidBapRowsUser)
            {
                be.DataChanged += _incidBapRowsUser_DataChanged;
            };
            //---------------------------------------------------------------------

            //---------------------------------------------------------------------
            // FIX: Show field errors on tab labels.
            // Check if there are any errors in the secondary BAP records to see
            // if the Details tab label should be flagged as also in error.
            if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
            {
                int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    UpdateErrorList(ref _detailsErrors, "BapUser", true);
                else
                    UpdateErrorList(ref _detailsErrors, "BapUser", false);

                // Check if there are any duplicates between the primary and 
                // secondary BAP records.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                                            group be by be.bap_habitat into g
                                            where g.Count() > 1
                                            select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                    if (beDups.Length > 2)
                        UpdateErrorList(ref _detailsErrors, "BapUser", true);
                    else
                        UpdateErrorList(ref _detailsErrors, "BapUser", false);
                }
            }
            else
                UpdateErrorList(ref _detailsErrors, "BapUser", false);

            OnPropertyChanged("DetailsTabLabel");
            //---------------------------------------------------------------------

            OnPropertyChanged("IncidBapHabitatsUser");
        }

        //---------------------------------------------------------------------
        // CHANGED: CR2 (Apply button)
        // Track when the BAP primary records have changed so that the apply
        // button will appear.
        private void _incidBapRowsAuto_DataChanged(bool BapChanged)
        {
            Changed = true;

            //---------------------------------------------------------------------
            // FIX: Show field errors on tab labels.
            // Check if there are any errors in the primary BAP records to see
            // if the Details tab label should be flagged as also in error.
            if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
            {
                int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    UpdateErrorList(ref _detailsErrors, "BapAuto", true);
                else
                    UpdateErrorList(ref _detailsErrors, "BapAuto", false);
            }
            OnPropertyChanged("DetailsTabLabel");
            //---------------------------------------------------------------------
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR2 (Apply button)
        // Track when the BAP secondary records have changed so that the apply
        // button will appear.
        private void _incidBapRowsUser_DataChanged(bool BapChanged)
        {
            Changed = true;

            //---------------------------------------------------------------------
            // FIX: Show field errors on tab labels.
            // Check if there are any errors in the secondary BAP records to see
            // if the Details tab label should be flagged as also in error.
            if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
            {
                int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    UpdateErrorList(ref _detailsErrors, "BapUser", true);
                else
                    UpdateErrorList(ref _detailsErrors, "BapUser", false);

                // Check if there are any duplicates between the primary and 
                // secondary BAP records.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                                            group be by be.bap_habitat into g
                                            where g.Count() > 1
                                            select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                    if (beDups.Length > 2)
                        UpdateErrorList(ref _detailsErrors, "BapUser", true);
                    else
                        UpdateErrorList(ref _detailsErrors, "BapUser", false);
                }
            }
            OnPropertyChanged("DetailsTabLabel");
            //---------------------------------------------------------------------
        }
        //---------------------------------------------------------------------

        internal IEnumerable<string> PrimaryBapEnvironments(string ihsHabitat, string ihsMatrix1,
            string ihsMatrix2, string ihsMatrix3, string ihsFormation1, string ihsFormation2,
            string ihsManagement1, string ihsManagement2, string ihsComplex1, string ihsComplex2)
        {
            IEnumerable<string> primaryBap = null;

            if (!String.IsNullOrEmpty(ihsHabitat))
            {
                try
                {
                    primaryBap = from r in HluDataset.lut_ihs_habitat
                                   where r.code == ihsHabitat && !r.Iscode_bap_priority_habitatNull()
                                   select r.code_bap_priority_habitat;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsMatrix1))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_matrix
                            where r.code == ihsMatrix1 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsMatrix2))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_matrix
                            where r.code == ihsMatrix2 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsMatrix3))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_matrix
                            where r.code == ihsMatrix3 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsFormation1))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_formation
                            where r.code == ihsFormation1 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsFormation2))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_formation
                            where r.code == ihsFormation2 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsManagement1))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_management
                            where r.code == ihsManagement1 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }
            if (!String.IsNullOrEmpty(ihsManagement2))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_management
                            where r.code == ihsManagement2 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsComplex1))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_complex
                            where r.code == ihsComplex1 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (!String.IsNullOrEmpty(ihsComplex2))
            {
                try
                {
                    var q = from r in HluDataset.lut_ihs_complex
                            where r.code == ihsComplex2 && !r.Isbap_habitatNull()
                            select r.bap_habitat;
                    primaryBap = primaryBap != null ? primaryBap.Concat(q) : q;
                }
                catch { }
            }

            if (primaryBap != null)
                return primaryBap.Distinct();
            else
                return new string[0];
        }

        private void _incidBapRowsAuto_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged("Error");
        }

        private void _incidBapRowsUser_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //---------------------------------------------------------------------
            // FIXED: KI108 (Deleting potential BAP habitats)
            // Deleting the rows from the _incidBapRows datatable here causes
            // problems if the same row number is deleted twice as the row is
            // marked as deleted (RowState = Deleted) and hence the bap_id
            // cannot be read.  The rows are deleted later anyway when the
            // record is updated so they are left alone here.
            //
            // The user interface source for the potential BAP habtiats is
            // _incidBapRowsUser which is updated automatically when a row
            // is deleted so the row deleted automatically disappears in
            // the user interface.
            //
            //if (e.Action == NotifyCollectionChangedAction.Remove)
            //{
            //    (from r in _incidBapRows
            //     join be in e.OldItems.Cast<BapEnvironment>() on r.bap_id equals be.bap_id
            //     select r).ToList().ForEach(delegate(HluDataSet.incid_bapRow row) { row.Delete(); });
            //}
            //---------------------------------------------------------------------

            OnPropertyChanged("Error");
        }

        #endregion

        #region General Comments

        public string IncidGeneralComments
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.general_commentsColumn))
                    return IncidCurrentRow.general_comments.Trim();
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.general_comments = value.Trim();
                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                    //---------------------------------------------------------------------
                }
            }
        }

        #endregion

        #region Maps

        public DataView BoundaryMapCodes
        {
            get
            {
                if (HluDataset.lut_boundary_map.IsInitialized && HluDataset.lut_boundary_map.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_boundary_mapTableAdapter == null)
                        _hluTableAdapterMgr.lut_boundary_mapTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_boundary_mapDataTable, HluDataSet.lut_boundary_mapRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_boundary_mapDataTable) }, false);
                    HluDataset.lut_boundary_map.DefaultView.Sort = HluDataset.lut_boundary_map.sort_orderColumn.ColumnName;
                }
                return HluDataset.lut_boundary_map.DefaultView;
            }
        }

        public string IncidBoundaryBaseMap
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.boundary_base_mapColumn))
                    return IncidCurrentRow.boundary_base_map;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.boundary_base_map = value;
                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                    //---------------------------------------------------------------------
                }
            }
        }

        public string IncidDigitisationBaseMap
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.digitisation_base_mapColumn))
                    return IncidCurrentRow.digitisation_base_map;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.digitisation_base_map = value;
                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                    //---------------------------------------------------------------------
                }
            }
        }

        #endregion

        #region Biosite

        public string IncidSiteName
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.Issite_nameNull())
                    return IncidCurrentRow.site_name;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.site_name = value;
                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                    //---------------------------------------------------------------------
                }
            }
        }

        #endregion

        #endregion

        #region Sources Tab

        //---------------------------------------------------------------------
        // FIX: Show field errors on tab labels.
        // Set the Sources tab label from here so that validation can be done.
        // This will enable tooltips to be shown so that validation errors
        // in any fields in the tab can be highlighted by flagging the tab
        // label as in error.
        public string SourcesTabLabel
        {
            get { return "Sources"; }
        }
        //---------------------------------------------------------------------

        private bool CheckSources()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidSourcesRows == null)
            {
                HluDataSet.incid_sourcesDataTable incidSourcesTable = _hluDS.incid_sources;
                _incidSourcesRows = GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_sourcesTableAdapter, ref incidSourcesTable);
            }
            return _incidSourcesRows != null;
        }

        public HluDataSet.lut_sourcesRow[] Source1Names
        {
            get
            {
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                if (IncidSource1Id != null)
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return HluDataset.lut_sources.AsEnumerable().Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ToArray();
                }
                else
                {
                    return HluDataset.lut_sources.AsEnumerable().ToArray();
                }
            }
        }

        public Nullable<int> IncidSource1Id
        {
            get
            {
                if (!CheckSources()) return null;

                if (_incidSourcesRows.Length < 3)
                {
                    HluDataSet.incid_sourcesRow[] tmpRows = new HluDataSet.incid_sourcesRow[3 - _incidSourcesRows.Length];
                    _incidSourcesRows = _incidSourcesRows.Concat(tmpRows).ToArray();
                }

                if (_incidSourcesRows[0] != null)
                    return _incidSourcesRows[0].source_id;
                else
                    return null;
            }
            set
            {
                if (value == -1)
                {
                    // delete the row
                    UpdateIncidSourcesRow(0, IncidSourcesTable.source_idColumn.Ordinal, (Nullable<int>)null);

                    // refresh source names list
                    OnPropertyChanged("Source1Names");

                    // clear all fields of Source 1
                    IncidSource1Date = null;
                    IncidSource1HabitatClass = null;
                    IncidSource1HabitatType = null;
                    IncidSource1BoundaryImportance = null;
                    IncidSource1HabitatImportance = null;
                }
                else if (value != null)
                {
                    bool wasNull = _incidSourcesRows[0] == null;
                    UpdateIncidSourcesRow(0, IncidSourcesTable.source_idColumn.Ordinal, value);
                    IncidSource1Date = DefaultSourceDate(IncidSource1Date, IncidSource1Id);
                    // if row added refresh source names list
                    if (wasNull && (_incidSourcesRows[0] != null)) OnPropertyChanged("Source1Names");
                }
                OnPropertyChanged("IncidSource1Date");
                OnPropertyChanged("IncidSource1HabitatClass");
                OnPropertyChanged("IncidSource1HabitatType");
                OnPropertyChanged("IncidSource1BoundaryImportance");
                OnPropertyChanged("IncidSource1HabitatImportance");
                OnPropertyChanged("IncidSource1Enabled");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public bool IncidSource1Enabled
        {
            get { return (IncidSource1Id != null); }
        }

        public Date.VagueDateInstance IncidSource1Date
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 0) && (_incidSourcesRows[0] != null) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_date_startColumn) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_date_endColumn) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_date_typeColumn))
                {
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(_incidSourcesRows[0].source_date_start,
                        _incidSourcesRows[0].source_date_end, _incidSourcesRows[0].source_date_type,
                        _incidSource1DateEntered != null ? _incidSource1DateEntered.UserEntry : null);
                    return vd;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                UpdateIncidSourcesRow(0, IncidSourcesTable.source_date_startColumn.Ordinal, value);
                _incidSource1DateEntered = value;
                OnPropertyChanged("IncidSource1Date");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public HluDataSet.lut_habitat_classRow[] SourceHabitatClassCodes
        {
            get
            {
                if (_sourceHabitatClassCodes == null)
                {
                    if (HluDataset.lut_habitat_class.IsInitialized && (HluDataset.lut_habitat_class.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_habitat_classTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_classTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_classDataTable) }, false);
                    }
                    _sourceHabitatClassCodes = HluDataset.lut_habitat_class.OrderBy(r => r.sort_order).ToArray();
                }
                return _sourceHabitatClassCodes;
            }
        }

        public string IncidSource1HabitatClass
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 0) && (_incidSourcesRows[0] != null) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_habitat_classColumn))
                    return _incidSourcesRows[0].source_habitat_class;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(0, IncidSourcesTable.source_habitat_classColumn.Ordinal, value);
                OnPropertyChanged("Source1HabitatTypeCodes");
                OnPropertyChanged("IncidSource1HabitatType");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public HluDataSet.lut_habitat_typeRow[] Source1HabitatTypeCodes
        {
            get
            {
                if (!String.IsNullOrEmpty(IncidSource1HabitatClass))
                {
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource1HabitatClass)
                        .OrderBy(r => r.sort_order).ToArray();

                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Don't pre-populate the habitat type value if there is only one
                    // possible value (as it triggers the data changed flag.
                    //
                    //if ((retArray.Length == 1) && (IncidSource1Id != null))
                    //{
                    //    IncidSource1HabitatType = retArray[0].code;
                    //    OnPropertyChanged("IncidSource1HabitatType");
                    //}
                    //---------------------------------------------------------------------

                    return retArray;
                }
                return null;
            }
        }

        public string IncidSource1HabitatType
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 0) && (_incidSourcesRows[0] != null) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_habitat_typeColumn))
                    return _incidSourcesRows[0].source_habitat_type;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(0, IncidSourcesTable.source_habitat_typeColumn.Ordinal, value);
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public HluDataSet.lut_importanceRow[] SourceImportanceCodes
        {
            get
            {
                if (_sourceImportanceCodes == null)
                {
                    if (HluDataset.lut_importance.IsInitialized && (HluDataset.lut_importance.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_importanceTableAdapter == null)
                            _hluTableAdapterMgr.lut_importanceTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_importanceDataTable, HluDataSet.lut_importanceRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_importanceDataTable) }, false);
                    }
                    _sourceImportanceCodes = HluDataset.lut_importance.OrderBy(r => r.sort_order).ToArray();
                }
                return _sourceImportanceCodes;
            }
        }

        public string IncidSource1BoundaryImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 0) && (_incidSourcesRows[0] != null) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_boundary_importanceColumn))
                    return _incidSourcesRows[0].source_boundary_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(0, IncidSourcesTable.source_boundary_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource1BoundaryImportance");
                OnPropertyChanged("IncidSource2BoundaryImportance");
                OnPropertyChanged("IncidSource3BoundaryImportance");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public string IncidSource1HabitatImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 0) && (_incidSourcesRows[0] != null) &&
                    !_incidSourcesRows[0].IsNull(HluDataset.incid_sources.source_habitat_importanceColumn))
                    return _incidSourcesRows[0].source_habitat_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(0, IncidSourcesTable.source_habitat_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource1HabitatImportance");
                OnPropertyChanged("IncidSource2HabitatImportance");
                OnPropertyChanged("IncidSource3HabitatImportance");
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
                //---------------------------------------------------------------------
            }
        }

        public HluDataSet.lut_sourcesRow[] Source2Names
        {
            get
            {
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                if (IncidSource2Id != null)
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return HluDataset.lut_sources.AsEnumerable().Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ToArray();
                }
                else
                {
                    return HluDataset.lut_sources.AsEnumerable().ToArray();
                }
            }
        }

        public Nullable<int> IncidSource2Id
        {
            get
            {
                if (!CheckSources()) return null;

                if (_incidSourcesRows.Length < 3)
                {
                    HluDataSet.incid_sourcesRow[] tmpRows = new HluDataSet.incid_sourcesRow[3 - _incidSourcesRows.Length];
                    _incidSourcesRows = _incidSourcesRows.Concat(tmpRows).ToArray();
                }
                if (_incidSourcesRows[1] != null)
                    return _incidSourcesRows[1].source_id;
                else
                    return null;
            }
            set
            {
                if (value == -1)
                {
                    // delete the row
                    UpdateIncidSourcesRow(1, IncidSourcesTable.source_idColumn.Ordinal, (Nullable<int>)null);

                    // refresh source names list
                    OnPropertyChanged("Source2Names");

                    // clear all fields of Source 2
                    IncidSource2Date = null;
                    IncidSource2HabitatClass = null;
                    IncidSource2HabitatType = null;
                    IncidSource2BoundaryImportance = null;
                    IncidSource2HabitatImportance = null;
                }
                else if (value != null)
                {
                    bool wasNull = _incidSourcesRows[1] == null;
                    UpdateIncidSourcesRow(1, IncidSourcesTable.source_idColumn.Ordinal, value);
                    IncidSource2Date = DefaultSourceDate(IncidSource2Date, IncidSource2Id);
                    // if row added refresh source names list
                    if (wasNull && (_incidSourcesRows[1] != null)) OnPropertyChanged("Source2Names");
                }
                OnPropertyChanged("IncidSource2Date");
                OnPropertyChanged("IncidSource2HabitatClass");
                OnPropertyChanged("IncidSource2HabitatType");
                OnPropertyChanged("IncidSource2BoundaryImportance");
                OnPropertyChanged("IncidSource2HabitatImportance");
                OnPropertyChanged("IncidSource2Enabled");
            }
        }

        public bool IncidSource2Enabled
        {
            get { return (IncidSource2Id != null); }
        }

        public Date.VagueDateInstance IncidSource2Date
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 1) && (_incidSourcesRows[1] != null) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_date_startColumn) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_date_endColumn) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_date_typeColumn))
                {
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(_incidSourcesRows[1].source_date_start,
                        _incidSourcesRows[1].source_date_end, _incidSourcesRows[1].source_date_type,
                        _incidSource2DateEntered != null ? _incidSource2DateEntered.UserEntry : null);
                    return vd;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                UpdateIncidSourcesRow(1, IncidSourcesTable.source_date_startColumn.Ordinal, value);
                _incidSource2DateEntered = value;
                OnPropertyChanged("IncidSource2Date");
            }
        }

        public string IncidSource2HabitatClass
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 1) && (_incidSourcesRows[1] != null) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_habitat_classColumn))
                    return _incidSourcesRows[1].source_habitat_class;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(1, IncidSourcesTable.source_habitat_classColumn.Ordinal, value);
                OnPropertyChanged("Source2HabitatTypeCodes");
                OnPropertyChanged("IncidSource2HabitatType");
            }
        }

        public DataRow[] Source2HabitatTypeCodes
        {
            get
            {
                if (!String.IsNullOrEmpty(IncidSource2HabitatClass))
                {
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource2HabitatClass)
                        .OrderBy(r => r.sort_order).ToArray();

                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Don't pre-populate the habitat type value if there is only one
                    // possible value (as it triggers the data changed flag.
                    //
                    //if ((retArray.Length == 1) && (IncidSource2Id != null))
                    //{
                    //    IncidSource2HabitatType = retArray[0].code;
                    //    OnPropertyChanged("IncidSource2HabitatType");
                    //}
                    //---------------------------------------------------------------------

                    return retArray;
                }
                return null;
            }
        }

        public string IncidSource2HabitatType
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 1) && (_incidSourcesRows[1] != null) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_habitat_typeColumn))
                    return _incidSourcesRows[1].source_habitat_type;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(1, IncidSourcesTable.source_habitat_typeColumn.Ordinal, value);
            }
        }

        public string IncidSource2BoundaryImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 1) && (_incidSourcesRows[1] != null) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_boundary_importanceColumn))
                    return _incidSourcesRows[1].source_boundary_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(1, IncidSourcesTable.source_boundary_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource2BoundaryImportance");
                OnPropertyChanged("IncidSource1BoundaryImportance");
                OnPropertyChanged("IncidSource3BoundaryImportance");
            }
        }

        public string IncidSource2HabitatImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 1) && (_incidSourcesRows[1] != null) &&
                    !_incidSourcesRows[1].IsNull(HluDataset.incid_sources.source_habitat_importanceColumn))
                    return _incidSourcesRows[1].source_habitat_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(1, IncidSourcesTable.source_habitat_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource2HabitatImportance");
                OnPropertyChanged("IncidSource1HabitatImportance");
                OnPropertyChanged("IncidSource3HabitatImportance");
            }
        }

        public HluDataSet.lut_sourcesRow[] Source3Names
        {
            get
            {
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                if (IncidSource3Id != null)
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return HluDataset.lut_sources.AsEnumerable().Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ToArray();
                }
                else
                {
                    return HluDataset.lut_sources.AsEnumerable().ToArray();
                }
            }
        }

        public Nullable<int> IncidSource3Id
        {
            get
            {
                if (!CheckSources()) return null;

                if (_incidSourcesRows.Length < 3)
                {
                    HluDataSet.incid_sourcesRow[] tmpRows = new HluDataSet.incid_sourcesRow[3 - _incidSourcesRows.Length];
                    _incidSourcesRows = _incidSourcesRows.Concat(tmpRows).ToArray();
                }
                if (_incidSourcesRows[2] != null)
                    return _incidSourcesRows[2].source_id;
                else
                    return null;
            }
            set
            {
                if (value == -1)
                {
                    // delete the row
                    UpdateIncidSourcesRow(2, IncidSourcesTable.source_idColumn.Ordinal, (Nullable<int>)null);

                    // refresh source names lists (all three)
                    OnPropertyChanged("Source3Names");

                    // clear all fields of Source 3
                    IncidSource3Date = null;
                    IncidSource3HabitatClass = null;
                    IncidSource3HabitatType = null;
                    IncidSource3BoundaryImportance = null;
                    IncidSource3HabitatImportance = null;
                }
                else if (value != null)
                {
                    bool wasNull = _incidSourcesRows[2] == null;
                    UpdateIncidSourcesRow(2, IncidSourcesTable.source_idColumn.Ordinal, value);
                    IncidSource3Date = DefaultSourceDate(IncidSource3Date, IncidSource3Id);
                    // if row added refresh source names lists (all three)
                    if (wasNull && (_incidSourcesRows[2] != null)) OnPropertyChanged("Source3Names");
                }
                OnPropertyChanged("IncidSource3Date");
                OnPropertyChanged("IncidSource3HabitatClass");
                OnPropertyChanged("IncidSource3HabitatType");
                OnPropertyChanged("IncidSource3BoundaryImportance");
                OnPropertyChanged("IncidSource3HabitatImportance");
                OnPropertyChanged("IncidSource3Enabled");
            }
        }

        public bool IncidSource3Enabled
        {
            get { return (IncidSource3Id != null); }
        }

        public Date.VagueDateInstance IncidSource3Date
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 2) && (_incidSourcesRows[2] != null) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_date_startColumn) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_date_endColumn) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_date_typeColumn))
                {
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(_incidSourcesRows[2].source_date_start,
                        _incidSourcesRows[2].source_date_end, _incidSourcesRows[2].source_date_type, 
                        _incidSource3DateEntered != null ? _incidSource3DateEntered.UserEntry : null);
                    return vd;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                UpdateIncidSourcesRow(2, IncidSourcesTable.source_date_startColumn.Ordinal, value);
                _incidSource3DateEntered = value;
                OnPropertyChanged("IncidSource3Date");
            }
        }

        public string IncidSource3HabitatClass
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 2) && (_incidSourcesRows[2] != null) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_habitat_classColumn))
                    return _incidSourcesRows[2].source_habitat_class;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(2, IncidSourcesTable.source_habitat_classColumn.Ordinal, value);
                OnPropertyChanged("Source3HabitatTypeCodes");
                OnPropertyChanged("IncidSource3HabitatType");
            }
        }

        public string IncidSource3HabitatType
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 2) && (_incidSourcesRows[2] != null) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_habitat_typeColumn))
                    return _incidSourcesRows[2].source_habitat_type;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(2, IncidSourcesTable.source_habitat_typeColumn.Ordinal, value);
            }
        }

        public DataRow[] Source3HabitatTypeCodes
        {
            get
            {
                if (!String.IsNullOrEmpty(IncidSource3HabitatClass))
                {
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource3HabitatClass)
                        .OrderBy(r => r.sort_order).ToArray();

                    //---------------------------------------------------------------------
                    // CHANGED: CR2 (Apply button)
                    // Don't pre-populate the habitat type value if there is only one
                    // possible value (as it triggers the data changed flag.
                    //
                    //if ((retArray.Length == 1) && (IncidSource3Id != null))
                    //{
                    //    IncidSource3HabitatType = retArray[0].code;
                    //    OnPropertyChanged("IncidSource3HabitatType");
                    //}
                    //---------------------------------------------------------------------

                    return retArray;
                }
                return null;
            }
        }

        public string IncidSource3BoundaryImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 2) && (_incidSourcesRows[2] != null) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_boundary_importanceColumn))
                    return _incidSourcesRows[2].source_boundary_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(2, IncidSourcesTable.source_boundary_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource3BoundaryImportance");
                OnPropertyChanged("IncidSource1BoundaryImportance");
                OnPropertyChanged("IncidSource2BoundaryImportance");
            }
        }

        public string IncidSource3HabitatImportance
        {
            get
            {
                if (!CheckSources()) return null;
                if ((_incidSourcesRows.Length > 2) && (_incidSourcesRows[2] != null) &&
                    !_incidSourcesRows[2].IsNull(HluDataset.incid_sources.source_habitat_importanceColumn))
                    return _incidSourcesRows[2].source_habitat_importance;
                else
                    return null;
            }
            set
            {
                UpdateIncidSourcesRow(2, IncidSourcesTable.source_habitat_importanceColumn.Ordinal, value);
                OnPropertyChanged("IncidSource3HabitatImportance");
                OnPropertyChanged("IncidSource1HabitatImportance");
                OnPropertyChanged("IncidSource2HabitatImportance");
            }
        }

        private Date.VagueDateInstance DefaultSourceDate(Date.VagueDateInstance currentDate, Nullable<int> sourceID)
        {
            if ((HluDataset == null) || (HluDataset.lut_sources == null)) return currentDate;

            EnumerableRowCollection<HluDataSet.lut_sourcesRow> rows =
                HluDataset.lut_sources.Where(r => r.source_id == sourceID &&
                    !r.IsNull(HluDataset.lut_sources.source_date_defaultColumn));

            if (rows.Count() > 0)
            {
                string defaultDate;
                string dateType = VagueDate.GetType(rows.ElementAt(0).source_date_default, out defaultDate);
                int startDate = VagueDate.ToTimeSpanDays(defaultDate, dateType, VagueDate.DateType.Start);
                int endDate = VagueDate.ToTimeSpanDays(defaultDate, dateType, VagueDate.DateType.End);
                return new Date.VagueDateInstance(startDate, endDate, dateType);
            }

            return currentDate;
        }

        private void UpdateIncidSourcesRow<T>(int rowNumber, int columnOrdinal, T newValue)
        {
            try
            {
                if (_incidSourcesRows == null) return;

                if (_incidSourcesRows[rowNumber] == null)
                {
                    if (columnOrdinal == HluDataset.incid_sources.source_idColumn.Ordinal)
                    {
                        HluDataSet.incid_sourcesRow newRow = IncidSourcesTable.Newincid_sourcesRow();
                        newRow.incid_source_id = NextIncidSourceId;
                        newRow.incid = IncidCurrentRow.incid;
                        newRow.sort_order = rowNumber + 1;
                        _incidSourcesRows[rowNumber] = newRow;
                    }
                    else
                    {
                        return;
                    }
                }
                else if ((columnOrdinal == HluDataset.incid_sources.source_idColumn.Ordinal) && (newValue == null))
                {
                    if (_incidSourcesRows[rowNumber].RowState != DataRowState.Detached)
                        _incidSourcesRows[rowNumber].Delete();
                    _incidSourcesRows[rowNumber] = null;
                    return;
                }

                if ((columnOrdinal == HluDataset.incid_sources.source_date_startColumn.Ordinal) ||
                    (columnOrdinal == HluDataset.incid_sources.source_date_endColumn.Ordinal))
                {
                    Date.VagueDateInstance vd = newValue as Date.VagueDateInstance;
                    if (vd != null)
                    {
                        _incidSourcesRows[rowNumber].source_date_start = vd.StartDate;
                        _incidSourcesRows[rowNumber].source_date_end = vd.EndDate;
                        _incidSourcesRows[rowNumber].source_date_type = vd.DateType;
                    }
                    else
                    {
                        _incidSourcesRows[rowNumber].source_date_start = VagueDate.DateUnknown;
                        _incidSourcesRows[rowNumber].source_date_end = VagueDate.DateUnknown;
                        _incidSourcesRows[rowNumber].source_date_type = null;
                    }
                }
                else if ((((_incidSourcesRows[rowNumber].IsNull(columnOrdinal) ^ (newValue == null)) || 
                    ((!_incidSourcesRows[rowNumber].IsNull(columnOrdinal) && (newValue != null)))) && 
                    !_incidSourcesRows[rowNumber][columnOrdinal].Equals(newValue)))
                {
                    _incidSourcesRows[rowNumber][columnOrdinal] = newValue;
                }

                if (columnOrdinal == HluDataset.incid_sources.source_idColumn.Ordinal)
                {
                    try
                    {
                        HluDataSet.lut_sourcesRow lutRow =
                            HluDataset.lut_sources.Single(r => r.source_id == _incidSourcesRows[rowNumber].source_id);
                        if (!String.IsNullOrEmpty(lutRow.source_date_default))
                        {
                            string defaultDateString;
                            string formatString = VagueDate.GetType(lutRow.source_date_default, out defaultDateString);
                            int defaultStartDate = VagueDate.ToTimeSpanDays(defaultDateString,
                                formatString, VagueDate.DateType.Start);
                            int defaultEndDate = VagueDate.ToTimeSpanDays(defaultDateString,
                                formatString, VagueDate.DateType.End);
                            _incidSourcesRows[rowNumber].source_date_start = defaultStartDate;
                            _incidSourcesRows[rowNumber].source_date_end = defaultEndDate;
                        }
                    }
                    catch { }
                }

                if ((_incidSourcesRows[rowNumber].RowState == DataRowState.Detached) &&
                    IsCompleteRow(_incidSourcesRows[rowNumber]))
                {
                    _incidSourcesRows[rowNumber].sort_order = rowNumber + 1;
                    IncidSourcesTable.Addincid_sourcesRow(_incidSourcesRows[rowNumber]);
                }
            }
            catch { }
        }

        #endregion

        # region History Tab

        public IEnumerable<string> IncidHistory
        {
            get
            {
                if (_incidHistoryRows == null)
                    return null;
                else
                {
                    //---------------------------------------------------------------------
                    // FIX: Always save all (both) of the history columns
                    // Figure out which history columns to display based on the user options
                    // now that all the available history columns are always update when
                    // creating history even if the user only wants to display some of them.
                    DataColumn[] displayHistoryColumns;
                    int result;
                    displayHistoryColumns = _gisIDColumns.Concat((from s in Settings.Default.HistoryColumnOrdinals.Cast<string>()
                                             where Int32.TryParse(s, out result) && (result >= 0) &&
                                                  (result < _hluDS.incid_mm_polygons.Columns.Count) &&
                                                  !_gisIDColumnOrdinals.Contains(result)
                                             select _hluDS.incid_mm_polygons.Columns[Int32.Parse(s)])).ToArray();
                    //---------------------------------------------------------------------

                    return (from r in _incidHistoryRows.OrderByDescending(r => r.history_id)
                            group r by new
                            {
                                modified_date = !r.Ismodified_dateNull() ?
                                    r.modified_date.ToShortDateString() : String.Empty,
                                //---------------------------------------------------------------------
                                // CHANGED: CR4 (Modified date)
                                // Display the modified_date column from the history wth both the
                                // date and time to avoid separate updates with identical details
                                // (except the time) being merged together when displayed.
                                modified_time = (!r.Ismodified_dateNull() && r.modified_date != r.modified_date.Date) ?
                                    @" at " + r.modified_date.ToLongTimeString() : String.Empty,
                                modified_user_id = r.lut_userRow != null ? r.lut_userRow.user_name :
                                    !r.Ismodified_user_idNull() ? r.modified_user_id : String.Empty,
                                modifid_process = r.lut_processRow != null ? r.lut_processRow.description : String.Empty,
                                modified_reason = r.lut_reasonRow != null ? r.lut_reasonRow.description : String.Empty,
                                modified_operation = r.lut_operationRow != null ? r.lut_operationRow.description : String.Empty,
                                modified_incid = !r.Ismodified_incidNull() ? r.modified_incid : String.Empty,
                                //---------------------------------------------------------------------
                                // FIX: Always save all (both) of the history columns
                                modified_ihs = r.Table.Columns.Cast<DataColumn>().Where(rc =>
                                    displayHistoryColumns.Count(hc => "modified_" + hc.ColumnName == rc.ColumnName) == 1 &&
                                    _gisIDColumns.Count(gc => "modified_" + gc.ColumnName == rc.ColumnName) == 0)
                                    .Aggregate(new StringBuilder(), (sb, hr) => sb.Append(String.Format("\n\t{0}: {1}",
                                    hr.ColumnName.Replace("ihs_", "IHS ").Replace("modified_", "").Replace("_", " "),
                                    r[hr.ColumnName].ToString()))).ToString()
                                //---------------------------------------------------------------------
                            } into g
                            select String.Format("Modified on {0}{1} by {2}:", g.Key.modified_date,
                                g.Key.modified_time, g.Key.modified_user_id) +
                            //---------------------------------------------------------------------
                                String.Format("\n\tProcess: {0}", g.Key.modifid_process) +
                                String.Format("\n\tReason: {0}", g.Key.modified_reason) +
                                String.Format("\n\tOperation: {0}", g.Key.modified_operation) +
                                g.Key.modified_ihs +
                                String.Format("\n\tINCID: {0}", g.Key.modified_incid) +
                                String.Format("\n\tLength Modified: {0}", g.Distinct(_histRowEqComp)
                                    .Sum(r => !r.Ismodified_lengthNull() ? r.modified_length : 0).ToString("f2")) +
                                String.Format("\n\tArea Modified: {0}", g.Distinct(_histRowEqComp)
                                    .Sum(r => !r.Ismodified_areaNull() ? r.modified_area : 0).ToString("f2")))
                                .Take(_historyDisplayLastN);
                }
            }
        }

        #endregion

        #region Record IDs

        public GeometryTypes GisLayerType { get { return _gisLayerType; } }

        public string SiteID { get { return _recIDs.SiteID; } }

        public string IhsVersion { get { return _recIDs.IhsVersion; } }

        public string CurrentIncid { get { return _recIDs.CurrentIncid; } }

        public string NextIncid { get { return _recIDs.NextIncid; } }

        private int NextIncidIhsMatrixId { get { return _recIDs.NextIncidIhsMatrixId; } }

        private int NextIncidIhsFormationId { get { return _recIDs.NextIncidIhsFormationId; } }

        private int NextIncidIhsManagementId { get { return _recIDs.NextIncidIhsManagementId; } }

        private int NextIncidIhsComplexId { get { return _recIDs.NextIncidIhsComplexId; } }

        private int CurrentIncidBapId { get { return _recIDs.CurrentIncidBapId; } }

        private int NextIncidBapId { get { return _recIDs.NextIncidBapId; } }

        private int NextIncidSourceId { get { return _recIDs.NextIncidSourceId; } }

        #endregion

        #region Validation

        internal bool HaveGisApp
        {
            get { return _gisApp != null && _gisApp.IsRunning; }
        }

        internal bool IsCompleteRow(DataRow r)
        {
            if (r == null) return false;

            foreach (DataColumn c in r.Table.Columns)
            {
                if (!c.AllowDBNull && r.IsNull(c)) return false;
            }

            return true;
        }

        private void FixIhsMatrixCodes()
        {
            if (_bulkUpdateMode == true) return;

            // check for null values
            if (String.IsNullOrEmpty(IncidIhsMatrix1))
            {
                if (String.IsNullOrEmpty(IncidIhsMatrix2))
                {
                    if (!String.IsNullOrEmpty(IncidIhsMatrix3))
                    {
                        string bak3 = IncidIhsMatrix3;
                        IncidIhsMatrix3 = null;
                        IncidIhsMatrix1 = bak3;
                    }
                }
                else
                {
                    if (String.IsNullOrEmpty(IncidIhsMatrix3))
                    {
                        string bak2 = IncidIhsMatrix2;
                        IncidIhsMatrix2 = null;
                        IncidIhsMatrix1 = bak2;
                    }
                    else
                    {
                        string bak3 = IncidIhsMatrix3;
                        string bak2 = IncidIhsMatrix2;
                        IncidIhsMatrix3 = null;
                        IncidIhsMatrix2 = bak3;
                        IncidIhsMatrix1 = bak2;
                    }
                }
            }
            else if (String.IsNullOrEmpty(IncidIhsMatrix2) && !String.IsNullOrEmpty(IncidIhsMatrix3))
            {
                string bak3 = IncidIhsMatrix3;
                IncidIhsMatrix3 = null;
                IncidIhsMatrix2 = bak3;
            }

            // check for duplicate values
            if ((IncidIhsMatrix2 != null) && (IncidIhsMatrix1 == IncidIhsMatrix2))
            {
                if (String.IsNullOrEmpty(IncidIhsMatrix3))
                {
                    IncidIhsMatrix2 = null;
                }
                else
                {
                    string bak3 = IncidIhsMatrix3;
                    IncidIhsMatrix3 = null;
                    IncidIhsMatrix2 = bak3;
                }
            }
            else if ((IncidIhsMatrix3 != null) &&
                ((IncidIhsMatrix1 == IncidIhsMatrix3) || (IncidIhsMatrix2 == IncidIhsMatrix3)))
            {
                IncidIhsMatrix3 = null;
            }
        }

        private void FixIhsFormationCodes()
        {
            // check for null values
            if (String.IsNullOrEmpty(IncidIhsFormation1) && !String.IsNullOrEmpty(IncidIhsFormation2))
            {
                string bak2 = IncidIhsFormation2;
                IncidIhsFormation2 = null;
                IncidIhsFormation1 = bak2;
            }

            // check for duplicate values
            if ((IncidIhsFormation1 == IncidIhsFormation2) && (IncidIhsFormation2 != null))
                IncidIhsFormation2 = null;
        }

        private void FixIhsManagementCodes()
        {
            // check for null values
            if (String.IsNullOrEmpty(IncidIhsManagement1) && !String.IsNullOrEmpty(IncidIhsManagement2))
            {
                string bak2 = IncidIhsManagement2;
                IncidIhsManagement2 = null;
                IncidIhsManagement1 = bak2;
            }

            // check for duplicate values
            if ((IncidIhsManagement1 == IncidIhsManagement2) && (IncidIhsManagement2 != null))
                IncidIhsManagement2 = null;
        }

        private void FixIhsComplexCodes()
        {
            // check for null values
            if (String.IsNullOrEmpty(IncidIhsComplex1) && !String.IsNullOrEmpty(IncidIhsComplex2))
            {
                string bak2 = IncidIhsComplex2;
                IncidIhsComplex2 = null;
                IncidIhsComplex1 = bak2;
            }

            // check for duplicate values
            if ((IncidIhsComplex1 == IncidIhsComplex2) && (IncidIhsComplex2 != null))
                IncidIhsComplex2 = null;
        }

        private string ValidateIhsFormation1Code()
        {
            if ((_bulkUpdateMode == true) || (_incidCurrentRow == null) ||
                String.IsNullOrEmpty(IncidIhsHabitat) || (IhsHabitatCodes == null)) 
                return null;

            var q = IhsHabitatCodes.Where(r => !String.IsNullOrEmpty(IncidIhsHabitat) && r.code == IncidIhsHabitat &&
                !r.IsNull(HluDataset.lut_ihs_habitat.man_formationColumn));

            if ((q.Count() > 0) && (q.ElementAt(0).man_formation == 1) && String.IsNullOrEmpty(IncidIhsFormation1))
                return String.Format("IHS formation code is mandatory for IHS habitat '{0} : {1}'", q.ElementAt(0).code,
                    q.ElementAt(0).description);
            else
                return null;
        }

        private string ValidateIhsManagement1Code()
        {
            if ((_bulkUpdateMode == true)|| (_incidCurrentRow == null) ||
                String.IsNullOrEmpty(IncidIhsHabitat) || (IhsHabitatCodes == null))
                return null;

            var q = IhsHabitatCodes.Where(r => !String.IsNullOrEmpty(IncidIhsHabitat) && r.code == IncidIhsHabitat &&
                !r.IsNull(HluDataset.lut_ihs_habitat.man_managementColumn));

            if ((q.Count() > 0) && (q.ElementAt(0).man_management == 1) && String.IsNullOrEmpty(IncidIhsManagement1))
                return String.Format("IHS management code is mandatory for IHS habitat '{0} : {1}'", q.ElementAt(0).code,
                    q.ElementAt(0).description);
            else
                return null;
        }

        private List<string[]> ValidateSource1()
        {
            List<string[]> errors = new List<string[]>();

            if (IncidSource1Id != null)
            {
                if (IncidSource1Date == null)
                    errors.Add(new string[] { "IncidSource1Date", "Date is mandatory for each source" });
                else if (IncidSource1Date.IsBad)
                    errors.Add(new string[] { "IncidSource1Date", "Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource1HabitatClass))
                    errors.Add(new string[] { "IncidSource1HabitatClass", "Habitat class is mandatory for each source" });
                else if ((IncidSource1HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource1HabitatType))
                    errors.Add(new string[] { "IncidSource1HabitatType", "Habitat type is mandatory if habitat class is filled in" });

                //---------------------------------------------------------------------
                // FIX: Use skip value from settings rather than hard-coded value
                string skipVal = Settings.Default.SourceImportanceSkip;
                //---------------------------------------------------------------------
                if (String.IsNullOrEmpty(IncidSource1BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource1BoundaryImportance", "Boundary importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+BoundaryImportance", @"\d+", "1", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+BoundaryImportance", @"\d+", "1", errors);
                    //---------------------------------------------------------------------
                }
                if (String.IsNullOrEmpty(IncidSource1HabitatImportance))
                {
                    errors.Add(new string[] { "IncidSource1HabitatImportance", 
                        "Habitat importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+HabitatImportance", @"\d+", "1", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+HabitatImportance", @"\d+", "1", errors);
                    //---------------------------------------------------------------------
                }
            }
            else
            {
                if (IncidSource1Date != null)
                    errors.Add(new string[] { "IncidSource1Date", 
                        "Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatClass))
                    errors.Add(new string[] { "IncidSource1HabitatClass", 
                        "Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatType))
                    errors.Add(new string[] { "IncidSource1HabitatType", 
                        "Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1BoundaryImportance))
                    errors.Add(new string[] { "IncidSource1BoundaryImportance", 
                        "Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatImportance))
                    errors.Add(new string[] { "IncidSource1HabitatImportance", 
                        "Habitat importance cannot be filled in if no source has been specified" });
            }

            return errors;
        }

        private List<string[]> ValidateSource2()
        {
            List<string[]> errors = new List<string[]>();

            if (IncidSource2Id != null)
            {
                if (IncidSource2Date == null)
                    errors.Add(new string[] { "IncidSource2Date", "Date is mandatory for each source" });
                else if (IncidSource2Date.IsBad)
                    errors.Add(new string[] { "IncidSource2Date", "Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource2HabitatClass))
                    errors.Add(new string[] { "IncidSource2HabitatClass", "Habitat class is mandatory for each source" });
                else if ((IncidSource2HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource2HabitatType))
                    errors.Add(new string[] { "IncidSource2HabitatType", "Habitat type is mandatory if habitat class is filled in" });

                //---------------------------------------------------------------------
                // FIX: Use skip value from settings rather than hard-coded value
                string skipVal = Settings.Default.SourceImportanceSkip;
                //---------------------------------------------------------------------
                if (String.IsNullOrEmpty(IncidSource2BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource2BoundaryImportance", "Boundary importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+BoundaryImportance", @"\d+", "2", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+BoundaryImportance", @"\d+", "2", errors);
                    //---------------------------------------------------------------------
                }
                if (String.IsNullOrEmpty(IncidSource2HabitatImportance))
                {
                    errors.Add(new string[] { "IncidSource2HabitatImportance", 
                        "Habitat importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+HabitatImportance", @"\d+", "2", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+HabitatImportance", @"\d+", "2", errors);
                    //---------------------------------------------------------------------
                }
            }
            else
            {
                if (IncidSource2Date != null)
                    errors.Add(new string[] { "IncidSource2Date", 
                        "Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatClass))
                    errors.Add(new string[] { "IncidSource2HabitatClass", 
                        "Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatType))
                    errors.Add(new string[] { "IncidSource2HabitatType", 
                        "Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2BoundaryImportance))
                    errors.Add(new string[] { "IncidSource2BoundaryImportance", 
                        "Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatImportance))
                    errors.Add(new string[] { "IncidSource2HabitatImportance", 
                        "Habitat importance cannot be filled in if no source has been specified" });
            }

            return errors;
        }

        private List<string[]> ValidateSource3()
        {
            List<string[]> errors = new List<string[]>();

            if (IncidSource3Id != null)
            {
                if (IncidSource3Date == null)
                    errors.Add(new string[] { "IncidSource3Date", "Date is mandatory for each source" });
                else if (IncidSource3Date.IsBad)
                    errors.Add(new string[] { "IncidSource3Date", "Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource3HabitatClass))
                    errors.Add(new string[] { "IncidSource3HabitatClass", "Habitat class is mandatory for each source" });
                else if ((IncidSource3HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource3HabitatType))
                    errors.Add(new string[] { "IncidSource3HabitatType", "Habitat type is mandatory if habitat class is filled in" });

                //---------------------------------------------------------------------
                // FIX: Use skip value from settings rather than hard-coded value
                string skipVal = Settings.Default.SourceImportanceSkip;
                //---------------------------------------------------------------------
                if (String.IsNullOrEmpty(IncidSource3BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource3BoundaryImportance", "Boundary importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+BoundaryImportance", @"\d+", "3", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+BoundaryImportance", @"\d+", "3", errors);
                    //---------------------------------------------------------------------
                }
                if (String.IsNullOrEmpty(IncidSource3HabitatImportance))
                {
                    errors.Add(new string[] { "IncidSource3HabitatImportance", 
                        "Habitat importance is mandatory for each source" });
                }
                else
                {
                    ValidateSourceDuplicates(@"IncidSource\d+HabitatImportance", @"\d+", "3", skipVal, errors);
                    //---------------------------------------------------------------------
                    // CHANGED: CR1 (Boundary and Habitat Importance)
                    /// Validates the source importances by ensuring that boundary and habitat importance
                    /// values are applied in order (as specified in the settings).
                    ValidateSourceImportances(@"IncidSource\d+HabitatImportance", @"\d+", "3", errors);
                    //---------------------------------------------------------------------
                }
            }
            else
            {
                if (IncidSource3Date != null)
                    errors.Add(new string[] { "IncidSource3Date", 
                        "Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatClass))
                    errors.Add(new string[] { "IncidSource3HabitatClass", 
                        "Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatType))
                    errors.Add(new string[] { "IncidSource3HabitatType", 
                        "Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3BoundaryImportance))
                    errors.Add(new string[] { "IncidSource3BoundaryImportance", 
                        "Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatImportance))
                    errors.Add(new string[] { "IncidSource3HabitatImportance", 
                        "Habitat importance cannot be filled in if no source has been specified" });
            }

            return errors;
        }

        //---------------------------------------------------------------------
        // CHANGED: CR1 (Boundary and Habitat Importance)
        //
        /// <summary>
        /// Validates the source importances by ensuring that boundary and habitat importance
        /// values are applied in order (as specified in the settings).
        /// </summary>
        /// <param name="propNamePat">Pattern for regular expression match of property name in current class.</param>
        /// <param name="propNamePatWildcard">Wildcard element in propNamePat.</param>
        /// <param name="propNameWildcardValCurrProp">Value of propNamePatWildcard for current property to be validated.</param>
        /// <param name="errors">List of errors, composed of name of property in error and error message.
        /// The error message is built by splitting propNamePat on propNamePatWildcard and prepending the 
        /// last element of the split array with blanks added in front of capital letters to the string
        /// "of two sources cannot be equal for the same INCID".</param>
        private void ValidateSourceImportances(string propNamePat, string propNamePatWildcard,
            string propNameWildcardValCurrProp, List<string[]> errors)
        {
            string propNameCheck = propNamePat.Replace(propNamePatWildcard, propNameWildcardValCurrProp);
            PropertyInfo propInf = this.GetType().GetProperty(propNameCheck);
            if (propInf == null) return;

            string skipVal = Settings.Default.SourceImportanceSkip;
            string ord1val = Settings.Default.SourceImportanceApply1;
            string ord2val = Settings.Default.SourceImportanceApply2;
            string ord3val = Settings.Default.SourceImportanceApply3;
            
            object checkVal = propInf.GetValue(this, null);
            if ((checkVal == null) || checkVal.Equals(skipVal)) return;

            string[] split = propNamePat.Split(new string[] { propNamePatWildcard }, StringSplitOptions.None);
            string errMsg = split[split.Length - 1];

            errMsg = Regex.Matches(errMsg, @"[A-Z][^A-Z\s]*").Cast<Match>()
                .Aggregate(new StringBuilder(), (sb, m) => sb.Append(errMsg.Substring(m.Index, m.Length)).Append(" "))
                .AppendFormat("must be applied in the order {0}, {1} then {2}", ord1val, ord2val, ord3val).ToString();

            if (!String.IsNullOrEmpty(ord1val))
            {
                int ord1Sources = 0;
                foreach (PropertyInfo pi in this.GetType().GetProperties().Where(pn => Regex.IsMatch(pn.Name, propNamePat)))
                {
                    object compVal = pi.GetValue(this, null);
                    if ((compVal != null) && !compVal.Equals(skipVal) && compVal.Equals(ord1val))
                    {
                        ord1Sources += 1;
                    }
                }
                if (ord1Sources == 0 && checkVal.Equals(ord2val))
                        errors.Add(new string[] { propNameCheck, errMsg });
            }

            if (!String.IsNullOrEmpty(ord2val))
            {
                int ord2Sources = 0;
                foreach (PropertyInfo pi in this.GetType().GetProperties().Where(pn => Regex.IsMatch(pn.Name, propNamePat)))
                {
                    object compVal = pi.GetValue(this, null);
                    if ((compVal != null) && !compVal.Equals(skipVal) && compVal.Equals(ord2val))
                    {
                        ord2Sources += 1;
                    }
                }
                if (ord2Sources == 0 && checkVal.Equals(ord3val))
                    errors.Add(new string[] { propNameCheck, errMsg });
            }

        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Checks all properties of current class whose names follow a specified pattern for duplicate values.
        /// </summary>
        /// <param name="propNamePat">Pattern for regular expression match of property name in current class.</param>
        /// <param name="propNamePatWildcard">Wildcard element in propNamePat.</param>
        /// <param name="propNameWildcardValCurrProp">Value of propNamePatWildcard for current property to be validated.</param>
        /// <param name="skipVal">Value that may occur repeatedly (e.g. "none").</param>
        /// <param name="errors">List of errors, composed of name of property in error and error message.
        /// The error message is built by splitting propNamePat on propNamePatWildcard and prepending the 
        /// last element of the split array with blanks added in front of capital letters to the string
        /// "of two sources cannot be equal for the same INCID".</param>
        private void ValidateSourceDuplicates(string propNamePat, string propNamePatWildcard,
            string propNameWildcardValCurrProp, object skipVal, List<string[]> errors)
        {
            string propNameCheck = propNamePat.Replace(propNamePatWildcard, propNameWildcardValCurrProp);
            PropertyInfo propInf = this.GetType().GetProperty(propNameCheck);
            if (propInf == null) return;

            object checkVal = propInf.GetValue(this, null);
            if ((checkVal == null) || checkVal.Equals(skipVal)) return;

            string[] split = propNamePat.Split(new string[] { propNamePatWildcard }, StringSplitOptions.None);
            string errMsg = split[split.Length - 1];
            errMsg = Regex.Matches(errMsg, @"[A-Z][^A-Z\s]*").Cast<Match>()
                .Aggregate(new StringBuilder(), (sb, m) => sb.Append(errMsg.Substring(m.Index, m.Length)).Append(" "))
                .Append("of two sources cannot be equal for the same INCID").ToString();

            foreach (PropertyInfo pi in this.GetType().GetProperties().Where(pn => pn.Name != propNameCheck && Regex.IsMatch(pn.Name, propNamePat)))
            {
                if (pi.Name == propNameCheck) continue;

                object compVal = pi.GetValue(this, null);
                if ((compVal != null) && !compVal.Equals(skipVal) && compVal.Equals(checkVal))
                {
                    errors.Add(new string[] { propNameCheck, errMsg });
                    errors.Add(new string[] { pi.Name, errMsg });
                }
            }
        }

        #endregion

        #region IDataErrorInfo Members

        private string ErrorMessage(string columnName, List<string[]> errorList)
        {
            if (errorList != null)
            {
                IEnumerable<string[]> err = errorList.Where(s => s[0] == columnName);
                if (err.Count() > 0) return err.ElementAt(0)[1];
            }
            return null;
        }

        private string ErrorMessageList(List<string[]> errors)
        {
            if ((errors == null) || (errors.Count == 0)) return null;

            StringBuilder sbMsg = new StringBuilder();

            foreach (string[] e in errors)
            {
                if ((e.Length == 2) && (!String.IsNullOrEmpty(e[1])))
                    sbMsg.Append(Environment.NewLine).Append(e[1]);
            }

            if (sbMsg.Length > 0)
            {
                sbMsg.Remove(0, 1);
                return sbMsg.ToString();
            }
            else
            {
                return null;
            }
        }

        public string Error
        {
            get
            {
                if ((_bulkUpdateMode == true) || (_incidCurrentRow == null) || 
                    (_incidCurrentRow.RowState == DataRowState.Detached)) return null;

                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(Reason))
                    error.Append(Environment.NewLine).Append("Reason is mandatory for the history trail of every INCID");

                if (String.IsNullOrEmpty(Process))
                    error.Append(Environment.NewLine).Append("Process is mandatory for the history trail of every INCID");

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // If there are any IHS field errors then show an error on the tab label.
                if (_ihsErrors != null && _ihsErrors.Count > 0)
                    error.Append(Environment.NewLine).Append("One or more IHS codes are in error");
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // No need to check errors again here as they are already being checked
                // whenever the IHS codes are changed.
                //
                //if (String.IsNullOrEmpty(IncidIhsHabitat))
                //    error.Append(Environment.NewLine).Append("IHS Habitat is mandatory for every INCID");

                //if (!String.IsNullOrEmpty(IncidIhsMatrix1) && (_lutIhsMatrixCodes != null) && 
                //    (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix1) == 0))
                //    error.Append(Environment.NewLine).Append("Matrix code 1 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsMatrix2) && (_lutIhsMatrixCodes != null) && 
                //    (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix2) == 0))
                //    error.Append(Environment.NewLine).Append("Matrix code 2 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsMatrix3) && (_lutIhsMatrixCodes != null) && 
                //    (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix3) == 0))
                //    error.Append(Environment.NewLine).Append("Matrix code 3 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsFormation1) && (_lutIhsFormationCodes != null) && 
                //    (_lutIhsFormationCodes.Count(c => c.code == IncidIhsFormation1) == 0))
                //    error.Append(Environment.NewLine).Append("Formation code 1 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsFormation2) && (_lutIhsFormationCodes != null) && 
                //    (_lutIhsFormationCodes.Count(c => c.code == IncidIhsFormation2) == 0))
                //    error.Append(Environment.NewLine).Append("Formation code 2 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsManagement1) && (_lutIhsManagementCodes != null) && 
                //    (_lutIhsManagementCodes.Count(c => c.code == IncidIhsManagement1) == 0))
                //    error.Append(Environment.NewLine).Append("Management code 1 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsManagement2) && (_lutIhsManagementCodes != null) && 
                //    (_lutIhsManagementCodes.Count(c => c.code == IncidIhsManagement2) == 0))
                //    error.Append(Environment.NewLine).Append("Management code 2 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsComplex1) && (_lutIhsComplexCodes != null) && 
                //    (_lutIhsComplexCodes.Count(c => c.code == IncidIhsComplex1) == 0))
                //    error.Append(Environment.NewLine).Append("Complex code 1 does not correspond to chosen IHS habitat");

                //if (!String.IsNullOrEmpty(IncidIhsComplex2) && (_lutIhsComplexCodes != null) && 
                //    (_lutIhsComplexCodes.Count(c => c.code == IncidIhsComplex2) == 0))
                //    error.Append(Environment.NewLine).Append("Complex code 2 does not correspond to chosen IHS habitat");

                //string s = ValidateIhsFormation1Code();
                //if (!String.IsNullOrEmpty(s)) error.Append(Environment.NewLine).Append(s);

                //s = ValidateIhsManagement1Code();
                //if (!String.IsNullOrEmpty(s)) error.Append(Environment.NewLine).Append(s);
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // If there are any Detail field errors then show an error on the tab label.
                if (_detailsErrors != null && _detailsErrors.Count > 0)
                    error.Append(Environment.NewLine).Append("One or more Detail codes are in error");
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // No need to check errors again here as they are already being checked
                // when the BAP fields are changed.
                //
                //if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                //{
                //    int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                //    if (countInvalid > 0)
                //        error.Append(Environment.NewLine).Append(String.Format(
                //            "Invalid primary priority habitat{0} in row{0} {1}.",
                //            countInvalid > 1 ? "s" : String.Empty, String.Join(", ", _incidBapRowsAuto
                //            .Where(be => !be.IsValid()).Select((be, index) => (index + 1).ToString()).ToArray())));
                //}

                //if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
                //{
                //    int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                //    if (countInvalid > 0)
                //        error.Append(Environment.NewLine).Append(String.Format(
                //            "Invalid secondary priority habitat{0} in row{0} {1}.",
                //            countInvalid > 1 ? "s" : String.Empty, String.Join(", ", _incidBapRowsUser
                //            .Where(be => !be.IsValid()).Select((be, index) => (index + 1).ToString()).ToArray())));

                //    if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                //    {
                //        StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                //                                group be by be.bap_habitat into g
                //                                where g.Count() > 1
                //                                select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                //        if (beDups.Length > 2) error.Append(beDups.Remove(0, 2));
                //    }
                //}
                //---------------------------------------------------------------------

                if (String.IsNullOrEmpty(IncidBoundaryBaseMap))
                    error.Append(Environment.NewLine).Append("Boundary basemap is mandatory for every INCID");

                if (String.IsNullOrEmpty(IncidDigitisationBaseMap))
                    error.Append(Environment.NewLine).Append("Digitisation basemap is mandatory for every INCID");

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // If there are any Source field errors then show an error on the tab label.
                if (((Source1Errors != null) && (Source1Errors.Count > 0)) ||
                    ((Source2Errors != null) && (Source2Errors.Count > 0)) ||
                    ((Source3Errors != null) && (Source3Errors.Count > 0)))
                    error.Append(Environment.NewLine).Append("One or more Sources are in error");
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // Store the Source field errors so that they can be checked
                // at the end to see if the Source tab label should also be flagged
                // as in error.
                //Source1Errors = ValidateSource1();
                if ((Source1Errors != null) && (Source1Errors.Count > 0))
                    error.Append(Environment.NewLine).Append(ErrorMessageList(Source1Errors));

                //Source2Errors = ValidateSource2();
                if ((Source2Errors != null) && (Source2Errors.Count > 0))
                    error.Append(Environment.NewLine).Append(ErrorMessageList(Source2Errors));

                //Source3Errors = ValidateSource3();
                if ((Source3Errors != null) && (Source3Errors.Count > 0))
                    error.Append(Environment.NewLine).Append(ErrorMessageList(Source3Errors));
                //---------------------------------------------------------------------

                if (error.Length > 1)
                    return error.Remove(0, 1).ToString();
                else
                    return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                if ((_bulkUpdateMode == true) || (_incidCurrentRow == null) || 
                    (_incidCurrentRow.RowState == DataRowState.Detached)) return null;

                string error = null;

                //---------------------------------------------------------------------
                // FIX: Show field errors on tab labels.
                // Check the individual field errors to see if their parent tab label
                // should be flagged as also in error.
                switch (columnName)
                {
                    case "Reason":
                        if (String.IsNullOrEmpty(Reason))
                            error = "Reason is mandatory for the history trail of every INCID";
                        break;
                    case "Process":
                        if (String.IsNullOrEmpty(Process))
                            error = "Process is mandatory for the history trail of every INCID";
                        break;
                    case "Incid":
                        break;
                    case "IhsTabLabel":
                        if (_ihsErrors != null && _ihsErrors.Count > 0)
                            error = "One or more IHS codes are in error";
                        break;
                    case "IncidIhsHabitat":
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (String.IsNullOrEmpty(IncidIhsHabitat))
                        {
                            error = "IHS Habitat is mandatory for every INCID";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsMatrix1":
                        FixIhsMatrixCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsMatrix1) && (_lutIhsMatrixCodes != null) && 
                            (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix1) == 0))
                        {
                            error = "Matrix code 1 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsMatrix2":
                        FixIhsMatrixCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsMatrix2) && (_lutIhsMatrixCodes != null) && 
                            (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix2) == 0))
                        {
                            error = "Matrix code 2 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsMatrix3":
                        FixIhsMatrixCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsMatrix3) && (_lutIhsMatrixCodes != null) &&
                            (_lutIhsMatrixCodes.Count(c => c.code == IncidIhsMatrix3) == 0))
                        {
                            error = "Matrix code 3 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsFormation1":
                        FixIhsFormationCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsFormation1) && (_lutIhsFormationCodes != null) &&
                            (_lutIhsFormationCodes.Count(c => c.code == IncidIhsFormation1) == 0))
                        {
                            error = "Formation code 1 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                            error = ValidateIhsFormation1Code();
                            if (error != null)
                            {
                                UpdateErrorList(ref _ihsErrors, columnName, true);
                            }
                            else
                            {
                                UpdateErrorList(ref _ihsErrors, columnName, false);
                            }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsFormation2":
                        FixIhsFormationCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsFormation2) && (_lutIhsFormationCodes != null) &&
                            (_lutIhsFormationCodes.Count(c => c.code == IncidIhsFormation2) == 0))
                        {
                            error = "Formation code 2 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsManagement1":
                        FixIhsManagementCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsManagement1) && (_lutIhsManagementCodes != null) &&
                            (_lutIhsManagementCodes.Count(c => c.code == IncidIhsManagement1) == 0))
                        {
                            error = "Management code 1 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                            error = ValidateIhsManagement1Code();
                            if (error != null)
                            {
                                UpdateErrorList(ref _ihsErrors, columnName, true);
                            }
                            else
                            {
                                UpdateErrorList(ref _ihsErrors, columnName, false);
                            }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsManagement2":
                        FixIhsManagementCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsManagement2) && (_lutIhsManagementCodes != null) && 
                            (_lutIhsManagementCodes.Count(c => c.code == IncidIhsManagement2) == 0))
                        {
                            error = "Management code 2 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsComplex1":
                        FixIhsComplexCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsComplex1) && (_lutIhsComplexCodes != null) &&
                            (_lutIhsComplexCodes.Count(c => c.code == IncidIhsComplex1) == 0))
                        {
                            error = "Complex code 1 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "IncidIhsComplex2":
                        FixIhsComplexCodes();
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (!String.IsNullOrEmpty(IncidIhsComplex2) && (_lutIhsComplexCodes != null) && 
                            (_lutIhsComplexCodes.Count(c => c.code == IncidIhsComplex2) == 0))
                        {
                            error = "Complex code 2 does not correspond to chosen IHS habitat";
                            UpdateErrorList(ref _ihsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _ihsErrors, columnName, false);
                        }
                        OnPropertyChanged("IhsTabLabel");
                        break;
                    case "DetailsTabLabel":
                        if (_detailsErrors != null && _detailsErrors.Count > 0)
                            error = "One or more Details are in error";
                        break;
                    case "IncidBoundaryBaseMap":
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (String.IsNullOrEmpty(IncidBoundaryBaseMap))
                        {
                            error = "Boundary basemap is mandatory for every INCID";
                            UpdateErrorList(ref _detailsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _detailsErrors, columnName, false);
                        }
                        OnPropertyChanged("DetailsTabLabel");
                        break;
                    case "IncidDigitisationBaseMap":
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (String.IsNullOrEmpty(IncidDigitisationBaseMap))
                        {
                            error = "Digitisation basemap is mandatory for every INCID";
                            UpdateErrorList(ref _detailsErrors, columnName, true);
                        }
                        else
                        {
                            UpdateErrorList(ref _detailsErrors, columnName, false);
                        }
                        OnPropertyChanged("DetailsTabLabel");
                        break;
                    case "SourcesTabLabel":
                        // Check the Source field errors to see if the Source tab label
                        // should be flagged as in error.
                        if ((Source1Errors != null && Source1Errors.Count > 0) ||
                            (Source2Errors != null && Source2Errors.Count > 0) ||
                            (Source3Errors != null && Source3Errors.Count > 0))
                            error = "One or more Sources are in error";
                        break;
                    case "IncidSource1Id":
                    case "IncidSource1Date":
                    case "IncidSource1HabitatClass":
                    case "IncidSource1HabitatType":
                    case "IncidSource1BoundaryImportance":
                    case "IncidSource1HabitatImportance":
                        // Store the Source1 field errors so that they can be checked
                        // later to see if the Source tab label should also be flagged
                        // as in error.
                        Source1Errors = ValidateSource1();
                        error = ErrorMessage(columnName, Source1Errors);
                        OnPropertyChanged("SourcesTabLabel");
                        break;
                    case "IncidSource2Id":
                    case "IncidSource2Date":
                    case "IncidSource2HabitatClass":
                    case "IncidSource2HabitatType":
                    case "IncidSource2BoundaryImportance":
                    case "IncidSource2HabitatImportance":
                        // Store the Source2 field errors so that they can be checked
                        // later to see if the Source tab label should also be flagged
                        // as in error.
                        Source2Errors = ValidateSource2();
                        error = ErrorMessage(columnName, Source2Errors);
                        OnPropertyChanged("SourcesTabLabel");
                        break;
                    case "IncidSource3Id":
                    case "IncidSource3Date":
                    case "IncidSource3HabitatClass":
                    case "IncidSource3HabitatType":
                    case "IncidSource3BoundaryImportance":
                    case "IncidSource3HabitatImportance":
                        // Store the Source3 field errors so that they can be checked
                        // later to see if the Source tab label should also be flagged
                        // as in error.
                        Source3Errors = ValidateSource3();
                        error = ErrorMessage(columnName, Source3Errors);
                        OnPropertyChanged("SourcesTabLabel");
                        break;
                }
                //---------------------------------------------------------------------

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                OnPropertyChanged("CanCopy");
                OnPropertyChanged("CanPaste");

                return error;
            }
        }

        public void UpdateErrorList(ref List<string> errorList, string columnName, bool addError)
        {
            if (addError == true)
            {
                if (!errorList.Contains(columnName))
                    errorList.Add(columnName);
            }
            else
            {
                if (errorList.Contains(columnName))
                    errorList.Remove(columnName);
            }
        }

        #endregion
    }
}
