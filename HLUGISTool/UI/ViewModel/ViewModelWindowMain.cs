// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013-2014, 2016 Thames Valley Environmental Records Centre
// Copyright © 2014, 2018 Sussex Biodiversity Record Centre
// Copyright © 2019 London & South East Record Centres (LaSER)
// Copyright © 2019-2022 Greenspace Information for Greater London CIC
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

    /// <summary>
    /// An enumeration of the different options for when to warn
    /// the user before performing a GIS selection.
    /// </summary>
    public enum WarnBeforeGISSelect
    {
        Always,
        Joins,
        Never
    };

    /// <summary>
    /// An enumeration of the different options for whether
    /// to auto zoom to the GIS selection.
    /// </summary>
    public enum AutoZoomSelection
    {
        Off,
        When,
        Always
    };

    /// <summary>
    /// An enumeration of the different options for whether
    /// to validate secondary codes against the primary code.
    /// </summary>
    public enum SecondaryCodeValidationOptions
    {
        Ignore,
        Error
    };

    /// <summary>
    /// An enumeration of the different options for whether
    /// to validate quality determination and interpretation.
    /// </summary>
    public enum QualityValidationOptions
    {
        Optional,
        Mandatory
    };

    public class ViewModelWindowMain : ViewModelBase, IDataErrorInfo
    {
        #region Enums

        /// <summary>
        /// Geometry types.
        /// </summary>
        public enum GeometryTypes { Point, Line, Polygon, Unknown };

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // Functionality to process proposed OSMM Updates.
        //    
        /// <summary>
        /// Update operations.
        /// </summary>
        public enum Operations { PhysicalMerge, PhysicalSplit, LogicalMerge, LogicalSplit, AttributeUpdate, BulkUpdate, OSMMUpdate };
        //---------------------------------------------------------------------

        /// <summary>
        /// User Interface control visibility values.
        /// </summary>
        //public enum Visibility { Visible, Hidden, Collapsed };

        #endregion

        #region Fields

        private ICommand _navigateFirstCommand;
        private ICommand _navigatePreviousCommand;
        private ICommand _navigateNextCommand;
        private ICommand _navigateLastCommand;
        private ICommand _filterByAttributesCommand;
        private ICommand _filterByAttributesOSMMCommand;
        private ICommand _filterByIncidCommand;
        private ICommand _selectOnMapCommand;
        private ICommand _selectAllOnMapCommand;
        private ICommand _clearFilterCommand;
        private ICommand _readMapSelectionCommand;
        private ICommand _editPriorityHabitatsCommand;
        private ICommand _editPotentialHabitatsCommand;
        private ICommand _addSecondaryHabitatCommand;
        private ICommand _addSecondaryHabitatListCommand;
        private ICommand _selectByIncidCommand;
        private ICommand _switchGISLayerCommand;
        private ICommand _logicalSplitCommand;
        private ICommand _physicalSplitCommand;
        private ICommand _logicalMergeCommand;
        private ICommand _physicalMergeCommand;
        private ICommand _updateCommand;
        private ICommand _bulkUpdateCommandMenu;
        private ICommand _cancelBulkUpdateCommand;
        private ICommand _osmmUpdateCommandMenu;
        private ICommand _osmmUpdateAcceptCommandMenu;
        private ICommand _osmmUpdateRejectCommandMenu;
        private ICommand _osmmSkipCommand;
        private ICommand _osmmAcceptCommand;
        private ICommand _osmmRejectCommand;
        private ICommand _osmmBulkUpdateCommandMenu;
        private ICommand _exportCommand;
        private ICommand _closeCommand;
        private ICommand _copyCommand;
        private ICommand _pasteCommand;
        private ICommand _appKeepOnTopCommand;
        private ICommand _autoZoomSelectedOffCommand;
        private ICommand _autoZoomSelectedWhenCommand;
        private ICommand _autoZoomSelectedAlwaysCommand;
        private ICommand _autoSelectOnGisCommand;
        private ICommand _zoomSelectionCommand;
        private ICommand _gisWinSideBySideCommand;
        private ICommand _gisWinSwitchToCommand;
        private ICommand _resetToolWindowCommand;
        private ICommand _optionsCommand;
        private ICommand _aboutCommand;

        private WindowMainCopySwitches _copySwitches = new WindowMainCopySwitches();
        private WindowAbout _windowAbout;
        private ViewModelWindowAbout _viewModelAbout;
        private WindowOptions _windowOptions;
        private ViewModelOptions _viewModelOptions;
        private WindowQueryIncid _windowQueryIncid;
        private ViewModelWindowQueryIncid _viewModelWinQueryIncid;
        private WindowQuerySecondaries _windowQuerySecondaries;
        private ViewModelWindowQuerySecondaries _viewModelWinQuerySecondaries;
        private WindowQueryAdvanced _windowQueryAdvanced;
        private ViewModelWindowQueryAdvanced _viewModelWinQueryAdvanced;
        private WindowQueryOSMM _windowQueryOSMM;
        private ViewModelWindowQueryOSMM _viewModelWinQueryOSMM;
        private WindowWarnOnGISSelect _windowWarnGISSelect;
        private ViewModelWindowWarnOnGISSelect _viewModelWinWarnGISSelect;
        private WindowNotifyOnSplitMerge _windowWarnSplitMerge;
        private ViewModelWindowNotifyOnSplitMerge _viewModelWinWarnSplitMerge;
        private WindowWarnOnSubsetUpdate _windowWarnSubsetUpdate;
        private WindowSwitchGISLayer _windowSwitchGISLayer;
        private ViewModelWindowSwitchGISLayer _viewModelSwitchGISLayer;
        private WindowCompletePhysicalSplit _windowCompSplit;
        private ViewModelCompletePhysicalSplit _vmCompSplit;
        private ViewModelWindowWarnOnSubsetUpdate _viewModelWinWarnSubsetUpdate;
        private ViewModelWindowMainBulkUpdate _viewModelBulkUpdate;
        private ViewModelWindowMainOSMMUpdate _viewModelOSMMUpdate;
        private ViewModelWindowMainUpdate _viewModelUpd;
        private WindowEditPriorityHabitats _windowEditPriorityHabitats;
        private ViewModelWindowEditPriorityHabitats _viewModelWinEditPriorityHabitats;
        private WindowEditPotentialHabitats _windowEditPotentialHabitats;
        private ViewModelWindowEditPotentialHabitats _viewModelWinEditPotentialHabitats;

        private bool haveSplashWin;
        private string _displayName = "HLU Tool";
        private int _windowHeight;
        private int _defaultWindowHeight;
        private int _windowWidth;
        private int _mapWindowsCount;
        private bool _showingReasonProcessGroup = false;
        private bool _showingOSMMPendingGroup = false;

        // Database options
        private int _dbConnectionTimeout = Settings.Default.DbConnectionTimeout;

        // GIS/Export options
        private int _minZoom = Settings.Default.MinAutoZoom;

        // History options
        private DataColumn[] _historyColumns;
        private int _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;

        // Interface options
        private string _preferredHabitatClass = Settings.Default.PreferredHabitatClass;
        private bool _showGroupHeaders = Settings.Default.ShowGroupHeaders;
        private bool _showNVCCodes = Settings.Default.ShowNVCCodes;
        private bool _showIHSTab = Settings.Default.ShowIHSTab;
        private string _showOSMMUpdates = Settings.Default.ShowOSMMUpdatesOption;
        private string _preferredSecondaryGroup = Settings.Default.PreferredSecondaryGroup;
        private string _secondaryCodeOrder = Settings.Default.SecondaryCodeOrder;
        private string _secondaryCodeDelimiter = Settings.Default.SecondaryCodeDelimiter;

        // Updates options
        private int _subsetUpdateAction = Settings.Default.SubsetUpdateAction;
        private string _clearIHSUpdateAction = Settings.Default.ClearIHSUpdateAction;
        private bool _notifyOnSplitMerge = Settings.Default.NotifyOnSplitMerge;
        private bool _resetOSMMUpdatesStatus = Settings.Default.ResetOSMMUpdatesStatus;
        private int _secondaryCodeValidation = Settings.Default.SecondaryCodeValidation;
        private int _qualityValidation = Settings.Default.QualityValidation;

        // Filter options
        private int _warnBeforeGISSelect = Settings.Default.WarnBeforeGISSelect;

        // Dates options
        // None

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
        private HluDataSet.incid_osmm_updatesRow[] _incidOSMMUpdatesRows;
        private HluDataSet.historyRow[] _incidHistoryRows;
        private HluDataSet.incid_conditionRow[] _incidConditionRows;
        private HluDataSet.incid_secondaryRow[] _incidSecondaryRows;

        private HluDataSet.lut_reasonRow[] _reasonCodes;
        private HluDataSet.lut_processRow[] _processCodes;
        private IEnumerable<HluDataSet.lut_ihs_habitatRow> _lutIhsHabitatCodes;
        private IEnumerable<HluDataSet.lut_ihs_matrixRow> _lutIhsMatrixCodes;
        private IEnumerable<HluDataSet.lut_ihs_formationRow> _lutIhsFormationCodes;
        private IEnumerable<HluDataSet.lut_ihs_managementRow> _lutIhsManagementCodes;
        private IEnumerable<HluDataSet.lut_ihs_complexRow> _lutIhsComplexCodes;
        private HluDataSet.lut_quality_determinationRow[] _qualityDeterminationCodes;
        private HluDataSet.lut_quality_interpretationRow[] _qualityInterpretationCodes;
        private HluDataSet.lut_sourcesRow[] _sourceNames;
        private HluDataSet.lut_habitat_classRow[] _sourceHabitatClassCodes;
        private HluDataSet.lut_importanceRow[] _sourceImportanceCodes;
        private HluDataSet.lut_habitat_typeRow[] _bapHabitatCodes;

        private HluDataSet.lut_conditionRow[] _conditionCodes;
        private HluDataSet.lut_condition_qualifierRow[] _conditionQualifierCodes;
        //private HluDataSet.lut_primary_categoryRow[] _primaryCategoryCodes;
        //private HluDataSet.lut_primaryRow[] _primaryCodes;
        private HluDataSet.lut_secondary_groupRow[] _secondaryGroupsValid;
        private HluDataSet.lut_secondary_groupRow[] _secondaryGroups;
        public static HluDataSet.lut_secondary_groupRow[] SecondaryGroupsAll;  // Used in the options window
        private HluDataSet.lut_secondaryRow[] _secondaryCodesAll;
        private HluDataSet.lut_secondaryRow[] _secondaryCodesValid;

        private ObservableCollection<SecondaryHabitat> _incidSecondaryHabitats;

        private ObservableCollection<BapEnvironment> _incidBapRowsAuto;
        private ObservableCollection<BapEnvironment> _incidBapRowsUser;

        private IEnumerable<HluDataSet.lut_primaryRow> _primaryCodes;
        private HistoryRowEqualityComparer _histRowEqComp = new HistoryRowEqualityComparer();
        private HluDataSet.lut_habitat_classRow[] _habitatClassCodes;
        public static HluDataSet.lut_habitat_classRow[] HabitatClasses; // Used in the options window
        private HluDataSet.lut_habitat_typeRow[] _habitatTypeCodes;
        private IEnumerable<HluDataSet.lut_osmm_habitat_xrefRow> _osmmHabitatXrefIds;
        //private IEnumerable<HluDataSet.lut_habitat_type_primaryRow> _xrefHabitatTypePrimaryCodes;
        //private IEnumerable<HluDataSet.lut_habitat_type_secondaryRow> _xrefHabitatTypeSecondaryCodes;

        private double _incidArea;
        private double _incidLength;
        private string _process;
        private string _reason;
        private string _habitatClass;
        private string _habitatType;
        private string _secondaryGroup;
        private string _secondaryHabitat;
        private bool _reasonProcessEnabled = true;
        private bool _tabControlDataEnabled = true;
        private int _tabItemSelected = 0;
        private bool _tabItemHabitatEnabled = true;
        private bool _tabItemIHSEnabled = true;
        private bool _tabItemPriorityEnabled = true;
        private bool _tabItemDetailsEnabled = true;
        private bool _tabItemSourcesEnabled = true;
        private bool _tabItemHistoryEnabled = true;
        private bool _tabHabitatControlsEnabled = true;
        private bool _tabIhsControlsEnabled = true;
        private bool _tabPriorityControlsEnabled = true;
        private bool _tabDetailsControlsEnabled = true;
        private bool _tabSourcesControlsEnabled = true;
        private bool _windowEnabled = true;
        private bool _editMode;
        private bool _pasting = false;
        private bool _changed = false;
        private bool _saving = false;
        private bool _closing = false;
        private bool _autoSplit = true;
        private bool _splitting = false;
        private bool _filterByMap = false;
        private bool _osmmUpdating = false;
        private Cursor _windowCursor = Cursors.Arrow;
        private DataColumn[] _gisIDColumns;
        private int[] _gisIDColumnOrdinals;
        private IEnumerable<string> _incidsSelectedMap;
        private IEnumerable<string> _toidsSelectedMap;
        private IEnumerable<string> _fragsSelectedMap;
        private int _incidsSelectedDBCount = 0;
        private int _toidsSelectedDBCount = 0;
        private int _fragsSelectedDBCount = 0;
        private int _incidsSelectedMapCount = 0;
        private int _toidsSelectedMapCount = 0;
        private int _fragsSelectedMapCount = 0;
        private int _toidsIncidGisCount = 0;
        private int _fragsIncidGisCount = 0;
        private int _toidsIncidDbCount = 0;
        private int _fragsIncidDbCount = 0;
        private int _origIncidSecondaryCount = 0;
        private int _origIncidConditionCount = 0;
        private int _origIncidIhsMatrixCount = 0;
        private int _origIncidIhsFormationCount = 0;
        private int _origIncidIhsManagementCount = 0;
        private int _origIncidIhsComplexCount = 0;
        private int _origIncidSourcesCount = 0;
        private SqlFilterCondition _incidMMPolygonsIncidFilter;
        private int _incidRowCount;
        private int _incidPageRowNoMin = 0;
        private int _incidPageRowNoMax = 0;
        private string _incidIhsHabitat;
        private string _incidPrimary;
        private string _incidPrimaryCategory;
        private string _incidSecondarySummary;
        private string _incidLastModifiedUser;
        private DateTime _incidLastModifiedDate;
        private string _incidLegacyHabitat;
        //private string _incidQualityDetermination;
        private int _incidOSMMUpdatesOSMMXref;
        private int _incidOSMMUpdatesProcessFlag;
        private string _incidOSMMUpdatesSpatialFlag;
        private string _incidOSMMUpdatesChangeFlag;
        private Nullable<int> _incidOSMMUpdatesStatus;
        private Dictionary<Type, List<SqlFilterCondition>> _childRowFilterDict;
        private Dictionary<Type, string> _childRowOrderByDict;
        private List<List<SqlFilterCondition>> _incidSelectionWhereClause;
        private string _osmmUpdateWhereClause;
        private List<string> _exportMdbs = new List<string>();
        private string _userName;
        private string _dbVersion;
        private Nullable<bool> _isAuthorisedUser;
        private Nullable<bool> _canBulkUpdate;
        private Nullable<bool> _bulkUpdateMode = false;
        private bool _bulkUpdatePrimaryBap = Settings.Default.BulkUpdatePotentialBap;
        private string _osmmAcceptTag = "A_ccept";
        private string _osmmRejectTag = "Re_ject";
        private Nullable<bool> _canOSMMUpdate;
        private Nullable<bool> _osmmUpdateMode = false;
        private Nullable<bool> _osmmBulkUpdateMode = false;
        private bool _osmmUpdatesEmpty = false;
        private bool _osmmUpdateCreateHistory;
        private string _codeAnyRow = Settings.Default.CodeAnyRow;

        private VagueDateInstance _incidConditionDateEntered;

        private VagueDateInstance _incidSource1DateEntered;
        private VagueDateInstance _incidSource2DateEntered;
        private VagueDateInstance _incidSource3DateEntered;
        private string _codeDeleteRow = Settings.Default.CodeDeleteRow;
        private string _processingMsg = "Processing ...";
        private bool _saved = false;
        private bool _savingAttempted;
        private List<string> _habitatWarnings = new List<string>();
        private List<string> _priorityWarnings = new List<string>();
        private List<string> _detailsWarnings = new List<string>();
        private List<string[]> _conditionWarnings = null;
        private List<string[]> _source1Warnings = null;
        private List<string[]> _source2Warnings = null;
        private List<string[]> _source3Warnings = null;
        private List<string> _habitatErrors = new List<string>();
        private List<string> _priorityErrors = new List<string>();
        private List<string> _detailsErrors = new List<string>();
        private List<string[]> _conditionErrors = null;
        private List<string[]> _source1Errors = null;
        private List<string[]> _source2Errors = null;
        private List<string[]> _source3Errors = null;

        private bool _updateCancelled = true;
        private bool _updateAllFeatures = true;
        private bool _refillIncidTable = false;
        private bool _appKeepOnTop = Settings.Default.AppKeepOnTop;
        private int _autoZoomSelection = Settings.Default.AutoZoomSelection;
        private bool _autoSelectOnGis = Settings.Default.AutoSelectOnGis;

        public static string HistoryGeometry1ColumnName = Settings.Default.HistoryGeometry1ColumnName;
        public static string HistoryGeometry2ColumnName = Settings.Default.HistoryGeometry2ColumnName;
        internal static string LutDescriptionFieldName = Settings.Default.LutDescriptionFieldName;
        internal static int LutDescriptionFieldOrdinal = Settings.Default.LutDescriptionFieldOrdinal;
        internal static string LutSourceFieldName = Settings.Default.LutSourceFieldName;
        internal static int LutSourceFieldOrdinal = Settings.Default.LutSourceFieldOrdinal;
        internal static string LutUserFieldName = Settings.Default.LutUserFieldName;
        internal static int LutUserFieldOrdinal = Settings.Default.LutUserFieldOrdinal;
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

                haveSplashWin = App.SplashViewModel != null;

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
                        // Clear the current database settings as they are clearly not valid.
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
                            errorMessage + "\n\nWould you like to connect to another database?", "HLU Dataset",
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
                    ChangeCursor(Cursors.Wait, "Initiating ...");
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

                //---------------------------------------------------------------------
                // CHANGED: CR30 (Database validation on start-up)
                // Check the assembly version is not earlier than the
                // minimum required dataset application version.
                if (!CheckVersion())
                    return false;
                //---------------------------------------------------------------------

                // wire up event handler for copy switches
                _copySwitches.PropertyChanged += new PropertyChangedEventHandler(_copySwitches_PropertyChanged);

                int result;
                // columns that identify map polygons and are returned by GIS
                _gisIDColumnOrdinals = (from s in Settings.Default.GisIDColumnOrdinals.Cast<string>()
                                        where Int32.TryParse(s, out result) && (result >= 0) &&
                                        (result < _hluDS.incid_mm_polygons.Columns.Count)
                                        select Int32.Parse(s)).ToArray();
                _gisIDColumns = _gisIDColumnOrdinals.Select(i => _hluDS.incid_mm_polygons.Columns[i]).ToArray();

                // columns to be displayed in history (always includes _gisIDColumns)
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

                // start the GIS application
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

                // Initialise the main update view model
                _viewModelUpd = new ViewModelWindowMainUpdate(this);

                // Get the BAP determination quality defaults
                GetBapDefaults();

                // Read the selected features from the map
                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = "Reading map selection...";
                    DispatcherHelper.DoEvents();
                }
                // Get the GIS layer selection
                ReadMapSelection(false);

                // Clear the splash window status bar (or reset the cursor to an arrow)
                if (haveSplashWin)
                {
                    App.SplashViewModel.ProgressText = String.Empty;
                    haveSplashWin = false;
                }
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
                // Make sure that all the available history columns are updated when
                // creating history even if the user only wants to display some of them.
                return _gisIDColumns.Concat(_hluDS.incid_mm_polygons.Columns.Cast<DataColumn>()
                    .Where(c => !_gisIDColumnOrdinals.Contains(c.Ordinal)
                        && !c.ColumnName.StartsWith("shape_"))).ToArray();
            }
            catch { return historyColumns; }
        }

        private void ProcessStartupArguments()
        {
            foreach (string s in App.StartupArguments)
            {
                switch (s.Replace(" ", String.Empty).ToLower())
                {
                    // Allow the user to reset only the database settings "/d"
                    // or only the GIS settings "/g" instead of always both "/c"
                    case "/d":
                        DbFactory.ClearSettings();
                        break;
                    case "/g":
                        GISAppFactory.ClearSettings();
                        break;
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

        //---------------------------------------------------------------------
        // CHANGED: CR30 (Database validation on start-up)
        /// <summary>
        /// Check the assembly version is greater than or equal to the
        /// application version from the lut_version table in the database.
        /// </summary>
        private bool CheckVersion()
        {
            // Get the assembly version.
            String assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // Get the application and database versions from the database.
            String lutAppVersion = "0.0.0";
            String lutDbVersion = "0";
            if (_hluDS.lut_version.Count > 0)
            {
                lutAppVersion = _hluDS.lut_version.ElementAt(_hluDS.lut_version.Count - 1).app_version;
                lutDbVersion = _hluDS.lut_version.ElementAt(_hluDS.lut_version.Count - 1).db_version;
            }
            else
            {
                //MessageBox.Show("The database has been updated to a later version than the application.\n\n" +
                //   "The application must be upgraded before it can be run.",
                //   "HLU", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }

            Version assVersion = new Version(assemblyVersion);
            Version appVersion = new Version(lutAppVersion);

            // Compare the assembly and application versions.
            if (assVersion.CompareTo(appVersion) < 0)
            {
                //MessageBox.Show(String.Format("The database has been updated to a later version than the application.\n\n" +
                //   "The minimum application version must be {0}.", appVersion.ToString()),
                //   "HLU", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                // Trap error if database requires a later application version.
                throw new Exception(String.Format("The minimum application version must be {0}.", appVersion.ToString()));
            }

            // Get the minimum database version.
            string minDbVersion = Settings.Default.MinimumDbVersion;

            // Compare the minimum database version.
            if (Base36.Base36ToNumber(lutDbVersion) < Base36.Base36ToNumber(minDbVersion))
            {
                //MessageBox.Show(String.Format("The database must be updated to a later version.\n\n" +
                //   "The minimum database version must be {0}.",minDbVersion),
                //   "HLU", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                // Trap error if application requires a later database version.
                throw new Exception(String.Format("The minimum database version must be {0}.", minDbVersion));
            }

            // Store the database version for displaying in the 'About' box.
            _dbVersion = lutDbVersion;

            return true;
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
                return String.Format("{0}{1}", DisplayName, _editMode ? String.Empty : " [READONLY]");
            }
        }

        /// <summary>
        /// Gets the name of the layer to display in the status bar.
        /// </summary>
        /// <value>
        /// The name of the layer.
        /// </value>
        public string LayerName
        {
            get
            {
                // If no HLU layer has been identified yet (GIS is still loading) then
                // don't return the layer name
                if (_gisApp.CurrentHluLayer == null)
                    return String.Empty;
                else
                {
                    //---------------------------------------------------------------------
                    // Do not display map window number with layer name
                    // if there is only one map window.
                    // 
                    if (_mapWindowsCount > 1)
                        // Include the layer name and active layer/map window number in the window title.
                        return String.Format("{0} [{1}]", _gisApp.CurrentHluLayer.LayerName, _gisApp.CurrentHluLayer.MapNum);
                    else
                        // Include only the layer name in the window title.
                        return String.Format("{0}", _gisApp.CurrentHluLayer.LayerName);
                    //---------------------------------------------------------------------
                }
                //---------------------------------------------------------------------
            }
        }
        /// <summary>
        /// Get and set the window height, adjusting it for any optional
        /// areas.
        /// </summary>
        public int WindowHeight
        {
            get
            {
                // Set the initial window height if not already set.
                if (_windowHeight == 0)
                    AdjustWindowHeight(true);

                return _windowHeight;
            }
            set
            {
                _windowHeight = value;
            }
        }

        /// <summary>
        /// Adjusts the window height for any optional groups/fields.
        /// </summary>
        public void AdjustWindowHeight(bool resetWindow)
        {
            int _newWindowHeight = 956;

            // Adjust the minimum height if the group headers are not showing.
            if (!_showGroupHeaders)
                _newWindowHeight -= 80;

            // Adjust the minimum height if the NVC codes text is not showing.
            if (!_showNVCCodes)
                _newWindowHeight -= 18;

            // Adjust the minimum height if the Reason and Process group is not showing.
            if (!_showingReasonProcessGroup)
                _newWindowHeight -= 47;

            // Adjust the minimum height if the OSMM Updates group is not showing.
            if (!_showingOSMMPendingGroup)
                _newWindowHeight -= 83;

            // Adjust the standard height for older versions of Windows.
            if (System.Environment.OSVersion.Version.Major < 10)
                _newWindowHeight += 10;

            // Only adjust the window height if the user hasn't
            // changed it manually or if the user has specified
            // if should be reset.
            if (_windowHeight == _defaultWindowHeight || resetWindow)
                _windowHeight = _newWindowHeight;

            // Set the new default window height.
            _defaultWindowHeight = _newWindowHeight;
        }

        /// <summary>
        /// Get and set the window width.
        /// </summary>
        public int WindowWidth
        {
            get
            {
                // Set the initial window width if not already set.
                if (_windowWidth == 0)
                {
                    _windowWidth = 492;
                }

                return _windowWidth;
            }
            set
            {
                _windowWidth = value;
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

        internal string OSMMUpdateWhereClause
        {
            get { return _osmmUpdateWhereClause; }
            set { _osmmUpdateWhereClause = value; }
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
        
        internal HluDataSet.incid_secondaryRow[] IncidSecondaryRows
        {
            get { return _incidSecondaryRows; }
            set { _incidSecondaryRows = value; }
        }

        internal HluDataSet.incid_conditionRow[] IncidConditionRows
        {
            get { return _incidConditionRows; }
            set { _incidConditionRows = value; }
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

        internal HluDataSet.incid_osmm_updatesRow[] IncidOSMMUpdatesRows
        {
            get { return _incidOSMMUpdatesRows; }
            set { _incidOSMMUpdatesRows = value; }
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

        internal List<string> HabitatWarnings
        {
            get { return _habitatWarnings; }
            set { _habitatWarnings = value; }
        }

        internal List<string> PriorityWarnings
        {
            get { return _priorityWarnings; }
            set { _priorityWarnings = value; }
        }

        internal List<string[]> ConditionWarnings
        {
            get { return _conditionWarnings; }
            set { _conditionWarnings = value; }
        }

        internal List<string> DetailsWarnings
        {
            get { return _detailsWarnings; }
            set { _detailsWarnings = value; }
        }

        internal List<string[]> Source1Warnings
        {
            get { return _source1Warnings; }
            set { _source1Warnings = value; }
        }

        internal List<string[]> Source2Warnings
        {
            get { return _source2Warnings; }
            set { _source2Warnings = value; }
        }

        internal List<string[]> Source3Warnings
        {
            get { return _source3Warnings; }
            set { _source3Warnings = value; }
        }

        internal List<string> HabitatErrors
        {
            get { return _habitatErrors; }
            set { _habitatErrors = value; }
        }

        internal List<string> PriorityErrors
        {
            get { return _priorityErrors; }
            set { _priorityErrors = value; }
        }

        internal List<string[]> ConditionErrors
        {
            get { return _conditionErrors; }
            set { _conditionErrors = value; }
        }

        internal List<string> DetailsErrors
        {
            get { return _detailsErrors; }
            set { _detailsErrors = value; }
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

        internal bool RefillIncidTable
        {
            get { return _refillIncidTable; }
            set { _refillIncidTable = true; }
        }

        internal int DBConnectionTimeout
        {
            get { return _dbConnectionTimeout; }
        }

        internal string ClearIHSUpdateAction
        {
            get { return _clearIHSUpdateAction; }
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

        #region Defaults

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process bulk OSMM Updates
        //
        /// <summary>
        /// Get the BAP determination quality default descriptions
        /// from the lookup table and update them in the settings.
        /// </summary>
        private void GetBapDefaults()
        {
            try
            {
                // Get the user added ('NP' = not present) description
                object result = _db.ExecuteScalar(String.Format("SELECT {0} FROM {1} WHERE {2} = {3}",
                    _db.QuoteIdentifier(_hluDS.lut_quality_determination.descriptionColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.lut_quality_determination.TableName),
                    _db.QuoteIdentifier(_hluDS.lut_quality_determination.codeColumn.ColumnName),
                    _db.QuoteValue(Settings.Default.BAPDeterminationQualityUserAdded)), _db.Connection.ConnectionTimeout, CommandType.Text);
                if (result != null)
                {
                    Settings.Default.BAPDeterminationQualityUserAddedDesc = (string)result;
                    Settings.Default.Save();
                }

                // Get the previous ('PP' = previously present) description
                result = _db.ExecuteScalar(String.Format("SELECT {0} FROM {1} WHERE {2} = {3}",
                    _db.QuoteIdentifier(_hluDS.lut_quality_determination.descriptionColumn.ColumnName),
                    _db.QualifyTableName(_hluDS.lut_quality_determination.TableName),
                    _db.QuoteIdentifier(_hluDS.lut_quality_determination.codeColumn.ColumnName),
                    _db.QuoteValue(Settings.Default.BAPDeterminationQualityPrevious)), _db.Connection.ConnectionTimeout, CommandType.Text);
                if (result != null)
                {
                    Settings.Default.BAPDeterminationQualityPreviousDesc = (string)result;
                    Settings.Default.Save();
                }
            }
            catch
            {
            }
        }
        //---------------------------------------------------------------------

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
            OnPropertyChanged("StatusBar");
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
                    _closeCommand = new RelayCommand(param => this.OnRequestClose(true));

                return _closeCommand;
            }
        }

        /// <summary>
        /// Raised when main window should be closed.
        /// </summary>
        public event EventHandler RequestClose;

        public void OnRequestClose(bool check)
        {
            // Set the event handler to close the application
            EventHandler handler = this.RequestClose;
            if (handler != null)
            {
                //---------------------------------------------------------------------
                // FIX: 106 Check if user is sure before closing application.
                //
                if ((check == false) || (MessageBox.Show("Close HLU Tool. Are you sure?", "HLU: Exit", MessageBoxButton.YesNo,
                        MessageBoxImage.Question) == MessageBoxResult.Yes))
                {
                    // Indicate the application is already closing.
                    _closing = true;
                    //---------------------------------------------------------------------

                    // Check there are no outstanding edits.
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

                    // Call the event handle to close the application
                    handler(this, EventArgs.Empty);
                }
            }
        }

        //---------------------------------------------------------------------
        // FIX: 106 Check if user is sure before closing application.
        //
        /// <summary>
        /// Is the application already in the process of closing.
        /// </summary>
        /// <returns></returns>
        public bool IsClosing
        {
            get { return _closing; }
        }
        //---------------------------------------------------------------------

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
                    Action<object> navigateFirstAction = new Action<object>(this.NavigateFirstClicked);
                    _navigateFirstCommand = new RelayCommand(navigateFirstAction, param => this.CanNavigateBackward);
                }
                return _navigateFirstCommand;
            }
        }

        private void NavigateFirstClicked(object param)
        {
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
                    Action<object> navigatePreviousAction = new Action<object>(this.NavigatePreviousClicked);
                    _navigatePreviousCommand = new RelayCommand(navigatePreviousAction, param => this.CanNavigateBackward);
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
            get { return IncidCurrentRowIndex > 1; }
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
                    Action<object> navigateNextAction = new Action<object>(this.NavigateNextClicked);
                    _navigateNextCommand = new RelayCommand(navigateNextAction, param => this.CanNavigateForward);
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
                return ((IsFiltered && (IncidCurrentRowIndex < _incidSelection.Rows.Count)) ||
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
                    Action<object> navigateLastAction = new Action<object>(this.NavigateLastClicked);
                    _navigateLastCommand = new RelayCommand(navigateLastAction, param => this.CanNavigateForward);
                }
                return _navigateLastCommand;
            }
        }

        private void NavigateLastClicked(object param)
        {
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
        /// Logical Split command.
        /// </summary>
        public ICommand LogicalSplitCommand
        {
            get
            {
                if (_logicalSplitCommand == null)
                {
                    Action<object> logicalSplitAction = new Action<object>(this.LogicalSplitClicked);
                    _logicalSplitCommand = new RelayCommand(logicalSplitAction, param => this.CanLogicallySplit);
                }
                return _logicalSplitCommand;
            }
        }

        /// <summary>
        /// Physical Split command.
        /// </summary>
        public ICommand PhysicalSplitCommand
        {
            get
            {
                if (_physicalSplitCommand == null)
                {
                    Action<object> physicalSplitAction = new Action<object>(this.PhysicalSplitClicked);
                    _physicalSplitCommand = new RelayCommand(physicalSplitAction, param => this.CanPhysicallySplit);
                }
                return _physicalSplitCommand;
            }
        }

        /// <summary>
        /// LogicalSplitCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void LogicalSplitClicked(object param)
        {
            _autoSplit = false;
            // Get the GIS layer selection again (just in case)
            ReadMapSelection(false);
            _autoSplit = true;

            // Check the selected rows are unique before attempting to split them.
            if (!_gisApp.SelectedRowsUnique())
            {
                MessageBox.Show("The map selection contains one or more features where a physical split has not been completed.\n\n" +
                    "Please select the features and invoke the Split command to prevent the map and database going out of sync.",
                    "HLU: Data Integrity", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                return;
            }

            ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
            //---------------------------------------------------------------------
            // CHANGED: CR39 (Split and merge complete messages)
            // Notify the user following the completion of the split.
            if (vmSplit.LogicalSplit()) NotifySplitMerge("Logical split completed.");
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// PhysicalSplitCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void PhysicalSplitClicked(object param)
        {
            _autoSplit = false;
            // Get the GIS layer selection again (just in case)
            ReadMapSelection(false);
            _autoSplit = true;

            ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
            //---------------------------------------------------------------------
            // CHANGED: CR39 (Split and merge complete messages)
            // Notify the user following the completion of the split.
            if (vmSplit.PhysicalSplit()) NotifySplitMerge("Physical split completed.");
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// At least one feature in selection that share the same incid, but *not* toid and toidfragid
        /// </summary>
        private bool CanLogicallySplit
        {
            get
            {
                return (_bulkUpdateMode == false && _osmmUpdateMode == false) &&
                    HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) &&
                    (_gisSelection != null) && (_incidsSelectedMapCount == 1) &&
                    ((_gisSelection.Rows.Count > 0) && ((_toidsSelectedMapCount > 1) || (_fragsSelectedMapCount > 0)) ||
                    (_gisSelection.Rows.Count == 1)) &&
                    // Only enable split/merge after select from map
                    (_filterByMap == true) &&
                    //---------------------------------------------------------------------
                    // CHANGED: CR7 (Split/merge options)
                    // Only enable logical split menu/button if a subset of all the
                    // features for the current incid have been selected.
                    ((_toidsIncidGisCount != _toidsIncidDbCount) ||
                    (_fragsIncidGisCount != _fragsIncidDbCount));
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// At least two features in selection that share the same incid, toid and toidfragid
        /// </summary>
        private bool CanPhysicallySplit
        {
            get
            {
                return (_bulkUpdateMode == false && _osmmUpdateMode == false) &&
                    HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) &&
                    (_gisSelection != null) && (_gisSelection.Rows.Count > 1) &&
                    // Only enable split/merge after select from map
                    (_filterByMap == true) &&
                    (_incidsSelectedMapCount == 1) && (_toidsSelectedMapCount == 1) && (_fragsSelectedMapCount == 1);
            }
        }

        #endregion

        #region Merge

        /// <summary>
        /// Logical Merge command.
        /// </summary>
        public ICommand LogicalMergeCommand
        {
            get
            {
                if (_logicalMergeCommand == null)
                {
                    Action<object> logcalMergeAction = new Action<object>(this.LogicalMergeClicked);
                    _logicalMergeCommand = new RelayCommand(logcalMergeAction, param => this.CanLogicallyMerge);
                }
                return _logicalMergeCommand;
            }
        }

        /// <summary>
        /// Physical Merge command.
        /// </summary>
        public ICommand PhysicalMergeCommand
        {
            get
            {
                if (_physicalMergeCommand == null)
                {
                    Action<object> logcalMergeAction = new Action<object>(this.PhysicalMergeClicked);
                    _physicalMergeCommand = new RelayCommand(logcalMergeAction, param => this.CanPhysicallyMerge);
                }
                return _physicalMergeCommand;
            }
        }

        /// <summary>
        /// LogicalMergeCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void LogicalMergeClicked(object param)
        {
            // Get the GIS layer selection again (just in case)
            ReadMapSelection(false);

            // Check the selected rows are unique before attempting to merge them.
            if (!_gisApp.SelectedRowsUnique())
            {
                MessageBox.Show("The map selection contains one or more features where a physical split has not been completed.\n\n" +
                    "Please select the features and invoke the Split command to prevent the map and database going out of sync.",
                    "HLU: Data Integrity", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                return;
            }

            ViewModelWindowMainMerge vmMerge = new ViewModelWindowMainMerge(this);
            //---------------------------------------------------------------------
            // CHANGED: CR39 (Split and merge complete messages)
            // Notify the user following the completion of the split.
            if (vmMerge.LogicalMerge()) NotifySplitMerge("Logical merge completed.");
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// PhysicalMergeCommand event handler.
        /// </summary>
        /// <param name="param"></param>
        private void PhysicalMergeClicked(object param)
        {
            // Get the GIS layer selection again (just in case)
            ReadMapSelection(false);

            // Check the selected rows are unique before attempting to merge them.
            if (!_gisApp.SelectedRowsUnique())
            {
                MessageBox.Show("The map selection contains one or more features where a physical split has not been completed.\n\n" +
                    "Please select the features and invoke the Split command to prevent the map and database going out of sync.",
                    "HLU: Data Integrity", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                return;
            }

            ViewModelWindowMainMerge vmMerge = new ViewModelWindowMainMerge(this);
            //---------------------------------------------------------------------
            // CHANGED: CR39 (Split and merge complete messages)
            // Notify the user following the completion of the split.
            if (vmMerge.PhysicalMerge()) NotifySplitMerge("Physical merge completed.");
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// At least one feature in selection that do not share the same incid or toidfragid
        /// </summary>
        private bool CanLogicallyMerge
        {
            get
            {
                return (_bulkUpdateMode == false && _osmmUpdateMode == false) &&
                    HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) &&
                    _gisSelection != null && _gisSelection.Rows.Count > 1 &&
                    // Only enable split/merge after select from map
                    (_filterByMap == true) &&
                    (_incidsSelectedMapCount > 1) && (_fragsSelectedMapCount > 1);
            }
        }

        /// <summary>
        /// At least one feature in selection that share the same incid and toid but *not* the same toidfragid
        /// </summary>
        private bool CanPhysicallyMerge
        {
            get
            {
                return (_bulkUpdateMode == false && _osmmUpdateMode == false) &&
                    HaveGisApp && EditMode && !String.IsNullOrEmpty(Reason) && !String.IsNullOrEmpty(Process) &&
                    _gisSelection != null && _gisSelection.Rows.Count > 1 &&
                    // Only enable split/merge after select from map
                    (_filterByMap == true) &&
                    (_incidsSelectedMapCount == 1) && (_toidsSelectedMapCount == 1) && (_fragsSelectedMapCount > 1);
            }
        }

        #endregion

        #region Notify SplitMerge

        //---------------------------------------------------------------------
        // CHANGED: CR39 (Split and merge complete messages)
        // Check the options to see if the user wants to be notified
        // following the completion of a split or merge, and display
        // the supplied message if they do.
        //
        /// <summary>
        /// Notify the user following the completion of a split of merge
        /// if the options specify they want to be notified.
        /// </summary>
        private void NotifySplitMerge(string msgText)
        {
            if (_notifyOnSplitMerge)
            {
                _windowWarnSplitMerge = new WindowNotifyOnSplitMerge();
                if ((_windowWarnSplitMerge.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowWarnSplitMerge.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinWarnSplitMerge = new ViewModelWindowNotifyOnSplitMerge(msgText);

                // when ViewModel asks to be closed, close window
                _viewModelWinWarnSplitMerge.RequestClose +=
                    new ViewModelWindowNotifyOnSplitMerge.RequestCloseEventHandler(_viewModelWinWarnSplitMerge_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowWarnSplitMerge.DataContext = _viewModelWinWarnSplitMerge;

                // show window
                _windowWarnSplitMerge.ShowDialog();
            }
        }

        void _viewModelWinWarnSplitMerge_RequestClose()
        {
            _viewModelWinWarnSplitMerge.RequestClose -= _viewModelWinWarnSplitMerge_RequestClose;
            _windowWarnSplitMerge.Close();

            // Update the user notify setting
            _notifyOnSplitMerge = Settings.Default.NotifyOnSplitMerge;

        }
        //---------------------------------------------------------------------

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
            _saving = true;
            _savingAttempted = false;

            // If there are no features selected in the GIS (because there is no
            // active filter).
            if (_incidsSelectedMapCount <= 0)
            {
                // Ask the user before re-selecting the current incid features in GIS.
                if (MessageBox.Show("There are no features selected in the GIS.\n" +
                            "Would you like to apply the changes to all features for this incid?", "HLU: Save Changes",
                            MessageBoxButton.YesNoCancel, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    // Set the status to processing and the cursor to wait.
                    ChangeCursor(Cursors.Wait, "Selecting ...");

                    // Select all features for current incid
                    SelectOnMap(false);

                    // If there are still no features selected in the GIS this suggests
                    // that the feature layer contains only a subset of the database
                    // features so this incid cannot be updated.
                    if (_incidsSelectedMapCount <= 0)
                    {
                        // Reset the cursor back to normal
                        ChangeCursor(Cursors.Arrow, null);

                        return;
                    }

                    // Count the number of toids and fragments for the current incid
                    // selected in the GIS and in the database.
                    CountToidFrags();

                    // Refresh all the status type fields.
                    RefreshStatus();

                }
                else
                {
                    // Reset the cursor back to normal
                    ChangeCursor(Cursors.Arrow, null);

                    return;
                }
            }

            // Reset the cursor back to normal
            ChangeCursor(Cursors.Arrow, null);

            // If in bulk update mode then perform the bulk update and exit.
            if (_bulkUpdateMode == true)
            {
                BulkUpdateClicked(param);
                return;
            }

            // Check if the record has changed and if it hasn't ask the user
            // if they still want to update the record (to create new history).
            //
            // Currently, in theory, this can't happen because the Apply button
            // shouldn't be enabled unless some changes have been made by the
            // user. But this logic is retained just in case.
            MessageBoxResult userResponse = CheckClean();
            switch (userResponse)
            {
                case MessageBoxResult.Yes:
                    break;
                case MessageBoxResult.No:
                    Changed = false;
                    return;
                case MessageBoxResult.Cancel:
                    return;
            }

            // If there is no filter active (and hence all the features for the
            // current incid are to be updated) or all of the features for the
            // current incid have been selected in GIS then update them all and exit.
            if ((!IsFiltered) || (_fragsIncidGisCount == _fragsIncidDbCount))
            {
                // If saving hasn't already been attempted, when the features for
                // the current incid were selected in the map (above), then
                // do the update now.
                if (!_savingAttempted)
                {
                    // Update the current incid.
                    _saving = true;
                    _savingAttempted = false;
                    _viewModelUpd.Update();
                }
                return;
            }

            ChangeCursor(Cursors.Wait, "Filtering ...");

            DispatcherHelper.DoEvents();

            // Recheck the selected features in GIS to make sure they
            // all belong to the current incid.
            _gisSelection = NewGisSelectionTable();
            _gisApp.ReadMapSelection(ref _gisSelection);

            // Count the number of toids and fragments for the current incid
            // selected in the GIS and in the database.
            CountToidFrags();

            // Refresh all the status type fields.
            RefreshStatus();

            ChangeCursor(Cursors.Arrow, null);

            // If there are no features for the current incid
            // selected in GIS then cancel the update.
            if (_fragsIncidGisCount < 1)
            {
                MessageBox.Show("No map features for the current incid are selected in the map.",
                    "HLU: Save Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // If all of the features for the current incid have been
            // selected in GIS then update them all.
            if (_fragsIncidGisCount == _fragsIncidDbCount)
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
                        ChangeCursor(Cursors.Wait, "Splitting ...");

                        // Logically split the features for the current incid into a new incid.
                        ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
                        _splitting = true;
                        if (!vmSplit.LogicalSplit())
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

                    // Recount the number of toids and fragments for the current incid
                    // selected in the GIS and in the database.
                    CountToidFrags();

                    // Refresh all the status type fields.
                    RefreshStatus();
                }
                else
                {
                    MessageBox.Show("The changes have not been applied - the update was cancelled.",
                        "HLU: Save Cancelled", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    //---------------------------------------------------------------------
                    // FIXED: KI116 (Cancelled attribute updates)
                    // 
                    // Clear the saving in progress flags so that the (still) pending
                    // changes aren't automatically applied when moving to another
                    // incid (ore refreshing the current incid).
                    _saving = false;
                    _savingAttempted = true;
                    //---------------------------------------------------------------------
                    return;
                }
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Update is disabled if not currently in edit mode, if no changes have been
        /// made by the user, if we're not currently in bulk update mode with no records
        /// selected, or if the current record is in error.
        /// </summary>
        private bool CanUpdate
        {
            get
            {
                return EditMode &&
                    (Changed == true) &&
                    (_bulkUpdateMode == false || _incidSelection != null) &&
                    String.IsNullOrEmpty(this.Error);
            }
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
                    OnPropertyChanged("CanBulkUpdateMode");
                    OnPropertyChanged("CanOSMMUpdateMode");
                    OnPropertyChanged("CanOSMMBulkUpdateMode");
                    OnPropertyChanged("ShowReasonProcessGroup");
                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    OnPropertyChanged("CanOSMMUpdateAccept");
                    OnPropertyChanged("CanOSMMUpdateReject");
                    //---------------------------------------------------------------------
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
        /// Confirms with the user if the update is to go ahead.
        /// </summary>
        /// <returns>
        /// True if the update is to go ahead, or false if it is cancelled.
        /// </returns>
        private bool ConfirmSubsetUpdate()
        {
            // The user settings indicate that only the selected features
            // should be updated (by logically splitting them first).
            if (_subsetUpdateAction == 1)
            {
                _updateAllFeatures = false;
                return true;
            }
            // The user settings indicate that all the features for the incid
            // should be updated.
            else if (_subsetUpdateAction == 2)
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
                _windowWarnSubsetUpdate.WindowStartupLocation = WindowStartupLocation.CenterScreen;

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

        private void BulkUpdateClicked(object param)
        {
            _saving = false;
            if (_viewModelBulkUpdate == null)
                _viewModelBulkUpdate = new ViewModelWindowMainBulkUpdate(this);

            // If already in bulk update mode then perform the bulk update
            // (only possible when this method was called after the 'Apply'
            // button was clicked.
            if (_bulkUpdateMode == true)
            {
                _viewModelBulkUpdate.BulkUpdate();
            }
            else
            {
                // Check there are no outstanding edits.
                MessageBoxResult userResponse = CheckDirty();

                // Ask the user if they want to apply the
                // outstanding edits.
                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                        // Apply the outstanding edits.
                        if (!_viewModelUpd.Update()) return;
                        break;
                    case MessageBoxResult.No:
                        break;
                    case MessageBoxResult.Cancel:
                        return;
                }

                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Functionality to process proposed OSMM Updates.
                //    
                // Start the bulk update process.
                _viewModelBulkUpdate.StartBulkUpdate(false);
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// Is the user authorised for bulk updates?
        /// </summary>
        public bool CanBulkUpdate
        {
            get
            {
                if (_canBulkUpdate == null) GetUserInfo();

                return _canBulkUpdate == true;
            }
        }

        /// <summary>
        /// Can bulk update mode be started?
        /// </summary>
        public bool CanBulkUpdateMode
        {
            get
            {
                if (_canBulkUpdate == null) GetUserInfo();

                return EditMode &&
                    _canBulkUpdate == true &&
                    _osmmUpdateMode == false &&
                    _osmmBulkUpdateMode == false &&
                    (IsFiltered || _bulkUpdateMode == true);
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
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Functionality to process proposed OSMM Updates.
                //    
                // If the Cancel button has been clicked then we need
                // to work out which mode was active and cancel the
                // right one
                if (_osmmBulkUpdateMode == true)
                    _viewModelBulkUpdate.CancelOSMMBulkUpdate();
                else
                    _viewModelBulkUpdate.CancelBulkUpdate();
                //---------------------------------------------------------------------

                _viewModelBulkUpdate = null;
            }
        }

        public bool CanCancelBulkUpdate { get { return _bulkUpdateMode == true; } }

        internal Nullable<bool> BulkUpdateMode
        {
            get { return _bulkUpdateMode; }
            set { _bulkUpdateMode = value; }
        }

        public Visibility HideInBulkUpdateMode
        {
            get
            {
                if (_bulkUpdateMode == true)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
            set { }
        }

        public Visibility ShowInBulkUpdateMode
        {
            get
            {
                if (_bulkUpdateMode == true)
                    return Visibility.Visible;
                else
                    return Visibility.Hidden;
            }
            set { }
        }

        public string BulkUpdateCommandHeader
        {
            get { return (_bulkUpdateMode == true && _osmmBulkUpdateMode == false) ? "Cancel _Bulk Apply Updates" : "_Bulk Apply Updates"; }
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

        public string TopControlsGroupHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "INCID";
                else
                    return null;
            }
        }

        public string TopControlsBulkUpdateGroupHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Bulk Update";
                else
                    return null;
            }
        }

        #endregion

        #region OSMM Update

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // Display and process proposed OSMM Updates.
        //
        /// <summary>
        /// Gets or sets the OSMM Accept button tag (which controls
        /// the text on the button (and whether the <Ctrl> button is
        /// pressed or not.
        /// </summary>
        /// <value>
        /// The osmm accept tag.
        /// </value>
        public string OSMMAcceptTag
        {
            get { return _osmmAcceptTag; }
            set
            {
                _osmmAcceptTag = value;
                if (_osmmUpdateMode == true)
                    OnPropertyChanged("OSMMAcceptText");
            }
        }

        /// <summary>
        /// Gets or sets the OSMM Reject button tag (which controls
        /// the text on the button (and whether the <Ctrl> button is
        /// pressed or not.
        /// </summary>
        /// <value>
        /// The osmm reject tag.
        /// </value>
        public string OSMMRejectTag
        {
            get { return _osmmRejectTag; }
            set
            {
                _osmmRejectTag = value;
                if (_osmmUpdateMode == true)
                    OnPropertyChanged("OSMMRejectText");
            }
        }

        /// <summary>
        /// Start the OSMM Update mode.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateClicked(object param)
        {
            _saving = false;
            if (_viewModelOSMMUpdate == null)
                _viewModelOSMMUpdate = new ViewModelWindowMainOSMMUpdate(this);

            // If the OSMM update mode is not already started.
            if (_osmmUpdateMode == false)
            {
                // Check there are no outstanding edits.
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

                // Prevent OSMM updates being actioned too quickly.
                _osmmUpdating = false;

                OnPropertyChanged("CanOSMMAccept");
                OnPropertyChanged("CanOSMMSkip");

                // Start the OSMM update mode
                _viewModelOSMMUpdate.StartOSMMUpdate();
            }
        }

        /// <summary>
        /// Can OSMM Update mode be started?
        /// </summary>
        public bool CanOSMMUpdateMode
        {
            get
            {
                if (_canOSMMUpdate == null)
                {
                    // Check if the user can process OSMM Updates.
                    if (CanBulkUpdate)
                    {
                        // Check if there are incid OSMM updates in the database
                        int incidOSMMUpdatesRowCount = (int)_db.ExecuteScalar(String.Format(
                            "SELECT COUNT(*) FROM {0}", _db.QualifyTableName(_hluDS.incid_osmm_updates.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text);

                        if (incidOSMMUpdatesRowCount > 0)
                            _canOSMMUpdate = true;
                        else
                            _canOSMMUpdate = false;
                    }
                    else
                        _canOSMMUpdate = false;
                }

                //---------------------------------------------------------------------
                // FIX: 103 Accept/Reject OSMM updates in edit mode.     
                //
                // Can start OSMM Update mode if in edit mode,
                // and user is authorised,
                // and not currently in bulk update mode.
                return EditMode &&
                        _canOSMMUpdate == true &&
                       _bulkUpdateMode == false;
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// Cancel the OSMM Update mode.
        /// </summary>
        /// <param name="param"></param>
        private void CancelOSMMUpdateClicked(object param)
        {
            if (_viewModelOSMMUpdate != null)
            {
                _osmmUpdatesEmpty = false;
                _viewModelOSMMUpdate.CancelOSMMUpdate();
                _viewModelOSMMUpdate = null;
                // Prevent OSMM updates being actioned too quickly.
                _osmmUpdating = false;
            }
        }

        public bool CanCancelOSMMUpdate { get { return _osmmUpdateMode == true; } }

        /// <summary>
        /// OSMM Skip command.
        /// </summary>
        public ICommand OSMMSkipCommand
        {
            get
            {
                if (_osmmSkipCommand == null)
                {
                    Action<object> osmmSkipAction = new Action<object>(this.OSMMSkipClicked);
                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    _osmmSkipCommand = new RelayCommand(osmmSkipAction, param => this.CanOSMMSkip);
                    //---------------------------------------------------------------------
                }
                return _osmmSkipCommand;
            }
        }

        /// <summary>
        /// Skip the OSMM Update for the current incid.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMSkipClicked(object param)
        {
            // Prevent OSMM updates being actioned too quickly.
            // Mark the OSMM Update row as skipped
            // If there are any OSMM Updates for this incid then store the values.
            if (_osmmUpdating == false && _osmmUpdatesEmpty == false)
            {
                if (IncidOSMMStatus  > 0)
                {
                    _osmmUpdating = true;

                    // Mark the OSMM Update row as skipped
                    _viewModelOSMMUpdate.OSMMUpdate(1);

                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    // Move to the next Incid
                    IncidCurrentRowIndex += 1;
                    //---------------------------------------------------------------------

                    _osmmUpdating = false;
                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    OnPropertyChanged("CanOSMMAccept");
                    OnPropertyChanged("CanOSMMSkip");
                    //---------------------------------------------------------------------
                }
                //---------------------------------------------------------------------
                // FIX: 103 Accept/Reject OSMM updates in edit mode.
                //
                else
                {
                    // Move to the next Incid
                    IncidCurrentRowIndex += 1;
                }
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// Can the proposed OSMM Update for the current incid
        /// be skipped?
        /// </summary>
        public bool CanOSMMSkip
        {
            get
            {
                // Check if there are proposed OSMM Updates
                // for the current filter.
                return (IsFiltered &&
                        _osmmUpdating == false &&
                        _osmmUpdatesEmpty == false);
            }
        }

        /// <summary>
        /// OSMM Accept command.
        /// </summary>
        public ICommand OSMMAcceptCommand
        {
            get
            {
                if (_osmmAcceptCommand == null)
                {
                    Action<object> osmmAcceptAction = new Action<object>(this.OSMMAcceptClicked);
                    _osmmAcceptCommand = new RelayCommand(osmmAcceptAction, param => this.CanOSMMAccept);
                }
                return _osmmAcceptCommand;
            }
        }

        /// <summary>
        /// Accept the OSMM Update for the current incid.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMAcceptClicked(object param)
        {
            // Prevent OSMM updates being actioned too quickly.
            if (_osmmUpdating == false)
            {
                _osmmUpdating = true;

                if (OSMMAcceptTag == "Ctrl")
                {
                    // Mark all the remaining OSMM Update rows as accepted
                    _viewModelOSMMUpdate.OSMMUpdateAll(0);

                    // Reset the button tags
                    OSMMAcceptTag = "";
                    OSMMRejectTag = "";
                }
                else
                {
                    // Mark the OSMM Update row as accepted
                    _viewModelOSMMUpdate.OSMMUpdate(0);

                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    // Move to the next Incid
                    IncidCurrentRowIndex += 1;
                    //---------------------------------------------------------------------
                }

                _osmmUpdating = false;
                //---------------------------------------------------------------------
                // FIX: 103 Accept/Reject OSMM updates in edit mode.
                //
                OnPropertyChanged("CanOSMMAccept");
                OnPropertyChanged("CanOSMMSkip");
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// Can the proposed OSMM Update for the current incid
        /// be processed?
        /// </summary>
        public bool CanOSMMAccept
        {
            get
            {
                //---------------------------------------------------------------------
                // FIX: 103 Accept/Reject OSMM updates in edit mode.
                // Prevent OSMM updates being actioned too quickly.
                // Check if there are no proposed OSMM Updates
                // for the current filter.
                return (_osmmUpdating == false && _osmmUpdatesEmpty == false &&
                    (_incidOSMMUpdatesStatus == null || (_incidOSMMUpdatesStatus > 0 || _incidOSMMUpdatesStatus < -1)));
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// OSMM Reject command.
        /// </summary>
        public ICommand OSMMRejectCommand
        {
            get
            {
                if (_osmmRejectCommand == null)
                {
                    Action<object> osmmRejectAction = new Action<object>(this.OSMMRejectClicked);
                    _osmmRejectCommand = new RelayCommand(osmmRejectAction, param => this.CanOSMMAccept);
                }
                return _osmmRejectCommand;
            }
        }

        /// <summary>
        /// Reject the OSMM Update for the current incid.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMRejectClicked(object param)
        {
            // Prevent OSMM updates being actioned too quickly.
            if (_osmmUpdating == false)
            {
                _osmmUpdating = true;

                if (OSMMRejectTag == "Ctrl")
                {
                    // Mark all the remaining OSMM Update rows as accepted
                    _viewModelOSMMUpdate.OSMMUpdateAll(-99);

                    // Reset the button tags
                    OSMMAcceptTag = "";
                    OSMMRejectTag = "";
                }
                else
                {
                    // Mark the OSMM Update row as rejected
                    _viewModelOSMMUpdate.OSMMUpdate(-99);

                    //---------------------------------------------------------------------
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    //
                    // Move to the next Incid
                    IncidCurrentRowIndex += 1;
                    //---------------------------------------------------------------------
                }

                _osmmUpdating = false;
                //---------------------------------------------------------------------
                // FIX: 103 Accept/Reject OSMM updates in edit mode.
                //
                OnPropertyChanged("CanOSMMAccept");
                OnPropertyChanged("CanOSMMSkip");
                //---------------------------------------------------------------------
            }
        }

        /// <summary>
        /// If the OSMM Update mode active?
        /// </summary>
        internal Nullable<bool> OSMMUpdateMode
        {
            get { return _osmmUpdateMode; }
            set { _osmmUpdateMode = value; }
        }

        /// <summary>
        /// Hide some controls when in OSMM Update mode.
        /// </summary>
        public Visibility HideInOSMMUpdateMode
        {
            get
            {
                if (_osmmUpdateMode == true)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
            set { }
        }

        /// <summary>
        /// Show some controls when in OSMM Update mode.
        /// </summary>
        public Visibility ShowInOSMMUpdateMode
        {
            get
            {
                // Show the group if in osmm update mode
                if (_osmmUpdateMode == true)
                {
                    return Visibility.Visible;
                }
                else
                {
                    return Visibility.Collapsed;
                }
            }
            set { }
        }

        /// <summary>
        /// Set the menu item text depending on whether in OSMM Update mode.
        /// </summary>
        public string OSMMUpdateCommandHeader
        {
            get { return _osmmUpdateMode == true ? "Cancel Review OSMM Updates" : "Review OSMM Updates"; }
        }

        /// <summary>
        /// OSMM Update menu command.
        /// </summary>
        public ICommand OSMMUpdateCommandMenu
        {
            get
            {
                if (_osmmUpdateCommandMenu == null)
                {
                    Action<object> osmmUpdateMenuAction = new Action<object>(this.OSMMUpdateCommandMenuClicked);
                    _osmmUpdateCommandMenu = new RelayCommand(osmmUpdateMenuAction);
                }
                return _osmmUpdateCommandMenu;
            }
        }

        /// <summary>
        /// Start or cancel the OSMM Update mode (as appropriate).
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateCommandMenuClicked(object param)
        {
            // If already in OSMM update mode
            if (_osmmUpdateMode == true)
                // Cancel the OSMM update mode
                CancelOSMMUpdateClicked(param);
            else
                // Start the OSMM update mode
                OSMMUpdateClicked(param);
        }

        /// <summary>
        /// Whether to create incid history for processing OSMM Updates.
        /// </summary>
        public bool OSMMUpdateCreateHistory
        {
            get { return _osmmUpdateCreateHistory; }
            set { _osmmUpdateCreateHistory = value; }
        }

        /// <summary>
        /// Set the Reject button caption depending on whether the Ctrl button
        /// is held down.
        /// </summary>
        public string OSMMRejectText
        {
            get { return OSMMRejectTag == "Ctrl" ? "Re_ject All" : "Re_ject"; }
        }

        /// <summary>
        /// Set the Accept button caption depending on whether the Ctrl button
        /// is held down.
        /// </summary>
        public string OSMMAcceptText
        {
            get { return OSMMAcceptTag == "Ctrl" ? "A_ccept All" : "A_ccept"; }
        }

        /// <summary>
        /// Get the row counter for the current incid.
        /// </summary>
        public int OSMMIncidCurrentRowIndex
        {
            get { return _osmmUpdatesEmpty ? 0 : _incidCurrentRowIndex; }
        }
        //---------------------------------------------------------------------

        #endregion

        #region OSMM Update Edits

        //---------------------------------------------------------------------
        // FIX: 103 Accept/Reject OSMM updates in edit mode.
        //
        /// <summary>
        /// Accept the proposed OSMM Update.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateAcceptClicked()
        {
            if (_viewModelOSMMUpdate == null)
                _viewModelOSMMUpdate = new ViewModelWindowMainOSMMUpdate(this);

            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            // Save the current incid
            int incidCurrRowIx = IncidCurrentRowIndex;
            //---------------------------------------------------------------------

            // Mark the OSMM Update row as accepted
            _viewModelOSMMUpdate.OSMMUpdate(0);

            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            // Reload the incid
            IncidCurrentRowIndex = incidCurrRowIx;

            //OnPropertyChanged("CanOSMMUpdateAccept");
            //---------------------------------------------------------------------

        }

        /// <summary>
        /// Can OSMM Update be accepted?
        /// </summary>
        public bool CanOSMMUpdateAccept
        {
            get
            {
                // If not in a bulk mode and a proposed OSMM update is showing
                if (EditMode &&
                    _osmmUpdateMode == false &&
                    _bulkUpdateMode == false &&
                    (_showOSMMUpdates == "Always" ||
                    _showOSMMUpdates == "When Outstanding") &&
                    IncidOSMMStatus > 0)
                    return true;
                else
                    return false;
            }
        }

        /// <summary>
        /// Reject the proposed OSMM Update.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateRejectClicked()
        {
            if (_viewModelOSMMUpdate == null)
                _viewModelOSMMUpdate = new ViewModelWindowMainOSMMUpdate(this);

            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            // Save the current incid
            int incidCurrRowIx = IncidCurrentRowIndex;
            //---------------------------------------------------------------------

            // Mark the OSMM Update row as rejected
            _viewModelOSMMUpdate.OSMMUpdate(-99);

            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            // Reload the incid
            IncidCurrentRowIndex = incidCurrRowIx;

            //OnPropertyChanged("CanOSMMUpdateReject");
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// Can OSMM Update be rejected?
        /// </summary>
        public bool CanOSMMUpdateReject
        {
            get
            {
                // If not in a bulk mode and a proposed OSMM update is showing
                if (EditMode &&
                    _osmmUpdateMode == false &&
                    _bulkUpdateMode == false &&
                    (_showOSMMUpdates == "Always" ||
                    _showOSMMUpdates == "When Outstanding") &&
                    IncidOSMMStatus > 0)
                    return true;
                else
                    return false;
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region OSMM Bulk Update

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // Bulk apply pending OSMM Updates.
        //
        /// <summary>
        /// Start the OSMM Bulk Update mode.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMBulkUpdateClicked(object param)
        {
            _saving = false;
            if (_viewModelBulkUpdate == null)
                _viewModelBulkUpdate = new ViewModelWindowMainBulkUpdate(this);

            // If the OSMM Bulk update mode is not already started.
            if (_osmmBulkUpdateMode == false)
            {
                // Check there are no outstanding edits.
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

                // Start the OSMM update mode
                _viewModelBulkUpdate.StartBulkUpdate(true);
            }
        }

        /// <summary>
        /// Can OSMM Bulk Update mode be started?
        /// </summary>
        public bool CanOSMMBulkUpdateMode
        {
            get
            {
                if (_canOSMMUpdate == null)
                {
                    // Check if the user can process OSMM Updates.
                    if (CanBulkUpdate)
                    {
                        // Check if there are incid OSMM updates in the database
                        int incidOSMMUpdatesRowCount = (int)_db.ExecuteScalar(String.Format(
                            "SELECT COUNT(*) FROM {0}", _db.QualifyTableName(_hluDS.incid_osmm_updates.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text);

                        if (incidOSMMUpdatesRowCount > 0)
                            _canOSMMUpdate = true;
                        else
                            _canOSMMUpdate = false;
                    }
                    else
                        _canOSMMUpdate = false;
                }

                // Can start OSMM Bulk Update mode if in edit mode,
                // and user is authorised,
                // and there are incid OSMM updates in the database,
                // and not currently in bulk update mode or osmm update mode.
                return EditMode &&
                    _canOSMMUpdate == true &&
                    _osmmUpdateMode == false &&
                    (_bulkUpdateMode == false || (_bulkUpdateMode == true && _osmmBulkUpdateMode == true));
            }
        }

        /// <summary>
        /// Cancel the OSMM Bulk Update mode.
        /// </summary>
        /// <param name="param"></param>
        private void CancelOSMMBulkUpdateClicked(object param)
        {
            if (_viewModelBulkUpdate != null)
            {
                _viewModelBulkUpdate.CancelOSMMBulkUpdate();
                _viewModelBulkUpdate = null;
            }
        }

        public bool CanCancelOSMMBulkUpdate { get { return _osmmBulkUpdateMode == true; } }

        /// <summary>
        /// If the OSMM Bulk Update mode active?
        /// </summary>
        internal Nullable<bool> OSMMBulkUpdateMode
        {
            get { return _osmmBulkUpdateMode; }
            set { _osmmBulkUpdateMode = value; }
        }

        /// <summary>
        /// Set the menu item text depending on whether in OSMM Bulk Update mode.
        /// </summary>
        public string OSMMBulkUpdateCommandHeader
        {
            get { return _osmmBulkUpdateMode == true ? "Cancel Bulk Apply OSMM Updates" : "Bulk Apply OSMM Updates"; }
        }

        /// <summary>
        /// OSMM Bulk Update menu command.
        /// </summary>
        public ICommand OSMMBulkUpdateCommandMenu
        {
            get
            {
                if (_osmmBulkUpdateCommandMenu == null)
                {
                    Action<object> osmmBulkUpdateMenuAction = new Action<object>(this.OSMMBulkUpdateCommandMenuClicked);
                    _osmmBulkUpdateCommandMenu = new RelayCommand(osmmBulkUpdateMenuAction);
                }
                return _osmmBulkUpdateCommandMenu;
            }
        }

        /// <summary>
        /// Start or cancel the OSMM Bulk Update mode (as appropriate).
        /// </summary>
        /// <param name="param"></param>
        private void OSMMBulkUpdateCommandMenuClicked(object param)
        {
            // If already in OSMM Bulk update mode
            if (_osmmBulkUpdateMode == true)
                // Cancel the OSMM Bulk update mode
                CancelOSMMBulkUpdateClicked(param);
            else
                // Start the OSMM Bulk update mode
                OSMMBulkUpdateClicked(param);
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIX: 103 Accept/Reject OSMM updates in edit mode.
        //
        /// <summary>
        /// OSMM Update Accept menu command.
        /// </summary>
        public ICommand OSMMUpdateAcceptCommandMenu
        {
            get
            {
                if (_osmmUpdateAcceptCommandMenu == null)
                {
                    Action<object> osmmUpdateAcceptMenuAction = new Action<object>(this.OSMMUpdateAcceptCommandMenuClicked);
                    _osmmUpdateAcceptCommandMenu = new RelayCommand(osmmUpdateAcceptMenuAction);
                }
                return _osmmUpdateAcceptCommandMenu;
            }
        }

        /// <summary>
        /// Accept the OSMM proposed update.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateAcceptCommandMenuClicked(object param)
        {
            // Accept the OSMM proposed update
            OSMMUpdateAcceptClicked();
        }

        /// <summary>
        /// OSMM Update Reject menu command.
        /// </summary>
        public ICommand OSMMUpdateRejectCommandMenu
        {
            get
            {
                if (_osmmUpdateRejectCommandMenu == null)
                {
                    Action<object> osmmUpdateRejectMenuAction = new Action<object>(this.OSMMUpdateRejectCommandMenuClicked);
                    _osmmUpdateRejectCommandMenu = new RelayCommand(osmmUpdateRejectMenuAction);
                }
                return _osmmUpdateRejectCommandMenu;
            }
        }

        /// <summary>
        /// Reject the OSMM proposed update.
        /// </summary>
        /// <param name="param"></param>
        private void OSMMUpdateRejectCommandMenuClicked(object param)
        {
            // Reject the OSMM proposed update
            OSMMUpdateRejectClicked();
        }
        //---------------------------------------------------------------------

        #endregion

        #region View

        public ICommand AppKeepOnTopCommand
        {
            get
            {
                if (_appKeepOnTopCommand == null)
                {
                    Action<object> appKeepOnTopAction = new Action<object>(this.AppKeepOnTopClicked);
                    _appKeepOnTopCommand = new RelayCommand(appKeepOnTopAction);
                }
                return _appKeepOnTopCommand;
            }
        }

        private void AppKeepOnTopClicked(object param)
        {
            // Update the keep app on top option.
            _appKeepOnTop = !_appKeepOnTop;

            // Save the keep app window on top option in the user settings.
            Settings.Default.AppKeepOnTop = _appKeepOnTop;
            Settings.Default.Save();

        }

        public ICommand AutoZoomSelectedOffCommand
        {
            get
            {
                if (_autoZoomSelectedOffCommand == null)
                {
                    Action<object> autoZoomSelectionOffAction = new Action<object>(this.AutoZoomSelectedOffClicked);
                    _autoZoomSelectedOffCommand = new RelayCommand(autoZoomSelectionOffAction, param => this.CanAutoZoomSelected);
                }
                return _autoZoomSelectedOffCommand;
            }
        }

        private void AutoZoomSelectedOffClicked(object param)
        {
            // Update the auto zoom on selected option.
            _autoZoomSelection = (int)AutoZoomSelection.Off;

            // Save the new auto zoom option in the user settings.
            Settings.Default.AutoZoomSelection = _autoZoomSelection;
            Settings.Default.Save();
        }

        public ICommand AutoZoomSelectedWhenCommand
        {
            get
            {
                if (_autoZoomSelectedWhenCommand == null)
                {
                    Action<object> autoZoomSelectionWhenAction = new Action<object>(this.AutoZoomSelectedWhenClicked);
                    _autoZoomSelectedWhenCommand = new RelayCommand(autoZoomSelectionWhenAction, param => this.CanAutoZoomSelected);
                }
                return _autoZoomSelectedWhenCommand;
            }
        }

        private void AutoZoomSelectedWhenClicked(object param)
        {
            // Update the auto zoom on selected option.
            _autoZoomSelection = (int)AutoZoomSelection.When;

            // Save the new auto zoom option in the user settings.
            Settings.Default.AutoZoomSelection = _autoZoomSelection;
            Settings.Default.Save();
        }

        public ICommand AutoZoomSelectedAlwaysCommand
        {
            get
            {
                if (_autoZoomSelectedAlwaysCommand == null)
                {
                    Action<object> autoZoomSelectionAlwaysAction = new Action<object>(this.AutoZoomSelectedAlwaysClicked);
                    _autoZoomSelectedAlwaysCommand = new RelayCommand(autoZoomSelectionAlwaysAction, param => this.CanAutoZoomSelected);
                }
                return _autoZoomSelectedAlwaysCommand;
            }
        }

        private void AutoZoomSelectedAlwaysClicked(object param)
        {
            // Update the auto zoom on selected option.
            _autoZoomSelection = (int)AutoZoomSelection.Always;

            // Save the new auto zoom option in the user settings.
            Settings.Default.AutoZoomSelection = _autoZoomSelection;
            Settings.Default.Save();
        }

        public bool CanAutoZoomSelected { get { return HaveGisApp; } }

        public ICommand AutoSelectOnGisCommand
        {
            get
            {
                if (_autoSelectOnGisCommand == null)
                {
                    Action<object> autoSelectiOnGisAction = new Action<object>(this.AutoSelectOnGisClicked);
                    _autoSelectOnGisCommand = new RelayCommand(autoSelectiOnGisAction, param => this.CanAutoSelectOnGis);
                }
                return _autoSelectOnGisCommand;
            }
        }

        private void AutoSelectOnGisClicked(object param)
        {
            // Update the auto select on GIS option.
            _autoSelectOnGis = !_autoSelectOnGis;

            // Save the new auto select on GIS option in the user settings.
            Settings.Default.AutoSelectOnGis = _autoSelectOnGis;
            Settings.Default.Save();

        }

        public bool CanAutoSelectOnGis { get { return HaveGisApp; } }

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
            // Get the minimum auto zoom value and map distance units.
            string distUnits = Settings.Default.MapDistanceUnits;

            _gisApp.ZoomSelected(_minZoom, distUnits, true);
        }

        public bool CanZoomSelection { get { return HaveGisApp && _gisSelection != null; } }

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

        public ICommand ResetToolWindowCommand
        {
            get
            {
                if (_resetToolWindowCommand == null)
                {
                    Action<object> resetToolWindowAction = new Action<object>(this.ResetToolWindowClicked);
                    _resetToolWindowCommand = new RelayCommand(resetToolWindowAction, param => this.CanResetToolWindow);
                }
                return _resetToolWindowCommand;
            }
        }

        private void ResetToolWindowClicked(object param)
        {
            // Adjust the window height.
            AdjustWindowHeight(true);
            OnPropertyChanged("WindowHeight");
            WindowWidth = 0;
            OnPropertyChanged("WindowWidth");
        }

        public bool CanResetToolWindow { get { return HaveGisApp; } }

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
            _windowOptions.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            _viewModelOptions = new ViewModelOptions(this);
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
                // Database options
                _dbConnectionTimeout = Settings.Default.DbConnectionTimeout;

                // GIS/Export options
                int _minZoom = Settings.Default.MinAutoZoom;

                // History options
                _historyDisplayLastN = Settings.Default.HistoryDisplayLastN;
                _historyColumns = InitializeHistoryColumns(_historyColumns);

                // Interface options
                _preferredHabitatClass = Settings.Default.PreferredHabitatClass;
                _showGroupHeaders = Settings.Default.ShowGroupHeaders;
                _showNVCCodes = Settings.Default.ShowNVCCodes;
                _showIHSTab = Settings.Default.ShowIHSTab;
                _showOSMMUpdates = Settings.Default.ShowOSMMUpdatesOption;
                _preferredSecondaryGroup = Settings.Default.PreferredSecondaryGroup;
                _secondaryCodeOrder = Settings.Default.SecondaryCodeOrder;
                _secondaryCodeDelimiter = Settings.Default.SecondaryCodeDelimiter;

                // Updates options
                _subsetUpdateAction = Settings.Default.SubsetUpdateAction;
                _clearIHSUpdateAction = Settings.Default.ClearIHSUpdateAction;
                _notifyOnSplitMerge = Settings.Default.NotifyOnSplitMerge;
                _resetOSMMUpdatesStatus = Settings.Default.ResetOSMMUpdatesStatus;
                _secondaryCodeValidation = Settings.Default.SecondaryCodeValidation;
                SecondaryHabitat.SecondaryCodeValidation = _secondaryCodeValidation; // Set in the secondary habitat environment
                _qualityValidation = Settings.Default.QualityValidation;

                // Filter options
                _warnBeforeGISSelect = Settings.Default.WarnBeforeGISSelect;

                // Dates options
                VagueDate.Delimiter = Settings.Default.VagueDateDelimiter;
                VagueDate.SeasonNames = Settings.Default.SeasonNames.Cast<string>().ToArray();

                // Bulk Update options
                //None

                // Refresh the user interface
                RefreshGroupHeaders();
                OnPropertyChanged("ShowNVCCodes");
                OnPropertyChanged("ShowIHSTab");
                OnPropertyChanged("ShowIncidOSMMPendingGroup");

                OnPropertyChanged("SecondaryGroupCodes");
                OnPropertyChanged("SecondaryHabitatCodes");
                RefreshSecondaryHabitats();
                OnPropertyChanged("HabitatTabLabel");
                OnPropertyChanged("IncidSecondarySummary");

                OnPropertyChanged("IncidCondition");
                OnPropertyChanged("IncidConditionQualifier");
                OnPropertyChanged("IncidConditionDate");

                OnPropertyChanged("IncidQualityDetermination");
                OnPropertyChanged("IncidQualityInterpretation");
                OnPropertyChanged("IncidQualityComments");

                // Adjust the window height.
                AdjustWindowHeight(false);
                OnPropertyChanged("WindowHeight");
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
                // Split the copyright statement at each full stop and
                // wrap it to a new line.
                String copyright = String.Join(Environment.NewLine, ((AssemblyCopyrightAttribute)attributes[0]).Copyright.Split('.'));
                // If there is a Copyright attribute, return its value
                return copyright;
            }
        }
        //---------------------------------------------------------------------

        private void AboutClicked(object param)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR30 (Database validation on startup)
            // Show the database version in the 'About' box.
            //
            // CHANGED: CR9 (Current userid)
            // Show the current userid and username together with the version
            // and copyright notice in the 'About' box.
            //
            string dbBackend;
            dbBackend = String.Format("{0}{1}{2}{3}",
                _db.Backend.ToString(),
                String.IsNullOrEmpty(_db.DefaultSchema) ? null : " (",
                _db.DefaultSchema,
                String.IsNullOrEmpty(_db.DefaultSchema) ? null : ")");
            string dbSettings;
            dbSettings = _db.ConnectionString.Replace(";", "\n");
            //---------------------------------------------------------------------

            _windowAbout = new WindowAbout();
            _windowAbout.Owner = App.Current.MainWindow;
            _windowAbout.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // Create ViewModel to which main window binds
            _viewModelAbout = new ViewModelWindowAbout();
            _viewModelAbout.AppVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            _viewModelAbout.DbVersion = _dbVersion;
            _viewModelAbout.ConnectionType = dbBackend;
            _viewModelAbout.ConnectionSettings = dbSettings;
            _viewModelAbout.UserId = UserID;
            _viewModelAbout.UserName = UserName;
            _viewModelAbout.Copyright = AssemblyCopyright;
            _viewModelAbout.UserGuideURL = "https://readthedocs.org/projects/hlutool-userguide/";
            _viewModelAbout.UserGuideText = "https://readthedocs.org/projects/hlutool-userguide/";
            _viewModelAbout.TechnicalGuideURL = "https://readthedocs.org/projects/hlutool-technicalguide/";
            _viewModelAbout.TechnicalGuideText = "https://readthedocs.org/projects/hlutool-technicalguide/";

            // When ViewModel asks to be closed, close window
            _viewModelAbout.RequestClose += new ViewModelWindowAbout.RequestCloseEventHandler(_viewModelAbout_RequestClose);

            // Allow all controls in window to bind to ViewModel by setting DataContext
            _windowAbout.DataContext = _viewModelAbout;

            // Show window
            _windowAbout.ShowDialog();
        }

        /// <summary>
        /// Closes help window and removes close window handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        internal void _viewModelAbout_RequestClose()
        {
            _viewModelAbout.RequestClose -= _viewModelAbout_RequestClose;
            _windowAbout.Close();
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

        /// <summary>
        /// Initiates the exports process.
        /// </summary>
        /// <param name="param">The parameter.</param>
        private void ExportClicked(object param)
        {
            ViewModelWindowMainExport viewModelExport = new ViewModelWindowMainExport(this);
            viewModelExport.InitiateExport();
        }

        public bool CanExport { get { return _bulkUpdateMode == false && _osmmUpdateMode == false && _hluDS != null && HaveGisApp; } }

        #endregion

        #region Filter by Attributes Command

        /// <summary>
        /// FilterByAttributes command.
        /// </summary>
        public ICommand FilterByAttributesCommand
        {
            get
            {
                if (_filterByAttributesCommand == null)
                {
                    Action<object> filterByAttributesAction = new Action<object>(this.FilterByAttributesClicked);
                    _filterByAttributesCommand = new RelayCommand(filterByAttributesAction, param => this.CanFilterByAttributes);
                }
                return _filterByAttributesCommand;
            }
        }

        /// <summary>
        /// Opens the relevant query window based on the mode/options.
        /// </summary>
        /// <param name="param">The parameter.</param>
        private void FilterByAttributesClicked(object param)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Open the OSMM Updates query window if in OSMM Update mode.
            //
            if (_osmmUpdateMode == true || _osmmBulkUpdateMode == true)
                OpenWindowQueryOSMM(false);
            else
            {
                // Open the select by attributes interface
                OpenWindowQueryAdvanced();
            }
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// Gets a value indicating whether the filter by attributes command can
        /// be executed.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance can filter by attributes; otherwise, <c>false</c>.
        /// </value>
        private bool CanFilterByAttributes
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Enable filter when in OSMM bulk update mode
            // 
            get
            {
                return (_bulkUpdateMode == false || (_bulkUpdateMode == true && _osmmBulkUpdateMode == true))
                && IncidCurrentRow != null;
            }
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// FilterByAttributesOSMM command.
        /// </summary>
        public ICommand FilterByAttributesOSMMCommand
        {
            get
            {
                if (_filterByAttributesOSMMCommand == null)
                {
                    Action<object> filterByAttributesOSMMAction = new Action<object>(this.FilterByAttributesOSMMClicked);
                    _filterByAttributesOSMMCommand = new RelayCommand(filterByAttributesOSMMAction, param => this.CanFilterByAttributesOSMM);
                }
                return _filterByAttributesOSMMCommand;
            }
        }

        /// <summary>
        /// Opens the relevant query window based on the mode/options.
        /// </summary>
        /// <param name="param">The parameter.</param>
        private void FilterByAttributesOSMMClicked(object param)
        {
            // Open the Advanced query window.
            OpenWindowQueryOSMMAdvanced(false);
        }

        /// <summary>
        /// Gets a value indicating whether the filter by attributes OSMM command can
        /// be executed.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance can filter by attributes OSMM; otherwise, <c>false</c>.
        /// </value>
        private bool CanFilterByAttributesOSMM
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Enable filter when in OSMM bulk update mode
            // 
            get
            {
                return (_osmmUpdateMode == true && IncidCurrentRow != null);
            }
            //---------------------------------------------------------------------
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Opens the new advanced query window.
        /// </summary>
        /// <exception cref="System.Exception">No parent window loaded</exception>
        private void OpenWindowQueryAdvanced()
        {
            try
            {
                _windowQueryAdvanced = new WindowQueryAdvanced();
                if ((_windowQueryAdvanced.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowQueryAdvanced.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinQueryAdvanced = new ViewModelWindowQueryAdvanced(HluDataset, _db);
                _viewModelWinQueryAdvanced.DisplayName = "Advanced Query Builder";

                // when ViewModel asks to be closed, close window
                _viewModelWinQueryAdvanced.RequestClose +=
                    new ViewModelWindowQueryAdvanced.RequestCloseEventHandler(_viewModelWinQueryAdvanced_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowQueryAdvanced.DataContext = _viewModelWinQueryAdvanced;

                // show window
                _windowQueryAdvanced.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        /// <summary>
        /// Process the sql when the advanced query window is closed.
        /// </summary>
        /// <param name="sqlFromTables">The tables to query.</param>
        /// <param name="sqlWhereClause">The where clause to apply in the query.</param>
        protected void _viewModelWinQueryAdvanced_RequestClose(string sqlFromTables, string sqlWhereClause)
        {
            _viewModelWinQueryAdvanced.RequestClose -= _viewModelWinQueryAdvanced_RequestClose;
            _windowQueryAdvanced.Close();

            if ((sqlFromTables != null) && (sqlWhereClause != null))
            {
                try
                {
                    ChangeCursor(Cursors.Wait, "Validating ...");

                    // Get a list of all the possible query tables.
                    List<DataTable> tables = new List<DataTable>();
                    if ((ViewModelWindowQueryAdvanced.HluDatasetStatic != null))
                    {
                        tables = ViewModelWindowQueryAdvanced.HluDatasetStatic.incid.ChildRelations
                            .Cast<DataRelation>().Select(r => r.ChildTable).ToList();
                        tables.Add(ViewModelWindowQueryAdvanced.HluDatasetStatic.incid);
                    }

                    // Split the string of query table names created by the
                    // user in the form into an array.
                    string[] fromTables = sqlFromTables.Split(',').Select(s => s.Trim(' ')).Distinct().ToArray();

                    // Select only the database tables that are in the query array.
                    List<DataTable> whereTables = tables.Where(t => fromTables.Contains(t.TableName)).ToList();

                    // Backup the current selection (filter).
                    DataTable incidSelectionBackup = _incidSelection;

                    // create a selection DataTable of PK values of IncidTable
                    if (whereTables.Count() > 0)
                    {

                        // Replace any connection type specific qualifiers and delimiters.
                        string newWhereClause = null;
                        if (sqlWhereClause != null)
                            newWhereClause = ReplaceStringQualifiers(sqlWhereClause);

                        // Create a selection DataTable of PK values of IncidTable.
                        _incidSelection = _db.SqlSelect(true, IncidTable.PrimaryKey, whereTables, newWhereClause);

                        // Get a list of all the incids in the selection.
                        _incidsSelectedMap = _incidSelection.AsEnumerable()
                            .GroupBy(r => r.Field<string>(_incidSelection.Columns[0].ColumnName)).Select(g => g.Key).OrderBy(s => s);

                        // Retrospectively set the where clause to match the list
                        // of selected incids (for possible use later).
                        _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                            IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);
                    }
                    else
                    {
                        // Clear the selection of incids.
                        _incidSelection = null;

                        // Clear the previous where clause (set when performing the
                        // original query builder or when reading the map selection)
                        // because otherwise it might be used in error later.
                        _incidSelectionWhereClause = null;
                    }

                    // If there are any records in the selection (and the tool is
                    // not currently in bulk update mode).
                    if (IsFiltered)
                    {
                        // Find the expected number of features to be selected in GIS.
                        _toidsSelectedDBCount = 0;
                        _fragsSelectedDBCount = 0;
                        ExpectedSelectionFeatures(whereTables, sqlWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                        //---------------------------------------------------------------------
                        // CHANGED: CR12 (Select by attribute performance)
                        // Store the number of incids found in the database
                        _incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;
                        //---------------------------------------------------------------------

                        ChangeCursor(Cursors.Wait, "Filtering ...");
                        // Select the required incid(s) in GIS.
                        if (PerformGisSelection(true, _fragsSelectedDBCount, _incidsSelectedDBCount))
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR21 (Select current incid in map)
                            // Analyse the results, set the filter and reset the cursor AFTER
                            // returning from performing the GIS selection so that other calls
                            // to the PerformGisSelection method can control if/when these things
                            // are done.
                            //
                            // Analyse the results of the GIS selection by counting the number of
                            // incids, toids and fragments selected.
                            AnalyzeGisSelectionSet(true);

                            // Indicate the selection didn't come from the map.
                            _filterByMap = false;

                            // Set the filter back to the first incid.
                            SetFilter();

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);

                            // Warn the user that no records were found.
                            if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                                MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "HLU Query",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            //---------------------------------------------------------------------
                        }
                        else
                        {
                            // Restore the previous selection (filter).
                            _incidSelection = incidSelectionBackup;

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);
                        }
                    }
                    else
                    {
                        // Restore the previous selection (filter).
                        _incidSelection = incidSelectionBackup;

                        // Reset the cursor back to normal
                        ChangeCursor(Cursors.Arrow, null);

                        // Warn the user that no records were found
                        MessageBox.Show(App.Current.MainWindow, "No records found.", "HLU Query",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    _incidSelection = null;
                    ChangeCursor(Cursors.Arrow, null);
                    MessageBox.Show(App.Current.MainWindow, ex.Message, "HLU Query",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally { RefreshStatus(); }
            }
        }

        /// <summary>
        /// Opens the warning on gis selection window to prompt the user
        /// for confirmation before proceeding.
        /// </summary>
        /// <param name="selectByjoin">if set to <c>true</c> [select byjoin].</param>
        /// <param name="expectedNumFeatures">The expected number features.</param>
        /// <param name="expectedNumIncids">The expected number incids.</param>
        /// <returns></returns>
        /// <exception cref="Exception">No parent window loaded</exception>
        private bool ConfirmGISSelect(bool selectByjoin, int expectedNumFeatures, int expectedNumIncids)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR12 (Select by attribute performance)
            // Warn the user either if the user option is set to
            // 'Always' or if a GIS table join will be used and
            // the user option is set to 'Joins'.
            if ((_warnBeforeGISSelect == 0) ||
                (selectByjoin && _warnBeforeGISSelect == 1))
            //---------------------------------------------------------------------
            {
                _windowWarnGISSelect = new WindowWarnOnGISSelect();
                if ((_windowWarnGISSelect.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowWarnGISSelect.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // Create ViewModel to which main window binds
                _viewModelWinWarnGISSelect = new ViewModelWindowWarnOnGISSelect(
                    expectedNumFeatures, expectedNumIncids, expectedNumFeatures > -1 ? _gisLayerType : GeometryTypes.Unknown, selectByjoin);

                // When ViewModel asks to be closed, close window
                _viewModelWinWarnGISSelect.RequestClose +=
                    new ViewModelWindowWarnOnGISSelect.RequestCloseEventHandler(_viewModelWinWarnGISSelect_RequestClose);

                // Allow all controls in window to bind to ViewModel by setting DataContext
                _windowWarnGISSelect.DataContext = _viewModelWinWarnGISSelect;

                // Show the window
                _windowWarnGISSelect.ShowDialog();

                return IsFiltered;
            }
            else
            {
                // Return true if the user has not been warned.
                return true;
            }
        }

        /// <summary>
        /// Closes the warning gis on selection window.
        /// </summary>
        /// <param name="proceed">if set to <c>true</c> [proceed].</param>
        void _viewModelWinWarnGISSelect_RequestClose(bool proceed)
        {
            _viewModelWinWarnGISSelect.RequestClose -= _viewModelWinWarnGISSelect_RequestClose;
            _windowWarnGISSelect.Close();

            // Update the user warning variable
            _warnBeforeGISSelect = Settings.Default.WarnBeforeGISSelect;

            // If the user doesn't wish to proceed then clear the
            // current incid filter.
            if (!proceed)
            {
                _incidSelectionWhereClause = null;
                _incidSelection = null;
                ChangeCursor(Cursors.Arrow, null);
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Filter by Incid Command

        /// <summary>
        /// FilterByIncid command.
        /// </summary>
        public ICommand FilterByIncidCommand
        {
            get
            {
                if (_filterByIncidCommand == null)
                {
                    Action<object> filterByIncidAction = new Action<object>(this.FilterByIncidClicked);
                    _filterByIncidCommand = new RelayCommand(filterByIncidAction, param => this.CanFilterByIncid);
                }
                return _filterByIncidCommand;
            }
        }

        /// <summary>
        /// Opens the by filter by incid window.
        /// </summary>
        /// <param name="param">The parameter.</param>
        private void FilterByIncidClicked(object param)
        {
            OpenQueryIncid();
        }

        /// <summary>
        /// Gets a value indicating whether the filter by incid command can
        /// be clicked.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance can filter by incid; otherwise, <c>false</c>.
        /// </value>
        private bool CanFilterByIncid
        {
            get { return _bulkUpdateMode == false && _osmmUpdateMode == false && IncidCurrentRow != null; }
        }

        /// <summary>
        /// Opens the query by incid window.
        /// </summary>
        /// <exception cref="Exception">No parent window loaded</exception>
        private void OpenQueryIncid()
        {
            try
            {
                _windowQueryIncid = new WindowQueryIncid();
                if ((_windowQueryIncid.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowQueryIncid.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinQueryIncid = new ViewModelWindowQueryIncid();
                _viewModelWinQueryIncid.DisplayName = "Filter By Incid";

                // when ViewModel asks to be closed, close window
                _viewModelWinQueryIncid.RequestClose +=
                    new ViewModelWindowQueryIncid.RequestCloseEventHandler(_viewModelWinQueryIncid_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowQueryIncid.DataContext = _viewModelWinQueryIncid;

                // show window
                _windowQueryIncid.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        /// <summary>
        /// Closes the query incid window and select the required incid.
        /// </summary>
        /// <param name="queryIncid">The query incid.</param>
        protected void _viewModelWinQueryIncid_RequestClose(String queryIncid)
        {
            _viewModelWinQueryIncid.RequestClose -= _viewModelWinQueryIncid_RequestClose;
            _windowQueryIncid.Close();

            if (!String.IsNullOrEmpty(queryIncid))
            {
                try
                {
                    ChangeCursor(Cursors.Wait, "Validating ...");

                    // Select only the incid database table to use in the query.
                    List<DataTable> whereTables = new List<DataTable>();
                    whereTables.Add(IncidTable);

                    // Replace any connection type specific qualifiers and delimiters.
                    string newWhereClause = null;

                    // Ensure predicted count of toids/fragment selected works with
                    // any query.
                    string sqlWhereClause = String.Format("[incid].incid = '{0}'", queryIncid);

                    newWhereClause = ReplaceStringQualifiers(sqlWhereClause);

                    // Create a selection DataTable of PK values of IncidTable.
                    _incidSelection = _db.SqlSelect(true, IncidTable.PrimaryKey, whereTables, newWhereClause);

                    // Get a list of all the incids in the selection.
                    _incidsSelectedMap = _incidSelection.AsEnumerable()
                        .GroupBy(r => r.Field<string>(_incidSelection.Columns[0].ColumnName)).Select(g => g.Key).OrderBy(s => s);

                    // Retrospectively set the where clause to match the list
                    // of selected incids (for possible use later).
                    _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                        IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);

                    // Backup the current selection (filter).
                    DataTable incidSelectionBackup = _incidSelection;

                    // If there are any records in the selection (and the tool is
                    // not currently in bulk update mode).
                    if (IsFiltered)
                    {
                        // Find the expected number of features to be selected in GIS.
                        _toidsSelectedDBCount = 0;
                        _fragsSelectedDBCount = 0;
                        ExpectedSelectionFeatures(whereTables, sqlWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                        //---------------------------------------------------------------------
                        // CHANGED: CR12 (Select by attribute performance)
                        // Store the number of incids found in the database
                        _incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;
                        //---------------------------------------------------------------------

                        ChangeCursor(Cursors.Wait, "Filtering ...");
                        // Select the required incid(s) in GIS.
                        if (PerformGisSelection(true, _fragsSelectedDBCount, _incidsSelectedDBCount))
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR21 (Select current incid in map)
                            // Analyse the results, set the filter and reset the cursor AFTER
                            // returning from performing the GIS selection so that other calls
                            // to the PerformGisSelection method can control if/when these things
                            // are done.
                            //
                            // Analyse the results of the GIS selection by counting the number of
                            // incids, toids and fragments selected.
                            AnalyzeGisSelectionSet(true);

                            // Indicate the selection didn't come from the map.
                            _filterByMap = false;

                            // Set the filter back to the first incid.
                            SetFilter();

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);

                            // Warn the user that no records were found.
                            if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                                MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "HLU Query",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            //---------------------------------------------------------------------
                        }
                        else
                        {
                            // Restore the previous selection (filter).
                            _incidSelection = incidSelectionBackup;

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);
                        }
                    }
                    else
                    {
                        // Restore the previous selection (filter).
                        _incidSelection = incidSelectionBackup;

                        // Reset the cursor back to normal
                        ChangeCursor(Cursors.Arrow, null);

                        // Warn the user that no records were found
                        MessageBox.Show(App.Current.MainWindow, "No records found.", "HLU Query",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    _incidSelection = null;
                    ChangeCursor(Cursors.Arrow, null);
                    MessageBox.Show(App.Current.MainWindow, ex.Message, "HLU Query",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally { RefreshStatus(); }
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Filter by OSMM Updates Command

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // Open the OSMM Updates query window when in OSMM Update mode.
        //
        public void OpenWindowQueryOSMM(bool initialise)
        {
            if (initialise)
            {
                // Clear the selection (filter).
                _incidSelection = null;

                // Indicate the selection didn't come from the map.
                _filterByMap = false;

                // Indicate there are no more OSMM updates to review.
                if (_osmmBulkUpdateMode == false)
                    _osmmUpdatesEmpty = true;

                // Clear all the form fields (except the habitat class
                // and habitat type).
                ClearForm();

                // Clear the map selection.
                _gisApp.ClearMapSelection();

                // Reset the map counters
                _incidsSelectedMapCount = 0;
                _toidsSelectedMapCount = 0;
                _fragsSelectedMapCount = 0;

                // Refresh all the controls
                RefreshAll();

                DispatcherHelper.DoEvents();
            }

            try
            {
                _windowQueryOSMM = new WindowQueryOSMM();
                if ((_windowQueryOSMM.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowQueryOSMM.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinQueryOSMM = new ViewModelWindowQueryOSMM(HluDataset, _db, this);
                _viewModelWinQueryOSMM.DisplayName = "OSMM Updates Filter";

                // when ViewModel asks to be closed, close window
                _viewModelWinQueryOSMM.RequestClose +=
                    new ViewModelWindowQueryOSMM.RequestCloseEventHandler(_viewModelWinQueryOSMM_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowQueryOSMM.DataContext = _viewModelWinQueryOSMM;

                // show window
                _windowQueryOSMM.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        /// <summary>
        /// Process the sql when the window is closed.
        /// </summary>
        /// <param name="processFlag">The process flag value.</param>
        /// <param name="spatialFlag">The spatial flag value.</param>
        /// <param name="changeFlag">The change flag value.</param>
        /// <param name="status">The OSMM status value.</param>
        /// <param name="apply">Whether to apply (or cancel) the query.</param>
        protected void _viewModelWinQueryOSMM_RequestClose(string processFlag, string spatialFlag, string changeFlag, string status, bool apply)
        {
            // Close the window
            _viewModelWinQueryOSMM.RequestClose -= _viewModelWinQueryOSMM_RequestClose;
            _windowQueryOSMM.Close();

            if (apply == true)
            {

                if (_osmmBulkUpdateMode == true)
                {
                    // Set the default source details
                    IncidSourcesRows[0].source_id = Settings.Default.BulkOSMMSourceId;
                    IncidSourcesRows[0].source_habitat_class = "N/A";
                    //_viewModelMain.IncidSourcesRows[0].source_habitat_type = "N/A";
                    //Date.VagueDateInstance defaultSourceDate = DefaultSourceDate(null, Settings.Default.BulkOSMMSourceId);
                    Date.VagueDateInstance defaultSourceDate = new Date.VagueDateInstance();
                    IncidSourcesRows[0].source_date_start = defaultSourceDate.StartDate;
                    IncidSourcesRows[0].source_date_end = defaultSourceDate.EndDate;
                    IncidSourcesRows[0].source_date_type = defaultSourceDate.DateType;
                    IncidSourcesRows[0].source_boundary_importance = Settings.Default.SourceImportanceApply1;
                    IncidSourcesRows[0].source_habitat_importance = Settings.Default.SourceImportanceApply1;
                }

                // Apply the OSMM Updates filter
                if (processFlag != null || spatialFlag != null || changeFlag != null || status != null)
                    ApplyOSMMUpdatesFilter(processFlag, spatialFlag, changeFlag, status);
            }
        }

        public void OpenWindowQueryOSMMAdvanced(bool initialise)
        {
            if (initialise)
            {
                // Clear the selection (filter).
                _incidSelection = null;

                // Indicate the selection didn't come from the map.
                _filterByMap = false;

                // Indicate there are no more OSMM updates to review.
                if (_osmmBulkUpdateMode == false)
                    _osmmUpdatesEmpty = true;

                // Clear all the form fields (except the habitat class
                // and habitat type).
                ClearForm();

                // Clear the map selection.
                _gisApp.ClearMapSelection();

                // Reset the map counters
                _incidsSelectedMapCount = 0;
                _toidsSelectedMapCount = 0;
                _fragsSelectedMapCount = 0;

                // Refresh all the controls
                RefreshAll();

                DispatcherHelper.DoEvents();
            }

            try
            {
                _windowQueryAdvanced = new WindowQueryAdvanced();
                if ((_windowQueryAdvanced.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowQueryAdvanced.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinQueryAdvanced = new ViewModelWindowQueryAdvanced(HluDataset, _db);
                _viewModelWinQueryAdvanced.DisplayName = "OSMM Updates Advanced Filter";

                // when ViewModel asks to be closed, close window
                _viewModelWinQueryAdvanced.RequestClose +=
                    new ViewModelWindowQueryAdvanced.RequestCloseEventHandler(_viewModelWinQueryOSMMAdvanced_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowQueryAdvanced.DataContext = _viewModelWinQueryAdvanced;

                // show window
                _windowQueryAdvanced.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        /// <summary>
        /// Process the sql when the advanced query window is closed.
        /// </summary>
        /// <param name="sqlFromTables">The tables to query.</param>
        /// <param name="sqlWhereClause">The where clause to apply in the query.</param>
        protected void _viewModelWinQueryOSMMAdvanced_RequestClose(string sqlFromTables, string sqlWhereClause)
        {
            _viewModelWinQueryAdvanced.RequestClose -= _viewModelWinQueryOSMMAdvanced_RequestClose;
            _windowQueryAdvanced.Close();

            if ((sqlFromTables != null) && (sqlWhereClause != null))
            {

                //if (_osmmBulkUpdateMode == true)
                //{
                //    // Set the default source details
                //    IncidSourcesRows[0].source_id = Settings.Default.BulkOSMMSourceId;
                //    IncidSourcesRows[0].source_habitat_class = "N/A";
                //    //_viewModelMain.IncidSourcesRows[0].source_habitat_type = "N/A";
                //    //Date.VagueDateInstance defaultSourceDate = DefaultSourceDate(null, Settings.Default.BulkOSMMSourceId);
                //    Date.VagueDateInstance defaultSourceDate = new Date.VagueDateInstance();
                //    IncidSourcesRows[0].source_date_start = defaultSourceDate.StartDate;
                //    IncidSourcesRows[0].source_date_end = defaultSourceDate.EndDate;
                //    IncidSourcesRows[0].source_date_type = defaultSourceDate.DateType;
                //    IncidSourcesRows[0].source_boundary_importance = Settings.Default.SourceImportanceApply1;
                //    IncidSourcesRows[0].source_habitat_importance = Settings.Default.SourceImportanceApply1;
                //}

                try
                {
                    ChangeCursor(Cursors.Wait, "Validating ...");
                    DispatcherHelper.DoEvents();

                    // Get a list of all the possible query tables.
                    List<DataTable> tables = new List<DataTable>();
                    if ((ViewModelWindowQueryAdvanced.HluDatasetStatic != null))
                    {
                        tables = ViewModelWindowQueryAdvanced.HluDatasetStatic.incid.ChildRelations
                            .Cast<DataRelation>().Select(r => r.ChildTable).ToList();
                        tables.Add(ViewModelWindowQueryAdvanced.HluDatasetStatic.incid);
                    }

                    // Split the string of query table names created by the
                    // user in the form into an array.
                    string[] fromTables = sqlFromTables.Split(',').Select(s => s.Trim(' ')).Distinct().ToArray();

                    // Include the incid_osmm_updates table to use in the query.
                    if (fromTables.Contains(IncidOSMMUpdatesTable.TableName) == false)
                        fromTables = fromTables.Concat(new string[] { IncidOSMMUpdatesTable.TableName }).ToArray();

                    // Select only the database tables that are in the query array.
                    List<DataTable> whereTables = tables.Where(t => fromTables.Contains(t.TableName)).ToList();

                    // If a status is included in the SQL then also filter out pending
                    // and applied updates, otherwise filter out everything
                    // except proposed updates.
                    if (sqlWhereClause.Contains("[incid_osmm_updates].status") == true)
                        sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].status <> 0 AND [incid_osmm_updates].status <> -1", sqlWhereClause);
                    else
                        sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].status > 0", sqlWhereClause);

                    // create a selection DataTable of PK values of IncidTable
                    if (whereTables.Count() > 0)
                    {

                        // Replace any connection type specific qualifiers and delimiters.
                        string newWhereClause = null;
                        if (sqlWhereClause != null)
                            newWhereClause = ReplaceStringQualifiers(sqlWhereClause);

                        // Store the where clause for updating the OSMM updates later.
                        _osmmUpdateWhereClause = null;

                        // Create a selection DataTable of PK values of IncidTable.
                        _incidSelection = _db.SqlSelect(true, IncidTable.PrimaryKey, whereTables, newWhereClause);

                        // Get a list of all the incids in the selection.
                        _incidsSelectedMap = _incidSelection.AsEnumerable()
                            .GroupBy(r => r.Field<string>(_incidSelection.Columns[0].ColumnName)).Select(g => g.Key).OrderBy(s => s);

                        // Retrospectively set the where clause to match the list
                        // of selected incids (for possible use later).
                        _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                            IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);

                        // Backup the current selection (filter).
                        DataTable incidSelectionBackup = _incidSelection;

                        // If there are any records in the selection (and the tool is
                        // not currently in bulk update mode).
                        if (IsFiltered)
                        {
                            // Find the expected number of features to be selected in GIS.
                            _toidsSelectedDBCount = 0;
                            _fragsSelectedDBCount = 0;
                            ExpectedSelectionFeatures(whereTables, sqlWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                            //---------------------------------------------------------------------
                            // CHANGED: CR12 (Select by attribute performance)
                            // Store the number of incids found in the database
                            _incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;
                            //---------------------------------------------------------------------

                            ChangeCursor(Cursors.Wait, "Filtering ...");
                            // Select the required incid(s) in GIS.
                            if (PerformGisSelection(true, _fragsSelectedDBCount, _incidsSelectedDBCount))
                            {
                                //---------------------------------------------------------------------
                                // CHANGED: CR21 (Select current incid in map)
                                // Analyse the results, set the filter and reset the cursor AFTER
                                // returning from performing the GIS selection so that other calls
                                // to the PerformGisSelection method can control if/when these things
                                // are done.
                                //
                                // Analyse the results of the GIS selection by counting the number of
                                // incids, toids and fragments selected.
                                AnalyzeGisSelectionSet(true);

                                // Indicate the selection didn't come from the map.
                                _filterByMap = false;

                                if (_osmmBulkUpdateMode == false)
                                {
                                    // Indicate there are more OSMM updates to review.
                                    _osmmUpdatesEmpty = false;
                                    //---------------------------------------------------------------------
                                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                                    //
                                    OnPropertyChanged("CanOSMMAccept");
                                    OnPropertyChanged("CanOSMMSkip");
                                    //---------------------------------------------------------------------

                                    // Set the filter to the first incid.
                                    SetFilter();
                                }

                                // Refresh all the controls
                                RefreshAll();

                                // Reset the cursor back to normal.
                                ChangeCursor(Cursors.Arrow, null);

                                // Warn the user that no records were found.
                                if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                                    MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "OSMM Updates",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                                //---------------------------------------------------------------------
                            }
                            else
                            {
                                if (_osmmBulkUpdateMode == false)
                                {
                                    // Clear the selection (filter).
                                    _incidSelection = null;

                                    // Indicate the selection didn't come from the map.
                                    _filterByMap = false;

                                    // Indicate there are no more OSMM updates to review.
                                    if (_osmmBulkUpdateMode == false)
                                        _osmmUpdatesEmpty = true;

                                    // Clear all the form fields (except the habitat class
                                    // and habitat type).
                                    ClearForm();

                                    // Clear the map selection.
                                    _gisApp.ClearMapSelection();

                                    // Reset the map counters
                                    _incidsSelectedMapCount = 0;
                                    _toidsSelectedMapCount = 0;
                                    _fragsSelectedMapCount = 0;

                                    // Refresh all the controls
                                    RefreshAll();
                                }
                            
                                // Reset the cursor back to normal.
                                ChangeCursor(Cursors.Arrow, null);

                                // Warn the user that no records were found.
                                MessageBox.Show(App.Current.MainWindow, "No records found.", "OSMM Updates",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                        else
                        {
                            if (_osmmBulkUpdateMode == false)
                            {
                                // Clear the selection (filter).
                                _incidSelection = null;

                                // Indicate the selection didn't come from the map.
                                _filterByMap = false;

                                // Indicate there are no more OSMM updates to review.
                                _osmmUpdatesEmpty = true;

                                // Clear all the form fields (except the habitat class
                                // and habitat type).
                                ClearForm();

                                // Clear the map selection.
                                _gisApp.ClearMapSelection();

                                // Reset the map counters
                                _incidsSelectedMapCount = 0;
                                _toidsSelectedMapCount = 0;
                                _fragsSelectedMapCount = 0;

                                // Refresh all the controls
                                RefreshAll();
                            }
                            else
                            {
                                // Restore the previous selection (filter).
                                _incidSelection = incidSelectionBackup;
                            }

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);

                            // Warn the user that no records were found.
                            MessageBox.Show(App.Current.MainWindow, "No records selected.", "OSMM Updates",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _incidSelection = null;
                    ChangeCursor(Cursors.Arrow, null);
                MessageBox.Show(App.Current.MainWindow, ex.Message, "OSMM Updates",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally { RefreshStatus(); }
            }
        }
        //---------------------------------------------------------------------

        public void ApplyOSMMUpdatesFilter(string processFlag, string spatialFlag, string changeFlag, string status)
        {
            try
            {
                ChangeCursor(Cursors.Wait, "Validating ...");
                DispatcherHelper.DoEvents();

                // Select only the incid_osmm_updates database table to use in the query.
                List<DataTable> whereTables = new List<DataTable>();
                whereTables.Add(IncidOSMMUpdatesTable);

                // Always filter out applied updates
                string sqlWhereClause;
                sqlWhereClause = "[incid_osmm_updates].status <> -1";

                // Add any other filter criteria.
                if (processFlag != null && processFlag != _codeAnyRow)
                    sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].process_flag = {1}", sqlWhereClause, processFlag);

                if (spatialFlag != null && spatialFlag != _codeAnyRow)
                    sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].spatial_flag = '{1}'", sqlWhereClause, spatialFlag);

                if (changeFlag != null && changeFlag != _codeAnyRow)
                    sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].change_flag = '{1}'", sqlWhereClause, changeFlag);

                if (status != null && status != _codeAnyRow)
                {
                    int newStatus = 999;
                    switch (status)
                    {
                        case "Rejected":
                            newStatus = -99;
                            break;
                        case "Ignored":
                            newStatus = -2;
                            break;
                        case "Applied":
                            newStatus = -1;
                            break;
                        case "Pending":
                            newStatus = 0;
                            break;
                        case "Proposed":
                            newStatus = 1;
                            break;
                        default:
                            newStatus = -999;
                            break;
                    }

                    if (newStatus == 1)
                        sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].status > 0", sqlWhereClause);
                    else
                        sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].status = {1}", sqlWhereClause, newStatus);
                }

                // Don't show pending or applied updates when no status filter is applied
                if (status == null || status == _codeAnyRow)
                    sqlWhereClause = String.Format("{0} AND [incid_osmm_updates].status <> 0  AND [incid_osmm_updates].status <> -1", sqlWhereClause);

                // Replace any connection type specific qualifiers and delimiters.
                string newWhereClause = null;
                newWhereClause = ReplaceStringQualifiers(sqlWhereClause);

                // Store the where clause for updating the OSMM updates later.
                _osmmUpdateWhereClause = newWhereClause;

                // Create a selection DataTable of PK values of IncidTable.
                _incidSelection = _db.SqlSelect(true, IncidTable.PrimaryKey, whereTables, newWhereClause);

                // Get a list of all the incids in the selection.
                _incidsSelectedMap = _incidSelection.AsEnumerable()
                    .GroupBy(r => r.Field<string>(_incidSelection.Columns[0].ColumnName)).Select(g => g.Key).OrderBy(s => s);

                // Retrospectively set the where clause to match the list
                // of selected incids (for possible use later).
                _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                    IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);

                // Backup the current selection (filter).
                DataTable incidSelectionBackup = _incidSelection;

                // If there are any records in the selection (and the tool is
                // not currently in bulk update mode).
                if (IsFiltered)
                {
                    // Find the expected number of features to be selected in GIS.
                    _toidsSelectedDBCount = 0;
                    _fragsSelectedDBCount = 0;
                    ExpectedSelectionFeatures(whereTables, sqlWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                    //---------------------------------------------------------------------
                    // CHANGED: CR12 (Select by attribute performance)
                    // Store the number of incids found in the database
                    _incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;
                    //---------------------------------------------------------------------

                    ChangeCursor(Cursors.Wait, "Filtering ...");
                    // Select the required incid(s) in GIS.
                    if (PerformGisSelection(true, _fragsSelectedDBCount, _incidsSelectedDBCount))
                    {
                        //---------------------------------------------------------------------
                        // CHANGED: CR21 (Select current incid in map)
                        // Analyse the results, set the filter and reset the cursor AFTER
                        // returning from performing the GIS selection so that other calls
                        // to the PerformGisSelection method can control if/when these things
                        // are done.
                        //
                        // Analyse the results of the GIS selection by counting the number of
                        // incids, toids and fragments selected.
                        AnalyzeGisSelectionSet(true);

                        // Indicate the selection didn't come from the map.
                        _filterByMap = false;

                        if (_osmmBulkUpdateMode == false)
                        {
                            // Indicate there are more OSMM updates to review.
                            _osmmUpdatesEmpty = false;

                            // Set the filter to the first incid.
                            SetFilter();

                            //---------------------------------------------------------------------
                            // FIX: 103 Accept/Reject OSMM updates in edit mode.
                            //
                            OnPropertyChanged("CanOSMMAccept");
                            OnPropertyChanged("CanOSMMSkip");
                            //---------------------------------------------------------------------
                        }

                        // Refresh all the controls
                        RefreshAll();

                        // Reset the cursor back to normal.
                        ChangeCursor(Cursors.Arrow, null);

                        // Warn the user that no records were found.
                        if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                            MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "OSMM Updates",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                        //---------------------------------------------------------------------
                    }
                    else
                    {
                        if (_osmmBulkUpdateMode == false)
                        {
                            // Clear the selection (filter).
                            _incidSelection = null;

                            // Indicate the selection didn't come from the map.
                            _filterByMap = false;

                            // Indicate there are no more OSMM updates to review.
                            if (_osmmBulkUpdateMode == false)
                                _osmmUpdatesEmpty = true;

                            // Clear all the form fields (except the habitat class
                            // and habitat type).
                            ClearForm();

                            // Clear the map selection.
                            _gisApp.ClearMapSelection();

                            // Reset the map counters
                            _incidsSelectedMapCount = 0;
                            _toidsSelectedMapCount = 0;
                            _fragsSelectedMapCount = 0;

                            // Refresh all the controls
                            RefreshAll();
                        }

                        // Reset the cursor back to normal.
                        ChangeCursor(Cursors.Arrow, null);

                        // Warn the user that no records were found.
                        MessageBox.Show(App.Current.MainWindow, "No map features found in current layer.", "OSMM Updates",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    if (_osmmBulkUpdateMode == false)
                    {
                        // Clear the selection (filter).
                        _incidSelection = null;

                        // Indicate the selection didn't come from the map.
                        _filterByMap = false;

                        // Indicate there are no more OSMM updates to review.
                        _osmmUpdatesEmpty = true;

                        // Clear all the form fields (except the habitat class
                        // and habitat type).
                        ClearForm();

                        // Clear the map selection.
                        _gisApp.ClearMapSelection();

                        // Reset the map counters
                        _incidsSelectedMapCount = 0;
                        _toidsSelectedMapCount = 0;
                        _fragsSelectedMapCount = 0;

                        // Refresh all the controls
                        RefreshAll();
                    }
                    else
                    {
                        // Restore the previous selection (filter).
                        _incidSelection = incidSelectionBackup;
                    }

                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow, null);

                    // Warn the user that no records were found.
                    MessageBox.Show(App.Current.MainWindow, "No records selected.", "OSMM Updates",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                _incidSelection = null;
                ChangeCursor(Cursors.Arrow, null);
                MessageBox.Show(App.Current.MainWindow, ex.Message, "OSMM Updates",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { RefreshStatus(); }
        }

        /// <summary>
        /// Clears the current record for the form.
        /// </summary>
        public void ClearForm()
        {
            // Disable the History tab
            TabItemHistoryEnabled = false;

            // Clear the habitat fields
            //TODO: Clear form - Check relevant fields are cleared
            ////---------------------------------------------------------------------
            _incidIhsHabitat = null;
            //IncidIhsHabitat = null;
            ////---------------------------------------------------------------------
            HabitatClass = null;
            HabitatType = null;

            IncidCurrentRow = HluDataset.incid.NewincidRow();

            //TODO: Remove and just set multiplex fields to null?
            ////---------------------------------------------------------------------
            //IncidIhsMatrixRows = new HluDataSet.incid_ihs_matrixRow[3]
            //    .Select(r => HluDataset.incid_ihs_matrix.Newincid_ihs_matrixRow()).ToArray();
            //for (int i = 0; i < IncidIhsMatrixRows.Length; i++)
            //{
            //    IncidIhsMatrixRows[i].matrix_id = i;
            //    IncidIhsMatrixRows[i].incid = RecIDs.CurrentIncid;
            //}

            //IncidIhsFormationRows = new HluDataSet.incid_ihs_formationRow[2]
            //    .Select(r => HluDataset.incid_ihs_formation.Newincid_ihs_formationRow()).ToArray();
            //for (int i = 0; i < IncidIhsFormationRows.Length; i++)
            //{
            //    IncidIhsFormationRows[i].formation_id = i;
            //    IncidIhsFormationRows[i].incid = RecIDs.CurrentIncid;
            //}

            //IncidIhsManagementRows = new HluDataSet.incid_ihs_managementRow[2]
            //    .Select(r => HluDataset.incid_ihs_management.Newincid_ihs_managementRow()).ToArray();
            //for (int i = 0; i < IncidIhsManagementRows.Length; i++)
            //{
            //    IncidIhsManagementRows[i].management_id = i;
            //    IncidIhsManagementRows[i].incid = RecIDs.CurrentIncid;
            //}

            //IncidIhsComplexRows = new HluDataSet.incid_ihs_complexRow[2]
            //    .Select(r => HluDataset.incid_ihs_complex.Newincid_ihs_complexRow()).ToArray();
            //for (int i = 0; i < IncidIhsComplexRows.Length; i++)
            //{
            //    IncidIhsComplexRows[i].complex_id = i;
            //    IncidIhsComplexRows[i].incid = RecIDs.CurrentIncid;
            //}
            ////---------------------------------------------------------------------

            IncidSecondaryRows = new HluDataSet.incid_secondaryRow[0]
                .Select(r => HluDataset.incid_secondary.Newincid_secondaryRow()).ToArray();
            //TODO: Clear form - Needed?
            //IncidSecondaryRows = new ObservableCollection<BapEnvironment>();

            IncidConditionRows = new HluDataSet.incid_conditionRow[0]
                .Select(r => HluDataset.incid_condition.Newincid_conditionRow()).ToArray();
            //TODO: Clear form - Needed?
            //for (int i = 0; i < IncidConditionRows.Length; i++)
            //{
            //    IncidConditionRows[i].incid_condition_id = i;
            //    IncidSecondaryRows[i].incid = RecIDs.CurrentIncid;
            //}

            IncidBapRows = new HluDataSet.incid_bapRow[0]
                .Select(r => HluDataset.incid_bap.Newincid_bapRow()).ToArray();
            IncidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            IncidBapRowsUser = new ObservableCollection<BapEnvironment>();

            IncidSourcesRows = new HluDataSet.incid_sourcesRow[3]
                .Select(r => HluDataset.incid_sources.Newincid_sourcesRow()).ToArray();
            for (int i = 0; i < IncidSourcesRows.Length; i++)
            {
                IncidSourcesRows[i].incid_source_id = i;
                IncidSourcesRows[i].source_id = Int32.MinValue;
                IncidSourcesRows[i].incid = RecIDs.CurrentIncid;
            }

            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            // Clear the OSMM Update fields
            _incidOSMMUpdatesOSMMXref = 0;
            _incidOSMMUpdatesProcessFlag = 0;
            _incidOSMMUpdatesSpatialFlag = null;
            _incidOSMMUpdatesChangeFlag = null;
            _incidOSMMUpdatesStatus = null;
            //---------------------------------------------------------------------


        }
        //---------------------------------------------------------------------

        #endregion

        #region Select On Map Command

        /// <summary>
        /// SelectOnMap command.
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
            // Set the status to processing and the cursor to wait.
            ChangeCursor(Cursors.Wait, "Selecting ...");

            SelectOnMap(false);

            // Count the number of toids and fragments for the current incid
            // selected in the GIS and in the database.
            CountToidFrags();

            // Refresh all the status type fields.
            RefreshStatus();

            // Reset the cursor back to normal
            ChangeCursor(Cursors.Arrow, null);
        }

        private bool CanSelectOnMap
        {
            get { return _bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp && IncidCurrentRow != null; }
        }

        /// <summary>
        /// Select current DB record on map when button pressed.
        /// </summary>
        public void SelectOnMap(bool updateIncidSelection)
        {
            if (IncidCurrentRow == null) return;

            //---------------------------------------------------------------------
            // CHANGED: CR21 (Select current incid in map)
            // Temporarily store the incid and GIS selections whilst
            // selecting the current incid in GIS so that the selections
            // can be restored again afterwards (so that the filter is
            // still active).
            //
            try
            {
                DataTable prevIncidSelection = NewIncidSelectionTable();
                DataTable prevGISSelection = NewGisSelectionTable();

                // Determine if a filter with more than one incid is currently active.
                bool multiIncidFilter = (IsFiltered && _incidSelection.Rows.Count > 1);

                // Save the current table of selected incids.
                prevIncidSelection = _incidSelection;

                // Save the current table of selected GIS features.
                prevGISSelection = _gisSelection;

                // Reset the table of selected incids.
                _incidSelection = NewIncidSelectionTable();

                // Set the table of selected incids to the current incid.
                DataRow selRow = _incidSelection.NewRow();
                foreach (DataColumn c in _incidSelection.Columns)
                    selRow[c] = IncidCurrentRow[c.ColumnName];
                _incidSelection.Rows.Add(selRow);

                // Select all the features for the current incid in GIS.
                PerformGisSelection(false, -1, -1);

                // If a multi-incid filter was previously active then restore it.
                if (multiIncidFilter)
                {
                    // Restore the previous table of selected incids.
                    _incidSelection = prevIncidSelection;

                    // Count the number of fragments previously selected for this incid.
                    int numFragsOld = 0;
                    if (prevGISSelection != null)
                    {
                        DataRow[] gisRows = prevGISSelection.AsEnumerable()
                            .Where(r => r[HluDataset.incid_mm_polygons.incidColumn.ColumnName].Equals(_incidCurrentRow.incid)).ToArray();
                        numFragsOld = gisRows.Length;
                    }

                    // Count the number of fragments now selected for this incid.
                    int numFragsNew = 0;
                    if (_gisSelection != null)
                    {
                        DataRow[] gisRows = _gisSelection.AsEnumerable()
                            .Where(r => r[HluDataset.incid_mm_polygons.incidColumn.ColumnName].Equals(_incidCurrentRow.incid)).ToArray();
                        numFragsNew = gisRows.Length;
                    }

                    // Check if the number of fragments now selected for this incid
                    // has changed.
                    if (numFragsNew == numFragsOld)
                    {
                        // If the same number of fragments for this incid has been
                        // selected then just restore the previous table.
                        _gisSelection = prevGISSelection;
                    }
                    else
                    {
                        // If the number of fragments selected has changed for this
                        // incid then add all the rows for all the other incids in
                        // the previous table of selected GIS features to the current
                        // table of selected GIS features (thereby replacing the previously
                        // selected features for the current incid with the new
                        // selection).
                        if (prevGISSelection != null)
                        {
                            selRow = _gisSelection.NewRow();
                            foreach (DataRow row in prevGISSelection.Rows)
                            {
                                if (row[HluDataset.incid.incidColumn.ColumnName].ToString() != _incidCurrentRow.incid)
                                    _gisSelection.ImportRow(row);
                            }
                        }
                    }

                    // Analyse the results of the GIS selection by counting
                    // the number of incids, toids and fragments selected.
                    AnalyzeGisSelectionSet(updateIncidSelection);

                }
                else
                {
                    // Restore the previous table of selected incids.
                    _incidSelection = prevIncidSelection;

                    // Restore the previous table of selected GIS features.
                    //_gisSelection = prevGISSelection;

                    // Analyse the results of the GIS selection by counting
                    // the number of incids, toids and fragments selected.
                    AnalyzeGisSelectionSet(false);

                    // Set the filter back to the first incid.
                    //SetFilter();
                }

                // Indicate the selection didn't come from the map.
                _filterByMap = false;

                // Zoom to the GIS selection if auto zoom is on.
                if (_gisSelection != null && _autoZoomSelection != 0)
                {
                    // Get the map distance units.
                    string distUnits = Settings.Default.MapDistanceUnits;

                    _gisApp.ZoomSelected(_minZoom, distUnits, _autoZoomSelection == 2);
                }

                // Warn the user that no records were found.
                if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                    MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "HLU Selection",
                        MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
            }
            //---------------------------------------------------------------------

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
            // Get the GIS layer selection
            ReadMapSelection(true);
        }

        private bool CanReadMapSelection
        {
                //---------------------------------------------------------------------
                // FIX: 101 Enable get map selection when in OSMM update mode.
                //--------------------------------------------------------------            
                //get { return _bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp; }
                get { return _bulkUpdateMode == false && HaveGisApp; }
                //---------------------------------------------------------------------
        }

        internal void ReadMapSelection(bool showMessage)
        {
            try
            {
                ChangeCursor(Cursors.Wait, "Filtering ...");

                DispatcherHelper.DoEvents();

                // Read which features are selected in GIS
                _gisSelection = NewGisSelectionTable();
                _gisApp.ReadMapSelection(ref _gisSelection);

                // Count how many incids, toids and fragments are selected in GIS
                _incidSelectionWhereClause = null;
                AnalyzeGisSelectionSet(true);

                // Update the number of features found in the database.
                _toidsSelectedDBCount = 0;
                _fragsSelectedDBCount = 0;
                ExpectedSelectionFeatures(_incidSelectionWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                // Store the number of incids found in the database
                _incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;

                // Indicate the selection came from the map.
                if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0))
                    _filterByMap = true;

                if (_gisSelection.Rows.Count > 0)
                {
                    //---------------------------------------------------------------------
                    // Prevent OSMM updates being actioned too quickly.
                    // FIX: 103 Accept/Reject OSMM updates in edit mode.
                    if (_osmmBulkUpdateMode == false && _osmmUpdateMode == true)
                    {
                        // Indicate there are more OSMM updates to review.
                        _osmmUpdatesEmpty = false;
                        OnPropertyChanged("CanOSMMAccept");
                        OnPropertyChanged("CanOSMMSkip");
                    }
                    //---------------------------------------------------------------------

                    // Set the filter to the first incid.
                    SetFilter();

                    if (_autoSplit && (_gisSelection != null) && (_gisSelection.Rows.Count > 1) && (_incidsSelectedMapCount == 1) &&
                        (_toidsSelectedMapCount == 1) && (_fragsSelectedMapCount == 1))
                    {
                        if (IsAuthorisedUser)
                        {
                            if (!CanPhysicallySplit)
                            {
                                _windowCompSplit = new WindowCompletePhysicalSplit();
                                _windowCompSplit.Owner = App.Current.MainWindow;
                                _windowCompSplit.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                                _vmCompSplit = new ViewModelCompletePhysicalSplit(_reason, _process, _reasonCodes, _processCodes);
                                _vmCompSplit.RequestClose += new ViewModelCompletePhysicalSplit.RequestCloseEventHandler(vmCompSplit_RequestClose);
                                _windowCompSplit.DataContext = _vmCompSplit;
                                _windowCompSplit.ShowDialog();
                            }
                            if (CanPhysicallySplit)
                            {
                                ViewModelWindowMainSplit vmSplit = new ViewModelWindowMainSplit(this);
                                if (vmSplit.PhysicalSplit()) NotifySplitMerge("Physical split completed.");
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
                    if (showMessage) MessageBox.Show("No map features selected in current layer.", "HLU Selection",
                        MessageBoxButton.OK, MessageBoxImage.Exclamation);

                    // Reset the incid and map selections and move
                    // to the first incid in the database.
                    if (!haveSplashWin)
                        ClearFilter(true);

                    //---------------------------------------------------------------------
                    // FIX: 107 Reset filter when no map features selected.
                    // 
                    // Indicate the selection didn't come from the map (but only after
                    // the filter has been cleared and the first incid selected so that
                    // the map doesn't auto zoom to the incid.
                    _filterByMap = false;
                    //---------------------------------------------------------------------
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

        #region Priority Habitats Command
        //---------------------------------------------------------------------
        // CHANGED: CR54 Add pop-out windows to show/edit priority habitats
        // New pop-out windows to view and edit priority and potential
        // priority habitats more clearly.
        //
        /// <summary>
        /// EditPriorityHabitats command.
        /// </summary>
        public ICommand EditPriorityHabitatsCommand
        {
            get
            {
                if (_editPriorityHabitatsCommand == null)
                {
                    Action<object> editPriorityHabitatsAction = new Action<object>(this.EditPriorityHabitatsClicked);
                    _editPriorityHabitatsCommand = new RelayCommand(editPriorityHabitatsAction, param => this.CanEditPriorityHabitats);
                }
                return _editPriorityHabitatsCommand;
            }
        }

        private void EditPriorityHabitatsClicked(object param)
        {
            try
            {
                _windowEditPriorityHabitats = new WindowEditPriorityHabitats();
                _windowEditPriorityHabitats.Owner = App.Current.MainWindow;
                _windowEditPriorityHabitats.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinEditPriorityHabitats = new ViewModelWindowEditPriorityHabitats(this, IncidBapHabitatsAuto);
                _viewModelWinEditPriorityHabitats.DisplayName = "Priority Habitats";

                // when ViewModel asks to be closed, close window
                _viewModelWinEditPriorityHabitats.RequestClose += new ViewModelWindowEditPriorityHabitats
                    .RequestCloseEventHandler(_viewModelWinEditPriorityHabitats_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowEditPriorityHabitats.DataContext = _viewModelWinEditPriorityHabitats;

                // show window
                _windowEditPriorityHabitats.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        protected void _viewModelWinEditPriorityHabitats_RequestClose(ObservableCollection<BapEnvironment> incidBapHabitatsAuto)
        {
            _viewModelWinEditPriorityHabitats.RequestClose -= _viewModelWinEditPriorityHabitats_RequestClose;
            _windowEditPriorityHabitats.Close();

            if (incidBapHabitatsAuto != null)
            {
                IncidBapHabitatsAuto = incidBapHabitatsAuto;

                // Check if there are any errors in the primary BAP records to see
                // if the Priority tab label should be flagged as also in error.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                    if (countInvalid > 0)
                        AddErrorList(ref _priorityErrors, "BapAuto");
                    else
                        DelErrorList(ref _priorityErrors, "BapAuto");
                }

                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("PriorityTabLabel");
            }
        }

        private bool CanEditPriorityHabitats
        {
            get { return _bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp && BapHabitatsAutoEnabled; }
        }

        #endregion

        #region Potential Priority Habitats Command

        /// <summary>
        /// EditPotentialHabitats command.
        /// </summary>
        public ICommand EditPotentialHabitatsCommand
        {
            get
            {
                if (_editPotentialHabitatsCommand == null)
                {
                    Action<object> editPotentialHabitatsAction = new Action<object>(this.EditPotentialHabitatsClicked);
                    _editPotentialHabitatsCommand = new RelayCommand(editPotentialHabitatsAction, param => this.CanEditPotentialHabitats);
                }
                return _editPotentialHabitatsCommand;
            }
        }

        private void EditPotentialHabitatsClicked(object param)
        {
            try
            {
                _windowEditPotentialHabitats = new WindowEditPotentialHabitats();
                _windowEditPotentialHabitats.Owner = App.Current.MainWindow;
                _windowEditPotentialHabitats.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinEditPotentialHabitats = new ViewModelWindowEditPotentialHabitats(this, IncidBapHabitatsUser);
                _viewModelWinEditPotentialHabitats.DisplayName = "Potential Priority Habitats";

                // when ViewModel asks to be closed, close window
                _viewModelWinEditPotentialHabitats.RequestClose += new ViewModelWindowEditPotentialHabitats
                    .RequestCloseEventHandler(_viewModelWinEditPotentialHabitats_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowEditPotentialHabitats.DataContext = _viewModelWinEditPotentialHabitats;

                // show window
                _windowEditPotentialHabitats.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        protected void _viewModelWinEditPotentialHabitats_RequestClose(ObservableCollection<BapEnvironment> incidBapHabitatsUser)
        {
            _viewModelWinEditPotentialHabitats.RequestClose -= _viewModelWinEditPotentialHabitats_RequestClose;
            _windowEditPotentialHabitats.Close();

            if (incidBapHabitatsUser != null)
            {
                IncidBapHabitatsUser = incidBapHabitatsUser;

                // Check if there are any errors in the secondary BAP records to see
                // if the Priority tab label should be flagged as also in error.
                if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
                {
                    int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                    if (countInvalid > 0)
                        AddErrorList(ref _priorityErrors, "BapUser");
                    else
                        DelErrorList(ref _priorityErrors, "BapUser");
                }

                OnPropertyChanged("IncidBapHabitatsUser");
                OnPropertyChanged("PriorityTabLabel");
            }
        }

        private bool CanEditPotentialHabitats
        {
            get { return _bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp && BapHabitatsUserEnabled; }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Add Secondary Habitat Command

        /// <summary>
        /// AddSecondaryHabitat command.
        /// </summary>
        public ICommand AddSecondaryHabitatCommand
        {
            get
            {
                if (_addSecondaryHabitatCommand == null)
                {
                    Action<object> addSecondaryHabitatAction = new Action<object>(this.AddSecondaryHabitatClicked);
                    _addSecondaryHabitatCommand = new RelayCommand(addSecondaryHabitatAction, param => this.CanAddSecondaryHabitat);
                }
                return _addSecondaryHabitatCommand;
            }
        }

        public bool CanAddSecondaryHabitat
        {
            get
            {
                // Check not in bulk update mode or OSMM update mode and GIS present
                // and primary code and secondary habitat group and code have been set.
                return (_bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp
                    && _incidPrimary != null && _secondaryGroup != null && _secondaryHabitat != null);
            }
        }

        /// <summary>
        /// Add a secondary habitat to the table.
        /// </summary>
        /// <param name="param">The parameter.</param>
        private void AddSecondaryHabitatClicked(object param)
        {
            try
            {
                // Double check secondary habitat group and code have been set.
                if (_secondaryGroup != null && _secondaryHabitat != null)
                {
                    string secondaryGroup = _secondaryGroup;
                    if (secondaryGroup == "<All>")
                    {
                        // Lookup the secondary group from the secondary code
                        IEnumerable<string> q = null;
                        q = (from s in SecondaryHabitatCodesAll
                             where s.is_local
                             && s.code == _secondaryHabitat
                             select s.code_group);
                        if ((q != null) && (q.Count() > 0)) secondaryGroup = q.First();
                    }

                    // Add secondary habitat to table if it isn't already in the table
                    if (SecondaryHabitat.SecondaryHabitatList == null ||
                        SecondaryHabitat.SecondaryHabitatList.Count(sh => sh.secondary_habitat == _secondaryHabitat) == 0)
                        AddSecondaryHabitat(false, -1, Incid, _secondaryHabitat, secondaryGroup);

                    // Refresh secondary table and summary.
                    RefreshSecondaryHabitats();
                    OnPropertyChanged("IncidSecondarySummary");

                    // Refresh the BAP habitat environments (in case secondary codes
                    // are, or should be, reflected).
                    GetBapEnvironments();
                    OnPropertyChanged("IncidBapHabitatsAuto");
                    OnPropertyChanged("IncidBapHabitatsUser");
                    OnPropertyChanged("BapHabitatsAutoEnabled");
                    OnPropertyChanged("BapHabitatsUserEnabled");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //throw;
            }
        }

        /// <summary>
        /// AddSecondaryHabitatList command.
        /// </summary>
        public ICommand AddSecondaryHabitatListCommand
        {
            get
            {
                if (_addSecondaryHabitatListCommand == null)
                {
                    Action<object> addSecondaryHabitatListAction = new Action<object>(this.AddSecondaryHabitatListClicked);
                    _addSecondaryHabitatListCommand = new RelayCommand(addSecondaryHabitatListAction, param => this.CanAddSecondaryHabitatList);
                }
                return _addSecondaryHabitatListCommand;
            }
        }

        public bool CanAddSecondaryHabitatList
        {
            get
            {
                // Check not in bulk update mode or OSMM update mode and GIS present
                // and primary code and secondary habitat group and code have been set.
                return (_bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp
                    && _incidPrimary != null);
            }
        }

        /// <summary>
        /// Opens the query secondaries window.
        /// </summary>
        /// <exception cref="Exception">No parent window loaded</exception>
        private void AddSecondaryHabitatListClicked(object param)
        {
            try
            {
                _windowQuerySecondaries = new WindowQuerySecondaries();
                if ((_windowQuerySecondaries.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _windowQuerySecondaries.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // create ViewModel to which main window binds
                _viewModelWinQuerySecondaries = new ViewModelWindowQuerySecondaries();
                _viewModelWinQuerySecondaries.DisplayName = "Add Secondary Habitats";

                // when ViewModel asks to be closed, close window
                _viewModelWinQuerySecondaries.RequestClose +=
                    new ViewModelWindowQuerySecondaries.RequestCloseEventHandler(_viewModelWinQuerySecondaries_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _windowQuerySecondaries.DataContext = _viewModelWinQuerySecondaries;

                // show window
                _windowQuerySecondaries.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        /// <summary>
        /// Closes the query add secondaries window and adds the list of 
        /// secondary habitats to the tble.
        /// </summary>
        /// <param name="querySecondaries">The list of secondaries to add.</param>
        protected void _viewModelWinQuerySecondaries_RequestClose(String querySecondaries)
        {
            _viewModelWinQuerySecondaries.RequestClose -= _viewModelWinQuerySecondaries_RequestClose;
            _windowQuerySecondaries.Close();

            if (!String.IsNullOrEmpty(querySecondaries))
            {
                try
                {
                    bool addedCodes = false;
                    List<string> errorCodes = new List<string>();

                    // Double check secondary habitat list has been entered.
                    if (querySecondaries != null)
                    {
                        // Split the list by spaces, commas or points
                        string pattern = @"\s|\.|\,";
                        Regex rgx = new Regex(pattern);

                        // Process each secondary habitat code
                        string[] secondaryHabitats = rgx.Split(querySecondaries);
                        for (int i = 0; i < secondaryHabitats.Length; i++)
                        {
                            string secondaryHabitat = secondaryHabitats[i];
                            if (secondaryHabitat != null)
                            {
                                // Lookup the secondary group for the secondary code
                                IEnumerable<string> q = null;
                                q = (from s in SecondaryHabitatCodesAll
                                     where s.is_local
                                     && s.code == secondaryHabitat
                                     select s.code_group);

                                // If the secondary group has been found
                                string secondaryGroup = null;
                                if ((q != null) && (q.Count() > 0))
                                {
                                    secondaryGroup = q.First();

                                    // Add secondary habitat if it isn't already in the table
                                    if (SecondaryHabitat.SecondaryHabitatList == null ||
                                        SecondaryHabitat.SecondaryHabitatList.Count(sh => sh.secondary_habitat == secondaryHabitat) == 0)
                                    {
                                        // Add secondary habitat to table if it isn't already in the table
                                        bool err;
                                        err = AddSecondaryHabitat(false, -1, Incid, secondaryHabitat, secondaryGroup);
                                        if (err == true)
                                            errorCodes.Add(secondaryHabitat);

                                        addedCodes = true;
                                    }
                                    else
                                        errorCodes.Add(secondaryHabitat);
                                }
                                else
                                    errorCodes.Add(secondaryHabitat);
                            }
                        }

                        // If any valid codes were entered and were added to the table
                        if (addedCodes == true)
                        {
                            // Refresh secondary table and summary.
                            RefreshSecondaryHabitats();
                            OnPropertyChanged("IncidSecondarySummary");

                            // Refresh the BAP habitat environments (in case secondary codes
                            // are, or should be, reflected).
                            GetBapEnvironments();
                            OnPropertyChanged("IncidBapHabitatsAuto");
                            OnPropertyChanged("IncidBapHabitatsUser");
                            OnPropertyChanged("BapHabitatsAutoEnabled");
                            OnPropertyChanged("BapHabitatsUserEnabled");
                        }

                        // If any codes were invalid then tell the user
                        if (errorCodes != null && errorCodes.Count > 0)
                        {
                            // Sort the distinct secondary codes in error numerically
                            errorCodes = errorCodes.Distinct().OrderBy(e => e.PadLeft(5, '0')).ToList();
                            // Message the user, depending on if there is one or more
                            if (errorCodes.Count() == 1)
                                MessageBox.Show(App.Current.MainWindow, "Code '" + errorCodes.FirstOrDefault() + "' is a duplicate or unknown and has not been added.", "Add Secondary Habitats",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            else
                                MessageBox.Show(App.Current.MainWindow, "Codes '" + String.Join(", ", errorCodes.Take(errorCodes.Count() - 1)) + " and " + errorCodes.Last() + "' are duplicates or unknown and have not been added.", "Add Secondary Habitats",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    ChangeCursor(Cursors.Arrow, null);
                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    //MessageBox.Show(App.Current.MainWindow, ex.Message, "HLU Query",
                    //    MessageBoxButton.OK, MessageBoxImage.Error);
                }
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
            get { return _bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp && _gisSelection != null && _incidsSelectedMapCount == 1; }
        }

        /// <summary>
        /// Select current DB record on map when button pressed.
        /// </summary>
        private void SelectByIncid()
        {
            if ((_gisSelection == null) || (_gisSelection.Rows.Count <= 0) || (_incidsSelectedMapCount != 1))
                return;

            //---------------------------------------------------------------------
            // CHANGED: CR21 (Select current incid in map)
            // Set the status message and cursor, analyse the results, set
            // the filter and reset the cursor AFTER performing the GIS selection
            // so that other calls to the PerformGisSelection method
            // can control if/when these things are done.
            //
            try
            {
                // Set the status to processing and the cursor to wait.
                ChangeCursor(Cursors.Wait, "Selecting ...");

                // Set the table of selected incids to the current incid.
                _incidSelection = NewIncidSelectionTable();
                DataRow selRow = _incidSelection.NewRow();
                foreach (DataColumn c in _incidSelection.Columns)
                {
                    if (_gisSelection.Columns.Contains(c.ColumnName))
                        selRow[c] = _gisSelection.Rows[0][c.ColumnName];
                }
                _incidSelection.Rows.Add(selRow);

                // Select all the features for the current incid in GIS.
                PerformGisSelection(false, -1, -1);

                // Analyse the results of the GIS selection by counting the number of
                // incids, toids and fragments selected.
                AnalyzeGisSelectionSet(true);

                // Store the number of incids found in the database
                //_incidsSelectedDBCount = _incidSelection != null ? _incidSelection.Rows.Count : 0;

                // Update the number of features found in the database.
                //_toidsSelectedDBCount = 0;
                //_fragsSelectedDBCount = 0;
                //ExpectedSelectionFeatures(_incidSelectionWhereClause, ref _toidsSelectedDBCount, ref _fragsSelectedDBCount);

                // Indicate the selection didn't come from the map.
                _filterByMap = false;

                // Set the filter back to the first incid.
                SetFilter();

                // Warn the user that no records were found.
                if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                    MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "HLU Selection",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                // Reset the cursor back to normal.
                ChangeCursor(Cursors.Arrow, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HLU", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            // Make sure the cursor is always reset.
            finally
            {
                // Reset the cursor back to normal
                ChangeCursor(Cursors.Arrow, null);
            }
            //---------------------------------------------------------------------

        }

        #endregion

        #region Select All On Map Command

        // Enable all the incids in the current filter to be selected
        // in GIS.
        //
        /// <summary>
        /// SelectAllOnMap command.
        /// </summary>
        public ICommand SelectAllOnMapCommand
        {
            get
            {
                if (_selectAllOnMapCommand == null)
                {
                    Action<object> selectAllOnMapAction = new Action<object>(this.SelectAllOnMapClicked);
                    _selectAllOnMapCommand = new RelayCommand(selectAllOnMapAction, param => this.CanSelectOnMap);
                }
                return _selectAllOnMapCommand;
            }
        }

        private void SelectAllOnMapClicked(object param)
        {
            // Select all the incids in the active filter in GIS.
            SelectAllOnMap();
        }

        /// <summary>
        /// Select all the incids in the active filter in GIS.
        /// </summary>
        public void SelectAllOnMap()
        {
            // If there are any records in the selection (and the tool is
            // not currently in bulk update mode).
            if (IsFiltered)
            {
                try
                {
                    // Set the status to processing and the cursor to wait.
                    ChangeCursor(Cursors.Wait, "Selecting ...");

                    // Backup the current selection (filter).
                    DataTable incidSelectionBackup = _incidSelection;

                    // Build a where clause list for the incids to be selected.
                    List<List<SqlFilterCondition>> whereClause = new List<List<SqlFilterCondition>>();
                    whereClause.Add(ScratchDb.GisWhereClause(_incidSelection, null, false));

                    // Find the expected number of features to be selected in GIS.
                    int expectedNumToids = -1;
                    int expectedNumFeatures = -1;
                    ExpectedSelectionFeatures(whereClause, ref expectedNumToids, ref expectedNumFeatures);

                    // Find the expected number of incids to be selected in GIS.
                    int expectedNumIncids = _incidSelection.Rows.Count;

                    // Select the required incid(s) in GIS.
                    if (PerformGisSelection(true, expectedNumFeatures, expectedNumIncids))
                    {
                        // Analyse the results of the GIS selection by counting the number of
                        // incids, toids and fragments selected.
                        AnalyzeGisSelectionSet(false);

                        // Indicate the selection came from the map.
                        if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0))
                            _filterByMap = true;

                        // Set the filter back to the first incid.
                        SetFilter();

                        // Zoom to the GIS selection if auto zoom is on.
                        if (_gisSelection != null && _autoZoomSelection != 0)
                        {
                            // Get the map distance units.
                            string distUnits = Settings.Default.MapDistanceUnits;

                            _gisApp.ZoomSelected(_minZoom, distUnits, _autoZoomSelection == 2);
                        }

                        // Warn the user that no records were found.
                        if ((_gisSelection == null) || (_gisSelection.Rows.Count == 0))
                            MessageBox.Show(App.Current.MainWindow, "No map features selected in current layer.", "HLU",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Restore the previous selection (filter).
                        _incidSelection = incidSelectionBackup;
                    }
                }
                catch (Exception ex)
                {
                    _incidSelection = null;
                    MessageBox.Show(ex.Message, "HLU", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                // Make sure the cursor is always reset.
                finally
                {
                    // Reset the cursor back to normal
                    ChangeCursor(Cursors.Arrow, null);
                }
                //RefreshStatus();
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
                    _clearFilterCommand = new RelayCommand(qryBuilderAction, param => this.CanClearFilter);
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

        public bool CanClearFilter
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Don't allow filter to be cleared when in OSMM Update mode or
            // OSMM Bulk Update mode.
            //
            get
            {
                return IsFiltered == true &&
                    _osmmUpdateMode == false &&
                    _osmmBulkUpdateMode == false;
            }
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// Clears any active incid filter and optionally moves to the first incid in the index.
        /// </summary>
        /// <param name="resetRowIndex">If set to <c>true</c> the first incid in the index is loaded.</param>
        internal void ClearFilter(bool resetRowIndex)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Reset the OSMM Updates filter when in OSMM Update mode.
            //
            if (_osmmUpdateMode == true)
                ApplyOSMMUpdatesFilter(null, null, null, null);
            else if (_osmmBulkUpdateMode == true)
                ApplyOSMMUpdatesFilter(null, null, null, "Pending");
            else
            //---------------------------------------------------------------------
            {
                //if (IsFiltered)
                //{
                _incidSelection = null;
                _incidSelectionWhereClause = null;
                _gisSelection = null;
                _incidsSelectedDBCount = 0;
                _toidsSelectedDBCount = 0;
                _fragsSelectedDBCount = 0;
                _incidsSelectedMapCount = 0;
                _toidsSelectedMapCount = 0;
                _fragsSelectedMapCount = 0;
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
                    //ChangeCursor(Cursors.Wait, "Processing ...");

                    _incidCurrentRowIndex = 1;
                    //IncidCurrentRowIndex = 1;

                    //ChangeCursor(Cursors.Arrow, null);
                    //---------------------------------------------------------------------
                }
                //---------------------------------------------------------------------
                //}

                //---------------------------------------------------------------------
                // FIX: 107 Reset filter when no map features selected.
                // 
                // Suggest the selection came from the map so that
                // the map doesn't auto zoom to the first incid.
                _filterByMap = true;
                //---------------------------------------------------------------------

                // Re-retrieve the current record
                if (resetRowIndex)
                    IncidCurrentRowIndex = _incidCurrentRowIndex;

                // Count the number of toids and fragments for the current incid
                // selected in the GIS and in the database.
                CountToidFrags();

                // Refresh all the status type fields.
                RefreshStatus();

                //---------------------------------------------------------------------
                // FIX: 107 Reset filter when no map features selected.
                // 
                // Indicate the selection didn't come from the map.
                _filterByMap = false;
                //---------------------------------------------------------------------
            }
        }

        #endregion

        #region Select Helpers

        /// <summary>
        /// Count how many incids, toids and fragments are selected in GIS.
        /// </summary>
        private void AnalyzeGisSelectionSet(bool updateIncidSelection)
        {
            _incidsSelectedMapCount = 0;
            _toidsSelectedMapCount = 0;
            _fragsSelectedMapCount = 0;
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
                        // Order the incids selected in the GIS so that the filter
                        // is sorted in incid order.
                        _incidsSelectedMap = _gisSelection.AsEnumerable()
                            .GroupBy(r => r.Field<string>(_gisIDColumns[0].ColumnName)).Select(g => g.Key).OrderBy(s => s);

                        // Count the number of incids selected in GIS.
                        _incidsSelectedMapCount = _incidsSelectedMap.Count();
                        break;
                }

                // Update the database Incid selection only if required.
                if ((updateIncidSelection) && (_incidsSelectedMapCount > 0))
                {
                    // Set the Incid selection where clause to match the list of
                    // selected incids (for possible use later).
                    if (_incidSelectionWhereClause == null)
                        _incidSelectionWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(
                            IncidPageSize, IncidTable.incidColumn.Ordinal, IncidTable, _incidsSelectedMap);

                    // Update the database Incid selection to the Incids selected in the map.
                    GisToDbSelection();
                }
            }
            else
            {
                if (updateIncidSelection)
                {
                    _incidSelection = null;
                    _incidSelectionWhereClause = null;

                    //---------------------------------------------------------------------
                    // FIX: 107 Reset filter when no map features selected.
                    // 
                    _incidPageRowNoMax = -1;
                    //---------------------------------------------------------------------
                }
            }
        }

        /// <summary>
        /// Set the database Incid selection based on the Incids selected in the map.
        /// </summary>
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

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // Calculate the expected number of GIS features to be selected
        // when using the original interface.
        //
        /// <summary>
        /// Calculates the expected number of GIS features to be selected
        /// by the sql query based upon a list of conditions.
        /// </summary>
        /// <param name="whereClause">The list of where clause conditions.</param>
        /// <returns>An integer of the number of GIS features to be selected.</returns>
        private void ExpectedSelectionFeatures(List<List<SqlFilterCondition>> whereClause, ref int numToids, ref int numFragments)
        {
            if (HaveGisApp && (_incidSelection != null) && (_incidSelection.Rows.Count > 0) &&
                (whereClause != null) && (whereClause.Count > 0))
            {
                try
                {
                    HluDataSet.incid_mm_polygonsDataTable t = new HluDataSet.incid_mm_polygonsDataTable();
                    DataTable[] selTables = new DataTable[] { t }.ToArray();

                    IEnumerable<DataTable> queryTables = whereClause.SelectMany(cond => cond.Select(c => c.Table)).Distinct();
                    //DataTable[] selTables = new DataTable[] { t }.Union(queryTables).ToArray();

                    var fromTables = queryTables.Distinct().Where(q => !selTables.Select(s => s.TableName).Contains(q.TableName));
                    DataTable[] whereTables = selTables.Concat(fromTables).ToArray();

                    DataRelation rel;
                    IEnumerable<SqlFilterCondition> joinCond = fromTables.Select(st =>
                        st.GetType() == typeof(HluDataSet.incidDataTable) ?
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", st.Columns[_hluDS.incid.incidColumn.Ordinal]) :
                        (rel = GetRelation(_hluDS.incid, st)) != null ?
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", rel.ChildColumns[0]) : null).Where(c => c != null);

                    // If there is only one long list then chunk
                    // it up into smaller lists.
                    if (whereClause.Count == 1)
                    {
                        try
                        {
                            List<SqlFilterCondition> whereCond = new List<SqlFilterCondition>();
                            whereCond = whereClause[0];
                            whereClause = whereCond.ChunkClause(IncidPageSize).ToList();
                        }
                        catch { }
                    }

                    numToids = 0;
                    numFragments = 0;
                    for (int i = 0; i < whereClause.Count; i++)
                    {
                        // If the where conditions are going to be appended
                        // to some join conditions then change the boolean
                        // operator before the first where condition to "AND"
                        // and wrap the where conditions in an extra set of
                        // parentheses.
                        if (joinCond.Count() != 0)
                        {
                            List<SqlFilterCondition> whereCond = new List<SqlFilterCondition>();
                            whereCond = whereClause[i];

                            SqlFilterCondition cond = new SqlFilterCondition();
                            cond = whereCond[0];
                            cond.BooleanOperator = "AND";
                            cond.OpenParentheses = "((";

                            cond = whereCond[whereCond.Count - 1];
                            cond.CloseParentheses = "))";
                        }

                        numToids += _db.SqlCount(selTables, String.Format("Distinct {0}", _hluDS.incid_mm_polygons.toidColumn.ColumnName), joinCond.Concat(whereClause[i]).ToList());
                        numFragments += _db.SqlCount(selTables, "*", joinCond.Concat(whereClause[i]).ToList());
                    }
                }
                catch { }
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Calculates the expected number of GIS features to be selected
        /// by the sql query based upon a list of data tables and a sql
        /// where clause.
        /// </summary>
        /// <param name="sqlFromTables">The list of data tables.</param>
        /// <param name="sqlWhereClause">The where clause string.</param>
        /// <returns>An integer of the number of GIS features to be selected.</returns>
        private void ExpectedSelectionFeatures(List<DataTable> sqlFromTables, string sqlWhereClause, ref int numToids, ref int numFragments)
        {
            if (HaveGisApp && (_incidSelection != null) && (_incidSelection.Rows.Count > 0) &&
                sqlFromTables.Count() > 0)
            {
                try
                {
                    HluDataSet.incid_mm_polygonsDataTable t = new HluDataSet.incid_mm_polygonsDataTable();
                    DataTable[] selTables = new DataTable[] { t }.ToArray();

                    var fromTables = sqlFromTables.Distinct().Where(q => !selTables.Select(s => s.TableName).Contains(q.TableName));
                    DataTable[] whereTables = selTables.Concat(fromTables).ToArray();

                    DataRelation rel;
                    IEnumerable<SqlFilterCondition> joinCond = fromTables.Select(st =>
                        st.GetType() == typeof(HluDataSet.incidDataTable) ?
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", st.Columns[_hluDS.incid.incidColumn.Ordinal]) :
                        (rel = GetRelation(_hluDS.incid, st)) != null ?
                        new SqlFilterCondition("AND", t, t.incidColumn, typeof(DataColumn), "(", ")", rel.ChildColumns[0]) : null).Where(c => c != null);

                    numToids = _db.SqlCount(whereTables, String.Format("Distinct {0}", _hluDS.incid_mm_polygons.toidColumn.ColumnName), joinCond.ToList(), sqlWhereClause);
                    numFragments = _db.SqlCount(whereTables, "*", joinCond.ToList(), sqlWhereClause);
                }
                catch { }
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR21 (Select current incid in map)
        // No longer set the filter or reset the cursor AFTER performing
        // the GIS selection so that methods that call this method
        // can control if/when these things are done.
        //
        private bool PerformGisSelection(bool confirmSelect, int expectedNumFeatures, int expectedNumIncids)
        {
            if (_gisApp != null)
            {
                //ChangeCursor(Cursors.Wait, "Processing ...");

                // Build a where clause list for the incids to be selected.
                List<SqlFilterCondition> whereClause = new List<SqlFilterCondition>();
                whereClause = ScratchDb.GisWhereClause(_incidSelection, _gisApp, false);

                //---------------------------------------------------------------------
                // CHANGED: CR12 (Select by attribute performance)
                // Calculate the length of the SQL statement to be sent to GIS.
                int sqlLen = _gisApp.SqlLength(_gisIDColumns, whereClause);
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // CHANGED: CR12 (Select by attribute performance)
                // If the length exceeds the maximum for the GIS application then
                // perform the selection using a join.
                bool selectByJoin = (sqlLen > _gisApp.MaxSqlLength);

                //---------------------------------------------------------------------
                // CHANGED: CR12 (Select by attribute performance)
                // If the length exceeds the maximum for the GIS application then
                // perform the selection using a join.
                if (selectByJoin)
                {
                    if ((!confirmSelect) || (ConfirmGISSelect(true, expectedNumFeatures, expectedNumIncids)))
                    {
                        // Save the incids to the selected to a temporary database
                        ScratchDb.WriteSelectionScratchTable(_gisIDColumns, _incidSelection);
                        DispatcherHelper.DoEvents();

                        // Select all features for incid selection in current layer.
                        _gisSelection = _gisApp.SqlSelect(ScratchDb.ScratchMdbPath,
                            ScratchDb.ScratchSelectionTable, _gisIDColumns);

                        // Check if any features found when applying filter.
                        if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0))
                            return true;
                        else
                            return false;
                    }
                }
                // Otherwise, perform the selection using a SQL query in GIS.
                else
                {
                    if ((!confirmSelect) || (ConfirmGISSelect(false, expectedNumFeatures, expectedNumIncids)))
                    {
                        DispatcherHelper.DoEvents();

                        // Select all features for incid selection in current layer.
                        _gisSelection = _gisApp.SqlSelect(true, false, _gisIDColumns,
                            whereClause);

                        // Check if any features found when applying filter.
                        if ((_gisSelection != null) && (_gisSelection.Rows.Count > 0))
                            return true;
                        else
                            return false;
                    }
                }
                //---------------------------------------------------------------------
            }

            // The selection didn't happen.
            return false;
        }
        //---------------------------------------------------------------------

        private void SetFilter()
        {
            try
            {
                if (IsFiltered && (((_incidsSelectedMapCount > 0) || (_gisApp == null)) || _osmmUpdateMode == true))
                    // If currently splitting a feature then go to the last incid
                    // in the filter (which will be the new incid).
                    if (_splitting)
                        IncidCurrentRowIndex = IsFiltered ? _incidSelection.Rows.Count : _incidRowCount;
                    else
                        IncidCurrentRowIndex = 1;
            }
            finally
            {
                // Count the number of toids and fragments for the current incid
                // selected in the GIS and in the database.
                CountToidFrags();

                // Refresh all the status type fields.
                RefreshStatus();
            }
        }

        #endregion

        #region Switch GIS Layer
        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        public ICommand SwitchGISLayerCommand
        {
            get
            {
                if (_switchGISLayerCommand == null)
                {
                    Action<object> SwitchGISLayerAction = new Action<object>(this.SwitchGISLayerClicked);
                    _switchGISLayerCommand = new RelayCommand(SwitchGISLayerAction, param => this.CanSwitchGISLayer);
                }
                return _switchGISLayerCommand;
            }
        }

        private void SwitchGISLayerClicked(object param)
        {
            if (_gisApp.ListHluLayers() > 0)
            {
                _windowSwitchGISLayer = new WindowSwitchGISLayer();
                _windowSwitchGISLayer.Owner = App.Current.MainWindow;
                _windowSwitchGISLayer.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                // Pass the total number of map windows to the view model
                _viewModelSwitchGISLayer = new ViewModelWindowSwitchGISLayer(_gisApp.ValidHluLayers, _gisApp.CurrentHluLayer, _mapWindowsCount);
                _viewModelSwitchGISLayer.RequestClose +=
                    new ViewModelWindowSwitchGISLayer.RequestCloseEventHandler(_viewModelSwitchGISLayer_RequestClose);

                _windowSwitchGISLayer.DataContext = _viewModelSwitchGISLayer;

                _windowSwitchGISLayer.ShowDialog();
            }
        }

        public bool CanSwitchGISLayer
        {
            get
            {
                if (_bulkUpdateMode == false && _osmmUpdateMode == false && HaveGisApp)
                {
                    // Get the total number of map layers
                    int mapLayersCount = _gisApp.ListHluLayers();

                    // Get the total number of map windows
                    int mapWindowsCount = _gisApp.MapWindowsCount;

                    // If the number of map windows has changed
                    if (mapWindowsCount != _mapWindowsCount)
                    {
                        _mapWindowsCount = mapWindowsCount;

                        // Refresh the layer name
                        OnPropertyChanged("LayerName");
                    }

                    // Return true if there is more than one map layer
                    return mapLayersCount > 1;
                }
                else
                    return false;
            }
        }

        void _viewModelSwitchGISLayer_RequestClose(bool switchGISLayer, GISLayer selectedHLULayer)
        {
            _viewModelSwitchGISLayer.RequestClose -= _viewModelSwitchGISLayer_RequestClose;
            _windowSwitchGISLayer.Close();

            // If the GIS layer has been switched
            if ((switchGISLayer) && (selectedHLULayer != _gisApp.CurrentHluLayer))
            {
                // Check if there are unsaved edits
                if (_editMode && (_bulkUpdateMode == false && _osmmUpdateMode == false) && IsDirty)
                {
                    // Inform the user to save the unsaved edits first.
                    MessageBox.Show("The current record has been changed." +
                        "\n\nYou must save any outstanding changes before switching GIS layer!",
                        "HLU: Switch GIS Layer", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
                else
                {
                    // Switch the GIS layer
                    if (_gisApp.IsHluLayer(selectedHLULayer))
                    {
                        //// Inform the user that the switch worked
                        //if (selectedHLULayer.MapName == null)
                        //    MessageBox.Show(string.Format("GIS Layer switched to {0} [{1}].", selectedHLULayer.LayerName, selectedHLULayer.MapNum),
                        //        "HLU: Switch GIS Layer",MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        //else
                        //    MessageBox.Show(string.Format("GIS Layer switched to {0} in {1} [{2}]", selectedHLULayer.LayerName, selectedHLULayer.MapName, selectedHLULayer.MapNum),
                        //        "HLU: Switch GIS Layer", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                        // Refresh the later name
                        OnPropertyChanged("LayerName");

                        // Get the new GIS layer selection
                        ReadMapSelection(false);
                    }
                }
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Data Tables

        public HluDataSet.incidDataTable IncidTable
        {
            get
            {
                // Load the data table if not already loaded.
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
                // Load the data table if not already loaded.
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
                // Load the data table if not already loaded.
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
                // Load the data table if not already loaded.
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
                // Load the lookup table if not already loaded.
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
                // Load the data table if not already loaded.
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
                // Load the data table if not already loaded.
                if (HluDataset.incid_sources.IsInitialized && (HluDataset.incid_sources.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.incid_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_sourcesDataTable, HluDataSet.incid_sourcesRow>(_db);
                }

                return _hluDS.incid_sources;
            }
        }

        public HluDataSet.incid_secondaryDataTable IncidSecondaryTable
        {
            get
            {
                // Load the data table if not already loaded.
                if (HluDataset.incid_secondary.IsInitialized && (HluDataset.incid_secondary.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_secondaryTableAdapter == null)
                        _hluTableAdapterMgr.incid_secondaryTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_secondaryDataTable, HluDataSet.incid_secondaryRow>(_db);
                }

                return _hluDS.incid_secondary;
            }
        }

        public HluDataSet.incid_conditionDataTable IncidConditionTable
        {
            get
            {
                // Load the data table if not already loaded.
                if (HluDataset.incid_condition.IsInitialized && (HluDataset.incid_condition.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_conditionTableAdapter == null)
                        _hluTableAdapterMgr.incid_conditionTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_conditionDataTable, HluDataSet.incid_conditionRow>(_db);
                }

                return _hluDS.incid_condition;
            }
        }

        public HluDataSet.incid_osmm_updatesDataTable IncidOSMMUpdatesTable
        {
            get
            {
                // Load the data table if not already loaded.
                if (HluDataset.incid_osmm_updates.IsInitialized && (HluDataset.incid_osmm_updates.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.incid_osmm_updatesTableAdapter == null)
                        _hluTableAdapterMgr.incid_osmm_updatesTableAdapter =
                            new HluTableAdapter<HluDataSet.incid_osmm_updatesDataTable, HluDataSet.incid_osmm_updatesRow>(_db);
                }

                return _hluDS.incid_osmm_updates;
            }
        }

        #endregion

        #region Data Rows

        public bool IsFiltered
        {
            get
            {
                return (_bulkUpdateMode != true || (_bulkUpdateMode == true && _osmmBulkUpdateMode == true))
                    && _incidSelection != null
                    && _incidSelection.Rows.Count > 0;
            }
        }

        /// <summary>
        /// Counts the rows in the Incid table.
        /// </summary>
        /// <param name="recount">if set to <c>true</c> [recount].</param>
        /// <returns></returns>
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
                if (OSMMUpdateMode == true && !IsFiltered)
                    return null;

                if (IsFiltered)
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR22 (Record selectors)
                    // Include the total toid and fragment counts for the current Incid
                    // in the status area in addition to the currently select toid and
                    // fragment counts.
                    //
                    if (_osmmUpdateMode == true)
                        return String.Format(" of {0}* [{1}:{2}]", _incidSelection.Rows.Count,
                            _toidsIncidDbCount.ToString(),
                            _fragsIncidDbCount.ToString());
                    else if (_osmmBulkUpdateMode == true)
                        return String.Format("[I:{0}] [T:{1}] [F:{2}]", _incidSelection.Rows.Count,
                            _toidsSelectedMapCount.ToString(),
                            _fragsSelectedMapCount.ToString());
                    else
                        return String.Format(" of {0}* [{1}:{2} of {3}:{4}]", _incidSelection.Rows.Count,
                            _toidsIncidGisCount.ToString(),
                            _fragsIncidGisCount.ToString(),
                            _toidsIncidDbCount.ToString(),
                            _fragsIncidDbCount.ToString());
                    //---------------------------------------------------------------------
                }
                else if (_bulkUpdateMode == true)
                {
                    if ((_incidSelection != null) && (_incidSelection.Rows.Count > 0))
                    {
                        return String.Format("[I:{0}] [T:{1}] [F:{2}]", _incidsSelectedMapCount,
                            _toidsSelectedMapCount.ToString(),
                            _fragsSelectedMapCount.ToString());
                    }
                    else
                    {
                        return String.Format("[I:{0}] [T:{1}] [F:{2}]", _incidsSelectedMapCount,
                            _toidsSelectedMapCount.ToString(),
                            _fragsSelectedMapCount.ToString());
                    }
                }
                else
                {
                    // Include the total toid and fragment counts for the current Incid
                    // in the status area, and the currently select toid and fragment
                    // counts, when auto selecting features on change of incid.
                    //
                    if (_osmmUpdateMode == true)
                        return String.Format(" of {0}* [{1}:{2}]", _incidRowCount,
                            _toidsIncidDbCount.ToString(),
                            _fragsIncidDbCount.ToString());
                    else
                        return String.Format(" of {0} [{1}:{2} of {3}:{4}]", _incidRowCount,
                            _toidsIncidGisCount.ToString(),
                            _fragsIncidGisCount.ToString(),
                            _toidsIncidDbCount.ToString(),
                            _fragsIncidDbCount.ToString());
                }
            }
        }

        public string StatusIncidToolTip { get { return IsFiltered ? "Double click to clear filter" : null; } }

        public string StatusBar
        {
            get { return _windowCursor == Cursors.Wait ? _processingMsg : String.Empty; }
        }

        public int NumIncidSelectedDB
        {
            get { return _incidsSelectedDBCount; }
            set { _incidsSelectedDBCount = value; }
        }

        public int NumToidSelectedDB
        {
            get { return _toidsSelectedDBCount; }
            set { _toidsSelectedDBCount = value; }
        }

        public int NumFragmentsSelectedDB
        {
            get { return _fragsSelectedDBCount; }
            set { _fragsSelectedDBCount = value; }
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

        /// <summary>
        /// Gets or sets the index of the incid current row.
        /// </summary>
        /// <value>
        /// The index of the incid current row.
        /// </value>
        public int IncidCurrentRowIndex
        {
            get { return _incidCurrentRowIndex; }
            set
            {
                // Check there are no outstanding edits.
                MessageBoxResult userResponse = CheckDirty();

                // Process based on the response ...
                // Yes = update the current record first then move to the new incid
                // No = move to the new incid
                // Cancel = don't move to the new incid
                switch (userResponse)
                {
                    case MessageBoxResult.Yes:
                        // Update the current incid
                        if (_viewModelUpd.Update())
                            goto case MessageBoxResult.No;
                        break;
                    case MessageBoxResult.No:
                        //---------------------------------------------------------------------
                        // CHANGED: CR49 Process proposed OSMM Updates
                        // Clear the form and warn the user when there are no more records
                        // when in OSMM Update mode.
                        //
                        if (_osmmUpdateMode == true && ((value > 0) &&
                            (IsFiltered && ((_incidSelection == null) || (value > _incidSelection.Rows.Count)))))
                        {
                            // Clear the selection (filter).
                            _incidSelection = null;

                            // Indicate the selection didn't come from the map.
                            _filterByMap = false;

                            // Indicate there are no more OSMM updates to review.
                            _osmmUpdatesEmpty = true;

                            // Clear all the form fields (except the habitat class
                            // and habitat type).
                            ClearForm();

                            // Clear the map selection.
                            _gisApp.ClearMapSelection();

                            // Reset the map counters
                            _incidsSelectedMapCount = 0;
                            _toidsSelectedMapCount = 0;
                            _fragsSelectedMapCount = 0;

                            // Refresh all the controls
                            RefreshAll();

                            // Reset the cursor back to normal.
                            ChangeCursor(Cursors.Arrow, null);

                            // Warn the user that no more records were found.
                            MessageBox.Show(App.Current.MainWindow, "No more records found.", "OSMM Updates",
                                MessageBoxButton.OK, MessageBoxImage.Information);

                            break;
                        }
                        //---------------------------------------------------------------------

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

            if (_editMode && (_bulkUpdateMode == false))
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

        /// <summary>
        /// Check there are any outstanding edits for the current incid.
        /// </summary>
        /// <returns>The user's response to save or not save the record.</returns>
        private MessageBoxResult CheckDirty()
        {
            MessageBoxResult userResponse = MessageBoxResult.No;

            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Don't check for edits when in OSMM Update mode (because the data
            // can't be edited by the user).
            //
            if (_editMode
                && (_splitting == false)
                && (_bulkUpdateMode == false)
                && (_osmmUpdateMode == false)
                && IsDirty)
            //---------------------------------------------------------------------
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

                // Restore the current row if the user doesn't want save
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

                //TODO: IsDirty - Check picking up changes to all fields
                //return IsDirtyIncid() || IsDirtyIncidIhsMatrix() || IsDirtyIncidIhsFormation() ||
                //    IsDirtyIncidIhsManagement() || IsDirtyIncidIhsComplex() || IsDirtyIncidBap() ||
                //    IsDirtyIncidSources();
                return IsDirtyIncid() || IsDirtyIncidSecondary() || IsDirtyIncidCondition() || IsDirtyIncidBap() ||
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
            // Re-check GIS selection in case it has changed.
            if (_gisApp != null)
            {
                _gisSelection = NewGisSelectionTable();
                _gisApp.ReadMapSelection(ref _gisSelection);
                _incidSelectionWhereClause = null;
                AnalyzeGisSelectionSet(false);
            }

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
                // Clone the current row to use to check for changes later
                CloneIncidCurrentRow();

                _incidArea = -1;
                _incidLength = -1;
                // Flag that the current record has not been changed yet so that the
                // apply button does not appear.
                Changed = false;

                //TODO: Check - Move incid - Needed?
                // without this IncidIhsHabitat becomes null, called from IhsHabitatCodes, when coming 
                // from a previous row with valid IHS habitat code 
                // (seemingly alternating rows when browsing, i.e. 1 ok, 2 wrong, 3 ok, ...)
                //_incidIhsHabitat = null;

                // Get the incid table values
                IncidCurrentRowDerivedValuesRetrieve();
                OnPropertyChanged("IncidPrimary");

                // Get the incid child rows
                GetIncidChildRows(IncidCurrentRow);

                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // If there are any OSMM Updates for this incid then store the values.
                if (_incidOSMMUpdatesRows.Length > 0)
                {
                    _incidOSMMUpdatesOSMMXref = _incidOSMMUpdatesRows[0].osmm_habitat_xref_id;
                    _incidOSMMUpdatesProcessFlag = _incidOSMMUpdatesRows[0].process_flag;
                    _incidOSMMUpdatesSpatialFlag = _incidOSMMUpdatesRows[0].Isspatial_flagNull() ? null : _incidOSMMUpdatesRows[0].spatial_flag;
                    _incidOSMMUpdatesChangeFlag = _incidOSMMUpdatesRows[0].Ischange_flagNull() ? null : _incidOSMMUpdatesRows[0].change_flag;
                    _incidOSMMUpdatesStatus = _incidOSMMUpdatesRows[0].status;
                }
                else
                {
                    _incidOSMMUpdatesOSMMXref = 0;
                    _incidOSMMUpdatesProcessFlag = 0;
                    _incidOSMMUpdatesSpatialFlag = null;
                    _incidOSMMUpdatesChangeFlag = null;
                    _incidOSMMUpdatesStatus = null;
                }
                //---------------------------------------------------------------------

                //TODO: Check - Move incid - Needed?
                // Clear the list of IHS codes.
                //_ihsHabitatCodes = null;

                // Clear the habitat type.
                HabitatType = null;
                OnPropertyChanged("HabitatType");

                // Enable auto select of features on change of incid.
                if (_gisApp != null && _autoSelectOnGis && _bulkUpdateMode == false && !_filterByMap)
                {
                    // Select the current DB record on the Map.
                    SelectOnMap(false);
                }

                // Count the number of toids and fragments for the current incid
                // selected in the GIS and in the database.
                CountToidFrags();

                OnPropertyChanged("IncidCurrentRowIndex");
                OnPropertyChanged("OSMMIncidCurrentRowIndex");
                OnPropertyChanged("IncidCurrentRow");

                // Refresh all statuses, headers adnd fields
                RefreshStatus();
                RefreshHeader();
                RefreshOSMMUpdate();
                RefreshHabitatTab();
                RefreshIHSTab();
                RefreshPriorityTab();
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
        public void CountToidFrags()
        {
            //---------------------------------------------------------------------
            // CHANGED: CR10 (Attribute updates for incid subsets)
            // Count the number of toids and fragments for this incid selected
            // in the GIS. They are counted here, once when the incid changes,
            // instead of in StatusIncid() which is constantly being called.
            _toidsIncidGisCount = 0;
            _fragsIncidGisCount = 0;
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
            _fragsIncidDbCount = 0;
            _toidsIncidDbCount = 0;

            // Count the number of toids and fragments in the database
            // regardless of whether a filter is currently applied.
            //
            // Count the total number of fragments in the database for
            // this incid.
            //if (IsFiltered)
            //{
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
            //}
            //---------------------------------------------------------------------

        }

        private void IncidCurrentRowDerivedValuesRetrieve()
        {
            _incidLastModifiedUser = _incidCurrentRow.last_modified_user_id;
            _incidLastModifiedDate = Convert.IsDBNull(_incidCurrentRow.last_modified_date) ? DateTime.MinValue : _incidCurrentRow.last_modified_date;

            //TODO: Check
            _incidPrimary = _incidCurrentRow.Ishabitat_primaryNull() ? null : _incidCurrentRow.habitat_primary;
            NewPrimaryHabitat(_incidPrimary);
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
            // If the row being sought is within the current page.
            if ((seekRowNumber >= _incidPageRowNoMin) && (seekRowNumber <= _incidPageRowNoMax))
            {
                return seekRowNumber - _incidPageRowNoMin;
            }
            // Otherwise, load a new page.
            else
            {
                int incidPageRowNoMinBak = _incidPageRowNoMin;
                int incidPageRowNoMaxBak = _incidPageRowNoMax;

                try
                {
                    int seekIncidNumber = seekRowNumber;

                    // Get the first incid number
                    if (seekRowNumber < 2)
                    {
                        seekIncidNumber = RecordIds.IncidNumber(_db.ExecuteScalar(
                            String.Format("SELECT {0} FROM {1} ORDER BY {0} ASC",
                            _db.QuoteIdentifier(_hluDS.incid.incidColumn.ColumnName),
                            _db.QualifyTableName(_hluDS.incid.TableName)),
                            _db.Connection.ConnectionTimeout, CommandType.Text).ToString());
                    }
                    // Get the last incid number
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

            // Enable the Incid table to be forced to refill if it has been
            // updated directly in the database rather than via the
            // local copy.
            if ((returnRow != null) && (!_refillIncidTable))
            {
                return returnRow;
            }
            else
            {
                _refillIncidTable = false;
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
                        // If the table has paged backwards (because the required incid
                        // is lower than the page minimum) and if the row number being
                        // sought is the first (i.e. zero) then return the lowest incid.
                        // Otherwise, return the lowest or highest as appropriate.
                        return (moveForward || seekRowNumber == 0) ? _hluDS.incid[0] : _hluDS.incid[_hluDS.incid.Count - 1];
                    }
                }
                catch { return null; }
            }
        }

        private Dictionary<Type, string> BuildChildRowOrderByClauses()
        {
            Dictionary<Type, string> childRowOrberByDict = new Dictionary<Type, string>();

            childRowOrberByDict.Add(typeof(HluDataSet.incid_secondaryDataTable), _hluDS.incid_secondary.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            childRowOrberByDict.Add(typeof(HluDataSet.incid_conditionDataTable), _hluDS.incid_condition.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + String.Format("{0} DESC", _db.QuoteIdentifier(c.ColumnName))))
                .Remove(0, 1).ToString());

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

            childRowOrberByDict.Add(typeof(HluDataSet.incid_osmm_updatesDataTable), _hluDS.incid_osmm_updates.PrimaryKey
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append("," + _db.QuoteIdentifier(c.ColumnName)))
                .Remove(0, 1).ToString());

            return childRowOrberByDict;
        }

        private Dictionary<Type, List<SqlFilterCondition>> BuildChildRowFilters()
        {
            Dictionary<Type, List<SqlFilterCondition>> childRowFilterDict =
                new Dictionary<Type, List<SqlFilterCondition>>();

            childRowFilterDict.Add(typeof(HluDataSet.incid_secondaryDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_secondary));

            childRowFilterDict.Add(typeof(HluDataSet.incid_conditionDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_condition));

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

            childRowFilterDict.Add(typeof(HluDataSet.incid_osmm_updatesDataTable),
                ChildRowFilter(_hluDS.incid, _hluDS.incid_osmm_updates));

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

            HluDataSet.incid_secondaryDataTable secondaryTable = _hluDS.incid_secondary;
            _incidSecondaryRows = GetIncidChildRowsDb(relValues,
               _hluTableAdapterMgr.incid_secondaryTableAdapter, ref secondaryTable);
            _origIncidSecondaryCount = _incidSecondaryRows.Length;

            GetSecondaryHabitats();

            HluDataSet.incid_conditionDataTable incidConditionTable = _hluDS.incid_condition;
            _incidConditionRows = GetIncidChildRowsDb(relValues,
                _hluTableAdapterMgr.incid_conditionTableAdapter, ref incidConditionTable);
            _origIncidConditionCount = _incidConditionRows.Length;

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

            // Get the BAP habitats and compare them to those relating to the
            // primary and secondary codes.
            GetBapEnvironments();

            HluDataSet.incid_sourcesDataTable incidSourcesTable = _hluDS.incid_sources;
            _incidSourcesRows = GetIncidChildRowsDb(relValues,
                _hluTableAdapterMgr.incid_sourcesTableAdapter, ref incidSourcesTable);
            _origIncidSourcesCount = _incidSourcesRows.Length;

            HluDataSet.historyDataTable historyTable = _hluDS.history;
            _incidHistoryRows = GetIncidChildRowsDb(relValues,
                _hluTableAdapterMgr.historyTableAdapter, ref historyTable);

            HluDataSet.incid_osmm_updatesDataTable incidOSMMUpdatesTable = _hluDS.incid_osmm_updates;
            _incidOSMMUpdatesRows = GetIncidChildRowsDb(relValues,
                _hluTableAdapterMgr.incid_osmm_updatesTableAdapter, ref incidOSMMUpdatesTable);

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

                //string orderByClause;
                //if (_childRowOrderByDict.TryGetValue(typeof(C), out orderByClause))
                //    adapter.Fill(childTable, String.Format("{0} ORDER BY {1}",
                //        _db.WhereClause(false, true, true, childConds), orderByClause));
                //else
                //    adapter.Fill(childTable, childConds);

                //return (R[])childTable.Select();

                // Sort after the Fill as the Select seems to be re-sorting
                // the rows after the Fill.
                string orderByClause;
                adapter.Fill(childTable, childConds);

                if (_childRowOrderByDict.TryGetValue(typeof(C), out orderByClause))
                    return (R[])childTable.Select(null, orderByClause);
                else
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

        internal void GetIncidOSMMUpdatesRows(List<List<SqlFilterCondition>> whereClause,
            ref HluDataSet.incid_osmm_updatesDataTable table)
        {
            if ((whereClause != null) && (whereClause.Count > 0))
            {
                if (_hluTableAdapterMgr.incid_osmm_updatesTableAdapter == null)
                    _hluTableAdapterMgr.incid_osmm_updatesTableAdapter =
                        new HluTableAdapter<HluDataSet.incid_osmm_updatesDataTable,
                            HluDataSet.incid_osmm_updatesRow>(_db);

                _hluTableAdapterMgr.incid_osmm_updatesTableAdapter.Fill(table, whereClause);
            }
        }

        /// <summary>
        /// Determines whether any of the incid tables are dirty].
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
        internal bool IsDirtyIncid()
        {
            // If anything has changed in any of the data tables
            return ((_incidCurrentRow != null) && (_incidCurrentRow.RowState != DataRowState.Detached) &&
                ((_incidCurrentRow.Ishabitat_primaryNull() && !String.IsNullOrEmpty(_incidPrimary)) ||
                (!_incidCurrentRow.Ishabitat_primaryNull() && String.IsNullOrEmpty(_incidPrimary)) ||
                (_incidPrimary != _incidCurrentRow.habitat_primary) ||
                !CompareIncidCurrentRowClone()));
        }

        /// <summary>
        /// Determines whether the incid ihs matrix table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Determines whether the incid ihs formation table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Determines whether the incid ihs management table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Determines whether the incid ihs complex table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
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
        
        //TODO: IsDirty IncidSecondary - Check
        /// <summary>
        /// Determines whether the incid secondary table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
        internal bool IsDirtyIncidSecondary()
        {
            if (_incidSecondaryRows.Count(r => r.RowState == DataRowState.Deleted) > 0) return true;

            if (_incidSecondaryHabitats != null)
            {
                if (_incidSecondaryHabitats.Count(sh => IncidSecondaryRowDirty(sh)) > 0) return true;
            }

            if ((_incidSecondaryRows != null) && (_incidSecondaryHabitats.Count !=
                _incidSecondaryRows.Count())) return true;

            if (_incidSecondaryRows != null)
            {
                foreach (DataRow r in _incidSecondaryRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;
            }
            return false;
        }

        /// <summary>
        /// Determines whether the incid condition table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
        internal bool IsDirtyIncidCondition()
        {
            if (_incidConditionRows != null)
            {
                if (_incidConditionRows.Count(r => r != null) != _origIncidConditionCount) return true;

                foreach (DataRow r in _incidConditionRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidConditionCount != 0;
        }

        /// <summary>
        /// Determines whether the incid bap table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Determines whether the incid sources table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
        private bool IsDirtyIncidSources()
        {
            if (_incidSourcesRows != null)
            {
                if (_incidSourcesRows.Count(r => r != null) != _origIncidSourcesCount) return true;

                foreach (DataRow r in _incidSourcesRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;

                return false;
            }
            return _origIncidSourcesCount != 0;
        }

        /// <summary>
        /// Determines whether the incid osmm updates table is dirty.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if dirty otherwise, <c>false</c>.
        /// </returns>
        internal bool IsDirtyIncidOSMMUpdates()
        {
            if (_incidOSMMUpdatesRows != null)
            {
                foreach (DataRow r in _incidOSMMUpdatesRows)
                    if (ViewModelWindowMainHelpers.RowIsDirty(r)) return true;
            }
            return false;
        }

        private bool IncidSecondaryRowDirty(SecondaryHabitat sh)
        {
            // deleted secondary habitat row
            var q = _incidSecondaryRows.Where(r => r.RowState != DataRowState.Deleted && r.secondary_id == sh.secondary_id);
            switch (q.Count())
            {
                case 0:
                    return true; // new row;
                case 1:
                    if (!sh.IsValid() && sh.IsAdded) return true;

                    HluDataSet.incid_secondaryRow oldRow = q.ElementAt(0);
                    object[] itemArray = sh.ToItemArray();
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

        private bool IncidBapRowDirty(BapEnvironment be)
        {
            // deleted user BAP row
            var q = _incidBapRows.Where(r => r.RowState != DataRowState.Deleted && r.bap_id == be.bap_id);
            switch (q.Count())
            {
                case 0:
                    return true; // new row;
                case 1:
                    // Only flag an incid_bap row that is invalid as dirty if it has
                    // been added by the user. This allows existing records to be
                    // viewed in the user interface without warning the user that
                    // the data has changed.
                    if (!be.IsValid() && be.IsAdded) return true;

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
            OnPropertyChanged("TabItemSelected");
            RefreshBulkUpdateControls();
            RefreshOSMMUpdateControls();
            RefreshStatus();
            RefreshHeader();
            RefreshOSMMUpdate();
            RefreshHabitatTab();
            RefreshIHSTab();
            RefreshPriorityTab();
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
            OnPropertyChanged("OSMMBulkUpdateCommandHeader");
            OnPropertyChanged("TopControlsGroupHeader");
            OnPropertyChanged("TabItemHistoryEnabled");

            OnPropertyChanged("NumIncidSelectedDB");
            OnPropertyChanged("NumToidSelectedDB");
            OnPropertyChanged("NumFragmentsSelectedDB");

            OnPropertyChanged("NumIncidSelectedMap");
            OnPropertyChanged("NumToidSelectedMap");
            OnPropertyChanged("NumFragmentsSelectedMap");

            OnPropertyChanged("BapHabitatsAutoEnabled");
            OnPropertyChanged("BapHabitatsUserEnabled");
        }

        private void RefreshOSMMUpdateControls()
        {
            OnPropertyChanged("ShowInOSMMUpdateMode");
            OnPropertyChanged("HideInOSMMUpdateMode");
            OnPropertyChanged("OSMMUpdateCommandHeader");
            OnPropertyChanged("TopControlsGroupHeader");
            OnPropertyChanged("ShowReasonProcessGroup");
        }

        private void RefreshStatus()
        {
            OnPropertyChanged("EditMode");
            OnPropertyChanged("IncidCurrentRowIndex");
            OnPropertyChanged("OSMMIncidCurrentRowIndex");
            OnPropertyChanged("StatusIncid");
            OnPropertyChanged("StatusIncidToolTip");
            OnPropertyChanged("StatusBar");
            OnPropertyChanged("CanZoomSelection");
            OnPropertyChanged("CanBulkUpdate");
            OnPropertyChanged("CanBulkUpdateMode");
            OnPropertyChanged("CanOSMMUpdateMode");
            OnPropertyChanged("CanOSMMBulkUpdateMode");
            //---------------------------------------------------------------------
            // FIX: 103 Accept/Reject OSMM updates in edit mode.
            //
            OnPropertyChanged("CanOSMMUpdateAccept");
            OnPropertyChanged("CanOSMMUpdateReject");
            //---------------------------------------------------------------------
            OnPropertyChanged("IsFiltered");
            OnPropertyChanged("CanClearFilter");
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

        private void RefreshGroupHeaders()
        {
            OnPropertyChanged("TopControlsGroupHeader");

            OnPropertyChanged("HabitatHeader");
            OnPropertyChanged("PrimaryHeader");
            OnPropertyChanged("SecondaryHabitatsHeader");
            OnPropertyChanged("SecondarySummaryHeader");
            OnPropertyChanged("LegacyHeader");

            OnPropertyChanged("IhsHabitatHeader");
            OnPropertyChanged("IhsMatrixHeader");
            OnPropertyChanged("IhsFormationHeader");
            OnPropertyChanged("IhsManagementHeader");
            OnPropertyChanged("IhsComplexHeader");
            OnPropertyChanged("IhsSummaryHeader");

            OnPropertyChanged("DetailsCommentsHeader");
            OnPropertyChanged("DetailsSiteHeader");
            OnPropertyChanged("DetailsMapsHeader");
            OnPropertyChanged("DetailsConditionHeader");
            OnPropertyChanged("DetailsQualityHeader");

            OnPropertyChanged("ShowSource1Number");
            OnPropertyChanged("Source1Header");
            OnPropertyChanged("ShowSource2Number");
            OnPropertyChanged("Source2Header");
            OnPropertyChanged("ShowSource3Number");
            OnPropertyChanged("Source3Header");
        }

        private void RefreshOSMMUpdate()
        {
            OnPropertyChanged("ShowIncidOSMMPendingGroup");
            //TODO: OSMM Update - Replace with primary?
            //OnPropertyChanged("IncidOSMMIhsSummary");
            //OnPropertyChanged("IncidOSMMXRefID");
            
            OnPropertyChanged("IncidOSMMProcessFlag");
            OnPropertyChanged("IncidOSMMSpatialFlag");
            OnPropertyChanged("IncidOSMMChangeFlag");
            OnPropertyChanged("IncidOSMMUpdateStatus");
        }

        private void RefreshHabitatTab()
        {
            OnPropertyChanged("TabItemHabitatEnabled");
            OnPropertyChanged("TabHabitatControlsEnabled");
            OnPropertyChanged("HabitatTabLabel");
            //OnPropertyChanged("HabitatClassCodes");
            OnPropertyChanged("HabitatTypeCodes");
            OnPropertyChanged("HabitatType");
            OnPropertyChanged("HabitatClass");
            OnPropertyChanged("IncidPrimary");
            OnPropertyChanged("NvcCodes");
            OnPropertyChanged("IncidSecondaryHabitats");
            OnPropertyChanged("IncidSecondarySummary");
            OnPropertyChanged("LegacyHabitatCodes");
            OnPropertyChanged("IncidLegacyHabitat");
        }

        private void RefreshIHSTab()
        {
            OnPropertyChanged("TabItemIHSEnabled");
            OnPropertyChanged("TabIhsControlsEnabled");
            OnPropertyChanged("IHSTabLabel");
            OnPropertyChanged("IncidIhsHabitat");
            RefreshIhsMultiplexValues();
        }

        private void RefreshIhsMultiplexValues()
        {
            OnPropertyChanged("IncidIhsHabitatText");
            OnPropertyChanged("IncidIhsMatrix1Text");
            OnPropertyChanged("IncidIhsMatrix2Text");
            OnPropertyChanged("IncidIhsMatrix3Text");
            OnPropertyChanged("IncidIhsFormation1Text");
            OnPropertyChanged("IncidIhsFormation2Text");
            OnPropertyChanged("IncidIhsManagement1Text");
            OnPropertyChanged("IncidIhsManagement2Text");
            OnPropertyChanged("IncidIhsComplex1Text");
            OnPropertyChanged("IncidIhsComplex2Text");
            OnPropertyChanged("IncidIhsSummary");
        }

        private void RefreshPriorityTab()
        {
            OnPropertyChanged("TabItemPriorityEnabled");
            OnPropertyChanged("TabPriorityControlsEnabled");
            OnPropertyChanged("PriorityTabLabel");
            OnPropertyChanged("IncidBapHabitatsAuto");
            OnPropertyChanged("IncidBapHabitatsUser");
            OnPropertyChanged("BapHabitatsUserEnabled");
        }

        private void RefreshDetailsTab()
        {
            OnPropertyChanged("TabItemDetailsEnabled");
            OnPropertyChanged("TabDetailsControlsEnabled");
            OnPropertyChanged("DetailsTabLabel");

            OnPropertyChanged("IncidGeneralComments");
            OnPropertyChanged("IncidBoundaryBaseMap");
            OnPropertyChanged("IncidDigitisationBaseMap");
            OnPropertyChanged("BapHabitatsAutoEnabled");
            OnPropertyChanged("BapHabitatsUserEnabled");

            OnPropertyChanged("IncidSiteRef");
            OnPropertyChanged("IncidSiteName");

            OnPropertyChanged("ConditionCodes");
            OnPropertyChanged("IncidCondition");
            OnPropertyChanged("IncidConditionQualifier");
            OnPropertyChanged("IncidConditionDate");
            OnPropertyChanged("IncidConditionEnabled");

            OnPropertyChanged("QualityDeterminationCodes");
            OnPropertyChanged("IncidQualityDetermination");
            OnPropertyChanged("QualityInterpretationCodes");
            OnPropertyChanged("IncidQualityInterpretation");
            OnPropertyChanged("IncidQualityComments");
        }

        private void RefreshSources()
        {
            OnPropertyChanged("TabItemSourcesEnabled");
            OnPropertyChanged("TabSourcesControlsEnabled");
            OnPropertyChanged("SourcesTabLabel");
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
            OnPropertyChanged("TabItemHistoryEnabled");
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
                if ((_bulkUpdateMode == false && _osmmUpdateMode == false) && IncidCurrentRow == null) _reasonProcessEnabled = false;
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

        public int TabItemSelected
        {
            get { return _tabItemSelected; }
            set { _tabItemSelected = value; }
        }

        public bool TabItemHabitatEnabled
        {
            get { return _tabItemHabitatEnabled; }
            set { _tabItemHabitatEnabled = value; }
        }

        public bool TabItemPriorityEnabled
        {
            get { return _tabItemPriorityEnabled; }
            set { _tabItemPriorityEnabled = value; }
        }

        public bool TabItemDetailsEnabled
        {
            get { return _tabItemDetailsEnabled; }
            set { _tabItemDetailsEnabled = value; }
        }

        public bool TabItemIHSEnabled
        {
            get { return _tabItemIHSEnabled; }
            set { _tabItemIHSEnabled = value; }
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

        public bool TabHabitatControlsEnabled
        {
            get { return _tabHabitatControlsEnabled; }
            set { _tabHabitatControlsEnabled = value; }
        }

        public bool TabIhsControlsEnabled
        {
            get { return _tabIhsControlsEnabled; }
            set { _tabIhsControlsEnabled = value; }
        }

        public bool TabPriorityControlsEnabled
        {
            get { return _tabPriorityControlsEnabled; }
            set { _tabPriorityControlsEnabled = value; }
        }

        public bool TabDetailsControlsEnabled
        {
            get { return _tabDetailsControlsEnabled; }
            set { _tabDetailsControlsEnabled = value; }
        }

        public bool TabSourcesControlsEnabled
        {
            get { return _tabSourcesControlsEnabled; }
            set { _tabSourcesControlsEnabled = value; }
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
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Don't show the area when in OSMM Update mode and there are no
                // updates to process.
                //
                if ((_bulkUpdateMode == false) && (_osmmUpdateMode == false || _osmmUpdatesEmpty == false))
                //---------------------------------------------------------------------
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
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Don't show the length when in OSMM Update mode and there are no
                // updates to process.
                //
                if ((_bulkUpdateMode == false) && (_osmmUpdateMode == false || _osmmUpdatesEmpty == false))
                //---------------------------------------------------------------------
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
                {
                    if (IncidCurrentRow.created_date.ToShortTimeString() == "00:00")
                        return IncidCurrentRow.created_date.ToShortDateString();
                    else
                        return String.Format("{0} {1}", IncidCurrentRow.created_date.ToShortDateString(), IncidCurrentRow.created_date.ToShortTimeString());
                }
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
            get
            {
                if (IncidCurrentRow != null && _osmmUpdatesEmpty == false && _incidLastModifiedDate != DateTime.MinValue)
                {
                    if (_incidLastModifiedDate.ToShortTimeString() == "00:00")
                        return _incidLastModifiedDate.ToShortDateString();
                    else
                        return String.Format("{0} {1}", _incidLastModifiedDate.ToShortDateString(), _incidLastModifiedDate.ToShortTimeString());
                }
                else
                    return null;
            }
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
                // Display the created user's name from the lut_user table
                // (if found) instead of the user_id
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.created_user_idColumn))
                    return String.IsNullOrEmpty(IncidCurrentRow.lut_userRowByfk_incid_user_created.user_name)
                        ? IncidCurrentRow.created_user_id : IncidCurrentRow.lut_userRowByfk_incid_user_created.user_name;
                else
                    return null;
            }
            set { if ((IncidCurrentRow != null) && (value != null)) IncidCurrentRow.created_user_id = value; }
        }

        public string IncidLastModifiedUser
        {
            get
            {
                // Display the last modified user's name from the lut_user table
                // (if found) instead of the user_id
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.last_modified_user_idColumn))
                    return String.IsNullOrEmpty(IncidCurrentRow.lut_userRowByfk_incid_user_modified.user_name)
                        ? IncidCurrentRow.last_modified_user_id : IncidCurrentRow.lut_userRowByfk_incid_user_modified.user_name;
                else
                    return null;
            }
            set { if ((IncidCurrentRow != null) && (value != null)) _incidLastModifiedUser = value; }
        }

        #endregion

        #region OSMM Updates

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // 
        /// <summary>
        /// Only show the OSMM Updates group if required, otherwise collapse it.
        /// </summary>
        public Visibility ShowIncidOSMMPendingGroup
        {
            get
            {
                // Show the group if not in osmm update mode and
                // show updates are "Always" required, or "When Outstanding"
                // (i.e. update flag is "Proposed" (> 0) or "Pending" = 0).
                if ((_osmmUpdateMode == true) ||
                    (_bulkUpdateMode == false &&
                    (_showOSMMUpdates == "Always" ||
                    (_showOSMMUpdates == "When Outstanding" && (IncidOSMMStatus >= 0)))))
                {
                    // Adjust the window height if not already showing the group.
                    if (!_showingOSMMPendingGroup)
                    {
                        _showingOSMMPendingGroup = true;

                        // Adjust the window height.
                        AdjustWindowHeight(false);
                        OnPropertyChanged("WindowHeight");
                    }

                    return Visibility.Visible;
                }
                else
                {
                    // Adjust the window height if currently showing the group.
                    if (_showingOSMMPendingGroup)
                    {
                        _showingOSMMPendingGroup = false;

                        // Adjust the window height.
                        AdjustWindowHeight(false);
                        OnPropertyChanged("WindowHeight");
                    }

                    return Visibility.Collapsed;
                }
            }
            set { }
        }

        /// <summary>
        /// Gets the OSMM process flag that relates to the selected incid.
        /// It is used to show how the latest OSMM translation was
        /// processed.
        /// </summary>
        /// <value>
        /// The string of Process Flag related to the current incid.
        /// </value>
        public string IncidOSMMProcessFlag
        {
            get
            {
                if (_incidOSMMUpdatesProcessFlag != 0)
                    return _incidOSMMUpdatesProcessFlag.ToString();
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the OSMM spatial flag that relates to the selected incid.
        /// It is used to show how the latest OSMM translation was
        /// processed.
        /// </summary>
        /// <value>
        /// The string of Spatial Flag related to the current incid.
        /// </value>
        public string IncidOSMMSpatialFlag
        {
            get
            {
                return _incidOSMMUpdatesSpatialFlag;
            }
        }

        /// <summary>
        /// Gets the OSMM change flag that relates to the selected incid.
        /// It is used to show how the latest OSMM translation was
        /// processed.
        /// </summary>
        /// <value>
        /// The string of Change Flag related to the current incid.
        /// </value>
        public string IncidOSMMChangeFlag
        {
            get
            {
                return _incidOSMMUpdatesChangeFlag;
            }
        }

        /// <summary>
        /// Gets the OSMM updates status that relates to the selected incid.
        /// It contains the update status of the latest OSMM translation.
        /// </summary>
        /// <value>
        /// The integer value of OSMM Updates Status related to the current incid.
        /// </value>
        public Nullable<int> IncidOSMMStatus
        {
            get
            {
                return _incidOSMMUpdatesStatus;
            }
        }

        /// <summary>
        /// Gets the OSMM Update proposed Habitat Primary that relates to the
        /// selected incid. It is used to show how the latest OSMM translates.
        /// </summary>
        /// <value>
        /// The string of OSMM Habitat Primary related to the current incid.
        /// </value>
        public string IncidOSMMHabitatPrimary
        {
            get
            {
                if (_osmmHabitatXrefIds == null)
                {
                    _osmmHabitatXrefIds = from x in HluDataset.lut_osmm_habitat_xref
                                          where x.is_local
                                          select x;
                }

                if (_osmmUpdatesEmpty == true || _incidOSMMUpdatesOSMMXref <= 0) return null;

                var q = _osmmHabitatXrefIds.Where(h => h.osmm_habitat_xref_id == _incidOSMMUpdatesOSMMXref);
                if (q.Count() > 0)
                    return q.ElementAt(0).habitat_primary;
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the OSMM Update proposed Habitat Secondaries that relates to the
        /// selected incid. It is used to show how the latest OSMM translates.
        /// </summary>
        /// <value>
        /// The string of OSMM Habitat Secondaries related to the current incid.
        /// </value>
        public string IncidOSMMHabitatSecondaries
        {
            get
            {
                if (_osmmHabitatXrefIds == null)
                {
                    _osmmHabitatXrefIds = from x in HluDataset.lut_osmm_habitat_xref
                                          where x.is_local
                                          select x;
                }

                if (_osmmUpdatesEmpty == true || _incidOSMMUpdatesOSMMXref <= 0) return null;

                var q = _osmmHabitatXrefIds.Where(h => h.osmm_habitat_xref_id == _incidOSMMUpdatesOSMMXref);
                if (q.Count() > 0)
                    return q.ElementAt(0).habitat_secondaries;
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the OSMM XRef ID that relates to the selected Incid.
        /// It shows the cross-reference ID of the latest OSMM and
        /// how it translates directly to the primary and secondary
        /// habitats.
        /// </summary>
        /// <value>
        /// The string of OSMM XRef ID related to the current incid.
        /// </value>
        public string IncidOSMMXRefID
        {
            get
            {
                if (_incidOSMMUpdatesOSMMXref != 0)
                    return _incidOSMMUpdatesOSMMXref.ToString();
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the OSMM update flag that relates to the selected incid.
        /// It is used to show the update status of the latest OSMM
        /// translation.
        /// </summary>
        /// <value>
        /// The string interpretation of Update Flag related to the
        /// current incid.
        /// </value>
        public string IncidOSMMUpdateStatus
        {
            get
            {
                if (_incidOSMMUpdatesStatus != null)
                {
                    // Values greater than zero indicate proposed changes
                    if (_incidOSMMUpdatesStatus > 0)
                        return "Proposed";
                    else
                    {
                        switch (_incidOSMMUpdatesStatus)
                        {
                            case 0:
                                return "Pending";
                            case -1:
                                return "Applied";
                            case -2:
                                return "Ignored";
                            case -99:
                                return "Rejected";
                            default:
                                return null;
                        }
                    }
                }
                else
                    return null;
            }
        }

        /// <summary>
        /// Indicates if the OSMM updates status should be reset when an incid
        /// is manually updated.
        /// </summary>
        /// <value>
        /// Reset the update status when an incid is manually updated.
        /// </value>
        public bool ResetOSMMUpdatesStatus
        {
            get { return _resetOSMMUpdatesStatus; }
            set { _resetOSMMUpdatesStatus = value; }
        }

        #endregion

        #region Reason and Process

        /// <summary>
        /// Only show the Reason and Process group if the data is editable
        /// and not in OSMM edit mode, otherwise collapse it.
        /// </summary>
        public Visibility ShowReasonProcessGroup
        {
            get
            {
                if (_editMode == true && _osmmUpdateMode == false)
                {
                    if (!_showingReasonProcessGroup)
                    {
                        _showingReasonProcessGroup = true;

                        // Adjust the window height.
                        AdjustWindowHeight(false);
                        OnPropertyChanged("WindowHeight");
                    }
                    return Visibility.Visible;
                }
                else
                {
                    if (_showingReasonProcessGroup)
                    {
                        _showingReasonProcessGroup = false;

                        // Adjust the window height.
                        AdjustWindowHeight(false);
                        OnPropertyChanged("WindowHeight");
                    }

                    return Visibility.Collapsed;
                }
            }
            set { }
        }

        public HluDataSet.lut_reasonRow[] ReasonCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_reasonCodes == null)
                {
                    // Load the data table if not already loaded.
                    if (HluDataset.lut_reason.IsInitialized && (HluDataset.lut_reason.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_reasonTableAdapter == null)
                            _hluTableAdapterMgr.lut_reasonTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_reasonDataTable, HluDataSet.lut_reasonRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_reasonDataTable) }, false);
                    }

                    _reasonCodes = HluDataset.lut_reason.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
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
                // Get the list of values from the lookup table
                if (_processCodes == null)
                {
                    // Load the data table if not already loaded.
                    if (HluDataset.lut_process.IsInitialized && (HluDataset.lut_process.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_processTableAdapter == null)
                            _hluTableAdapterMgr.lut_processTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_processDataTable, HluDataSet.lut_processRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_processDataTable) }, false);
                    }

                    _processCodes = HluDataset.lut_process.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
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

        #region Habitat Tab

        // Set the Habitat tab label from here so that validation can be done.
        // This will enable tooltips to be shown so that validation errors
        // in any fields in the tab can be highlighted by flagging the tab
        // label as in error.
        public string HabitatTabLabel
        {
            get { return "Habitats"; }
        }

        #region Habitat Class

        public string HabitatHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Source Habitat";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the list of all local habitat class codes that have at
        /// least one habitat type that is local.
        /// </summary>
        /// <value>
        /// A list of habitat class codes.
        /// </value>
        public HluDataSet.lut_habitat_classRow[] HabitatClassCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_habitatClassCodes == null)
                {
                    // Load the lookup table is not already loaded.
                    if (HluDataset.lut_habitat_class.IsInitialized && (HluDataset.lut_habitat_class.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_habitat_classTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_classTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_classDataTable) }, false);
                    }

                    // Set the static variable (used in the options window) for all
                    // local habitat classes with local habitat types.
                    HabitatClasses = (from c in HluDataset.lut_habitat_class
                                          join t in HluDataset.lut_habitat_type on c.code equals t.habitat_class_code
                                          where c.is_local && t.is_local
                                          select c).Distinct().OrderBy(c => c.sort_order).ThenBy(c => c.description).ToArray();

                    // Set the habitat classes for all local habitat classes with
                    // local habitat types that relate to at least one primary
                    // habitat type that is local.
                    _habitatClassCodes = (from c in HluDataset.lut_habitat_class
                                                 join t in HluDataset.lut_habitat_type on c.code equals t.habitat_class_code
                                                 join p in HluDataset.lut_habitat_type_primary on t.code equals p.code_habitat_type
                                                 where c.is_local && t.is_local && p.is_local
                                                 select c).Distinct().OrderBy(c => c.sort_order).ThenBy(c => c.description).ToArray();
                }

                return _habitatClassCodes;
            }
        }

        /// <summary>
        /// Gets or sets the habitat class which will then load the list
        /// of habitat types related to that class.
        /// </summary>
        /// <value>
        /// The habitat class.
        /// </value>
        public string HabitatClass
        {
            get
            {
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Don't set the habitat class when there are no OSMM updates to process.
                //
                if (_habitatClass == null && _osmmUpdatesEmpty == false)
                    _habitatClass = _preferredHabitatClass;
                //---------------------------------------------------------------------
                return _habitatClass;
            }
            set
            {
                _habitatClass = value;

                if (!String.IsNullOrEmpty(_habitatClass))
                {
                    // Clear the habitat type and then reload the list of
                    // possible habitat types that relate to the selected
                    // habitat class.
                    _habitatType = null;
                    OnPropertyChanged("HabitatTypeCodes");

                    HabitatType = null;
                    OnPropertyChanged("HabitatType");
                }
                else
                {
                    _habitatTypeCodes = null;
                    OnPropertyChanged("HabitatTypeCodes");
                }

                if ((_habitatTypeCodes != null) && (_habitatTypeCodes.Count() == 1))
                    OnPropertyChanged("HabitatType");
            }
        }

        /// <summary>
        /// Gets the list of all local habitat type codes related to
        /// the selected habitat class that have at least one
        /// cross reference to a primary habitat.
        /// </summary>
        /// <value>
        /// A list of habitat type codes.
        /// </value>
        public HluDataSet.lut_habitat_typeRow[] HabitatTypeCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (!String.IsNullOrEmpty(HabitatClass))
                {
                    // Load the lookup table is not already loaded.
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    // Only load habitat classes that are flagged as local.
                    _habitatTypeCodes = (from t in HluDataset.lut_habitat_type
                                         // Only load codes with a local primary habitat type.
                                         //join i in HluDataset.lut_habitat_type_primary on t.code equals i.code_habitat_type
                                         where t.is_local && t.habitat_class_code == HabitatClass
                                         select t).Distinct().OrderBy(c => c.sort_order).ThenBy(c => c.description).ToArray();

                    return _habitatTypeCodes;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the habitat type which will then load the list
        /// of primary habitats related to that type.
        /// </summary>
        /// <value>
        /// The habitat type.
        /// </value>
        public string HabitatType
        {
            get { return _habitatType; }
            set
            {
                _habitatType = value;

                if (!String.IsNullOrEmpty(_habitatType))
                {
                    // Load all primary habitat codes where the primary habitat code
                    // and primary habitat category are both flagged as local and
                    // are related as local to the current habitat type.
                    _primaryCodes = from p in HluDataset.lut_primary
                                    join c in HluDataset.lut_primary_category on p.category equals c.code
                                    from t in HluDataset.lut_habitat_type_primary
                                    where t.code_habitat_type == _habitatType
                                    && p.is_local && c.is_local && t.is_local
                                    && (p.code == t.code_primary
                                    || (t.code_primary.EndsWith("*") && Regex.IsMatch(p.code, @"\A" + t.code_primary.TrimEnd('*') + @"") == true))
                                    select p;
                }
                else
                {
                    // Load all primary habitat codes where the primary habitat code
                    // and primary habitat category are both flagged as local.
                    _primaryCodes = from p in HluDataset.lut_primary
                                    join c in HluDataset.lut_primary_category on p.category equals c.code
                                    where p.is_local && c.is_local
                                    select p;
                }

                OnPropertyChanged("PrimaryCodes");
                OnPropertyChanged("PrimaryEnabled");
                OnPropertyChanged("NvcCodes");
            }
        }

        #endregion

        #region Primary Habitat

        public string PrimaryHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Primary Habitat";
                else
                    return null;
            }
        }

        public bool PrimaryEnabled
        {
            get
            {
                return true;
            }
        }
        
        /// <summary>
        /// Gets the primary codes.
        /// </summary>
        /// <value>
        /// The primary codes.
        /// </value>
        public HluDataSet.lut_primaryRow[] PrimaryCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (HluDataset.lut_primary.IsInitialized && HluDataset.lut_primary.Count == 0)
                {
                    // Load the lookup table if not already loaded.
                    if (_hluTableAdapterMgr.lut_primaryTableAdapter == null)
                        _hluTableAdapterMgr.lut_primaryTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_primaryDataTable, HluDataSet.lut_primaryRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_primaryDataTable) }, false);
                }

                // Return the primary codes if already set, otherwise get the
                // primary codes from the lookup table.
                if (_primaryCodes != null)
                {
                    return _primaryCodes.ToArray();
                }
                else if (!String.IsNullOrEmpty(IncidPrimary))
                {
                    // Load all primary habitat codes where the primary habitat code
                    // and primary habitat category are both flagged as local.
                    _primaryCodes = HluDataset.lut_primary.Where(r => r.code == IncidPrimary);
                    _primaryCodes = from p in HluDataset.lut_primary
                                    join c in HluDataset.lut_primary_category on p.category equals c.code
                                    where p.is_local && c.is_local && p.code == IncidPrimary
                                    select p;
                    return _primaryCodes.ToArray();
                }
                else
                {
                    return null;
                }
            }
        }

        public string IncidPrimary
        {
            get { return _incidPrimary; }
            set
            {
                if (IncidCurrentRow != null)
                {
                    if (_pasting && (_primaryCodes == null || _primaryCodes.Count(r => r.code == value) == 0))
                    {
                        _pasting = false;
                    }

                    _incidPrimary = value;

                    // Set the list of secondary codes for the primary habitat.
                    NewPrimaryHabitat(_incidPrimary);

                    // Refresh the BAP habitat environments (in case secondary codes
                    // are, or should be, reflected).
                    GetBapEnvironments();
                    OnPropertyChanged("IncidBapHabitatsAuto");
                    OnPropertyChanged("IncidBapHabitatsUser");
                    OnPropertyChanged("BapHabitatsAutoEnabled");
                    OnPropertyChanged("BapHabitatsUserEnabled");

                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }

        private void NewPrimaryHabitat(string incidPrimary)
        {
            if (incidPrimary != null)
            {
                // Set the primary habitat category.
                _incidPrimaryCategory = HluDataset.lut_primary.Where(p => p.code == incidPrimary).ElementAt(0).category;

                // Store all secondary habitat codes that are flagged as local for
                // all secondary groups that relate to the primary habitat category.
                _secondaryCodesValid = (from s in SecondaryHabitatCodesAll
                                        join p in HluDataset.lut_primary_secondary on s.code equals p.code_secondary
                                        where p.category == _incidPrimaryCategory
                                        && p.is_local
                                        select s).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();

                // Store the list of valid secondary codes.
                SecondaryHabitat.ValidSecondaryCodes = _secondaryCodesValid.Select(s => s.code);
            }
            else
            {
                _incidPrimaryCategory = null;
                _secondaryCodesValid = null;

                // Clear the list of valid secondary codes.
                SecondaryHabitat.ValidSecondaryCodes = null;
            }

            // Refresh the related fields
            OnPropertyChanged("NvcCodes");

            OnPropertyChanged("SecondaryGroupCodes");
            _secondaryGroup = _preferredSecondaryGroup;
            OnPropertyChanged("SecondaryGroup");
            OnPropertyChanged("SecondaryHabitatCodes");

            OnPropertyChanged("SecondaryGroupEnabled");
            OnPropertyChanged("SecondaryHabitatEnabled");

            OnPropertyChanged("CanAddSecondaryHabitat");
            OnPropertyChanged("CanAddSecondaryHabitatList");

            // Refresh secondary table to re-trigger the validation.
            RefreshSecondaryHabitats();

        }

        public string IncidPrimaryCategory
        {
            get { return _incidPrimaryCategory; }
        }

        /// <summary>
        /// Only show the NVC Codes if the option is set, otherwise collapse it.
        /// </summary>
        public Visibility ShowNVCCodes
        {
            get
            {
                // If should be showing NVC codes
                if (_showNVCCodes)
                {
                    return Visibility.Visible;
                }
                else  // If shouldn't be showing NVC codes
                {
                    return Visibility.Collapsed;
                }
            }
            set { }
        }

        /// <summary>
        /// Gets the string of NVC codes that are related to the selected
        /// primary habitat. It is used as an aid to the user to help double-
        /// check they have selected the correct primary habitat.
        /// </summary>
        /// <value>
        /// The string of NVC codes related to the current primary habitat.
        /// </value>
        public string NvcCodes
        {
            get
            {
                if (String.IsNullOrEmpty(_incidPrimary)) return null;

                // Select NVC codes based on current primary habitat
                var q = _primaryCodes.Where(h => h.code == _incidPrimary);
                if (q.Count() > 0)
                    return q.ElementAt(0).nvc_codes;
                else
                    return null;
            }
        }

        #endregion

        #region Secondary Habitats

        public string SecondaryHabitatsHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Secondary Habitats";
                else
                    return null;
            }
        }

        public bool SecondaryGroupEnabled
        {
            get
            {
                return (!String.IsNullOrEmpty(_incidPrimary));
            }
        }

        public bool SecondaryHabitatEnabled
        {
            get
            {
                return (!String.IsNullOrEmpty(_incidPrimary) && !String.IsNullOrEmpty(_secondaryGroup));
            }
        }

        public bool SecondaryHabitatsEnabled
        {
            get
            {
                return true;
            }
        }

        public HluDataSet.lut_secondary_groupRow[] SecondaryGroupCodes
        {
            get
            {
                // Load the lookup table if not already loaded.
                if (HluDataset.lut_secondary_group.IsInitialized && HluDataset.lut_secondary_group.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_secondary_groupTableAdapter == null)
                        _hluTableAdapterMgr.lut_secondary_groupTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_secondary_groupDataTable,
                                HluDataSet.lut_secondary_groupRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_secondary_groupDataTable) }, false);
                }

                // Set the public and static variables
                if (SecondaryGroupsAll == null || SecondaryGroupsAll.Count() == 0)
                {
                    // Set the full list of local secondary groups.
                    _secondaryGroups = (from sg in HluDataset.lut_secondary_group
                                            where sg.is_local
                                            select sg).OrderBy(r => r.sort_order).ThenBy(r => r.description).Distinct().ToArray();

                    // Set the full list of secondary groups including an <All> group.
                    HluDataSet.lut_secondary_groupRow[] secondaryGroupsAll;
                    secondaryGroupsAll = _secondaryGroups;
                    if (secondaryGroupsAll != null)
                    {
                        HluDataSet.lut_secondary_groupRow allRow = HluDataset.lut_secondary_group.Newlut_secondary_groupRow();
                        allRow.code = "<All>";
                        allRow.description = "<All>";
                        allRow.sort_order = -1;
                        secondaryGroupsAll = secondaryGroupsAll.Concat(
                            new HluDataSet.lut_secondary_groupRow[] { allRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                    }

                    // Set the static variable
                    SecondaryGroupsAll = secondaryGroupsAll;

                    // Set the dictionary of local secondary group codes.
                    SecondaryHabitat.SecondaryGroupCodes = (from sg in HluDataset.lut_secondary
                                                            where sg.is_local
                                                            select sg).OrderBy(r => r.code).ThenBy(r => r.code_group).ToDictionary(r => r.code, r => r.code_group);
                }

                if (!String.IsNullOrEmpty(IncidPrimary))
                {
                    // Set the valid list of secondary codes for the primary category.
                    _secondaryGroupsValid = (from sg in HluDataset.lut_secondary_group
                                            join s in HluDataset.lut_secondary on sg.code equals s.code_group
                                            join ps in HluDataset.lut_primary_secondary on s.code equals ps.code_secondary
                                             where ps.category == IncidPrimaryCategory
                                             && sg.is_local
                                             && s.is_local
                                             && ps.is_local
                                             select sg).OrderBy(r => r.sort_order).ThenBy(r => r.description).Distinct().ToArray();

                    if (_secondaryGroupsValid != null)
                    {
                        // Add the <ALL> group containing all secondary codes
                        HluDataSet.lut_secondary_groupRow allRow = HluDataset.lut_secondary_group.Newlut_secondary_groupRow();
                        allRow.code = "<All>";
                        allRow.description = "<All>";
                        allRow.sort_order = -1;
                        _secondaryGroupsValid = _secondaryGroupsValid.Concat(
                            new HluDataSet.lut_secondary_groupRow[] { allRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                    }
                }
                else
                {
                    // Set the valid list of secondary codes to all codes (rather than clearing the list)
                    _secondaryGroupsValid = _secondaryGroups;

                    // Add the <ALL> group containing all secondary codes
                    HluDataSet.lut_secondary_groupRow allRow = HluDataset.lut_secondary_group.Newlut_secondary_groupRow();
                    allRow.code = "<All>";
                    allRow.description = "<All>";
                    allRow.sort_order = -1;
                    _secondaryGroupsValid = _secondaryGroupsValid.Concat(
                        new HluDataSet.lut_secondary_groupRow[] { allRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();

                    // Set the combo box list to null (it will also be disabled).
                    return null;
                }

                return _secondaryGroupsValid;
            }
        }

        public HluDataSet.lut_secondary_groupRow[] SecondaryGroupCodesAll
        {
            get
            {
                return _secondaryGroups;
            }
        }

        /// <summary>
        /// Gets or sets the secondary group which will then load the list
        /// of secondary habitats related to that group.
        /// </summary>
        /// <value>
        /// The secondary group.
        /// </value>
        public string SecondaryGroup
        {
            get { return _secondaryGroup; }
            //get
            //{
            //    if (_secondaryGroup == null)
            //        _secondaryGroup = _preferredSecondaryGroup;
            //    return _secondaryGroup;
            //}
            set
            {
                _secondaryGroup = value;
                OnPropertyChanged("SecondaryHabitatEnabled");
                OnPropertyChanged("SecondaryHabitatCodes");
                OnPropertyChanged("CanAddSecondaryHabitat");
            }
        }

        public HluDataSet.lut_secondaryRow[] SecondaryHabitatCodes
        {
            get
            {
                if (!String.IsNullOrEmpty(_incidPrimary) && !String.IsNullOrEmpty(_secondaryGroup))
                {
                    // If the secondary codes must be valid
                    if (_secondaryCodeValidation > 0)
                    {
                        if (_secondaryGroup == "<All>")
                        {
                            // Load all secondary habitat codes that are flagged as local for
                            // all secondary groups that relate to the primary habitat.
                            return _secondaryCodesValid;
                        }
                        else
                        {
                            // Load all secondary habitat codes that are flagged as local and
                            // relate to the primary habitat and selected secondary group.
                            return _secondaryCodesValid.Where(s => s.code_group == _secondaryGroup).ToArray();
                        }
                    }
                    else
                    {
                        if (_secondaryGroup == "<All>")
                        {
                            // Load all secondary habitat codes that are flagged as local
                            // regardless of the primary habitat.
                            return _secondaryCodesAll;
                        }
                        else
                        {
                            // Load all secondary habitat codes that are flagged as local
                            // regardless of the primary habitat but relater to the
                            // selected secondary group.
                            return _secondaryCodesAll.Where(s => s.code_group == _secondaryGroup).ToArray();
                        }
                    }
                }
                else
                {
                    // Set the combo box list to null (it will also be disabled).
                    return null;
                }
            }
        }

        public HluDataSet.lut_secondaryRow[] SecondaryHabitatCodesAll
        {
            get
            {
                // Load the lookup table if not already loaded.
                if (HluDataset.lut_secondary.IsInitialized && (HluDataset.lut_secondary.Rows.Count == 0))
                {
                    if (_hluTableAdapterMgr.lut_secondaryTableAdapter == null)
                        _hluTableAdapterMgr.lut_secondaryTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_secondaryDataTable,
                                HluDataSet.lut_secondaryRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_secondaryDataTable) }, false);
                }

                if (_secondaryCodesAll == null)
                {
                    _secondaryCodesAll = (from s in HluDataset.lut_secondary
                                            where s.is_local
                                            select s).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _secondaryCodesAll;
            }
        }

        /// <summary>
        /// Gets or sets the secondary habitat.
        /// </summary>
        /// <value>
        /// The secondary habitat.
        /// </value>
        public string SecondaryHabitatCode
        {
            get { return _secondaryHabitat; }
            set
            {
                _secondaryHabitat = value;
                OnPropertyChanged("CanAddSecondaryHabitat");
            }
        }

        /// <summary>
        /// The collection of secondary habitats.
        /// </summary>
        public ObservableCollection<SecondaryHabitat> IncidSecondaryHabitats
        {
            get { return _incidSecondaryHabitats; }
            set
            {
                _incidSecondaryHabitats = value;

                // Set the new list of secondary habitat rows for the class.
                SecondaryHabitat.SecondaryHabitatList = _incidSecondaryHabitats;

                // Refresh the secondary habitat table (as they have been pasted).
                RefreshSecondaryHabitats();
                //OnPropertyChanged("IncidSecondarySummary");   // Doesn't seem to be needed.

                // Refresh the BAP habitat environments (in case secondary codes
                // are, or should be, reflected).
                GetBapEnvironments();
                OnPropertyChanged("IncidBapHabitatsAuto");
                OnPropertyChanged("IncidBapHabitatsUser");
                OnPropertyChanged("BapHabitatsAutoEnabled");
                OnPropertyChanged("BapHabitatsUserEnabled");
                
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public string SecondarySummaryHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Secondary Habitats Summary";
                else
                    return null;
            }
        }

        public string IncidSecondarySummary
        {
            get
            {
                _incidSecondarySummary = String.Join(_secondaryCodeDelimiter, _incidSecondaryHabitats
                    .OrderBy(s => s.secondary_habitat_int)
                    .ThenBy(s => s.secondary_habitat)
                    .Select(s => s.secondary_habitat_int)
                    .Distinct().ToList());
                return _incidSecondarySummary == String.Empty ? null : _incidSecondarySummary;
            }
        }

        /// <summary>
        /// Gets the secondary habitats.
        /// </summary>
        public void GetSecondaryHabitats()
        {
            // Remove any existing handlers before assigning a new collection.
            if (_incidSecondaryHabitats != null)
                _incidSecondaryHabitats.CollectionChanged -= _incidSecondaryHabitats_CollectionChanged;

            // Identify any secondary habitat rows that have not been marked as deleted.
            IEnumerable<HluDataSet.incid_secondaryRow> incidSecondaryRowsUndel =
                _incidSecondaryRows.Where(r => r.RowState != DataRowState.Deleted);

            // If there are any rows not marked as deleted add them to the collection.
            if (incidSecondaryRowsUndel != null)
            {
                // Order the secondary codes as required
                switch (_secondaryCodeOrder)
                {
                    case "As entered":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                           incidSecondaryRowsUndel.OrderBy(r => r.secondary_id).Select(r => new SecondaryHabitat(_bulkUpdateMode == true, r)));
                        break;
                    case "By group then code":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                           incidSecondaryRowsUndel.OrderBy(r => r.secondary_group).ThenBy(r => r.secondary).Select(r => new SecondaryHabitat(_bulkUpdateMode == true, r)));
                        break;
                    case "By code":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                           incidSecondaryRowsUndel.OrderBy(r => r.secondary).Select(r => new SecondaryHabitat(_bulkUpdateMode == true, r)));
                        break;
                    default:
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                           incidSecondaryRowsUndel.OrderBy(r => r.secondary_id).Select(r => new SecondaryHabitat(_bulkUpdateMode == true, r)));
                        break;
                }

                //_incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                //   incidSecondaryRowsUndel.Select(r => new SecondaryHabitat(_bulkUpdateMode == true, r)));
            }
            else
            {
                // Otherwise there can't be any secondary habitat rows so
                // set a new collection.
                _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>();
            }

            // Track any changes to the user rows collection.
            _incidSecondaryHabitats.CollectionChanged += _incidSecondaryHabitats_CollectionChanged;

            // Set the new list of secondary habitat rows for the class.
            SecondaryHabitat.SecondaryHabitatList = _incidSecondaryHabitats;

            // Set the validation type in the secondary habitat environment.
            SecondaryHabitat.SecondaryCodeValidation = _secondaryCodeValidation;

            // Check if there are any errors in the secondary habitat records to see
            // if the Habitats tab label should be flagged as also in error.
            if (_incidSecondaryHabitats != null && _incidSecondaryHabitats.Count > 0)
            {
                int countInvalid = _incidSecondaryHabitats.Count(sh => !sh.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _habitatErrors, "SecondaryHabitat");
                else
                    DelErrorList(ref _habitatErrors, "SecondaryHabitat");
            }
            else
                DelErrorList(ref _habitatErrors, "SecondaryHabitat");

            OnPropertyChanged("IncidSecondaryHabitats");
            OnPropertyChanged("HabitatTabLabel");
        }

        /// <summary>
        /// Add a secondary habitat.
        /// </summary>
        public bool AddSecondaryHabitat(bool bulkUpdateMode, int secondary_id, string incid, string secondary_habitat, string secondary_group)
        {
            // Store old secondary habitats list
            ObservableCollection<SecondaryHabitat> oldSecondaryHabs = _incidSecondaryHabitats;

            // Remove any existing handlers before assigning a new collection.
            if (_incidSecondaryHabitats != null)
                _incidSecondaryHabitats.CollectionChanged -= _incidSecondaryHabitats_CollectionChanged;

            // If there are any existing rows add the new row the collection
            // and then sort them.
            if (_incidSecondaryHabitats != null)
                _incidSecondaryHabitats.Add(new SecondaryHabitat(false, -1, Incid, secondary_habitat, secondary_group));
            else
            {
                // Otherwise there can't be any secondary habitat rows so
                // just create a new collection and add the new row.
                _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>();
                _incidSecondaryHabitats.Add(new SecondaryHabitat(false, -1, Incid, secondary_habitat, secondary_group));
            }

            // Track any changes to the user rows collection.
            _incidSecondaryHabitats.CollectionChanged += _incidSecondaryHabitats_CollectionChanged;

            // Set the new list of secondary habitat rows for the class.
            SecondaryHabitat.SecondaryHabitatList = _incidSecondaryHabitats;

            // Track when the secondary habitat records have changed so that the apply
            // button will appear.
            Changed = true;

            return (_incidSecondaryHabitats == null || (oldSecondaryHabs != null && _incidSecondaryHabitats != oldSecondaryHabs));
        }

        /// <summary>
        /// Refresh the secondary habitat table.
        /// </summary>
        public void RefreshSecondaryHabitats()
        {
            // If there are any existing rows then (re)sort them.
            if (_incidSecondaryHabitats != null)
            {
                // Remove any existing handlers before assigning a new collection.
                _incidSecondaryHabitats.CollectionChanged -= _incidSecondaryHabitats_CollectionChanged;

                // Order the secondary codes as required
                switch (_secondaryCodeOrder)
                {
                    case "As entered":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                            _incidSecondaryHabitats.OrderBy(r => r.secondary_id));
                        break;
                    case "By group then code":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                            _incidSecondaryHabitats.OrderBy(r => r.secondary_group).ThenBy(r => r.secondary_habitat_int));
                        break;
                    case "By code":
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                            _incidSecondaryHabitats.OrderBy(r => r.secondary_habitat_int));
                        break;
                    default:
                        _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(
                            _incidSecondaryHabitats.OrderBy(r => r.secondary_id));
                        break;
                }

                // Track any changes to the user rows collection.
                _incidSecondaryHabitats.CollectionChanged += _incidSecondaryHabitats_CollectionChanged;

                // Check if there are any errors in the secondary habitat records to see
                // if the Habitats tab label should be flagged as also in error.
                if (_incidSecondaryHabitats.Count > 0)
                {
                    int countInvalid = _incidSecondaryHabitats.Count(sh => !sh.IsValid());
                    if (countInvalid > 0)
                        AddErrorList(ref _habitatErrors, "SecondaryHabitat");
                    else
                        DelErrorList(ref _habitatErrors, "SecondaryHabitat");
                }
                else
                    DelErrorList(ref _habitatErrors, "SecondaryHabitat");

                OnPropertyChanged("IncidSecondaryHabitats");
                OnPropertyChanged("HabitatTabLabel");
            }
        }

        private void _incidSecondaryHabitats_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged("Error");

            // Track when the secondary habitat records have changed so that the apply
            // button will appear.
            Changed = true;

            // Check if there are any errors in the secondary habitat records to see
            // if the Habitats tab label should be flagged as also in error.
            if (_incidSecondaryHabitats != null && _incidSecondaryHabitats.Count > 0)
            {
                int countInvalid = _incidSecondaryHabitats.Count(sh => !sh.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _habitatErrors, "SecondaryHabitat");
                else
                    DelErrorList(ref _habitatErrors, "SecondaryHabitat");
            }
            else
            {
                DelErrorList(ref _habitatErrors, "SecondaryHabitat");
            }

            // Update the list of secondary habitat rows for the class.
            SecondaryHabitat.SecondaryHabitatList = _incidSecondaryHabitats;

            // Refresh secondary table and summary.
            RefreshSecondaryHabitats();
            OnPropertyChanged("IncidSecondarySummary");

            // Refresh the BAP habitat environments (in case secondary codes
            // are, or should be, reflected).
            GetBapEnvironments();
            OnPropertyChanged("IncidBapHabitatsAuto");
            OnPropertyChanged("IncidBapHabitatsUser");
            OnPropertyChanged("BapHabitatsAutoEnabled");
            OnPropertyChanged("BapHabitatsUserEnabled");

            OnPropertyChanged("HabitatTabLabel");
        }

        #endregion

        #region Legacy

        public string LegacyHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Legacy";
                else
                    return null;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR44 (Editable Legacy Habitat field)
        //
        /// <summary>
        /// Looks up the legacy habitat codes and descriptions from
        /// the database table 'lut_legacy_habitat'.
        /// </summary>
        /// <value>
        /// The sorted rows of all Legacy Habitats from the database.
        /// </value>
        public HluDataSet.lut_legacy_habitatRow[] LegacyHabitatCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (HluDataset.lut_legacy_habitat.IsInitialized && HluDataset.lut_legacy_habitat.Count == 0)
                {
                    // Load the lookup table if not already loaded.
                    if (_hluTableAdapterMgr.lut_legacy_habitatTableAdapter == null)
                        _hluTableAdapterMgr.lut_legacy_habitatTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_legacy_habitatDataTable, HluDataSet.lut_legacy_habitatRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_legacy_habitatDataTable) }, false);
                }

                if (!String.IsNullOrEmpty(IncidLegacyHabitat))
                {
                    HluDataSet.lut_legacy_habitatRow clearRow = HluDataset.lut_legacy_habitat.Newlut_legacy_habitatRow();
                    clearRow.code = _codeDeleteRow;
                    clearRow.description = _codeDeleteRow;
                    clearRow.sort_order = -1;

                    return HluDataset.lut_legacy_habitat.AsEnumerable().Concat(
                        new HluDataSet.lut_legacy_habitatRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }
                else
                {
                    return HluDataset.lut_legacy_habitat.AsEnumerable().ToArray();
                }
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR44 (Editable Legacy Habitat field)
        //
        /// <summary>
        /// Gets the Legacy Habitat code for the current incid.
        /// </summary>
        /// <value>
        /// The Legacy Habitat code for the current incid.
        /// </value>
        public string IncidLegacyHabitat
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.legacy_habitatColumn))
                    _incidLegacyHabitat = IncidCurrentRow.legacy_habitat;
                else
                    _incidLegacyHabitat = null;

                return _incidLegacyHabitat;
            }
            set
            {
                if (IncidCurrentRow != null)
                {
                    bool clearCode = value == _codeDeleteRow;
                    bool newCode = false;
                    if (clearCode)
                        value = null;
                    else
                        newCode = ((String.IsNullOrEmpty(_incidLegacyHabitat)) && (!String.IsNullOrEmpty(value)));

                    _incidLegacyHabitat = value;
                    IncidCurrentRow.legacy_habitat = _incidLegacyHabitat;

                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;

                    // Refresh legacy habitat list
                    if (clearCode || newCode)
                        OnPropertyChanged("LegacyHabitatCodes");
                }
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #endregion

        #region IHS Tab

        /// <summary>
        /// Gets the IHS tab label.
        /// </summary>
        /// <value>
        /// The IHS tab label.
        /// </value>
        public string IHSTabLabel
        {
            get { return "IHS"; }
        }

        /// <summary>
        /// Show or hide the IHS tab.
        /// </summary>
        /// <value>
        /// The visibility of the IHS tab.
        /// </value>
        public Visibility ShowIHSTab
        {
            get
            {
                if ((bool)_showIHSTab)
                    return Visibility.Visible;
                else
                    return Visibility.Collapsed;
            }
        }

        #region IHS Habitat

        /// <summary>
        /// Gets the IHS habitat group header.
        /// </summary>
        /// <value>
        /// The IHS habitat header.
        /// </value>
        public string IhsHabitatHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "IHS Habitat";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the incid IHS habitat.
        /// </summary>
        /// <value>
        /// The incid IHS habitat.
        /// </value>
        public string IncidIhsHabitat
        {
            get { return _incidIhsHabitat; }
        }

        /// <summary>
        /// Gets the incid ihs habitat code and description.
        /// </summary>
        /// <value>
        /// The incid ihs habitat code and description.
        /// </value>
        public string IncidIhsHabitatText
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_lutIhsHabitatCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_habitat.IsInitialized && (HluDataset.lut_ihs_habitat.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_habitatTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_habitatTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_habitatDataTable, HluDataSet.lut_ihs_habitatRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_habitatDataTable) }, false);
                    }

                    _lutIhsHabitatCodes = from x in HluDataset.lut_ihs_habitat
                                          //where x.is_local
                                          select x;
                }

                // Return the concatenated habitat code and description.
                if (_incidIhsHabitat == null)
                    return null;
                else
                {
                    var q = _lutIhsHabitatCodes.Where(h => h.code == _incidIhsHabitat);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        #endregion

        #region IHS Matrix

        /// <summary>
        /// Gets the ihs matrix group header.
        /// </summary>
        /// <value>
        /// The ihs matrix group header.
        /// </value>
        public string IhsMatrixHeader
        {
            get
            {
                // Optionally hide group headers to reduce window height.
                if ((bool)_showGroupHeaders)
                    return "IHS Matrix";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the first matrix code for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs matrix value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the first combined matrix code and test for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs matrix text.
        /// </value>
        public string IncidIhsMatrix1Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsMatrixCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_matrix.IsInitialized && (HluDataset.lut_ihs_matrix.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_matrixTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_matrixTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_matrixDataTable) }, false);
                    }

                    _lutIhsMatrixCodes = from x in HluDataset.lut_ihs_matrix
                                            //where x.is_local
                                            select x;
                }

                // Return the combined code and text value
                if (IncidIhsMatrix1 == null)
                    return null;
                else
                {
                    var q = _lutIhsMatrixCodes.Where(m => m.code == _incidIhsMatrixRows[0].matrix);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets the second matrix code for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs matrix value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the second combined matrix code and test for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs matrix text.
        /// </value>
        public string IncidIhsMatrix2Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsMatrixCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_matrix.IsInitialized && (HluDataset.lut_ihs_matrix.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_matrixTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_matrixTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_matrixDataTable) }, false);
                    }

                    _lutIhsMatrixCodes = from x in HluDataset.lut_ihs_matrix
                                            //where x.is_local
                                            select x;
                }

                // Return the combined code and text value.
                if (IncidIhsMatrix2 == null)
                    return null;
                else
                {
                    var q = _lutIhsMatrixCodes.Where(m => m.code == _incidIhsMatrixRows[1].matrix);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets the third matrix code for this incid.
        /// </summary>
        /// <value>
        /// The third incid ihs matrix value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the third combined matrix code and test for this incid.
        /// </summary>
        /// <value>
        /// The third incid ihs matrix text.
        /// </value>
        public string IncidIhsMatrix3Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsMatrixCodes == null)
                {
                // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_matrix.IsInitialized && (HluDataset.lut_ihs_matrix.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_matrixTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_matrixTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_matrixDataTable) }, false);
                    }

                    _lutIhsMatrixCodes = from x in HluDataset.lut_ihs_matrix
                                         //where x.is_local
                                         select x;
                }

                // Return the combined code and text value.
                if (IncidIhsMatrix3 == null)
                    return null;
                else
                {
                    var q = _lutIhsMatrixCodes.Where(m => m.code == _incidIhsMatrixRows[2].matrix);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        //TODO: Check IHS matrix - Needed?        
        /// <summary>
        /// Checks if there are any valid ihs matrix rows.
        /// </summary>
        /// <returns></returns>
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
            //return true;
        }

        /// <summary>
        /// Removes the incid ihs matrix rows.
        /// </summary>
        public void RemoveIncidIhsMatrixRows()
        {
            //TODO: Update remove IHS codes - check needed
            if (CheckIhsMatrix())   // Needed?
            {
                for (int i = 0; i < _incidIhsMatrixRows.Length; i++)
                {
                    if (_incidIhsMatrixRows[i].RowState != DataRowState.Detached)
                        _incidIhsMatrixRows[i].Delete();
                    _incidIhsMatrixRows[i] = null;
                }
            }
        }

        #endregion

        #region IHS Formation

        /// <summary>
        /// Gets the ihs formation group header.
        /// </summary>
        /// <value>
        /// The ihs formation group header.
        /// </value>
        public string IhsFormationHeader
        {
            get
            {
                // Optionally hide group headers to reduce window height.
                if ((bool)_showGroupHeaders)
                    return "IHS Formation";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the first formation code for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs formation value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the first combined formation code and test for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs formation text.
        /// </value>
        public string IncidIhsFormation1Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsFormationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_formation.IsInitialized && (HluDataset.lut_ihs_formation.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_formationTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_formationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_formationDataTable) }, false);
                    }

                    _lutIhsFormationCodes = from x in HluDataset.lut_ihs_formation
                                         //where x.is_local
                                         select x;
                }

                // Return the combined code and text value.
                if (IncidIhsFormation1 == null)
                    return null;
                else
                {
                    var q = _lutIhsFormationCodes.Where(m => m.code == _incidIhsFormationRows[0].formation);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets the second formation code for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs formation value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the second combined formation code and test for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs formation text.
        /// </value>
        public string IncidIhsFormation2Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsFormationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_formation.IsInitialized && (HluDataset.lut_ihs_formation.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_formationTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_formationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_formationDataTable) }, false);
                    }

                    _lutIhsFormationCodes = from x in HluDataset.lut_ihs_formation
                                            //where x.is_local
                                            select x;
                }

                // Return the combined code and text value.
                if (IncidIhsFormation2 == null)
                    return null;
                else
                {
                    var q = _lutIhsFormationCodes.Where(m => m.code == _incidIhsFormationRows[1].formation);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        //TODO: Check IHS formation - Needed?
        /// <summary>
        /// Checks if there are any valid ihs formation rows.
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Removes the incid ihs formation rows.
        /// </summary>
        public void RemoveIncidIhsFormationRows()
        {
            //TODO: Update remove IHS codes - check needed
            if (CheckIhsFormation())   // Needed?
            {
                for (int i = 0; i < _incidIhsFormationRows.Length; i++)
                {
                    if (_incidIhsFormationRows[i].RowState != DataRowState.Detached)
                        _incidIhsFormationRows[i].Delete();
                    _incidIhsFormationRows[i] = null;
                }
            }
        }

        #endregion

        #region IHS Management

        /// <summary>
        /// Gets the ihs management group header.
        /// </summary>
        /// <value>
        /// The ihs management group header.
        /// </value>
        public string IhsManagementHeader
        {
            get
            {
                // Optionally hide group headers to reduce window height.
                if ((bool)_showGroupHeaders)
                    return "IHS Management";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the first management code for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs management value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the first combined management code and test for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs management text.
        /// </value>
        public string IncidIhsManagement1Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsManagementCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_management.IsInitialized && (HluDataset.lut_ihs_management.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_managementTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_managementTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_managementDataTable) }, false);
                    }

                    _lutIhsManagementCodes = from x in HluDataset.lut_ihs_management
                                         //where x.is_local
                                         select x;
                }

                // Return the combined code and text value.
                if (IncidIhsManagement1 == null)
                    return null;
                else
                {
                    var q = _lutIhsManagementCodes.Where(m => m.code == _incidIhsManagementRows[0].management);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets the second management code for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs management value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the second combined management code and test for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs management text.
        /// </value>
        public string IncidIhsManagement2Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsManagementCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_management.IsInitialized && (HluDataset.lut_ihs_management.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_managementTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_managementTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_managementDataTable) }, false);
                    }

                    _lutIhsManagementCodes = from x in HluDataset.lut_ihs_management
                                             //where x.is_local
                                             select x;
                }

                // Return the combined code and text value.
                if (IncidIhsManagement2 == null)
                    return null;
                else
                {
                    var q = _lutIhsManagementCodes.Where(m => m.code == _incidIhsManagementRows[1].management);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        //TODO: Check IHS management - Needed?
        /// <summary>
        /// Checks if there are any valid ihs management rows.
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Removes the incid ihs management rows.
        /// </summary>
        public void RemoveIncidIhsManagementRows()
        {
            //TODO: Update remove IHS codes - check needed
            if (CheckIhsManagement())   // Needed?
            {
                for (int i = 0; i < _incidIhsManagementRows.Length; i++)
                {
                    if (_incidIhsManagementRows[i].RowState != DataRowState.Detached)
                        _incidIhsManagementRows[i].Delete();
                    _incidIhsManagementRows[i] = null;
                }
            }
        }

        #endregion

        #region IHS Complex

        /// <summary>
        /// Gets the ihs complex group header.
        /// </summary>
        /// <value>
        /// The ihs complex group header.
        /// </value>
        public string IhsComplexHeader
        {
            get
            {
                // Optionally hide group headers to reduce window height.
                if ((bool)_showGroupHeaders)
                    return "IHS Complex";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the first complex code for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs complex value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the first combined complex code and test for this incid.
        /// </summary>
        /// <value>
        /// The first incid ihs complex text.
        /// </value>
        public string IncidIhsComplex1Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsComplexCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_complex.IsInitialized && (HluDataset.lut_ihs_complex.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_complexTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_complexTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_complexDataTable) }, false);
                    }

                    _lutIhsComplexCodes = from x in HluDataset.lut_ihs_complex
                                         //where x.is_local
                                         select x;
                }

                // Return the combined code and text value.
                if (IncidIhsComplex1 == null)
                    return null;
                else
                {
                    var q = _lutIhsComplexCodes.Where(m => m.code == _incidIhsComplexRows[0].complex);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        /// <summary>
        /// Gets the second complex code for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs complex value.
        /// </value>
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
        }

        /// <summary>
        /// Gets the second combined complex code and test for this incid.
        /// </summary>
        /// <value>
        /// The second incid ihs complex text.
        /// </value>
        public string IncidIhsComplex2Text
        {
            get
            {
                // Get the value from the lookup table.
                if (_lutIhsComplexCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_ihs_complex.IsInitialized && (HluDataset.lut_ihs_complex.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_ihs_complexTableAdapter == null)
                            _hluTableAdapterMgr.lut_ihs_complexTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_ihs_complexDataTable) }, false);
                    }

                    _lutIhsComplexCodes = from x in HluDataset.lut_ihs_complex
                                          //where x.is_local
                                          select x;
                }

                // Return the combined code and text value.
                if (IncidIhsComplex2 == null)
                    return null;
                else
                {
                    var q = _lutIhsComplexCodes.Where(m => m.code == _incidIhsComplexRows[1].complex);
                    if (q.Count() > 0)
                        return String.Concat(q.ElementAt(0).code, " : ", q.ElementAt(0).description);
                    else
                        return null;
                }
            }
        }

        //TODO: Check IHS complex - Needed?
        /// <summary>
        /// Checks if there are any valid ihs complex rows.
        /// </summary>
        /// <returns></returns>
        private bool CheckIhsComplex()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidIhsComplexRows == null)
            {
                HluDataSet.incid_ihs_complexDataTable ihsComplexTable = _hluDS.incid_ihs_complex;
                GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_ihs_complexTableAdapter, ref ihsComplexTable);
            }

            return _incidIhsManagementRows != null;
        }

        /// <summary>
        /// Removes the incid ihs complex rows.
        /// </summary>
        public void RemoveIncidIhsComplexRows()
        {
            //TODO: Update remove IHS codes - check needed
            if (CheckIhsComplex())   // Needed?
            {
                for (int i = 0; i < _incidIhsComplexRows.Length; i++)
                {
                    if (_incidIhsComplexRows[i].RowState != DataRowState.Detached)
                        _incidIhsComplexRows[i].Delete();
                    _incidIhsComplexRows[i] = null;
                }
            }
        }

        #endregion

        #region IHS Summary

        /// <summary>
        /// Gets the ihs summary group header.
        /// </summary>
        /// <value>
        /// The ihs summary group header.
        /// </value>
        public string IhsSummaryHeader
        {
            get
            {
                // Optionally hide group headers to reduce window height.
                if ((bool)_showGroupHeaders)
                    return "IHS Summary";
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the concatenated incid ihs summary.
        /// </summary>
        /// <value>
        /// The concatenated incid ihs summary.
        /// </value>
        public string IncidIhsSummary
        {
            get
            {
                return ViewModelWindowMainHelpers.IhsSummary(new string[] { 
                    IncidIhsHabitat,
                    IncidIhsMatrix1,
                    IncidIhsMatrix2,
                    IncidIhsMatrix3,
                    IncidIhsFormation1,
                    IncidIhsFormation2,
                    IncidIhsManagement1,
                    IncidIhsManagement2,
                    IncidIhsComplex1,
                    IncidIhsComplex2 });
            }
        }

        #endregion

        #endregion

        #region Priority Tab

        /// <summary>
        /// Gets the Priority tab label.
        /// </summary>
        /// <value>
        /// The Priority tab label.
        /// </value>
        public string PriorityTabLabel
        {
            get { return "Priority"; }
        }
        //---------------------------------------------------------------------

        #region Priority Habitat
        
        /// <summary>
        /// Gets the array of all bap habitat codes.
        /// </summary>
        /// <value>
        /// The array of all bap habitat codes.
        /// </value>
        public HluDataSet.lut_habitat_typeRow[] BapHabitatCodes
        {
            get
            {
                // Get the value from the lookup table.
                if (_bapHabitatCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_habitat_type.IsInitialized && (HluDataset.lut_habitat_type.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    //---------------------------------------------------------------------
                    // CHANGED: CR52 Enable support for multiple priority habitat classifications
                    // Enable multiple priority habitat types (from the same or different
                    // classifications) to be assigned
                    //
                    //_bapHabitatCodes = (from r in HluDataset.lut_habitat_type
                    //                    where r.habitat_class_code == "PHAP"
                    //                    select r).ToArray();
                    _bapHabitatCodes = (from r in HluDataset.lut_habitat_type
                                        where r.bap_priority == true && r.is_local
                                        select r).ToArray();
                    //---------------------------------------------------------------------
                }

                return _bapHabitatCodes;
            }
        }
        
        /// <summary>
        /// Gets the array of all determination quality codes.
        /// </summary>
        /// <value>
        /// The array of all determination quality codes.
        /// </value>
        public HluDataSet.lut_quality_determinationRow[] DeterminationQualityCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_qualityDeterminationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_quality_determination.IsInitialized &&
                        HluDataset.lut_quality_determination.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_quality_determinationTableAdapter == null)
                            _hluTableAdapterMgr.lut_quality_determinationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_quality_determinationDataTable,
                                    HluDataSet.lut_quality_determinationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_quality_determinationDataTable) }, false);
                    }

                    _qualityDeterminationCodes =
                        HluDataset.lut_quality_determination.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _qualityDeterminationCodes;
            }
        }
        
        /// <summary>
        /// Gets the array of bap determination quality codes valid for automatically
        /// assigned priority habitats.
        /// </summary>
        /// <value>
        /// The array of automatic bap determination quality codes.
        /// </value>
        public HluDataSet.lut_quality_determinationRow[] BapDeterminationQualityCodesAuto
        {
            get
            {
                if (DeterminationQualityCodes != null)
                    return DeterminationQualityCodes.Where(r => r.code != BapEnvironment.BAPDetQltyUserAdded
                        && r.code != BapEnvironment.BAPDetQltyPrevious)
                        .OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                else
                    return null;
            }
        }

        /// <summary>
        /// Gets the array of bap determination quality codes valid for user
        /// assigned priority habitats.
        /// </summary>
        /// <value>
        /// The array of user assigned bap determination quality codes.
        /// </value>
        public HluDataSet.lut_quality_determinationRow[] BapDeterminationQualityCodesUser
        {
            get
            {
                // Show all determination quality values in the drop-down list (instead
                // of just 'Not present but close to definition') but validate the
                // selected value later.
                return DeterminationQualityCodes;
            }
        }
        
        /// <summary>
        /// Gets the  array of all interpretation quality codes.
        /// </summary>
        /// <value>
        /// The  array of all interpretation quality codes.
        /// </value>
        public HluDataSet.lut_quality_interpretationRow[] InterpretationQualityCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_qualityInterpretationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_quality_interpretation.IsInitialized &&
                        HluDataset.lut_quality_interpretation.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_quality_interpretationTableAdapter == null)
                            _hluTableAdapterMgr.lut_quality_interpretationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_quality_interpretationDataTable,
                                    HluDataSet.lut_quality_interpretationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_quality_interpretationDataTable) }, false);
                    }

                    //TODO: Add is_local flag
                    _qualityInterpretationCodes =
                        HluDataset.lut_quality_interpretation.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _qualityInterpretationCodes;
            }
        }
        
        /// <summary>
        /// Gets the  array of all bap related interpretation quality codes.
        /// </summary>
        /// <value>
        /// The array of all bap related interpretation quality codes.
        /// </value>
        public HluDataSet.lut_quality_interpretationRow[] BapInterpretationQualityCodes
        {
            get
            {
                return InterpretationQualityCodes;
            }
        }
        
        /// <summary>
        /// Gets or sets the collection of incid bap habitats automatically assigned.
        /// </summary>
        /// <value>
        /// The collection of incid bap habitats automatically assigned.
        /// </value>
        public ObservableCollection<BapEnvironment> IncidBapHabitatsAuto
        {
            get { return _incidBapRowsAuto; }
            set
            {
                _incidBapRowsAuto = value;
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        /// <summary>
        /// Gets or sets the collection of incid bap habitats assigned by the user.
        /// The bap_id of existing secondary priority habitats is multiplied by -1 (and same again when 
        /// saving back to DB) to distinguish them from primary priority habitats in UI validation methods.
        /// </summary>
        /// <value>
        /// The collection of incid bap habitats assigned by the user.
        /// </value>
        public ObservableCollection<BapEnvironment> IncidBapHabitatsUser
        {
            get { return _incidBapRowsUser; }
            set
            {
                _incidBapRowsUser = value;
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
        /// Gets all of the automatically assigned and user assigned bap environments.
        /// </summary>
        public void GetBapEnvironments()
        {
            // Remove any existing handlers before assigning a new collection.
            if (_incidBapRowsAuto != null)
                _incidBapRowsAuto.CollectionChanged -= _incidBapRowsAuto_CollectionChanged;
            if (_incidBapRowsUser != null)
                _incidBapRowsUser.CollectionChanged -= _incidBapRowsUser_CollectionChanged;

            IEnumerable<string> mandatoryBap = null;
            IEnumerable<HluDataSet.incid_bapRow> incidBapRowsUndel = null;
            if (IncidPrimary != null)
            {
                // Identify which primary BAP rows there should be from the
                // primary and secondary codes.
                //TODO: Check if mandatory priority habitats are set from primary and secondary
                mandatoryBap = MandatoryBapEnvironments(IncidPrimary, IncidSecondaryHabitats);

                // Identify any BAP rows (both auto generated and user added) that
                // have not been marked as deleted.
                incidBapRowsUndel = _incidBapRows.Where(r => r.RowState != DataRowState.Deleted);
            }

            // If there are any undeleted rows and the IHS codes indicate
            // that there should be some primary BAP (auto) rows then sort out
            // which of the undeleted rows are the auto rows.
            if ((incidBapRowsUndel != null) && (mandatoryBap != null))
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
                    // Which of the undeleted rows are auto rows that
                    // already existed.
                    prevBapRowsAuto = from r in incidBapRowsUndel
                                      join pot in mandatoryBap on r.bap_habitat equals pot
                                      where _incidCurrentRow.incid != null && r.incid == _incidCurrentRow.incid
                                      select new BapEnvironment(false, false, r);

                    // Which of the undeleted rows were previously user
                    // added rows but should now be promoted to auto
                    // rows as a result of changes to the IHS codes.
                    newBapRowsAuto = from r in incidBapRowsUndel
                                     join pot in mandatoryBap on r.bap_habitat equals pot
                                     where prevBapRowsAuto.Count(p => p.bap_habitat == r.bap_habitat) == 0
                                     select new BapEnvironment(false, false, r);
                }

                // Determine if there are any potential BAP rows that should
                // be added as a result of changes to the IHS codes.
                var potBap = from p in mandatoryBap
                             where prevBapRowsAuto.Count(a => a.bap_habitat == p) == 0
                             where incidBapRowsUndel.Count(row => row.bap_habitat == p) == 0
                             select new BapEnvironment(false, false, -1, Incid, p, null, null, null);

                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // Concatenate the previous auto rows, the newly promoted auto
                // rows and the potential BAP rows.
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>(
                    prevBapRowsAuto.Concat(newBapRowsAuto).Concat(potBap));
            }
            else if (incidBapRowsUndel != null)
            {
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // As there should be no primary BAP rows according to the
                // IHS codes then the auto rows should be blank (because any
                // undeleted rows must therefore now be considered as user rows.
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            }
            else if ((mandatoryBap != null) && (mandatoryBap.Count() > 0))
            {
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // If there should be some primary BAP rows according to the
                // IHS codes, but there are no existing undeleted rows, then
                // all the primrary BAP codes must become new auto rows.
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>(
                    mandatoryBap.Select(p => new BapEnvironment(false, false, -1, Incid, p, null, null, null)));
            }
            else
            {
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsAuto != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsAuto)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // There shouldn't be any primary BAP rows according to the IHS
                // codes, and there are no existing undeleted rows, so there are
                // no auto rows.
                _incidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            }

            // Track any changes to the auto rows collection.
            _incidBapRowsAuto.CollectionChanged += _incidBapRowsAuto_CollectionChanged;

            // Track when the auto data has been changed so that the apply button
            // will appear.
            foreach (BapEnvironment be in _incidBapRowsAuto)
            {
                be.DataChanged += _incidBapRowsAuto_DataChanged;
            };

            // Check if there are any errors in the auto BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
            {
                int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapAuto");
                else
                    DelErrorList(ref _priorityErrors, "BapAuto");
            }
            else
                DelErrorList(ref _priorityErrors, "BapAuto");

            OnPropertyChanged("IncidBapHabitatsAuto");

            // If there are undeleted rows and there are some auto rows
            // then sort them out to determine which of the undeleted rows
            // are considered as user added.
            if ((incidBapRowsUndel != null) && (_incidBapRowsAuto != null))
            {
                List<BapEnvironment> prevBapRowsUser = null;
                // If there were no user added rows before then there
                // are no previous user added rows.
                if (_incidBapRowsUser == null)
                {
                    prevBapRowsUser = new List<BapEnvironment>();
                }
                else
                {
                    // If there were user added rows before then determine
                    // which of them have not been promoted to auto rows.
                    prevBapRowsUser = (from r in _incidBapRowsUser
                                       where _incidCurrentRow.incid != null && r.incid == _incidCurrentRow.incid
                                       where _incidBapRowsAuto.Count(row => row.bap_habitat == r.bap_habitat) == 0
                                       select r).ToList();
                    prevBapRowsUser.ForEach(delegate(BapEnvironment be)
                    {
                        // Don't overwrite the determination quality value loaded from the
                        // database with 'Not present but close to definition' as other
                        // values may be valid and will be validated later.
                        //
                        //be.quality_determination = BapEnvironment.BAPDetQltyUserAdded;
                        be.BulkUpdateMode = _bulkUpdateMode == true;
                    });
                }

                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // Concatenate the previous user added rows with any remaining
                // undeleted rows that are not auto rows.
                _incidBapRowsUser = new ObservableCollection<BapEnvironment>(prevBapRowsUser.Concat(
                    from r in incidBapRowsUndel
                    where _incidBapRowsAuto.Count(a => a.bap_habitat == r.bap_habitat) == 0
                    where prevBapRowsUser.Count(p => p.bap_habitat == r.bap_habitat) == 0
                    select new BapEnvironment(_bulkUpdateMode == true, true, r)));
            }
            // If thereare undeleted rows but no auto rows then all the
            // undeleted rows must be considered user added rows.
            else if (incidBapRowsUndel != null)
            {
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                _incidBapRowsUser = new ObservableCollection<BapEnvironment>(
                   incidBapRowsUndel.Select(r => new BapEnvironment(_bulkUpdateMode == true, true, r)));
            }
            else
            {
                // Remove any existing handlers before assigning a new collection.
                if (_incidBapRowsUser != null)
                {
                    foreach (BapEnvironment be in _incidBapRowsUser)
                    {
                        be.DataChanged -= _incidBapRowsUser_DataChanged;
                    }
                }

                // Otherwise there can't be any user added rows.
                _incidBapRowsUser = new ObservableCollection<BapEnvironment>();
            }

            // Track any changes to the user rows collection.
            _incidBapRowsUser.CollectionChanged += _incidBapRowsUser_CollectionChanged;

            // Track when the user data has been changed so that the apply button
            // will appear.
            foreach (BapEnvironment be in _incidBapRowsUser)
            {
                be.DataChanged += _incidBapRowsUser_DataChanged;
            };

            // Check if there are any errors in the user BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
            {
                int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapUser");
                else
                    DelErrorList(ref _priorityErrors, "BapUser");

                // Check if there are any duplicates between the primary and 
                // secondary BAP records.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                                            group be by be.bap_habitat into g
                                            where g.Count() > 1
                                            select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                    if (beDups.Length > 2)
                        AddErrorList(ref _priorityErrors, "BapUserDup");
                    else
                        DelErrorList(ref _priorityErrors, "BapUserDup");
                }
            }
            else
                DelErrorList(ref _priorityErrors, "BapUser");

            OnPropertyChanged("IncidBapHabitatsUser");

            // Concatenate the auto rows and the user rows to become the new list
            // of BAP rows.
            BapEnvironment.BapEnvironmentList = _incidBapRowsAuto.Concat(_incidBapRowsUser);

            OnPropertyChanged("PriorityTabLabel");

        }

        /// <summary>
        /// Track when the BAP primary records have changed so that the apply
        /// button will appear.        
        /// </summary>
        private void _incidBapRowsAuto_DataChanged(bool BapChanged)
        {
            Changed = true;

            // Check if there are any errors in the primary BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
            {
                int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapAuto");
                else
                    DelErrorList(ref _priorityErrors, "BapAuto");
            }
            OnPropertyChanged("PriorityTabLabel");
        }

        /// <summary>
        /// Track when the BAP secondary records have changed so that the apply
        /// button will appear.
        /// </summary>
        private void _incidBapRowsUser_DataChanged(bool BapChanged)
        {
            Changed = true;

            // Check if there are any errors in the secondary BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
            {
                int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapUser");
                else
                    DelErrorList(ref _priorityErrors, "BapUser");

                // Check if there are any duplicates between the primary and 
                // secondary BAP records.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                                            group be by be.bap_habitat into g
                                            where g.Count() > 1
                                            select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                    if (beDups.Length > 2)
                        AddErrorList(ref _priorityErrors, "BapUserDup");
                    else
                        DelErrorList(ref _priorityErrors, "BapUserDup");
                }
            }
            OnPropertyChanged("PriorityTabLabel");
        }

        /// <summary>
        /// Build a enumerable of the mandatory bap habitats
        /// based on the primary habitat and all the secondary habitats.
        /// </summary>
        /// <param name="primaryHabitat">The primary habitat.</param>
        /// <param name="secondaryHabitats">The secondary habitats.</param>
        /// <returns></returns>
        internal IEnumerable<string> MandatoryBapEnvironments(string primaryHabitat, ObservableCollection<SecondaryHabitat> secondaryHabitats)
        {
            IEnumerable<string> primaryBap = null;
            IEnumerable<string> secondaryBap = null;
            string[] q = null;

            // Get the BAP habitats associated with the primary habitat
            if (!String.IsNullOrEmpty(primaryHabitat))
            {
                try
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR52 Enable support for multiple priority habitat classifications
                    // Enable multiple priority habitat types (from the same or different
                    // classifications) to be assigned
                    //
                    //primaryBap = from r in HluDataset.lut_ihs_habitat
                    //             where r.code == ihsHabitat && !r.Iscode_bap_priority_habitatNull()
                    //             select r.code_bap_priority_habitat;

                    //q = (from r in HluDataset.lut_primary
                    //     join b in HluDataset.lut_primary_bap_habitat on r.code equals b.code_primary
                    //     where r.code == primaryHabitat
                    //     select b.bap_habitat).ToArray();
                    q = (from b in HluDataset.lut_primary_bap_habitat
                         where b.code_primary == primaryHabitat
                         select b.bap_habitat).ToArray();

                    // If any primary bap habitats have been found
                    primaryBap = null;
                    if ((q != null) && (q.Count() > 0))
                        primaryBap = q;
                    //---------------------------------------------------------------------
                }
                catch { }
            }

            //TODO: Secondary BAP habitats - check
            // Get the BAP habitats associated with the secondary habitats
            if (secondaryHabitats != null)
            {
                try
                {
                    q = (from b in HluDataset.lut_secondary_bap_habitat
                         join s in secondaryHabitats on b.code_secondary equals s.secondary_habitat
                         select b.bap_habitat).ToArray();

                    // If any secondary bap habitats have been found
                    secondaryBap = null;
                    if ((q != null) && (q.Count() > 0))
                        secondaryBap = q;
                }
                catch { }
            }

            IEnumerable<string> allBap = null;
            allBap = primaryBap != null ? secondaryBap != null ? primaryBap.Concat(secondaryBap) : primaryBap : secondaryBap;
            if (allBap != null)
                return allBap.Distinct();
            else
                return new string[0];
        }

        private void _incidBapRowsAuto_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged("Error");

            // Track when the BAP primary records have changed so that the apply
            // button will appear.
            Changed = true;

            // Check if there are any errors in the primary BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
            {
                int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapAuto");
                else
                    DelErrorList(ref _priorityErrors, "BapAuto");
            }
            OnPropertyChanged("PriorityTabLabel");
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

            // Track when the BAP secondary records have changed so that the apply
            // button will appear.
            Changed = true;

            // Check if there are any errors in the secondary BAP records to see
            // if the Priority tab label should be flagged as also in error.
            if (_incidBapRowsUser != null && _incidBapRowsUser.Count > 0)
            {
                int countInvalid = _incidBapRowsUser.Count(be => !be.IsValid());
                if (countInvalid > 0)
                    AddErrorList(ref _priorityErrors, "BapUser");
                else
                    DelErrorList(ref _priorityErrors, "BapUser");

                // Check if there are any duplicates between the primary and 
                // secondary BAP records.
                if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                {
                    StringBuilder beDups = (from be in _incidBapRowsAuto.Concat(_incidBapRowsUser)
                                            group be by be.bap_habitat into g
                                            where g.Count() > 1
                                            select g.Key).Aggregate(new StringBuilder(), (sb, code) => sb.Append(", " + code));
                    if (beDups.Length > 2)
                        AddErrorList(ref _priorityErrors, "BapUserDup");
                    else
                        DelErrorList(ref _priorityErrors, "BapUserDup");
                }
            }
            else
            {
                DelErrorList(ref _priorityErrors, "BapUser");
            }

            OnPropertyChanged("PriorityTabLabel");

            foreach (BapEnvironment be in _incidBapRowsUser)
            {
                if (be == null)
                    be.DataChanged -= _incidBapRowsUser_DataChanged;
                else if (be.bap_id == -1)
                    be.DataChanged += _incidBapRowsUser_DataChanged;
            }
        }
        
        #endregion

        #endregion

        #region Details Tab

        // Set the Details tab label from here so that validation can be done.
        // This will enable tooltips to be shown so that validation errors
        // in any fields in the tab can be highlighted by flagging the tab
        // label as in error.
        public string DetailsTabLabel
        {
            get { return "Details"; }
        }

        #region General Comments

        public string DetailsCommentsHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "General Comments";
                else
                    return null;
            }
        }

        public string IncidGeneralComments
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.general_commentsColumn))
                    return IncidCurrentRow.general_comments;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.general_comments = value;
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }

        #endregion

        #region Maps

        public string DetailsMapsHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Maps";
                else
                    return null;
            }
        }
        
        /// <summary>
        /// Gets the list of boundary map codes.
        /// </summary>
        /// <value>
        /// The list of boundary map codes.
        /// </value>
        public DataView BoundaryMapCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (HluDataset.lut_boundary_map.IsInitialized && HluDataset.lut_boundary_map.Count == 0)
                {
                    // Load the lookup table if not already loaded.
                    if (_hluTableAdapterMgr.lut_boundary_mapTableAdapter == null)
                        _hluTableAdapterMgr.lut_boundary_mapTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_boundary_mapDataTable, HluDataSet.lut_boundary_mapRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_boundary_mapDataTable) }, false);

                    string sortCols = String.Concat(HluDataset.lut_boundary_map.sort_orderColumn.ColumnName, ", ", HluDataset.lut_boundary_map.descriptionColumn.ColumnName);
                    HluDataset.lut_boundary_map.DefaultView.Sort = sortCols;
                }

                return HluDataset.lut_boundary_map.DefaultView;
            }
        }
        
        /// <summary>
        /// Gets or sets the incid boundary base map.
        /// </summary>
        /// <value>
        /// The incid boundary base map.
        /// </value>
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
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }
        
        /// <summary>
        /// Gets or sets the incid digitisation base map.
        /// </summary>
        /// <value>
        /// The incid digitisation base map.
        /// </value>
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
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }

        #endregion

        #region Site

        public string DetailsSiteHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Site";
                else
                    return null;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR37 (Site reference and site name)
        // Display the site reference with the site name in the interface.
        public string IncidSiteRef
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.Issite_refNull())
                    return IncidCurrentRow.site_ref;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.site_ref = value;
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }
        //---------------------------------------------------------------------

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
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }

        #endregion

        #region Condition
        
        /// <summary>
        /// Gets the details condition group header.
        /// </summary>
        /// <value>
        /// The details condition group header.
        /// </value>
        public string DetailsConditionHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Condition";
                else
                    return null;
            }
        }
        
        /// <summary>
        /// Check if there are any valid condition rows.
        /// </summary>
        /// <returns></returns>
        private bool CheckCondition()
        {
            if (_bulkUpdateMode == true) return true;

            if (_incidConditionRows == null)
            {
                HluDataSet.incid_conditionDataTable incidConditionTable = _hluDS.incid_condition;
                _incidConditionRows = GetIncidChildRowsDb(new object[] { Incid },
                    _hluTableAdapterMgr.incid_conditionTableAdapter, ref incidConditionTable);
            }

            return _incidConditionRows != null;
        }
        
        /// <summary>
        /// Gets the list of condition codes.
        /// </summary>
        /// <value>
        /// The list of condition codes.
        /// </value>
        public HluDataSet.lut_conditionRow[] ConditionCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_conditionCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_condition.IsInitialized && (HluDataset.lut_condition.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_conditionTableAdapter == null)
                            _hluTableAdapterMgr.lut_conditionTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_conditionDataTable, HluDataSet.lut_conditionRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_conditionDataTable) }, false);
                    }

                    // Load all local condition codes.
                    _conditionCodes = HluDataset.lut_condition.Where(r => r.is_local).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                // Return the list of condition codes, with the clear row if applicable.
                if (_incidConditionRows.Length >= 1 &&
                    _incidConditionRows[0] != null &&
                    !String.IsNullOrEmpty(_incidConditionRows[0].condition))
                {
                    HluDataSet.lut_conditionRow clearRow = HluDataset.lut_condition.Newlut_conditionRow();
                    clearRow.code = "";
                    clearRow.description = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return _conditionCodes.Concat(
                        new HluDataSet.lut_conditionRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }
                else
                {
                    return _conditionCodes;
                }
            }
        }

        /// <summary>
        /// Gets or sets the incid condition.
        /// </summary>
        /// <value>
        /// The incid condition.
        /// </value>
        public string IncidCondition
        {
            get
            {
                if (!CheckCondition()) return null;

                if (_incidConditionRows.Length < 1)
                {
                    _incidConditionRows = new HluDataSet.incid_conditionRow[1];
                }

                if ((_incidConditionRows[0] != null) &&
                    !_incidConditionRows[0].IsNull(HluDataset.incid_condition.conditionColumn))
                    return _incidConditionRows[0].condition;
                else
                    return null;
            }
            set
            {
                //TODO: Condition set - Check works, especially in bulk update mode
                if (IncidCurrentRow != null)
                {
                    bool clearCode = value == "";
                    bool newCode = false;
                    if (clearCode)
                        value = null;
                    else
                        newCode = ((_incidConditionRows.Length < 1 ||
                            _incidConditionRows[0] == null ||
                            String.IsNullOrEmpty(_incidConditionRows[0].condition)) && (!String.IsNullOrEmpty(value)));

                    UpdateIncidConditionRow(0, IncidConditionTable.conditionColumn.Ordinal, value);

                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;

                    // If the code has been cleared or it was previously clear and
                    // has been set then refresh the condition codes list and other
                    // condition fields
                    if (clearCode || newCode)
                    {
                        OnPropertyChanged("ConditionCodes");
                        OnPropertyChanged("IncidConditionQualifier");
                        OnPropertyChanged("IncidConditionDate");
                        OnPropertyChanged("IncidConditionEnabled");
                    }
                }
            }
        }
        
        /// <summary>
        /// Gets the list of condition qualifier codes.
        /// </summary>
        /// <value>
        /// The list of condition qualifier codes.
        /// </value>
        public HluDataSet.lut_condition_qualifierRow[] ConditionQualifierCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_conditionQualifierCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_condition_qualifier.IsInitialized && (HluDataset.lut_condition_qualifier.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_condition_qualifierTableAdapter == null)
                            _hluTableAdapterMgr.lut_condition_qualifierTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_condition_qualifierDataTable, HluDataSet.lut_condition_qualifierRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_condition_qualifierDataTable) }, false);
                    }

                    // Load all local condition qualifier codes.
                    _conditionQualifierCodes = HluDataset.lut_condition_qualifier.Where(r => r.is_local).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _conditionQualifierCodes;
            }
        }
        
        /// <summary>
        /// Gets or sets the incid condition qualifier.
        /// </summary>
        /// <value>
        /// The incid condition qualifier.
        /// </value>
        public string IncidConditionQualifier
        {
            get
            {
                if (!CheckCondition()) return null;
                if ((_incidConditionRows.Length > 0) && (_incidConditionRows[0] != null) &&
                    !_incidConditionRows[0].IsNull(HluDataset.incid_condition.condition_qualifierColumn))
                    return _incidConditionRows[0].condition_qualifier;
                else
                    return null;
            }
            set
            {
                //TODO: Condition set - Check works
                UpdateIncidConditionRow(0, IncidConditionTable.condition_qualifierColumn.Ordinal, value);
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }
        
        /// <summary>
        /// Gets or sets the incid condition date.
        /// </summary>
        /// <value>
        /// The incid condition date.
        /// </value>
        public Date.VagueDateInstance IncidConditionDate
        {
            get
            {
                if (!CheckCondition()) return null;
                if ((_incidConditionRows.Length > 0) && (_incidConditionRows[0] != null) &&
                    !_incidConditionRows[0].IsNull(HluDataset.incid_condition.condition_date_startColumn) &&
                    !_incidConditionRows[0].IsNull(HluDataset.incid_condition.condition_date_endColumn) &&
                    !_incidConditionRows[0].IsNull(HluDataset.incid_condition.condition_date_typeColumn))
                {
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(_incidConditionRows[0].condition_date_start,
                        _incidConditionRows[0].condition_date_end, _incidConditionRows[0].condition_date_type,
                        _incidConditionDateEntered != null ? _incidConditionDateEntered.UserEntry : null);
                    return vd;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                //TODO: Condition set - Check works
                UpdateIncidConditionRow(0, IncidConditionTable.condition_date_startColumn.Ordinal, value);
                _incidConditionDateEntered = value;
                OnPropertyChanged("IncidConditionDate");
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }
        
        /// <summary>
        /// Updates the incid condition row.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="columnOrdinal">The column ordinal.</param>
        /// <param name="newValue">The new value.</param>
        private void UpdateIncidConditionRow<T>(int rowNumber, int columnOrdinal, T newValue)
        {
            //TODO: Condition set - Check works
            try
            {
                if (_incidConditionRows == null) return;

                // If the row is blank
                if (_incidConditionRows[rowNumber] == null)
                {
                    if ((columnOrdinal == HluDataset.incid_condition.conditionColumn.Ordinal) && newValue != null)
                    {
                        // Set the row id
                        HluDataSet.incid_conditionRow newRow = IncidConditionTable.Newincid_conditionRow();
                        newRow.incid_condition_id = NextIncidConditionId;
                        newRow.incid = IncidCurrentRow.incid;
                        //newRow.sort_order = rowNumber + 1;    //TODO: Not needed but left if for now just in case
                        _incidConditionRows[rowNumber] = newRow;
                    }
                    else
                    {
                        return;
                    }
                }
                // If the new condition is null
                else if ((columnOrdinal == HluDataset.incid_condition.conditionColumn.Ordinal) && (newValue == null))
                {
                    // Delete the row unless during a bulk update
                    if (_bulkUpdateMode == false)
                    {
                        if (_incidConditionRows[rowNumber].RowState != DataRowState.Detached)
                            _incidConditionRows[rowNumber].Delete();
                        _incidConditionRows[rowNumber] = null;
                    }
                    else
                    {
                        _incidConditionRows[rowNumber] = IncidConditionTable.Newincid_conditionRow();
                        IncidConditionRows[rowNumber].incid_condition_id = rowNumber;
                        IncidConditionRows[rowNumber].condition = null;
                        IncidConditionRows[rowNumber].incid = RecIDs.CurrentIncid;
                    }
                    return;
                }

                // Update the date columns
                if ((columnOrdinal == HluDataset.incid_condition.condition_date_startColumn.Ordinal) ||
                    (columnOrdinal == HluDataset.incid_condition.condition_date_endColumn.Ordinal))
                {
                    Date.VagueDateInstance vd = newValue as Date.VagueDateInstance;
                    if (vd != null)
                    {
                        _incidConditionRows[rowNumber].condition_date_start = vd.StartDate;
                        _incidConditionRows[rowNumber].condition_date_end = vd.EndDate;
                        _incidConditionRows[rowNumber].condition_date_type = vd.DateType;
                    }
                    else
                    {
                        _incidConditionRows[rowNumber].condition_date_start = VagueDate.DateUnknown;
                        _incidConditionRows[rowNumber].condition_date_end = VagueDate.DateUnknown;
                        _incidConditionRows[rowNumber].condition_date_type = null;
                    }
                }
                // Update all other columns if they have changed
                else if ((((_incidConditionRows[rowNumber].IsNull(columnOrdinal) ^ (newValue == null)) ||
                    ((!_incidConditionRows[rowNumber].IsNull(columnOrdinal) && (newValue != null)))) &&
                    !_incidConditionRows[rowNumber][columnOrdinal].Equals(newValue)))
                {
                    _incidConditionRows[rowNumber][columnOrdinal] = newValue;
                }

                if ((_incidConditionRows[rowNumber].RowState == DataRowState.Detached) &&
                    IsCompleteRow(_incidConditionRows[rowNumber]))
                {
                    IncidConditionTable.Addincid_conditionRow(_incidConditionRows[rowNumber]);
                }
            }
            catch { }
        }

        //
        public bool IncidConditionEnabled
        {
            // Disable remaining condition fields when condition is blank
            get { return (IncidCondition != null); }
        }

        #endregion

        #region Quality
        
        /// <summary>
        /// Gets the details quality group header.
        /// </summary>
        /// <value>
        /// The details quality group header.
        /// </value>
        public string DetailsQualityHeader
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Quality";
                else
                    return null;
            }
        }
        
        /// <summary>
        /// Gets the list of quality determination codes.
        /// </summary>
        /// <value>
        /// The list of quality determination codes.
        /// </value>
        public HluDataSet.lut_quality_determinationRow[] QualityDeterminationCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_qualityDeterminationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_quality_determination.IsInitialized &&
                        HluDataset.lut_quality_determination.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_quality_determinationTableAdapter == null)
                            _hluTableAdapterMgr.lut_quality_determinationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_quality_determinationDataTable,
                                    HluDataSet.lut_quality_determinationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_quality_determinationDataTable) }, false);
                    }

                    _qualityDeterminationCodes =
                        HluDataset.lut_quality_determination.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                // Return the list of determination codes, with the clear row if applicable.
                if (!String.IsNullOrEmpty(IncidQualityDetermination))
                {
                    HluDataSet.lut_quality_determinationRow clearRow = HluDataset.lut_quality_determination.Newlut_quality_determinationRow();
                    clearRow.code = _codeDeleteRow;
                    clearRow.description = _codeDeleteRow;
                    clearRow.sort_order = -1;

                    return _qualityDeterminationCodes.Concat(
                        new HluDataSet.lut_quality_determinationRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }
                else
                {
                    return _qualityDeterminationCodes;
                }
            }
        }
        
        /// <summary>
        /// Gets or sets the incid quality determination.
        /// </summary>
        /// <value>
        /// The incid quality determination.
        /// </value>
        public string IncidQualityDetermination
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.quality_determinationColumn))
                    return IncidCurrentRow.quality_determination;
                else
                    return null;
            }
            set
            {
                if (IncidCurrentRow != null)
                {
                    bool clearCode = value == _codeDeleteRow;
                    bool newCode = false;
                    if (clearCode)
                        value = null;
                    else
                        newCode = ((IncidCurrentRow.IsNull(HluDataset.incid.quality_determinationColumn)) && (!String.IsNullOrEmpty(value)));

                    IncidCurrentRow.quality_determination = value;

                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;

                    // Refresh quality determnation list
                    if (clearCode || newCode)
                        OnPropertyChanged("QualityDeterminationCodes");
                }
            }
        }
        
        /// <summary>
        /// Gets the list of quality interpretation codes.
        /// </summary>
        /// <value>
        /// The list of quality interpretation codes.
        /// </value>
        public HluDataSet.lut_quality_interpretationRow[] QualityInterpretationCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_qualityInterpretationCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_quality_interpretation.IsInitialized &&
                        HluDataset.lut_quality_interpretation.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_quality_interpretationTableAdapter == null)
                            _hluTableAdapterMgr.lut_quality_interpretationTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_quality_interpretationDataTable,
                                    HluDataSet.lut_quality_interpretationRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_quality_interpretationDataTable) }, false);
                    }

                    //TODO: Add is_local flag???
                    _qualityInterpretationCodes =
                        HluDataset.lut_quality_interpretation.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                // Return the list of interpretation codes, with the clear row if applicable.
                if (!String.IsNullOrEmpty(IncidQualityInterpretation))
                {
                    HluDataSet.lut_quality_interpretationRow clearRow = HluDataset.lut_quality_interpretation.Newlut_quality_interpretationRow();
                    clearRow.code = _codeDeleteRow;
                    clearRow.description = _codeDeleteRow;
                    clearRow.sort_order = -1;

                    return _qualityInterpretationCodes.Concat(
                        new HluDataSet.lut_quality_interpretationRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }
                else
                {
                    return _qualityInterpretationCodes;
                }
            }
        }
        
        /// <summary>
        /// Gets or sets the incid quality interpretation.
        /// </summary>
        /// <value>
        /// The incid quality interpretation.
        /// </value>
        public string IncidQualityInterpretation
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.quality_interpretationColumn))
                    return IncidCurrentRow.quality_interpretation;
                else
                    return null;
            }
            set
            {
                if (IncidCurrentRow != null)
                {
                    bool clearCode = value == _codeDeleteRow;
                    bool newCode = false;
                    if (clearCode)
                        value = null;
                    else
                        newCode = ((IncidCurrentRow.IsNull(HluDataset.incid.quality_interpretationColumn)) && (!String.IsNullOrEmpty(value)));

                    IncidCurrentRow.quality_interpretation = value;

                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;

                    // Refresh quality determnation list
                    if (clearCode || newCode)
                        OnPropertyChanged("QualityInterpretationCodes");

                    // Revalidate the comments
                    OnPropertyChanged("IncidQualityComments");
                }
            }
        }
        
        /// <summary>
        /// Gets or sets the incid quality comments.
        /// </summary>
        /// <value>
        /// The incid quality comments.
        /// </value>
        public string IncidQualityComments
        {
            get
            {
                if ((IncidCurrentRow != null) && !IncidCurrentRow.IsNull(HluDataset.incid.interpretation_commentsColumn))
                    return IncidCurrentRow.interpretation_comments;
                else
                    return null;
            }
            set
            {
                if ((IncidCurrentRow != null) && (value != null))
                {
                    IncidCurrentRow.interpretation_comments = value;
                    // Flag that the current record has changed so that the apply button
                    // will appear.
                    Changed = true;
                }
            }
        }

        #endregion

        #endregion

        #region Sources Tab
        
        /// <summary>
        /// Gets the sources tab group label.
        /// </summary>
        /// <value>
        /// The sources tab group label.
        /// </value>
        public string SourcesTabLabel
        {
            get { return "Sources"; }
        }

        #region Sources
        
        /// <summary>
        /// Checks if there are any valid source rows.
        /// </summary>
        /// <returns></returns>
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
        
        /// <summary>
        /// Returns the default date for a given source.
        /// </summary>
        /// <param name="currentDate">The current date.</param>
        /// <param name="sourceID">The source identifier.</param>
        /// <returns></returns>
        public Date.VagueDateInstance DefaultSourceDate(Date.VagueDateInstance currentDate, Nullable<int> sourceID)
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
        
        /// <summary>
        /// Updates the incid sources row.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="columnOrdinal">The column ordinal.</param>
        /// <param name="newValue">The new value.</param>
        private void UpdateIncidSourcesRow<T>(int rowNumber, int columnOrdinal, T newValue)
        {
            try
            {
                if (_incidSourcesRows == null) return;

                // If the row is blank
                if (_incidSourcesRows[rowNumber] == null)
                {
                    if (columnOrdinal == HluDataset.incid_sources.source_idColumn.Ordinal)
                    {
                        // Set the row id
                        HluDataSet.incid_sourcesRow newRow = IncidSourcesTable.Newincid_sourcesRow();
                        newRow.incid_source_id = NextIncidSourcesId;
                        newRow.incid = IncidCurrentRow.incid;
                        newRow.sort_order = rowNumber + 1;
                        _incidSourcesRows[rowNumber] = newRow;
                    }
                    else
                    {
                        return;
                    }
                }
                // If the new source_id is null
                else if ((columnOrdinal == HluDataset.incid_sources.source_idColumn.Ordinal) && (newValue == null))
                {
                    // Delete the row unless during a bulk update
                    if (_bulkUpdateMode == false)
                    {
                        if (_incidSourcesRows[rowNumber].RowState != DataRowState.Detached)
                            _incidSourcesRows[rowNumber].Delete();
                        _incidSourcesRows[rowNumber] = null;
                    }
                    else
                    {
                        _incidSourcesRows[rowNumber] = IncidSourcesTable.Newincid_sourcesRow();
                        IncidSourcesRows[rowNumber].incid_source_id = rowNumber;
                        IncidSourcesRows[rowNumber].source_id = Int32.MinValue;
                        IncidSourcesRows[rowNumber].incid = RecIDs.CurrentIncid;
                    }
                    return;
                }

                // Update the date columns
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
                // Update all other columns if they have changed
                else if ((((_incidSourcesRows[rowNumber].IsNull(columnOrdinal) ^ (newValue == null)) ||
                    ((!_incidSourcesRows[rowNumber].IsNull(columnOrdinal) && (newValue != null)))) &&
                    !_incidSourcesRows[rowNumber][columnOrdinal].Equals(newValue)))
                {
                    _incidSourcesRows[rowNumber][columnOrdinal] = newValue;
                }

                // If updating the source_id get the default date
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
        
        /// <summary>
        /// Gets the list of source names.
        /// </summary>
        /// <value>
        /// The list of source names.
        /// </value>
        public HluDataSet.lut_sourcesRow[] SourceNames
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_sourceNames == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_sources.IsInitialized &&
                        HluDataset.lut_sources.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                            _hluTableAdapterMgr.lut_sourcesTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset,
                            new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                    }
                    _sourceNames = HluDataset.lut_sources.OrderBy(r => r.sort_order).ThenBy(r => r.source_name).ToArray();
                }

                return _sourceNames;
            }
        }
        
        /// <summary>
        /// Gets the list of source habitat class codes.
        /// </summary>
        /// <value>
        /// The list of source habitat class codes.
        /// </value>
        public HluDataSet.lut_habitat_classRow[] SourceHabitatClassCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_sourceHabitatClassCodes == null)
                {
                    // Load the kookup table is not already loaded.
                    if (HluDataset.lut_habitat_class.IsInitialized &&
                        HluDataset.lut_habitat_class.Rows.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_classTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_classTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_classDataTable) }, false);
                    }

                    // Only load habitat classes that are flagged as local.
                    _sourceHabitatClassCodes = HluDataset.lut_habitat_class.Where(r => r.is_local)
                        .OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _sourceHabitatClassCodes;
            }
        }
        
        /// <summary>
        /// Gets the list of source importance codes.
        /// </summary>
        /// <value>
        /// The list of source importance codes.
        /// </value>
        public HluDataSet.lut_importanceRow[] SourceImportanceCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (_sourceImportanceCodes == null)
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_importance.IsInitialized && (HluDataset.lut_importance.Rows.Count == 0))
                    {
                        if (_hluTableAdapterMgr.lut_importanceTableAdapter == null)
                            _hluTableAdapterMgr.lut_importanceTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_importanceDataTable, HluDataSet.lut_importanceRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_importanceDataTable) }, false);
                    }

                    _sourceImportanceCodes = HluDataset.lut_importance.OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();
                }

                return _sourceImportanceCodes;
            }
        }

        #endregion

        #region Source1
        
        public string Source1Header
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Source 1";
                else
                    return null;
            }
        }

        public Visibility ShowSource1Number
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        public HluDataSet.lut_sourcesRow[] Source1Names
        {
            get
            {
                // Load the lookup table if not already loaded.
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                // Return the list of source names, with the clear row if applicable, but
                // exclude the clear row in bulk update mode
                if ((IncidSource1Id != null) && (IncidSource1Id != Int32.MinValue))
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return SourceNames.Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.source_name).ToArray();
                }
                else
                {
                    return SourceNames;
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
                    // Check for equivalent null value when in bulk update mode
                    bool wasNull = (_incidSourcesRows[0] == null || (int)_incidSourcesRows[0]["source_id"] == Int32.MinValue);

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public bool IncidSource1Enabled
        {
            // Disable remaining source fields when source name is blank
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public HluDataSet.lut_habitat_typeRow[] Source1HabitatTypeCodes
        {
            get
            {
                // Get the list of values from the lookup table.
                if (!String.IsNullOrEmpty(IncidSource1HabitatClass))
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    // Only load habitat types that are flagged as local.
                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource1HabitatClass && r.is_local)
                        .OrderBy(r => r.sort_order).ThenBy(r => r.name).ToArray();

                    if ((retArray.Length == 1) && (IncidSource1Id != null))
                    {
                        IncidSource1HabitatType = retArray[0].code;
                        OnPropertyChanged("IncidSource1HabitatType");
                    }

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        #endregion

        #region Source2

        public string Source2Header
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Source 2";
                else
                    return null;
            }
        }

        public Visibility ShowSource2Number
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        public HluDataSet.lut_sourcesRow[] Source2Names
        {
            get
            {
                // Load the lookup table if not already loaded.
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                // Return the list of source names, with the clear row if applicable, but
                // exclude the clear row in bulk update mode
                if ((IncidSource2Id != null) && (IncidSource2Id != Int32.MinValue))
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return SourceNames.Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.source_name).ToArray();
                }
                else
                {
                    return SourceNames;
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
                    // Check for equivalent null value when in bulk update mode
                    bool wasNull = (_incidSourcesRows[1] == null || (int)_incidSourcesRows[1]["source_id"] == Int32.MinValue);

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public bool IncidSource2Enabled
        {
            // Disable remaining source fields when source name is blank
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public DataRow[] Source2HabitatTypeCodes
        {
            get
            {
                // Get the list of values from the lookup table
                if (!String.IsNullOrEmpty(IncidSource2HabitatClass))
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    // Only load habitat types that are flagged as local.
                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource2HabitatClass && r.is_local)
                        .OrderBy(r => r.sort_order).ThenBy(r => r.name).ToArray();

                    if ((retArray.Length == 1) && (IncidSource2Id != null))
                    {
                        IncidSource2HabitatType = retArray[0].code;
                        OnPropertyChanged("IncidSource2HabitatType");
                    }

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        #endregion

        #region Source3

        public string Source3Header
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return "Source 3";
                else
                    return null;
            }
        }

        public Visibility ShowSource3Number
        {
            get
            {
                if ((bool)_showGroupHeaders)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        public HluDataSet.lut_sourcesRow[] Source3Names
        {
            get
            {
                // Get the list of values from the lookup table.
                if (HluDataset.lut_sources.IsInitialized &&
                    HluDataset.lut_sources.Count == 0)
                {
                    if (_hluTableAdapterMgr.lut_sourcesTableAdapter == null)
                        _hluTableAdapterMgr.lut_sourcesTableAdapter =
                            new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                    _hluTableAdapterMgr.Fill(HluDataset,
                        new Type[] { typeof(HluDataSet.lut_sourcesDataTable) }, false);
                }

                // Return the list of source names, with the clear row if applicable, but
                // exclude the clear row in bulk update mode
                if ((IncidSource3Id != null) && (IncidSource3Id != Int32.MinValue))
                {
                    HluDataSet.lut_sourcesRow clearRow = HluDataset.lut_sources.Newlut_sourcesRow();
                    clearRow.source_id = -1;
                    clearRow.source_name = _codeDeleteRow;
                    clearRow.sort_order = -1;
                    return SourceNames.Concat(
                        new HluDataSet.lut_sourcesRow[] { clearRow }).OrderBy(r => r.sort_order).ThenBy(r => r.source_name).ToArray();
                }
                else
                {
                    return SourceNames;
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
                    // Check for equivalent null value when in bulk update mode
                    bool wasNull = (_incidSourcesRows[2] == null || (int)_incidSourcesRows[2]["source_id"] == Int32.MinValue);

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public bool IncidSource3Enabled
        {
            // Disable remaining source fields when source name is blank
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        public DataRow[] Source3HabitatTypeCodes
        {
            get
            {
                // Get the list of values from the lookup table
                if (!String.IsNullOrEmpty(IncidSource3HabitatClass))
                {
                    // Load the lookup table if not already loaded.
                    if (HluDataset.lut_habitat_type.IsInitialized && HluDataset.lut_habitat_type.Count == 0)
                    {
                        if (_hluTableAdapterMgr.lut_habitat_typeTableAdapter == null)
                            _hluTableAdapterMgr.lut_habitat_typeTableAdapter =
                                new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                        _hluTableAdapterMgr.Fill(HluDataset, new Type[] { typeof(HluDataSet.lut_habitat_typeDataTable) }, false);
                    }

                    // Only load habitat types that are flagged as local.
                    HluDataSet.lut_habitat_typeRow[] retArray = HluDataset.lut_habitat_type
                        .Where(r => r.habitat_class_code == IncidSource3HabitatClass && r.is_local)
                        .OrderBy(r => r.sort_order).ThenBy(r => r.name).ToArray();

                    if ((retArray.Length == 1) && (IncidSource3Id != null))
                    {
                        IncidSource3HabitatType = retArray[0].code;
                        OnPropertyChanged("IncidSource3HabitatType");
                    }

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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
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
                // Flag that the current record has changed so that the apply button
                // will appear.
                Changed = true;
            }
        }

        #endregion

        #endregion

        # region History Tab

        //TODO: IncidHistory - display new history fields and labels
        public IEnumerable<string> IncidHistory
        {
            get
            {
                if (_incidHistoryRows == null)
                    return null;
                else
                {
                    // Figure out which history columns to display based on the user options
                    // now that all the available history columns are always updated when
                    // creating history even if the user only wants to display some of them.
                    DataColumn[] displayHistoryColumns;
                    int result;
                    displayHistoryColumns = _gisIDColumns.Concat((from s in Settings.Default.HistoryColumnOrdinals.Cast<string>()
                                                                  where Int32.TryParse(s, out result) && (result >= 0) &&
                                                                       (result < _hluDS.incid_mm_polygons.Columns.Count) &&
                                                                       !_gisIDColumnOrdinals.Contains(result)
                                                                  select _hluDS.incid_mm_polygons.Columns[Int32.Parse(s)])).ToArray();

                    HabitatClasses = (from c in HluDataset.lut_habitat_class
                                      join t in HluDataset.lut_habitat_type on c.code equals t.habitat_class_code
                                      where c.is_local && t.is_local
                                      select c).Distinct().OrderBy(c => c.sort_order).ThenBy(c => c.description).ToArray();

                    //return (from r in _incidHistoryRows.OrderByDescending(r => r.history_id)
                    //        group r by new
                    return (from r in _incidHistoryRows.OrderByDescending(r => r.history_id)
                            group r by new
                            {
                                incid = r.incid,
                                // Display the modified_date column from the history with both the
                                // date and time to avoid separate updates with identical details
                                // (except the time) being merged together when displayed.
                                modified_date = !r.Ismodified_dateNull() ?
                                    r.modified_date.ToShortDateString() : String.Empty,
                                modified_time = (!r.Ismodified_dateNull() && r.modified_date != r.modified_date.Date) ?
                                    @" at " + r.modified_date.ToLongTimeString() : String.Empty,
                                modified_user_id = r.lut_userRow != null ? r.lut_userRow.user_name :
                                    !r.Ismodified_user_idNull() ? r.modified_user_id : String.Empty,

                                modified_process = r.lut_processRow != null ? r.lut_processRow.description : String.Empty,
                                modified_reason = r.lut_reasonRow != null ? r.lut_reasonRow.description : String.Empty,
                                modified_operation = r.lut_operationRow != null ? r.lut_operationRow.description : String.Empty,

                                // Only show the previous incid if it was different
                                modified_incid = !r.Ismodified_incidNull() ? String.Format("{0}", r.modified_incid == r.incid ? null : "\n\tPrevious INCID: " + r.modified_incid) : String.Empty,

                                //// Only show the previous values if they are not null and different
                                //modified_primary = displayHistoryColumns.Count(hc => hc.ColumnName == "habprimary") == 1 ?
                                //    //!r.Ismodified_habitat_primaryNull() ? String.Format("\n\tPrevious Primary: {0}", r.modified_habitat_primary) : String.Empty : String.Empty,
                                //    !r.Ismodified_habitat_primaryNull() ? String.Format("{0}", r.modified_habitat_primary == IncidCurrentRow.habitat_primary ? null : "\n\tPrevious Primary: " + r.modified_habitat_primary) : String.Empty : String.Empty,
                                //modified_secondaries = displayHistoryColumns.Count(hc => hc.ColumnName == "habsecond") == 1 ?
                                //    //!r.Ismodified_habitat_secondariesNull() ? String.Format("\n\tPrevious Secondaries: {0}", r.modified_habitat_secondaries) : String.Empty : String.Empty,
                                //    !r.Ismodified_habitat_secondariesNull() ? String.Format("{0}", r.modified_habitat_secondaries == IncidCurrentRow.habitat_secondaries ? null : "\n\tPrevious Secondaries: " + r.modified_habitat_secondaries) : String.Empty : String.Empty,
                                //modified_determination = displayHistoryColumns.Count(hc => hc.ColumnName == "determqty") == 1 ?
                                //    //r.lut_quality_determinationRow != null ? String.Format("\n\tPrevious Determination: {0}", r.lut_quality_determinationRow.description) : String.Empty : String.Empty,
                                //    r.lut_quality_determinationRow != null ? String.Format("{0}", r.modified_habitat_determination == IncidCurrentRow.quality_determination ? null : "\n\tPrevious Determination: " + r.lut_quality_determinationRow.description) : String.Empty : String.Empty,
                                //modified_intepretation = displayHistoryColumns.Count(hc => hc.ColumnName == "interpqty") == 1 ?
                                //    //r.lut_quality_interpretationRow != null ? String.Format("\n\tPrevious Interpretation: {0}", r.lut_quality_interpretationRow.description) : String.Empty : String.Empty,
                                //    r.lut_quality_interpretationRow != null ? String.Format("{0}", r.modified_habitat_interpretation == IncidCurrentRow.quality_interpretation ? null : "\n\tPrevious Interpretation: " + r.lut_quality_interpretationRow.description) : String.Empty : String.Empty,

                                // Only show the previous values if they are not null
                                modified_primary = displayHistoryColumns.Count(hc => hc.ColumnName == "habprimary") == 1 ?
                                    !r.Ismodified_habitat_primaryNull() ? String.Format("\n\tPrevious Primary: {0}", r.modified_habitat_primary) : String.Empty : String.Empty,
                                modified_secondaries = displayHistoryColumns.Count(hc => hc.ColumnName == "habsecond") == 1 ?
                                    !r.Ismodified_habitat_secondariesNull() ? String.Format("\n\tPrevious Secondaries: {0}", r.modified_habitat_secondaries) : String.Empty : String.Empty,
                                modified_determination = displayHistoryColumns.Count(hc => hc.ColumnName == "determqty") == 1 ?
                                    r.lut_quality_determinationRow != null ? String.Format("\n\tPrevious Determination: {0}", r.lut_quality_determinationRow.description) : String.Empty : String.Empty,
                                modified_intepretation = displayHistoryColumns.Count(hc => hc.ColumnName == "interpqty") == 1 ?
                                    r.lut_quality_interpretationRow != null ? String.Format("\n\tPrevious Interpretation: {0}", r.lut_quality_interpretationRow.description) : String.Empty : String.Empty,

                            } into g
                            select
                                String.Format("{0} on {1}{2} by {3}:", g.Key.modified_operation, g.Key.modified_date, g.Key.modified_time, g.Key.modified_user_id) +

                                String.Format("\n\tProcess: {0}", g.Key.modified_process) +
                                String.Format("\n\tReason: {0}", g.Key.modified_reason) +
                                //String.Format("\n\tOperation: {0}", g.Key.modified_operation) +

                                g.Key.modified_incid +
                                g.Key.modified_primary +
                                g.Key.modified_secondaries +
                                g.Key.modified_determination +
                                g.Key.modified_intepretation +

                                // Show the area and length values in the history as hectares and metres.
                                String.Format("\n\tModified Length: {0} [km]", g.Distinct(_histRowEqComp)
                                    .Sum(r => !r.Ismodified_lengthNull() ? Math.Round(r.modified_length / 1000, 3) : 0).ToString("f3")) +
                                String.Format("\n\tModified Area: {0} [ha]", g.Distinct(_histRowEqComp)
                                    .Sum(r => !r.Ismodified_areaNull() ? Math.Round(r.modified_area / 10000, 4) : 0).ToString("f4")))
                                .Take(_historyDisplayLastN);

                }
            }
        }

        #endregion

        #region Record IDs

        public GeometryTypes GisLayerType { get { return _gisLayerType; } }

        public string SiteID { get { return _recIDs.SiteID; } }

        public string HabitatVersion { get { return _recIDs.HabitatVersion; } }

        public string CurrentIncid { get { return _recIDs.CurrentIncid; } }

        public string NextIncid { get { return _recIDs.NextIncid; } }

        private int CurrentIncidBapId { get { return _recIDs.CurrentIncidBapId; } }

        private int NextIncidBapId { get { return _recIDs.NextIncidBapId; } }

        private int NextIncidSourcesId { get { return _recIDs.NextIncidSourcesId; } }

        private int NextIncidSecondaryId { get { return _recIDs.NextIncidSecondaryId; } }

        private int NextIncidConditionId { get { return _recIDs.NextIncidConditionId; } }

        #endregion

        #region SQLUpdater

        /// <summary>
        /// Replaces any string or date delimiters with connection type specific
        /// versions and qualifies any table names.
        /// </summary>
        /// <param name="words">The words.</param>
        /// <returns></returns>
        internal String ReplaceStringQualifiers(String sqlcmd)
        {
            // Check if a table name (delimited by '[]' characters) is found
            // in the sql command.
            int i1 = 0;
            int i2 = 0;
            String start = String.Empty;
            String end = String.Empty;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("[", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("]", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the table name.
                        string table = sqlcmd.Substring(i1 + 1, i2 - i1 - 1);

                        // Split the table name from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace the table name with a qualified table name.
                        sqlcmd = start + _db.QualifyTableName(table) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }

            // Check if any strings are found (delimited by single quotes)
            // in the sql command.
            i1 = 0;
            i2 = 0;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("'", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("'", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the text string.
                        string text = sqlcmd.Substring(i1 + 1, i2 - i1 - 1);

                        // Split the text string from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace any wild characters found in the text.
                        if (start.TrimEnd().EndsWith(" LIKE"))
                        {
                            text.Replace("_", _db.WildcardSingleMatch);
                            text.Replace("%", _db.WildcardManyMatch);
                        }

                        // Replace the text delimiters with the correct delimiters.
                        sqlcmd = start + _db.QuoteValue(text) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }

            // Check if any dates are found (delimited by '#' characters)
            // in the sql command.
            i1 = 0;
            i2 = 0;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("#", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("#", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the date string.
                        DateTime dt;
                        DateTime.TryParse(sqlcmd.Substring(i1 + 1, i2 - i1 - 1), out dt);

                        // Split the date string from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace the date delimiters with the correct delimiters.
                        sqlcmd = start + _db.QuoteValue(dt) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }
            return sqlcmd;
        }

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

        private List<string[]> ValidateCondition()
        {
            List<string[]> errors = new List<string[]>();

            // Validate the condition fields if no condition has been entered
            if (IncidCondition == null)
            {
                if (IncidConditionQualifier != null)
                    errors.Add(new string[] { "IncidConditionQualifier", "Error: Condition qualifier is not valid without a condition" });
                if (IncidConditionDate != null)
                    errors.Add(new string[] { "IncidConditionDate", "Error: Condition date is not valid without a condition" });
            }
            else
            {
                // Check the condition fields if a condition has been entered
                if (IncidConditionQualifier == null)
                    errors.Add(new string[] { "IncidConditionQualifier", "Error: Condition qualifier is mandatory for a condition" });
                if (IncidConditionDate == null)
                    errors.Add(new string[] { "IncidConditionDate", "Error: Condition date is mandatory for a condition" });
                else if (IncidConditionDate.IsBad)
                    errors.Add(new string[] { "IncidConditionDate", "Error: Invalid condition vague date" });
            }

            return errors;
        }

        private List<string[]> ValidateSource1()
        {
            List<string[]> errors = new List<string[]>();

            // Validate the source if it is real
            if (IncidSource1Id != null && IncidSource1Id != Int32.MinValue)
            {
                // Check the source id is found in the lookup table
                EnumerableRowCollection<HluDataSet.lut_sourcesRow> rows =
                    HluDataset.lut_sources.Where(r => r.source_id == IncidSource1Id);
                if (rows.Count() == 0)
                    errors.Add(new string[] { "IncidSource1Id", "Error: Source name is mandatory for each source" });
                if (IncidSource1Date == null)
                    errors.Add(new string[] { "IncidSource1Date", "Error: Date is mandatory for each source" });
                else if (IncidSource1Date.IsBad)
                    errors.Add(new string[] { "IncidSource1Date", "Error: Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource1HabitatClass))
                    errors.Add(new string[] { "IncidSource1HabitatClass", "Error: Habitat class is mandatory for each source" });
                else if ((IncidSource1HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource1HabitatType))
                    errors.Add(new string[] { "IncidSource1HabitatType", "Error: Habitat type is mandatory if habitat class is filled in" });

                // Use the skip value from settings
                string skipVal = Settings.Default.SourceImportanceSkip;
                if (String.IsNullOrEmpty(IncidSource1BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource1BoundaryImportance", "Error: Boundary importance is mandatory for each source" });
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
                        "Error: Habitat importance is mandatory for each source" });
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
                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // Validation for OSMM Bulk Update mode.
                //    
                if ((OSMMBulkUpdateMode == true) &&
                    (IncidSource2Id == null || IncidSource2Id == Int32.MinValue) &&
                    (IncidSource3Id == null || IncidSource3Id == Int32.MinValue))
                    errors.Add(new string[] { "IncidSource1Id", 
                        "Error: At least one source must be specified" });
                if (IncidSource1Date != null)
                    errors.Add(new string[] { "IncidSource1Date", 
                        "Error: Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatClass))
                    errors.Add(new string[] { "IncidSource1HabitatClass", 
                        "Error: Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatType))
                    errors.Add(new string[] { "IncidSource1HabitatType", 
                        "Error: Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1BoundaryImportance))
                    errors.Add(new string[] { "IncidSource1BoundaryImportance", 
                        "Error: Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource1HabitatImportance))
                    errors.Add(new string[] { "IncidSource1HabitatImportance", 
                        "Error: Habitat importance cannot be filled in if no source has been specified" });
                //---------------------------------------------------------------------
            }

            return errors;
        }

        private List<string[]> ValidateSource2()
        {
            List<string[]> errors = new List<string[]>();

            // Validate the source if it is real
            if (IncidSource2Id != null && IncidSource2Id != Int32.MinValue)
            {
                // Check the source id is found in the lookup table
                EnumerableRowCollection<HluDataSet.lut_sourcesRow> rows =
                    HluDataset.lut_sources.Where(r => r.source_id == IncidSource2Id);
                if (rows.Count() == 0)
                    errors.Add(new string[] { "IncidSource1Id", "Error: Source name is mandatory for each source" });
                if (IncidSource2Date == null)
                    errors.Add(new string[] { "IncidSource2Date", "Error: Date is mandatory for each source" });
                else if (IncidSource2Date.IsBad)
                    errors.Add(new string[] { "IncidSource2Date", "Error: Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource2HabitatClass))
                    errors.Add(new string[] { "IncidSource2HabitatClass", "Error: Habitat class is mandatory for each source" });
                else if ((IncidSource2HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource2HabitatType))
                    errors.Add(new string[] { "IncidSource2HabitatType", "Error: Habitat type is mandatory if habitat class is filled in" });

                // Use the skip value from settings
                string skipVal = Settings.Default.SourceImportanceSkip;
                if (String.IsNullOrEmpty(IncidSource2BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource2BoundaryImportance", "Error: Boundary importance is mandatory for each source" });
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
                        "Error: Habitat importance is mandatory for each source" });
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
                        "Error: Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatClass))
                    errors.Add(new string[] { "IncidSource2HabitatClass", 
                        "Error: Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatType))
                    errors.Add(new string[] { "IncidSource2HabitatType", 
                        "Error: Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2BoundaryImportance))
                    errors.Add(new string[] { "IncidSource2BoundaryImportance", 
                        "Error: Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource2HabitatImportance))
                    errors.Add(new string[] { "IncidSource2HabitatImportance", 
                        "Error: Habitat importance cannot be filled in if no source has been specified" });
            }

            return errors;
        }

        private List<string[]> ValidateSource3()
        {
            List<string[]> errors = new List<string[]>();

            // Validate the source if it is real
            if (IncidSource3Id != null && IncidSource3Id != Int32.MinValue)
            {
                // Check the source id is found in the lookup table
                EnumerableRowCollection<HluDataSet.lut_sourcesRow> rows =
                    HluDataset.lut_sources.Where(r => r.source_id == IncidSource3Id);
                if (rows.Count() == 0)
                    errors.Add(new string[] { "IncidSource1Id", "Error: Source name is mandatory for each source" });
                if (IncidSource3Date == null)
                    errors.Add(new string[] { "IncidSource3Date", "Error: Date is mandatory for each source" });
                else if (IncidSource3Date.IsBad)
                    errors.Add(new string[] { "IncidSource3Date", "Error: Invalid vague date" });
                if (String.IsNullOrEmpty(IncidSource3HabitatClass))
                    errors.Add(new string[] { "IncidSource3HabitatClass", "Error: Habitat class is mandatory for each source" });
                else if ((IncidSource3HabitatClass.ToLower() == "none") != String.IsNullOrEmpty(IncidSource3HabitatType))
                    errors.Add(new string[] { "IncidSource3HabitatType", "Error: Habitat type is mandatory if habitat class is filled in" });

                // Use the skip value from settings
                string skipVal = Settings.Default.SourceImportanceSkip;
                if (String.IsNullOrEmpty(IncidSource3BoundaryImportance))
                {
                    errors.Add(new string[] { "IncidSource3BoundaryImportance", "Error: Boundary importance is mandatory for each source" });
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
                        "Error: Habitat importance is mandatory for each source" });
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
                        "Error: Date cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatClass))
                    errors.Add(new string[] { "IncidSource3HabitatClass", 
                        "Error: Habitat class cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatType))
                    errors.Add(new string[] { "IncidSource3HabitatType", 
                        "Error: Habitat type cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3BoundaryImportance))
                    errors.Add(new string[] { "IncidSource3BoundaryImportance", 
                        "Error: Boundary importance cannot be filled in if no source has been specified" });
                if (!String.IsNullOrEmpty(IncidSource3HabitatImportance))
                    errors.Add(new string[] { "IncidSource3HabitatImportance", 
                        "Error: Habitat importance cannot be filled in if no source has been specified" });
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
            string errMsg = String.Format("Error: {0}", split[split.Length - 1]);

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
            string errMsg = String.Format("Error: {0}", split[split.Length - 1]);

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
                // Show errors in bulk update mode.
                if ((_incidCurrentRow == null) ||
                    (_incidCurrentRow.RowState == DataRowState.Detached && _bulkUpdateMode == false)) return null;

                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(Reason))
                    error.Append(Environment.NewLine).Append("Reason is mandatory for the history trail of every update");

                if (String.IsNullOrEmpty(Process))
                    error.Append(Environment.NewLine).Append("Process is mandatory for the history trail of every update");

                if (_bulkUpdateMode == false)
                {
                    if (String.IsNullOrEmpty(IncidBoundaryBaseMap))
                        error.Append(Environment.NewLine).Append("Boundary basemap is mandatory for every INCID");

                    if (String.IsNullOrEmpty(IncidDigitisationBaseMap))
                        error.Append(Environment.NewLine).Append("Digitisation basemap is mandatory for every INCID");

                    if (_qualityValidation == 1)
                    {
                        if (String.IsNullOrEmpty(IncidQualityDetermination))
                            error.Append(Environment.NewLine).Append("Quality determination is mandatory for every INCID");

                        if (String.IsNullOrEmpty(IncidQualityInterpretation))
                            error.Append(Environment.NewLine).Append("Quality interpretation is mandatory for every INCID");

                        if ((!String.IsNullOrEmpty(IncidQualityComments) && String.IsNullOrEmpty(IncidQualityInterpretation)))
                            error.Append(Environment.NewLine).Append("Interpretation comments are invalid without interpretation quality");
                    }
                }

                if (String.IsNullOrEmpty(IncidPrimary) && _bulkUpdateMode == false)
                    error.Append(Environment.NewLine).Append("Primary Habitat is mandatory for every INCID");

                // If there are any IHS field errors then show an error on the tab label.
                if (HabitatErrors != null && HabitatErrors.Count > 0)
                    error.Append(Environment.NewLine).Append("One or more habitat fields are in error");

                // If there are any Priority field errors then show an error on the tab label.
                if (PriorityErrors != null && PriorityErrors.Count > 0)
                    error.Append(Environment.NewLine).Append("One or more priority fields are in error");

                // If there are any Detail field errors then show an error on the tab label.
                if (DetailsErrors != null && DetailsErrors.Count > 0)
                    error.Append(Environment.NewLine).Append("One or more detail fields are in error");

                if ((ConditionErrors != null) && (ConditionErrors.Count > 0))
                    error.Append(Environment.NewLine).Append(ErrorMessageList(ConditionErrors));

                // If there are any Source field errors then show an error on the tab label.
                if (((Source1Errors != null) && (Source1Errors.Count > 0)) ||
                    ((Source2Errors != null) && (Source2Errors.Count > 0)) ||
                    ((Source3Errors != null) && (Source3Errors.Count > 0)))
                    error.Append(Environment.NewLine).Append("One or more source fields are in error");

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
                // Show errors in bulk update mode.
                if ((_incidCurrentRow == null) ||
                    (_incidCurrentRow.RowState == DataRowState.Detached && _bulkUpdateMode == false)) return null;

                string error = null;

                switch (columnName)
                {
                    case "Reason":
                        if (String.IsNullOrEmpty(Reason))
                            error = "Error: Reason is mandatory for the history trail of every INCID";
                        break;
                    case "Process":
                        if (String.IsNullOrEmpty(Process))
                            error = "Error: Process is mandatory for the history trail of every INCID";
                        break;

                }

                // Check the individual field errors to see if their parent tab label
                // should be flagged as also in error.
                switch (columnName)
                {
                    case "Incid":
                        break;
                    case "HabitatTabLabel":
                        if (HabitatWarnings != null && HabitatWarnings.Count > 0)
                            error = "Warning: One or more habitat fields have a warning";
                        if (HabitatErrors != null && HabitatErrors.Count > 0)
                            error = "Error: One or more habitat fields are in error";
                        break;
                    case "IncidPrimary":
                        // If the field is in error add the field name to the list of errors
                        // for the parent tab. Otherwise remove the field from the list.
                        if (String.IsNullOrEmpty(IncidPrimary) && _bulkUpdateMode == false)
                        {
                            error = "Error: Primary Habitat is mandatory for every INCID";
                            AddErrorList(ref _habitatErrors, columnName);
                        }
                        else
                        {
                            DelErrorList(ref _habitatErrors, columnName);
                        }
                        OnPropertyChanged("HabitatTabLabel");
                        break;
                    case "IncidCondition":
                    case "IncidConditionQualifier":
                    case "IncidConditionDate":
                        // Store the Source1 field errors so that they can be checked
                        // later to see if the Source tab label should also be flagged
                        // as in error.
                        ConditionErrors = ValidateCondition();
                        error = ErrorMessage(columnName, ConditionErrors);
                        OnPropertyChanged("DetailsTabLabel");
                        break;
                    case "PriorityTabLabel":
                        if (PriorityWarnings != null && PriorityWarnings.Count > 0)
                            error = "Warning: One or more priority fields have a warning";

                        if (PriorityErrors != null && PriorityErrors.Count > 0)
                            error = "Error: One or more priority fields are in error";
                        break;
                    case "DetailsTabLabel":
                        if ((DetailsWarnings != null && DetailsWarnings.Count > 0) ||
                            (ConditionWarnings != null && ConditionWarnings.Count > 0))
                            error = "Warning: One or more detail fields have a warning";

                        if ((DetailsErrors != null && DetailsErrors.Count > 0) ||
                            (ConditionErrors != null && ConditionErrors.Count > 0))
                            error = "Error: One or more details fields are in error";
                        break;
                    case "SourcesTabLabel":
                        if ((Source1Warnings != null && Source1Warnings.Count > 0) ||
                            (Source2Warnings != null && Source2Warnings.Count > 0) ||
                            (Source3Warnings != null && Source3Warnings.Count > 0))
                            error = "Warning: One or more source fields have a warning";

                        if ((Source1Errors != null && Source1Errors.Count > 0) ||
                            (Source2Errors != null && Source2Errors.Count > 0) ||
                            (Source3Errors != null && Source3Errors.Count > 0))
                            error = "Error: One or more source fields are in error";
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

                switch (columnName)
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR49 Process proposed OSMM Updates
                    // Warnings with in Bulk Update mode.
                    //    
                    case "NumIncidSelectedMap":
                        if (_incidsSelectedMapCount < _incidsSelectedDBCount)
                            error = "Warning: Not all database incids are selected in map";
                        break;
                    case "NumToidSelectedMap":
                        if (_toidsSelectedMapCount < _toidsSelectedDBCount)
                            error = "Warning: Not all database toids are selected in map";
                        break;
                    case "NumFragmentsSelectedMap":
                        if (_fragsSelectedMapCount < _fragsSelectedDBCount)
                            error = "Warning: Not all database fragments are selected in map";
                        break;
                    case "IncidOSMMUpdateStatus":
                        if (_incidOSMMUpdatesStatus != null & _incidOSMMUpdatesStatus >= 0)
                            error = "Warning: OSMM Update is outstanding";
                        break;
                }

                if (_osmmUpdateMode == true) return null;

                if (_bulkUpdateMode == false)
                {
                    switch (columnName)
                    {
                        case "IncidBoundaryBaseMap":
                            // If the field is in error add the field name to the list of errors
                            // for the parent tab. Otherwise remove the field from the list.
                            if (String.IsNullOrEmpty(IncidBoundaryBaseMap))
                            {
                                error = "Error: Boundary basemap is mandatory for every INCID";
                                AddErrorList(ref _detailsErrors, columnName);
                            }
                            else
                            {
                                DelErrorList(ref _detailsErrors, columnName);
                            }
                            OnPropertyChanged("DetailsTabLabel");
                            break;
                        case "IncidDigitisationBaseMap":
                            // If the field is in error add the field name to the list of errors
                            // for the parent tab. Otherwise remove the field from the list.
                            if (String.IsNullOrEmpty(IncidDigitisationBaseMap))
                            {
                                error = "Error: Digitisation basemap is mandatory for every INCID";
                                AddErrorList(ref _detailsErrors, columnName);
                            }
                            else
                            {
                                DelErrorList(ref _detailsErrors, columnName);
                            }
                            OnPropertyChanged("DetailsTabLabel");
                            break;
                        case "IncidQualityDetermination":
                            if ((_qualityValidation == 1)
                                && (String.IsNullOrEmpty(IncidQualityDetermination)))
                            {
                                error = "Error: Determination quality is mandatory for every INCID";
                                AddErrorList(ref _detailsErrors, columnName);
                            }
                            else
                            {
                                DelErrorList(ref _detailsErrors, columnName);
                            }
                            OnPropertyChanged("DetailsTabLabel");
                            break;
                        case "IncidQualityInterpretation":
                            if ((_qualityValidation == 1)
                                && (String.IsNullOrEmpty(IncidQualityInterpretation)))
                            {
                                error = "Error: Interpretation quality is mandatory for every INCID";
                                AddErrorList(ref _detailsErrors, columnName);
                            }
                            else
                            {
                                DelErrorList(ref _detailsErrors, columnName);
                            }
                            OnPropertyChanged("DetailsTabLabel");
                            break;
                        case "IncidQualityComments":
                            if ((_qualityValidation == 1)
                                && (!String.IsNullOrEmpty(IncidQualityComments))
                                && String.IsNullOrEmpty(IncidQualityInterpretation))
                            {
                                error = "Error: Interpretation comments are invalid without interpretation quality";
                                AddErrorList(ref _detailsErrors, columnName);
                            }
                            else
                            {
                                DelErrorList(ref _detailsErrors, columnName);
                            }
                            OnPropertyChanged("DetailsTabLabel");
                            break;
                    }
                }

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                OnPropertyChanged("CanCopy");
                OnPropertyChanged("CanPaste");

                return error;
            }
        }

        public void AddErrorList(ref List<string> errorList, string columnName)
        {
            if (!errorList.Contains(columnName))
                errorList.Add(columnName);
        }

        public void DelErrorList(ref List<string> errorList, string columnName)
        {
            if (errorList.Contains(columnName))
                errorList.Remove(columnName);
        }

        public void ResetWarningsErrors()
        {
            _habitatWarnings = new List<string>();
            _habitatErrors = new List<string>();
            _priorityWarnings = new List<string>();
            _priorityErrors = new List<string>();
            _detailsWarnings = new List<string>();
            _detailsErrors = new List<string>();
            _conditionWarnings = null;
            _conditionErrors = null;
            _source1Warnings = null;
            _source2Warnings = null;
            _source3Warnings = null;
            _source1Errors = null;
            _source2Errors = null;
            _source3Errors = null;
        }

        #endregion

    }
}
