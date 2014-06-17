// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013-2014 Thames Valley Environmental Records Centre
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

//Comment out the line below if compiling for versions of ArcGIS 9.3 or
//earlier and remove the comments for versions of ArcGIS 10.0 or later.
#define ARC10

using System;
using System.Collections.Generic;
using System.Configuration.Install;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AppModule.InterProcessComm;
using ESRI.ArcGIS.ADF.CATIDs;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.CartoUI;
using ESRI.ArcGIS.DataSourcesOleDB;
using ESRI.ArcGIS.Editor;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;
using ESRI.ArcGIS.Geometry;
using HLU.Data.Model;
using HLU.GISApplication;
using HLU.GISApplication.ArcGIS;
using HLU.UI.ViewModel;
using Microsoft.Win32;
using Server;

using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.Display;

namespace HLU
{
    [Guid("c61db89f-7118-4a10-a5c1-d4a375867a02")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("HLU.HluArcMapExtension")]
    public class HluArcMapExtension : IExtension, IExtensionConfig, IPersistVariant
    {
        #region Delegates

        public delegate void PipeSelectionDelegate(IFeatureSelection featureSelection);
        public delegate void FlashSelectedFeatureDelegate(IQueryFilter queryFilter);
        public delegate void SplitFeatureDelegate(IQueryFilter selectionQueryFilter,
            string lastToidFragmentID, string[] historyColumns);
        public delegate void SplitFeaturesLogicallyDelegate(string oldIncid, string newIncid, string[] historyColumns);
        public delegate void MergeFeaturesDelegate(IQueryFilter resultFeatureQueryFilter,
            string newToidFragmentID, string[] historyColumns);
        public delegate void MergeFeaturesLogicallyDelegate(string keepIncid, string[] historyColumns);
        public delegate void UpdateAttributesSelectionDelegate(List<int> updateFieldOrdinals,
            List<object> updateValues, string[] historyColumns);
        public delegate void UpdateAttributesDelegate(IQueryFilter queryFilter, 
            List<int> updateFieldOrdinals, List<object> updateValues, string[] historyColumns);
        public delegate void UpdateAttributesBulkDelegate(string scratchMdbPath, string selectionDatasetName,
            List<int> updateFieldOrdinals, List<object> updateValues, string[] historyColumns);
        public delegate void SelectByQueryFilterDelegate(IQueryFilter queryFilter);
        public delegate void SelectByQueryDefDelegate(IQueryDef queryDef, string oidColumnAlias);
        public delegate void SelectByJoinDelegate(string scratchMdbPath, string selectionDatasetName);
        public delegate void ZoomSelectedDelegate();
        public delegate void ZoomSelectedCursorDelegate(IQueryFilter queryFilter);
        public delegate void ExportDelegate(string mdbPathName, string attributeDatasetName, int exportRowCount);
        public delegate void IsHluWorkspaceDelegate();
        public delegate void ListHluLayersDelegate();
        public delegate void IsHluLayerDelegate(int ixMap, int ixLayer);
        public delegate void IsEditingDelegate();

        #endregion

        #region Events

        public static event EventHandler OutgoingDataReady;

        #endregion

        #region Fields

        private IApplication _application;
        private esriExtensionState _enableState;
        private object[] _validWorkspaces = null;
        private HluGISLayer.incid_mm_polygonsDataTable _hluLayerStructure;
        private int[] _hluUidFieldOrdinals;
        private int[] _hluFieldMap;
        private string[] _hluFieldNames;
        private int[] _selectFieldOrdinals;
        private string[] _selectFields;
        private string[] _selectColumns;
        private ISQLSyntax _hluSqlSyntax;
        private string _quotePrefix;
        private string _quoteSuffix;
        private Dictionary<string, esriSQLPredicates> _sqlPredicates;
        private Dictionary<Type, int> _typeMapSystemToSql;
        private Dictionary<int, Type> _typeMapSQLToSystem;
        private string[] _hluFieldSysTypeNames;
        private IFeatureWorkspace _hluWS;
        private IWorkspaceEdit _hluWorkspaceEdit;
        private bool _hluWSisSDE = false;
        private IFeatureLayer _hluLayer;
        private List<string> _hluLayerList;
        private string _hluTableName;
        private IFeatureClass _hluFeatureClass;
        private IFeatureSelection _hluFeatureSelection;
        private IActiveView _hluView;
        private IEditor _editor;
        private ICommandItem _saveEditsCmd;
        private IMap _focusMap;
        private IDocumentEvents_Event _docEvents = null;
        private IActiveViewEvents_FocusMapChangedEventHandler _activeViewEventsFocusMapChanged;
        private IActiveViewEvents_ItemAddedEventHandler _activeViewEventsItemAdded;
        private IActiveViewEvents_ItemDeletedEventHandler _activeViewEventsItemDeleted;
        private IFeatureLayerSelectionEvents_Event _hluLayerSelectionChangedHandler;
        private string _pipeName;
        private int _pipeMaxReadBytes = ArcMapApp.PipeMaxReadBytes;
        private string _pipeTransmissionInterrupt = ArcMapApp.PipeTransmissionInterrupt;
        private string _pipeTransmissionEnd = ArcMapApp.PipeTransmissionEnd;
        private char _pipeFieldDelimiter = ArcMapApp.PipeFieldDelimiter;
        private char _pipeErrorSymbol = ArcMapApp.PipeErrorSymbol;
        private string _historyGeometry1ColumnName = ViewModelWindowMain.HistoryGeometry1ColumnName;
        private string _historyGeometry2ColumnName = ViewModelWindowMain.HistoryGeometry2ColumnName;
        private bool _sendColumnHeaders;
        private bool _joinedTable = false;
        private bool _pipeCalling = false;
        private static List<string> _pipeData;
        private System.Windows.Forms.Control _dummyControl;
        private static int _whereClauseLengthMax = Properties.Settings.Default.WhereClauseMaxLength;
        private static PipeSelectionDelegate _pipeSelDel;
        private static FlashSelectedFeatureDelegate _flashSelFeatDel;
        private static SplitFeatureDelegate _splitFeatDel;
        private static SplitFeaturesLogicallyDelegate _splitFeatLogDel;
        private static MergeFeaturesLogicallyDelegate _mergeFeatLogDel;
        private static MergeFeaturesDelegate _mergeFeatDel;
        private static UpdateAttributesDelegate _updAttsDel;
        private static UpdateAttributesBulkDelegate _updAttsBulkDel;
        private static UpdateAttributesSelectionDelegate _updAttsSelDel;
        private static SelectByQueryFilterDelegate _selByQFilterDel;
        private static SelectByQueryDefDelegate _selByQDefDel;
        private static SelectByJoinDelegate _selByJoinDel;
        private static ZoomSelectedDelegate _zoomSelDel;
        private static ZoomSelectedCursorDelegate _zoomSelCursorDel;
        private static ExportDelegate _exportDel;
        private static IsHluWorkspaceDelegate _isHluWorkspaceDel;
        private static ListHluLayersDelegate _ListHluLayersDel;
        private static IsHluLayerDelegate _isHluLayerDel;
        private static IsEditingDelegate _isEditingDel;
        
        #endregion

        #region COM Registration Function(s)

        [ComRegisterFunction()]
        [ComVisible(false)]
        static void RegisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryRegistration(registerType);
            #if ARC10
            #else
                EsriRegasm(true, registerType);
            #endif
        }

        [ComUnregisterFunction()]
        [ComVisible(false)]
        static void UnregisterFunction(Type registerType)
        {
            // Required for ArcGIS Component Category Registrar support
            ArcGISCategoryUnregistration(registerType);
            #if ARC10
            #else
                EsriRegasm(false, registerType);
            #endif
        }

        #region ArcGIS Component Category Registrar generated code

        /// <summary>
        /// Required method for ArcGIS Component Category registration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryRegistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxExtension.Register(regKey);
        }

        /// <summary>
        /// Required method for ArcGIS Component Category unregistration -
        /// Do not modify the contents of this method with the code editor.
        /// </summary>
        private static void ArcGISCategoryUnregistration(Type registerType)
        {
            string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", registerType.GUID);
            MxExtension.Unregister(regKey);
        }

        private static void EsriRegasm(bool install, Type registerType)
        {
            try
            {
                int arcVersion = -1;
                string regCmd = String.Empty, args = String.Empty;

                RegistryKey rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ESRI\ArcGIS");
                object rkVal = rk.GetValue("RealVersion");
                if ((rkVal == null) || !Int32.TryParse(rkVal.ToString().Split('.')[0], out arcVersion)) arcVersion = -1;

                if (arcVersion > 9)
                {
                    ProcessStartInfo psi = new ProcessStartInfo(System.IO.Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.CommonProgramFiles), @"ArcGIS\bin\ESRIRegasm.exe"));
                    psi.Arguments = String.Format(@"{0} /p:Desktop{1} /s",
                        System.IO.Path.GetFileName(registerType.Assembly.Location), install ? String.Empty : @" /u");
                    psi.WorkingDirectory = System.IO.Path.GetDirectoryName(registerType.Assembly.Location);
                    psi.CreateNoWindow = true;
                    psi.UseShellExecute = false;

                    Process p = new Process();
                    p.StartInfo = psi;
                    p.Start();
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("{0}\n{1}", ex.Source, ex.Message), "HLU GIS Tool Installer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        #endregion

        #endregion

        #region IExtension Members

        /// <summary>
        /// Name of extension. Do not exceed 31 characters
        /// </summary>
        public string Name
        {
            get { return "HluArcMapExtension"; }
        }

        public void Shutdown()
        {
            _hluLayer = null;
            _hluFeatureClass = null;
            _hluFeatureSelection = null;
            _hluWS = null;
            _hluSqlSyntax = null;
            _hluView = null;
            _hluFieldMap = null;
            _hluFieldNames = null;
            _hluUidFieldOrdinals = null;
            _selectFieldOrdinals = null;

            _pipeSelDel = null;
            _flashSelFeatDel = null;
            _splitFeatDel = null;
            _splitFeatLogDel = null;
            _mergeFeatDel = null;
            _mergeFeatLogDel = null;
            _updAttsDel = null;
            _updAttsSelDel = null;
            _updAttsBulkDel = null;
            _selByQDefDel = null;
            _selByQFilterDel = null;
            _selByJoinDel = null;
            _zoomSelDel = null;
            _zoomSelCursorDel = null;
            _exportDel = null;
            _isHluWorkspaceDel = null;
            _ListHluLayersDel = null;
            _isHluLayerDel = null;
            _isEditingDel = null;

            if (PipeManager != null)
            {
                PipeManager = null;
            }

            RemoveActiveViewEvents(_focusMap);
        }

        public void Startup(ref object initializationData)
        {
            _application = initializationData as IApplication;
            if (_application == null) return;

            ArcMapAppHelperClass.GetValidWorkspaces(out _validWorkspaces);
            ArcMapAppHelperClass.GetTypeMaps(out _sqlPredicates, out _typeMapSystemToSql, out _typeMapSQLToSystem);
            _hluLayerStructure = new HluGISLayer.incid_mm_polygonsDataTable();

            SetupDocumentEvents(_application.Document);

            if (PipeManager == null)
            {
                _pipeName = String.Format("{0}.{1}", ArcMapApp.PipeBaseName, _application.hWnd);
                PipeManager = new PipeManager();
                PipeManager.Initialize(_pipeName, _pipeMaxReadBytes);
                PipeManager.IncomingDataReady += new EventHandler(_pipeManager_IncomingDataReady);

                _dummyControl = new System.Windows.Forms.Control();
                _dummyControl.CreateControl();

                _pipeSelDel = new PipeSelectionDelegate(PipeSelection);
                _flashSelFeatDel = new FlashSelectedFeatureDelegate(FlashFeature);
                _splitFeatDel = new SplitFeatureDelegate(SplitFeature);
                _splitFeatLogDel = new SplitFeaturesLogicallyDelegate(SplitFeaturesLogically);
                _mergeFeatDel = new MergeFeaturesDelegate(MergeFeatures);
                _mergeFeatLogDel = new MergeFeaturesLogicallyDelegate(MergeFeaturesLogically);
                _updAttsDel = new UpdateAttributesDelegate(UpdateAttributes);
                _updAttsSelDel = new UpdateAttributesSelectionDelegate(UpdateAttributes);
                _updAttsBulkDel = new UpdateAttributesBulkDelegate(UpdateAttributes);
                _selByQDefDel = new SelectByQueryDefDelegate(SelectByQueryDef);
                _selByQFilterDel = new SelectByQueryFilterDelegate(SelectByQueryFilter);
                _selByJoinDel = new SelectByJoinDelegate(SelectByJoin);
                _zoomSelDel = new ZoomSelectedDelegate(ZoomSelected);
                _zoomSelCursorDel = new ZoomSelectedCursorDelegate(ZoomSelectedCursor);
                _exportDel = new ExportDelegate(Export);
                _isHluWorkspaceDel = new IsHluWorkspaceDelegate(IsHluWorkspace);
                _ListHluLayersDel = new ListHluLayersDelegate(ListHluLayers);
                _isHluLayerDel = new IsHluLayerDelegate(IsHluLayer);
                _isEditingDel = new IsEditingDelegate(IsEditing);
            }
        }

        #endregion

        #region IExtensionConfig Members

        public string Description
        {
            get { return "HLU ArcMap Extension\r\n\r\nProvides interface to HLU GIS Tool."; }
        }

        /// <summary>
        /// Friendly name shown in the Extension dialog
        /// </summary>
        public string ProductName
        {
            get { return "HLU ArcMap Extension"; }
        }

        public esriExtensionState State
        {
            get
            {
                return _enableState;
            }
            set
            {
                if (_enableState != 0 && value == _enableState)
                    return;

                // check if ok to enable or disable extension
                esriExtensionState requestState = value;
                if (requestState == esriExtensionState.esriESEnabled)
                {
                    // cannot enable if it's already in unavailable state
                    if (_enableState == esriExtensionState.esriESUnavailable)
                    {
                        throw new COMException("Cannot enable extension.");
                    }

                    // determine if state can be changed
                    esriExtensionState checkState = StateCheck(true);
                    _enableState = checkState;
                    if (_enableState == esriExtensionState.esriESEnabled)
                    {
                        object initData = _application;
                        Startup(ref initData);
                        OnOpenDocument();
                    }
                }
                else if (requestState == 0 || requestState == esriExtensionState.esriESDisabled)
                {
                    // determine if state can be changed
                    esriExtensionState checkState = StateCheck(false);
                    if (checkState != _enableState)
                    {
                        _enableState = checkState;
                        Shutdown();
                    }
                }
            }
        }

        /// <summary>
        /// Determine extension state 
        /// </summary>
        /// <param name="requestEnable">true if to enable; false to disable</param>
        private esriExtensionState StateCheck(bool requestEnable)
        {
            if (requestEnable)
                return esriExtensionState.esriESEnabled;
            else
                return esriExtensionState.esriESDisabled;
        }

        #endregion

        #region IPersistVariant Members

        public UID ID
        {
            get
            {
                UID typeID = new UIDClass();
                typeID.Value = GetType().GUID.ToString("B");
                return typeID;
            }
        }

        public void Load(IVariantStream Stream)
        {
            Marshal.ReleaseComObject(Stream);
        }

        public void Save(IVariantStream Stream)
        {
            Marshal.ReleaseComObject(Stream);
        }

        #endregion

        #region Setup Events

        private void SetupDocumentEvents(IDocument document)
        {
            _focusMap = ((IMxDocument)_application.Document).FocusMap;
            _docEvents = document as IDocumentEvents_Event;

            _docEvents.NewDocument -= OnNewDocument;
            _docEvents.OpenDocument -= OnOpenDocument;

            _docEvents.NewDocument += new IDocumentEvents_NewDocumentEventHandler(OnNewDocument);
            _docEvents.OpenDocument += new IDocumentEvents_OpenDocumentEventHandler(OnOpenDocument);
        }

        private void SetupActiveViewEvents(IMap map)
        {
            if (map == null) return;

            try
            {
                IActiveViewEvents_Event activeViewEvents = map as IActiveViewEvents_Event;

                _activeViewEventsFocusMapChanged = 
                    new IActiveViewEvents_FocusMapChangedEventHandler(OnActiveViewEventsFocusMapChanged);
                activeViewEvents.FocusMapChanged += _activeViewEventsFocusMapChanged;

                _activeViewEventsItemAdded = new IActiveViewEvents_ItemAddedEventHandler(OnActiveViewEventsItemAdded);
                activeViewEvents.ItemAdded += _activeViewEventsItemAdded;

                _activeViewEventsItemDeleted = new IActiveViewEvents_ItemDeletedEventHandler(OnActiveViewEventsItemDeleted);
                activeViewEvents.ItemDeleted += _activeViewEventsItemDeleted;
            }
            catch { }
        }

        private void RemoveActiveViewEvents(IMap map)
        {
            if (map == null) return;

            try
            {
                IActiveViewEvents_Event activeViewEvents = map as IActiveViewEvents_Event;
                activeViewEvents.FocusMapChanged -= _activeViewEventsFocusMapChanged;
                activeViewEvents.ItemAdded -= _activeViewEventsItemAdded;
                activeViewEvents.ItemDeleted -= _activeViewEventsItemDeleted;
            }
            catch { }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        // Pass the layer to be switched to so that the event handlers
        // can be linked to that layer.
        private void SetupSelectionChangedEvent(IFeatureLayer layer)
        {
            if (layer == null) return;
            try
            {

                // Remove the previous event handler if present
                if (_hluLayerSelectionChangedHandler != null)
                    _hluLayerSelectionChangedHandler.FeatureLayerSelectionChanged -=
                        _hluLayerSelectionEvent_FeatureLayerSelectionChanged;
                
                _hluFeatureSelection = (IFeatureSelection)layer;
                _hluLayerSelectionChangedHandler = (IFeatureLayerSelectionEvents_Event)layer;
                _hluLayerSelectionChangedHandler.FeatureLayerSelectionChanged +=
                    new IFeatureLayerSelectionEvents_FeatureLayerSelectionChangedEventHandler(
                        _hluLayerSelectionEvent_FeatureLayerSelectionChanged);
            }
            catch (Exception ex) { string s = ex.Message; }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Event Handlers

        private void OnNewDocument()
        {
            IDocument document = _docEvents as IDocument;

            _hluWS = null;
            _hluLayer = null;
            _hluFeatureClass = null;
            _hluView = null;
            _hluWSisSDE = false;
            RemoveActiveViewEvents(_focusMap);
            _focusMap = ((IMxDocument)document).FocusMap;
            SetupActiveViewEvents(_focusMap);
        }

        private void OnOpenDocument()
        {
            IDocument document = _docEvents as IDocument;

            _hluWS = null;
            _hluLayer = null;
            _hluFeatureClass = null;
            _hluView = null;
            RemoveActiveViewEvents(_focusMap);
            _focusMap = ((IMxDocument)document).FocusMap;

            if (IsHluWorkspace(((IMxDocument)document).Maps))
                SetupActiveViewEvents(_focusMap);
        }

        private void OnActiveViewEventsFocusMapChanged()
        {
            RemoveActiveViewEvents(_focusMap);
            _focusMap = ((IMxDocument)_application.Document).FocusMap;
            SetupActiveViewEvents(_focusMap);
        }

        private void OnActiveViewEventsItemAdded(object Item)
        {
            if (_hluLayer != null) return;

            IFeatureLayer layer = Item as IFeatureLayer;
            if (layer == null) return;

            if (IsHluLayer(layer))
                _hluView = _focusMap as IActiveView;
        }

        private void OnActiveViewEventsItemDeleted(object Item)
        {
            if (_hluLayer == null) return;

            IFeatureLayer layer = Item as IFeatureLayer;
            if (layer == null) return;

            if (layer.Equals(_hluLayer))
            {
                try
                {
                    _hluLayerSelectionChangedHandler.FeatureLayerSelectionChanged -=
                        _hluLayerSelectionEvent_FeatureLayerSelectionChanged;
                }
                catch { }
                _hluLayer = null;
                _hluFeatureClass = null;
                _hluWS = null;
                _hluView = null;
                _hluFieldMap = null;
                _hluFieldNames = null;
            }
        }

        private void _hluLayerSelectionEvent_FeatureLayerSelectionChanged()
        {
            if (_hluLayer != null)
                _hluFeatureSelection = (IFeatureSelection)_hluLayer;

            if ((_selectFieldOrdinals == null) || (_selectFieldOrdinals.Length == 0)) return;
            
            if (_joinedTable)
            {
                ((IDisplayRelationshipClass)_hluLayer).DisplayRelationshipClass(
                    null, esriJoinType.esriLeftInnerJoin);
                _joinedTable = false;
            }

            PipeSelection(_hluFeatureSelection);

            // flag the new selection
            _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, _hluLayer, _hluView.Extent);
        }

        #endregion

        #region IPC

        public static IChannelManager PipeManager;

        /// <summary>
        /// List of incoming and outgoing IPC messages.
        /// </summary>
        public static List<string> PipeData
        {
            get { return _pipeData; }
            set { _pipeData = value; }
        }

        /// <summary>
        /// Main event handler for incoming messages from DB.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _pipeManager_IncomingDataReady(object sender, EventArgs e)
        {
            if (_pipeData.Count > 0)
            {
                _selectColumns = null;
                _sendColumnHeaders = false;
                string cmd = _pipeData[0];

                if ((_hluLayer == null) && (cmd != "iw") && (cmd != "il") && (cmd != "ll")) _pipeData.Clear();

                switch (cmd)
                {
                    case "qd": // QueryDef: cmd, tables, columns, whereClause, oidColumnAlias, sendColumnHeaders
                        if (_pipeData.Count == 5)
                        {
                            try
                            {
                                IFeatureClass hluDisplayFeatureClass = ((IGeoFeatureLayer)_hluLayer).DisplayFeatureClass;
                                CreateSelectionFieldList(_pipeData[2]);
                                
                                int[] selFieldIxs = _selectFields.Select((c, index) =>
                                    new { Index = index, Ordinal = hluDisplayFeatureClass.FindField(c) })
                                    .Where(a => a.Ordinal != -1).Select(a => a.Index).ToArray();
                                _selectColumns = selFieldIxs.Select(i => _selectColumns[i]).ToArray();
                                _selectFields = selFieldIxs.Select(i => _selectFields[i]).ToArray();
                                _selectFieldOrdinals = _selectFields.Select(c =>
                                    hluDisplayFeatureClass.FindField(c)).ToArray();

                                IQueryDef queryDef = ArcMapAppHelperClass.CreateQueryDef(
                                    _hluWS, _pipeData[1], _pipeData[2], _pipeData[3]);
                                _sendColumnHeaders = _pipeData[5] == "true";
                                _pipeData.Clear();
                                _dummyControl.Invoke(_selByQDefDel, new object[] { queryDef, _pipeData[4] });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "qf": // QueryFilter: cmd, columns, whereClause, sendColumnHeaders 
                        if (_pipeData.Count == 4)
                        {
                            try
                            {
                                CreateSelectionFieldList(_pipeData[1]);
                                IQueryFilter queryFilter = new QueryFilterClass();
                                queryFilter.SubFields = String.Join(",", _selectFields);
                                queryFilter.WhereClause = _pipeData[2];
                                _sendColumnHeaders = _pipeData[3] == "true";
                                _pipeData.Clear();

                                _dummyControl.Invoke(_selByQFilterDel, new object[] { queryFilter });
                            }
                            catch (Exception ex) { PipeException(ex); }
                        }
                        break;
                    case "sj": // select by join: cmd, scratchMdbPath, selectionTableName, subFields, "false"
                        if (_pipeData.Count == 5)
                        {
                            try
                            {
                                CreateSelectionFieldList(_pipeData[3]);
                                _sendColumnHeaders = _pipeData[4] == "true";
                                _dummyControl.Invoke(_selByJoinDel, new object[] { _pipeData[1], _pipeData[2] });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "rs": // read selection: cmd, columns (list)
                        if (_pipeData.Count >= 2)
                        {
                            try
                            {
                                CreateSelectionFieldList(_pipeData.Skip(1).ToArray());

                                _dummyControl.Invoke(_pipeSelDel, new object[] { _hluFeatureSelection });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "us": // update selection: cmd, columns, values, historyColumns [last 3 lists]
                        try
                        {
                            List<int> updateFieldOrdinals;
                            List<object> updateValues;
                            List<string> historyColumnNames;
                            ReadUpdateParameters(1, out updateFieldOrdinals, out updateValues, out historyColumnNames);

                            _dummyControl.Invoke(_updAttsSelDel, new object[] { updateFieldOrdinals, 
                                updateValues, historyColumnNames.ToArray() });
                        }
                        catch { _pipeData.Clear(); }
                        break;
                    case "up": // update: cmd, whereClause, columns, values, historyColumns [last 3 lists]
                        try
                        {
                            IQueryFilter queryFilter = new QueryFilterClass();
                            queryFilter.WhereClause = _pipeData[1];

                            List<int> updateFieldOrdinals;
                            List<object> updateValues;
                            List<string> historyColumnNames;
                            ReadUpdateParameters(2, out updateFieldOrdinals, out updateValues, out historyColumnNames);

                            _dummyControl.Invoke(_updAttsDel, new object[] { queryFilter, 
                                updateFieldOrdinals, updateValues, historyColumnNames.ToArray() });
                        }
                        catch { _pipeData.Clear(); }
                        break;
                    case "ub": // update bulk: cmd, mdbPathName, attributeDatasetName, columns, values, historyColumns [last 3 lists] 
                        try
                        {
                            List<int> updateFieldOrdinals;
                            List<object> updateValues;
                            List<string> historyColumnNames;
                            ReadUpdateParameters(3, out updateFieldOrdinals, out updateValues, out historyColumnNames);

                            _dummyControl.Invoke(_updAttsBulkDel, new object[] {_pipeData[1], _pipeData[2], 
                                updateFieldOrdinals, updateValues, historyColumnNames.ToArray()});
                        }
                        catch { _pipeData.Clear(); }
                        break;
                    case "fl": // flash feature: cmd, whereClause
                        if (_pipeData.Count == 2)
                        {
                            try
                            {
                                IQueryFilter queryFilter = new QueryFilterClass();
                                queryFilter.WhereClause = _pipeData[1];
                                _dummyControl.Invoke(_flashSelFeatDel, new object[] { queryFilter });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "mg": // merge features: cmd, resultWhereClause, newToidFragmentID, historyColumns
                        if (_pipeData.Count == 4)
                        {
                            try
                            {
                                IQueryFilter resultFeatureQueryFilter = new QueryFilterClass();
                                resultFeatureQueryFilter.WhereClause = _pipeData[1];

                                _dummyControl.Invoke(_mergeFeatDel, new object[] { resultFeatureQueryFilter, 
                                    _pipeData[2], _pipeData[3].Split(',') });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "ml": // merge features logically: cmd, keepIncid, historyColumns, updateTable (as List<string>)
                        if (_pipeData.Count == 3) // > 3)
                        {
                            try
                            {
                                string keepIncid = _pipeData[1];
                                string[] historyColumns = _pipeData[2].Split(',');

                                _dummyControl.Invoke(_mergeFeatLogDel, new object[] { keepIncid, historyColumns });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "sp": // split feature: cmd, whereClause, lastToidFragmentID, historyColumns
                        if (_pipeData.Count == 4)
                        {
                            try
                            {
                                IQueryFilter queryFilter = new QueryFilterClass();
                                queryFilter.WhereClause = _pipeData[1];

                                _dummyControl.Invoke(_splitFeatDel, new object[] { queryFilter, 
                                    _pipeData[2], _pipeData[3].Split(',') });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // The old incid number is passed together with the new incid
                    // number so that only features belonging to the old incid are
                    // updated.
                    case "sl": // split features logically: cmd, newIncid, historyColumns
                        if (_pipeData.Count == 4)
                        {
                            try
                            {
                                _dummyControl.Invoke(_splitFeatLogDel, new object[] { _pipeData[1], 
                                    _pipeData[2], _pipeData[3].Split(',') });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    //---------------------------------------------------------------------
                    case "zs": // zoom selected: cmd [, queryFilter]
                        try
                        {
                            switch (_pipeData.Count)
                            {
                                case 1:
                                    _pipeData.Clear();
                                    _dummyControl.Invoke(_zoomSelDel, null);
                                    break;
                                case 2:
                                    IQueryFilter queryFilter = new QueryFilterClass();
                                    queryFilter.WhereClause = _pipeData[1];
                                    _pipeData.Clear();
                                    _dummyControl.Invoke(_zoomSelCursorDel, new object[] { queryFilter });
                                    break;
                            }
                        }
                        catch { _pipeData.Clear(); }
                        break;
                    case "ex": // export: cmd, mdbPathName, attributeDatasetName, exportRowCount
                        if (_pipeData.Count == 4)
                        {
                            try
                            {
                                string mdbPathName = _pipeData[1];
                                string attributeDatasetName = _pipeData[2];
                                int exportRowCount = Int32.Parse(_pipeData[3]);
                                _pipeData.Clear();

                                _dummyControl.Invoke(_exportDel,
                                    new object[] { mdbPathName, attributeDatasetName, exportRowCount });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "iw": // IsHluWorkspace: cmd
                        if (_pipeData.Count == 1)
                        {
                            try { _dummyControl.Invoke(_isHluWorkspaceDel, null); }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "ll": // ListHLULayers: cmd
                        if (_pipeData.Count == 1)
                        {
                            try { _dummyControl.Invoke(_ListHluLayersDel, null); }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "il": // IsHluLayer: cmd, ixMap, ixLayer
                        if (_pipeData.Count == 3)
                        {
                            try
                            {
                                int ixMap = Int32.Parse(_pipeData[1]);
                                int ixLayer = Int32.Parse(_pipeData[2]);
                                _pipeData.Clear();
                                _dummyControl.Invoke(_isHluLayerDel, new object[] { ixMap, ixLayer });
                            }
                            catch { _pipeData.Clear(); }
                        }
                        break;
                    case "ie": // IsEditing: cmd
                        try { _dummyControl.Invoke(_isEditingDel, null); }
                        catch { _pipeData.Clear(); }
                        break;
                    default:
                        _pipeData.Clear();
                        break;
                }
            }

            // notify PipeManager of data ready to be sent
            OutgoingDataReady.Invoke(null, null);
        }

        private void PipeException(Exception ex)
        {
            _pipeData.Clear();
            _pipeData.Add(_pipeErrorSymbol.ToString());
            _pipeData.Add(ex.Message);
        }

        private void ReadUpdateParameters(int startAt, out List<int> updateFieldOrdinals,
            out List<object> updateValues, out List<string> historyColumnNames)
        {
            updateFieldOrdinals = new List<int>();
            updateValues = new List<object>();
            historyColumnNames = new List<string>();

            int i = startAt;
            string s;
            int o;
            while ((i < _pipeData.Count) && ((s = _pipeData[i++]) != _pipeTransmissionInterrupt))
                if ((o = MapField(s)) != -1) updateFieldOrdinals.Add(o);

            while ((i < _pipeData.Count - 1) && ((s = _pipeData[i++]) != _pipeTransmissionInterrupt))
                updateValues.Add(s);

            while (i < _pipeData.Count)
                historyColumnNames.Add(_pipeData[i++]);
        }

        #endregion

        #region Select

        private void SelectByQueryDef(IQueryDef queryDef, string oidColumnAlias)
        {
            // Evaluate queryDef to execute a database query and return a cursor.
            ICursor selCursor = queryDef.Evaluate();

            // check if the cursors contains the OID field
            int oidOrdinalCursor = selCursor.FindField(oidColumnAlias);

            // if OID field present select features on map
            if (oidOrdinalCursor != -1)
            {
                IQueryFilter queryFilter = new QueryFilterClass();
                queryFilter.WhereClause = ArcMapAppHelperClass.WhereClauseFromCursor(oidOrdinalCursor, oidColumnAlias, selCursor);
                SelectByQueryFilter(queryFilter);
            }
        }

        private void SelectByQueryFilter(IQueryFilter queryFilter)
        {
            _hluFeatureSelection = (IFeatureSelection)_hluLayer; // _hluLayer.DisplayFeatureClass ??

            // invalidate only the selection cache. Flag the original selection
            _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
            
            // perform selection
            _hluFeatureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
        }

        private void SelectByJoin(string scratchMdbPath, string selectionDatasetName)
        {
            bool restoreEditSession = InEditingSession && HluLayerBeingEdited;
            if (restoreEditSession) CloseEditSession(true);

            try
            {
                SetCursor(true);

                IDataset joinDataset;
                ITable joinTable;
                OpenOleDbWorkspace(scratchMdbPath, selectionDatasetName, out joinDataset, out joinTable);

                if ((joinTable.Fields.FieldCount == 1) && (joinTable.Fields.get_Field(0).Name ==
                    _hluLayerStructure.incidColumn.ColumnName)) // single column incid: use a join
                {
                    IDisplayTable hluDisplayTable = (IDisplayTable)_hluLayer;
                    IFeatureClass hluDisplayTableFeatureClass = (IFeatureClass)hluDisplayTable.DisplayTable;
                    ITable hluLayerTable = (ITable)hluDisplayTableFeatureClass;

                    string originPKJoinField = _hluLayerStructure.incidColumn.ColumnName;
                    string originFKJoinField =
                        _hluFeatureClass.Fields.get_Field(_hluFieldMap[_hluLayerStructure.incidColumn.Ordinal]).Name;

                    // create virtual relate
                    IMemoryRelationshipClassFactory memoryRelFactory = new MemoryRelationshipClassFactoryClass();
                    IRelationshipClass relClass = memoryRelFactory.Open("SelRelClass", (IObjectClass)joinTable,
                        originPKJoinField, (IObjectClass)hluLayerTable, originFKJoinField, "forward", "backward",
                        esriRelCardinality.esriRelCardinalityOneToMany);

                    // use Relate to perform a join
                    IDisplayRelationshipClass displayRelClass = (IDisplayRelationshipClass)_hluLayer;
                    displayRelClass.DisplayRelationshipClass(relClass, esriJoinType.esriLeftOuterJoin);

                    _joinedTable = true;

                    Regex fldNameRegex = new Regex(@"((""|\[)*" + selectionDatasetName +
                        @"(""|\[)*\.)*(""|\])*" + originPKJoinField + @"(""|\])*", RegexOptions.IgnoreCase);

                    string joinExpr = originPKJoinField;
                    string s = null;
                    for (int i = hluDisplayTable.DisplayTable.Fields.FieldCount - 1; i > 0; i--)
                    {
                        s = hluDisplayTable.DisplayTable.Fields.get_Field(i).Name;
                        if (fldNameRegex.IsMatch(s))
                        {
                            joinExpr = s;
                            break;
                        }
                    }

                    IQueryFilter queryFilter = new QueryFilterClass();
                    queryFilter.WhereClause = String.Format("{0} IS NOT NULL", joinExpr);

                    // invalidate only the selection cache. Flag the original selection
                    _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);

                    // perform selection
                    ISelectionSet selSet = hluDisplayTable.SelectDisplayTable(queryFilter,
                        esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, null);

                    _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    _hluFeatureSelection.SelectionSet = selSet;
                }
                else // multi-column join: use cumulative selection sets
                {
                    SelectByJoinMultiColumn(joinTable);
                }
            }
            catch { }
            finally
            {
                if (restoreEditSession) OpenEditSession();
                SetCursor(false);
            }
        }

        private void SelectByJoinMultiColumn(ITable joinTable)
        {
            ICursor joinCursor = joinTable.Search(null, true);
            IRow joinRow;
            List<string> whereCondList = new List<string>();
            StringBuilder whereCond = new StringBuilder();
            int whereClauseLength = 0;
            string quotedHluTableName = QuoteIdentifier(_hluTableName);

            while ((joinRow = joinCursor.NextRow()) != null)
            {
                StringBuilder cond = new StringBuilder();
                for (int i = 0; i < joinTable.Fields.FieldCount; i++)
                {
                    IField currFld = joinTable.Fields.get_Field(i);
                    cond.Append(String.Format(" AND {0} = {1}",
                        _hluSqlSyntax.QualifyColumnName(quotedHluTableName, 
                        QuoteIdentifier(_hluFeatureClass.Fields.get_Field(
                        _hluFieldMap[_hluLayerStructure.Columns[currFld.Name].Ordinal]).Name)),
                        QuoteValue(currFld, joinRow.get_Value(i))));
                }
                cond.Remove(0, 5).Insert(0, "(").Append(")");
                whereClauseLength += cond.Length + 4;

                if (whereClauseLength < _whereClauseLengthMax)
                {
                    whereCond.Append(" OR ").Append(cond);
                }
                else
                {
                    whereCondList.Add(whereCond.Remove(0, 4).ToString());
                    whereCond = new StringBuilder(" OR ").Append(cond);
                    whereClauseLength = whereCond.Length;
                }
            }

            if (whereCond.Length > 0) whereCondList.Add(whereCond.Remove(0, 4).ToString());
 
            // make sure selection changed event handler won't intervene
            int[] selectFieldOrdinalsBak = _selectFieldOrdinals;
            _selectFieldOrdinals = null;

            IQueryFilter queryFilter = new QueryFilterClass();

            _hluFeatureSelection = (IFeatureSelection)_hluLayer;
            _hluFeatureSelection.Clear();

            for (int i = 0; i < whereCondList.Count - 1; i++)
            {
                queryFilter.WhereClause = whereCondList[i];
                _hluFeatureSelection.SelectFeatures(queryFilter,
                    esriSelectionResultEnum.esriSelectionResultAdd, false);
            }

            // before last selection reenable event handler
            _selectFieldOrdinals = selectFieldOrdinalsBak;
            queryFilter.WhereClause = whereCondList[whereCondList.Count - 1];
            _hluFeatureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultAdd, false);
        }

        private string QuoteValue(IField field, object value)
        {
            switch (field.Type)
            {
                case esriFieldType.esriFieldTypeString:
                case esriFieldType.esriFieldTypeXML:
                case esriFieldType.esriFieldTypeGUID:
                case esriFieldType.esriFieldTypeDate:
                    return "'" + value.ToString() + "'";
                default:
                    return value.ToString();
            }
        }

        private string QuoteIdentifier(string identifier)
        {
            if (!identifier.StartsWith(_quotePrefix)) identifier = _quotePrefix + identifier;
            if (!identifier.EndsWith(_quoteSuffix)) identifier += _quoteSuffix;
            return identifier;
        }

        private void PipeSelection(IFeatureSelection featureSelection)
        {
            try
            {
                if ((_selectFieldOrdinals == null) || (_selectFieldOrdinals.Length == 0)) return;

                if (_hluFeatureSelection == null)
                    _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                if (_hluFeatureSelection.SelectionSet.Count == 0)
                {
                    _pipeData.Clear();
                    return;
                }

                if (_pipeData == null) _pipeData = new List<string>();

                lock (_pipeData)
                {
                    _pipeData.Clear();
                    ICursor resultCursor;
                    IRow selectRow;
                    featureSelection.SelectionSet.Search(null, true, out resultCursor);

                    if (_sendColumnHeaders && (_selectColumns != null))
                        _pipeData.Add(String.Join(",", _selectColumns));

                    while ((selectRow = resultCursor.NextRow()) != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        for (int i = 0; i < _selectFieldOrdinals.Length; i++)
                            sb.Append(String.Format("{0}{1}", _pipeFieldDelimiter,
                                selectRow.get_Value(_selectFieldOrdinals[i])));
                        _pipeData.Add(sb.Remove(0, 1).ToString());
                    }

                    Marshal.FinalReleaseComObject(resultCursor);
                }
            }
            catch (Exception ex)
            {
                _pipeData.Clear();
                _pipeData.Add(_pipeErrorSymbol.ToString());
                _pipeData.Add(ex.Message);
            }
            finally { _selectFieldOrdinals = null; }
        }

        private void CreateSelectionFieldList(string columnList)
        {
            if (!String.IsNullOrEmpty(columnList))
                CreateSelectionFieldList(columnList.Split(','));
        }

        private void CreateSelectionFieldList(string[] columns)
        {
            if (columns != null)
            {
                try
                {
                    _selectColumns = (from fieldName in columns
                                      let columnOrdinal = ColumnOrdinal(fieldName)
                                      where columnOrdinal != -1
                                      select _hluLayerStructure.Columns[columnOrdinal].ColumnName).ToArray();
                }
                catch { }

                if ((_selectColumns == null) || (_selectColumns.Count() != columns.Length))
                {
                    _selectColumns = (from c in columns
                                      let colName = c.Trim()
                                      where _hluLayerStructure.Columns.Contains(colName)
                                      select colName).ToArray();
                }

                _selectFieldOrdinals = _selectColumns.Select(c => FieldOrdinal(c)).ToArray();

                _selectFields = _selectFieldOrdinals.Select((f, index) => _hluFieldNames[index]).ToArray();
            }
        }

        #endregion

        #region View

        private void ZoomSelectedCursor(IQueryFilter queryFilter)
        {
            if ((queryFilter == null) || (_hluFeatureClass == null) || (_hluView == null)) return;
            IEnumGeometryBind enumGeometryBind = new EnumFeatureGeometryClass();
            enumGeometryBind.BindGeometrySource(queryFilter, _hluFeatureClass);
            IGeometryFactory geometryFactory = new GeometryEnvironmentClass();
            IGeometry geom = geometryFactory.CreateGeometryFromEnumerator((IEnumGeometry)enumGeometryBind);
            _hluView.Extent = geom.Envelope;
            _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeography, _hluLayer, _hluView.Extent);
        }

        private void ZoomSelected()
        {
            if ((_hluFeatureClass == null) || (_hluView == null)) return;

            if (_hluFeatureSelection == null)
                _hluFeatureSelection = (IFeatureSelection)_hluFeatureClass;

            if (_hluFeatureSelection.SelectionSet.Count == 0) return;

            IEnumGeometryBind enumGeometryBind = new EnumFeatureGeometryClass();
            enumGeometryBind.BindGeometrySource(null, _hluFeatureSelection.SelectionSet);
            IGeometryFactory geometryFactory = new GeometryEnvironmentClass();
            IGeometry geom = geometryFactory.CreateGeometryFromEnumerator((IEnumGeometry)enumGeometryBind);
            _hluView.Extent = geom.Envelope;
            _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeography, _hluLayer, _hluView.Extent);
        }

        private void FlashFeature(IQueryFilter queryFilter)
        {
            if (_pipeData == null)
                _pipeData = new List<string>();
            else
                _pipeData.Clear();

            try
            {
                //---------------------------------------------------------------------
                // CHANGED: CR23 (Merged features)
                // Flash all the features relating to the selected incid at once.
                // This method may be triggered more than once if there are too
                // many to pass via the Named Pipes in one go.
                //
                IEnumGeometryBind enumGeometryBind = new EnumFeatureGeometryClass();
                enumGeometryBind.BindGeometrySource(null, _hluFeatureSelection.SelectionSet.Select(queryFilter, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, null));
                IGeometryFactory geometryFactory = new GeometryEnvironmentClass();
                IGeometry geom = geometryFactory.CreateGeometryFromEnumerator((IEnumGeometry)enumGeometryBind);

                IMxDocument mxDoc = (IMxDocument)_application.Document;
                IActiveView activeView = mxDoc.FocusMap as IActiveView;
                IScreenDisplay screenDisplay = activeView.ScreenDisplay;
                //---------------------------------------------------------------------
                //---------------------------------------------------------------------
                // FIX: 018 Bring ArcGIS and MapInfo into line by flashing all features twice
                FlashGeometry(geom, screenDisplay, 300, 2);
                //---------------------------------------------------------------------
            }
            catch { }
        }


        ///<summary>Flash geometry on the display.</summary>
        ///<param name="geometry"> The input IGeometry to flash.  Supported geometry types are GeometryBag, Polygon, Polyline, Point and Multipoint.</param>
        ///<param name="screenDisplay">An IScreenDisplay reference</param>
        ///<param name="delay">An integer that is the time in milliseconds to wait.</param>
        public static void FlashGeometry(IGeometry geometry, IScreenDisplay screenDisplay, int delay, int times)
        {
            if (geometry == null || screenDisplay == null)
            {
                return;
            }
            bool continueFlashing = true;

            using (ComReleaser comReleaser = new ComReleaser())
            {
                ITrackCancel cancelTracker = new CancelTrackerClass();
                comReleaser.ManageLifetime(cancelTracker);
                screenDisplay.CancelTracker = cancelTracker;
                short cacheID = screenDisplay.AddCache();
                int cacheMemDC = screenDisplay.get_CacheMemDC(cacheID);
                
                IRgbColor fillColor = new RgbColorClass();
                comReleaser.ManageLifetime(fillColor);
                fillColor.Green = 128;
                IRgbColor lineColor = new RgbColorClass();
                comReleaser.ManageLifetime(lineColor);

                screenDisplay.StartDrawing(cacheMemDC, cacheID);
                DrawGeometry(geometry, fillColor, lineColor, (IDisplay)screenDisplay, cancelTracker);
                ESRI.ArcGIS.esriSystem.tagRECT RECT = new tagRECT();
                screenDisplay.FinishDrawing();

                for (int j = 0; j < times; j++)
                {
                    if (continueFlashing == true)
                    {
                        screenDisplay.DrawCache(screenDisplay.hDC, cacheID, ref RECT, ref RECT);
                        if (delay > 0)
                        {
                            System.Threading.Thread.Sleep(delay);
                            screenDisplay.Invalidate(null, true, cacheID);
                            screenDisplay.UpdateWindow();
                            System.Threading.Thread.Sleep(delay);
                        }
                    }
                }
                //---------------------------------------------------------------------

                screenDisplay.RemoveCache(cacheID);
                cancelTracker.Reset();
            }
        }

        /// <summary>
        /// Draws the input geometry using the specified colors.
        /// </summary>
        /// <param name="geometry">The input IGeometry to draw. Supported geometry types are GeometryBag, Polygon, Polyline, Point and Multipoint.</param>
        /// <param name="fillColor">An IRgbColor reference for the fill color</param>
        /// <param name="lineColor">An IRgbColor reference for the line or outline color</param>
        /// <param name="display">An IDisplay reference</param>
        /// <param name="cancelTracker">An ITrackCancel reference</param>
        private static void DrawGeometry(IGeometry geometry, IRgbColor fillColor, IRgbColor lineColor, IDisplay display, ITrackCancel cancelTracker)
        {
            bool continueDrawing = true;
            switch (geometry.GeometryType)
            {
                case esriGeometryType.esriGeometryBag:
                    {
                        IEnumGeometry enumGeometry = (IEnumGeometry)geometry;
                        IGeometry innerGeometry = enumGeometry.Next();
                        while (innerGeometry != null && continueDrawing == true)
                        {
                            DrawGeometry(innerGeometry, fillColor, lineColor, display, cancelTracker); // Recursive method call
                            innerGeometry = enumGeometry.Next();
                            if (cancelTracker != null)
                            {
                                continueDrawing = cancelTracker.Continue();
                            }
                        }
                        break;
                    }
                case esriGeometryType.esriGeometryPolygon:
                    {
                        // Set the input polygon geometry's symbol.
                        ISimpleFillSymbol fillSymbol = new SimpleFillSymbolClass();
                        fillSymbol.Color = (IColor)fillColor;
                        ILineSymbol lineSymbol = new SimpleLineSymbolClass();
                        lineSymbol.Color = lineColor;
                        fillSymbol.Outline = lineSymbol;

                        // Draw the input polygon geometry.
                        display.SetSymbol((ISymbol)fillSymbol);
                        display.DrawPolygon(geometry);
                        break;
                    }
                case esriGeometryType.esriGeometryPolyline:
                    {
                        // Set the input polyline geometry's symbol.
                        IMultiLayerLineSymbol multiLineSymbol = new MultiLayerLineSymbolClass();
                        ISimpleLineSymbol simpleLineSymbol1 = new SimpleLineSymbolClass();
                        ISimpleLineSymbol simpleLineSymbol2 = new SimpleLineSymbolClass();
                        simpleLineSymbol1.Width = 3;
                        simpleLineSymbol1.Color = fillColor;
                        simpleLineSymbol2.Width = 5;
                        simpleLineSymbol2.Color = lineColor;
                        multiLineSymbol.AddLayer((ILineSymbol)simpleLineSymbol2);
                        multiLineSymbol.AddLayer((ILineSymbol)simpleLineSymbol1);

                        // Draw the input polyline geometry.
                        display.SetSymbol((ISymbol)multiLineSymbol);
                        display.DrawPolyline(geometry);
                        break;
                    }
                case esriGeometryType.esriGeometryPoint:
                    {
                        // Set the input point geometry's symbol.
                        ISimpleMarkerSymbol simpleMarkerSymbol = new SimpleMarkerSymbolClass();
                        simpleMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCircle;
                        simpleMarkerSymbol.Size = 12;
                        simpleMarkerSymbol.Color = fillColor;
                        simpleMarkerSymbol.Outline = true;
                        simpleMarkerSymbol.OutlineColor = lineColor;

                        // Draw the input point geometry.
                        display.SetSymbol((ISymbol)simpleMarkerSymbol);
                        display.DrawPoint(geometry);
                        break;
                    }
                case esriGeometryType.esriGeometryMultipoint:
                    {
                        // Set the input multipoint geometry's symbol.
                        ISimpleMarkerSymbol simpleMarkerSymbol = new SimpleMarkerSymbolClass();
                        simpleMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCircle;
                        simpleMarkerSymbol.Size = 8;
                        simpleMarkerSymbol.Color = fillColor;
                        simpleMarkerSymbol.Outline = true;
                        simpleMarkerSymbol.OutlineColor = lineColor;

                        // Draw the input multipoint geometry.
                        display.SetSymbol((ISymbol)simpleMarkerSymbol);
                        display.DrawMultipoint(geometry);
                        break;
                    }
            }
        }

        #endregion

        #region Split

        private void SplitFeature(IQueryFilter selectionQueryFilter, 
            string lastToidFragmentID, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // check if HLU layer is being edited
                    if (!HluLayerBeingEdited) return;

                    // make sure at least two features are selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count < 2) return;

                    // write history field names and types to pipe
                    int[] historyFieldOrdinals = HistorySchema(historyColumns);

                    // get the lowest OID in selection set (using query filter for safety)
                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(selectionQueryFilter, true, out cursor);
                    ITableSort tableSort = new TableSortClass();
                    tableSort.Table = (ITable)_hluFeatureClass;
                    tableSort.Cursor = cursor;
                    tableSort.Fields = _hluFeatureClass.OIDFieldName;
                    tableSort.set_Ascending(_hluFeatureClass.OIDFieldName, true);
                    tableSort.Sort(null);
                    ICursor sortCursor = tableSort.Rows;
                    IRow row = sortCursor.NextRow();
                    int minOID = row.OID;

                    FlushCursor(false, ref cursor);

                    // start workspace edit session and operation
                    StartEditing();

                    try
                    {
                        // get features to be updated
                        IFeatureCursor updateCursor = _hluFeatureClass.Update(selectionQueryFilter, false);

                        string numFormat = String.Format("D{0}", lastToidFragmentID.Length);
                        int newToidFragmentIDnum = Int32.Parse(lastToidFragmentID);

                        int toidFragOrdinal = _hluFieldMap[_hluLayerStructure.toid_fragment_idColumn.Ordinal];

                        // temporary history list, sorted before adding to _pipeData
                        List<string> historyList = new List<string>();
                        int originalFeatureIx = -1;
                        int i = 0;
                        IFeature updateFeature;

                        //---------------------------------------------------------------------
                        // FIXED: KI106 (Shape area and length values)
                        // Includes updates for the geom1 and geom2 columns as the features
                        // have changed in size

                        // Check if the current layer is a shapefile
                        bool isShp = IsShp(_hluWS as IWorkspace);
                        //---------------------------------------------------------------------

                        while ((updateFeature = updateCursor.NextFeature()) != null)
                        {
                            if (updateFeature.OID != minOID)
                            {
                                // Set the toid_fragment_id to the next available number
                                updateFeature.set_Value(toidFragOrdinal, (++newToidFragmentIDnum).ToString(numFormat));
                            }
                            else
                            {
                                // Store the original feature number
                                originalFeatureIx = i;
                            }

                            //---------------------------------------------------------------------
                            // FIXED: KI106 (Shape area and length values)
                            // Includes updates for the geom1 and geom2 columns as the features
                            // have changed in size

                            // If it is a shapefile then update the geometry fields
                            if (isShp)
                            {
                                double geom1;
                                double geom2;
                                GetGeometryProperties(updateFeature, out geom1, out geom2);

                                int ixGeom1 = updateFeature.Fields.FindField("shape_leng");
                                int ixGeom2 = updateFeature.Fields.FindField("shape_area");
                                updateFeature.set_Value(ixGeom1, geom1);
                                updateFeature.set_Value(ixGeom2, geom2);
                            }

                            // Update the feature
                            updateCursor.UpdateFeature(updateFeature);
                            //updateFeature.Store();
                            //---------------------------------------------------------------------

                            historyList.Add(History(updateFeature, historyFieldOrdinals, null));
                            i++;
                        }

                        // make sure original feature is on top of the list
                        _pipeData.Add(historyList[originalFeatureIx]);
                        _pipeData.AddRange(historyList.Where((h, index) => index != originalFeatureIx));

                        FlushCursor(false, ref updateCursor);

                        // stop edit operation and session
                        CommitEdits();

                        // redraw selected HLU features
                        _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, _hluLayer, _hluView.Extent);
                    }
                    catch
                    {
                        DiscardEdits();
                        throw;
                    }
                }
            }
            catch { _pipeData.Clear(); }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // The old incid number is passed together with the new incid
        // number so that only features belonging to the old incid are
        // updated.
        private void SplitFeaturesLogically(string oldIncid, string newIncid, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // make sure at least one feature is selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count < 1) return;

                    // write history field names and types to pipe
                    int[] historyFieldOrdinals = HistorySchema(historyColumns);

                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(null, false, out cursor);
                    IFeatureCursor featCursor = (IFeatureCursor)cursor;
                    IFeature splitFeature = featCursor.NextFeature();

                    if (splitFeature == null) return;

                    // start workspace edit session and operation
                    StartEditing();

                    try
                    {
                        string numFormat = String.Format("D{0}", _hluLayerStructure.toid_fragment_idColumn.MaxLength);
                        int incidOrdinal = _hluFieldMap[_hluLayerStructure.incidColumn.Ordinal];
                        int fragOrdinal = _hluFieldMap[_hluLayerStructure.toid_fragment_idColumn.Ordinal];

                        var q = historyColumns.Where(n => _hluLayerStructure.Columns.Cast<DataColumn>()
                            .Count(c => c.ColumnName == n) == 0);
                        string newToidFragmentColumnName = q.Count() == 1 ? q.ElementAt(0) : null;

                        while (splitFeature != null)
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR10 (Attribute updates for incid subsets)
                            // Only collect the history details and update the incid number if
                            // the each feature belongs to the old incid.
                            if (splitFeature.get_Value(incidOrdinal).ToString() == oldIncid)
                            {
                                if (!String.IsNullOrEmpty(newToidFragmentColumnName))
                                    _pipeData.Add(History(splitFeature, historyFieldOrdinals,
                                        new string[] { splitFeature.get_Value(fragOrdinal).ToString() }));
                                splitFeature.set_Value(incidOrdinal, newIncid);
                                splitFeature.Store();
                            }
                            //---------------------------------------------------------------------
                            splitFeature = featCursor.NextFeature();
                        }
                        FlushCursor(false, ref featCursor);

                        // stop edit operation and session
                        CommitEdits();

                        FlushCursor(true, ref cursor);
                    }
                    catch
                    {
                        DiscardEdits();
                        throw;
                    }
                }
            }
            catch { _pipeData.Clear(); }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Merge

        /// <summary>
        /// Merges the geometries of features in the HLU layer's current selection set into one result feature
        /// located through the queryFilter passed in and deletes the other features from the HLU feature class.
        /// </summary>
        /// <param name="resultFeatureQueryFilter">Query filter that identifies the result feature in the selection set.</param>
        /// <param name="mergeFeaturesQueryFilter">Query filter that identifies the features in the selection set 
        /// that are to be merged to the result feature.</param>
        private void MergeFeatures(IQueryFilter resultFeatureQueryFilter,
            string newToidFragmentID, string[] historyColumns) 
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // make sure at least two features are selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count < 2) return;

                    // write history field names and types to pipe
                    int[] historyFieldOrdinals = HistorySchema(historyColumns);

                    // get OID of result feature
                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(resultFeatureQueryFilter, true, out cursor);
                    IFeatureCursor featCursor = (IFeatureCursor)cursor;
                    IFeature resultFeature = featCursor.NextFeature();

                    if (resultFeature != null)
                    {
                        IGeometry resultGeom = resultFeature.ShapeCopy;

                        Marshal.ReleaseComObject(featCursor);
                        Marshal.ReleaseComObject(cursor);

                        // start workspace edit session and operation
                        StartEditing();

                        try
                        {
                            // get features to be merged and deleted
                            IQueryFilter mergeFeaturesQueryFilter = new QueryFilterClass();
                            mergeFeaturesQueryFilter.WhereClause = String.Format("{0} <> {1}", 
                                _hluFeatureClass.OIDFieldName, resultFeature.OID);
                            _hluFeatureSelection.SelectionSet.Search(mergeFeaturesQueryFilter,  false, out cursor);
                            IFeatureCursor updateCursor = (IFeatureCursor)cursor;

                            // merge geometries and delete features
                            IFeature mergeFeature;
                            while ((mergeFeature = updateCursor.NextFeature()) != null)
                            {
                                _pipeData.Add(History(mergeFeature, historyFieldOrdinals, null));
                                resultGeom = ((ITopologicalOperator4)resultGeom).Union(mergeFeature.Shape);
                                mergeFeature.Delete();
                            }
                            FlushCursor(false, ref updateCursor);

                            // get the result feature for update
                            resultFeatureQueryFilter.WhereClause = String.Format("{0} = {1}",
                                _hluFeatureClass.OIDFieldName, resultFeature.OID);
                            updateCursor = _hluFeatureClass.Update(resultFeatureQueryFilter, false);

                            resultFeature = updateCursor.NextFeature();

                            // Update the shape to the new merged geometry
                            resultFeature.Shape = resultGeom;

                            // Set the toid_fragment_id to the same value (passed to this function) for all fragments
                            resultFeature.set_Value(
                                _hluFieldMap[_hluLayerStructure.toid_fragment_idColumn.Ordinal], newToidFragmentID);

                            //---------------------------------------------------------------------
                            // FIXED: KI106 (Shape area and length values)
                            // Includes updates for the geom1 and geom2 columns as the features
                            // have changed in size

                            // Check if the current layer is a shapefile
                            bool isShp = IsShp(_hluWS as IWorkspace);

                            // If it is a shapefile then update the geometry fields
                            if (isShp)
                            {
                                double geom1;
                                double geom2;
                                GetGeometryProperties(resultFeature, out geom1, out geom2);

                                int ixGeom1 = resultFeature.Fields.FindField("shape_leng");
                                int ixGeom2 = resultFeature.Fields.FindField("shape_area");
                                resultFeature.set_Value(ixGeom1, geom1);
                                resultFeature.set_Value(ixGeom2, geom2);
                            }

                            // Update the feature
                            updateCursor.UpdateFeature(resultFeature);
                            //resultFeature.Store();
                            //---------------------------------------------------------------------

                            // add result feature to history (for updating incid_mm_polygon geometry data)
                            _pipeData.Add(History(resultFeature, historyFieldOrdinals, null));

                            FlushCursor(false, ref updateCursor);
                            FlushCursor(true, ref cursor);

                            // stop edit operation and session
                            CommitEdits();

                            // redraw HLU layer within current map extent
                            _hluView.PartialRefresh(esriViewDrawPhase.esriViewGeography, _hluLayer, _hluView.Extent);
                        }
                        catch
                        {
                            DiscardEdits();
                            throw;
                        }
                    }
                }
            }
            catch { _pipeData.Clear(); }
        }

        private void MergeFeaturesLogically(string keepIncid, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // make sure at least two features are selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count < 2) return;

                    // write history field names and types to pipe
                    int[] historyFieldOrdinals = HistorySchema(historyColumns);

                    IField incidField = _hluFeatureClass.Fields.get_Field(
                        _hluFieldMap[_hluLayerStructure.incidColumn.Ordinal]);
                    IQueryFilter mergeQueryFilter = new QueryFilterClass();
                    mergeQueryFilter.WhereClause = String.Format("{0} = {1}",
                        incidField.Name, QuoteValue(incidField, keepIncid));
                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(mergeQueryFilter, false, out cursor);
                    IFeatureCursor featCursor = (IFeatureCursor)cursor;
                    IFeature keepFeature = featCursor.NextFeature();
                    FlushCursor(false, ref featCursor);
                    FlushCursor(true, ref cursor);

                    mergeQueryFilter.WhereClause = String.Format("{0} <> {1}",
                        incidField.Name, QuoteValue(incidField, keepIncid));
                    _hluFeatureSelection.SelectionSet.Search(mergeQueryFilter, false, out cursor);
                    featCursor = (IFeatureCursor)cursor;
                    IFeature mergeFeature = featCursor.NextFeature();

                    if (mergeFeature == null) return;

                    // start workspace edit session and operation
                    StartEditing();

                    try
                    {
                        string numFormat = String.Format("D{0}", _hluLayerStructure.toid_fragment_idColumn.MaxLength);
                        int incidOrdinal = _hluFieldMap[_hluLayerStructure.incidColumn.Ordinal];
                        int toidOrdinal = _hluFieldMap[_hluLayerStructure.toidColumn.Ordinal];
                        int fragOrdinal = _hluFieldMap[_hluLayerStructure.toid_fragment_idColumn.Ordinal];

                        List<int> updateColumns = new List<int>();
                        for (int i = 0; i < mergeFeature.Fields.FieldCount; i++)
                        {
                            IField f = mergeFeature.Fields.get_Field(i);
                            if ((i != incidOrdinal) && (i != toidOrdinal) && (i != fragOrdinal) && f.Editable &&
                                (f.Type != esriFieldType.esriFieldTypeOID) && (f.Type != esriFieldType.esriFieldTypeGeometry))
                                updateColumns.Add(i);
                        }

                        while (mergeFeature != null)
                        {
                            _pipeData.Add(History(mergeFeature, historyFieldOrdinals, null));
                            mergeFeature.set_Value(incidOrdinal, keepIncid);

                            for (int i = 0; i < updateColumns.Count; i++)
                                mergeFeature.set_Value(updateColumns[i], keepFeature.get_Value(updateColumns[i]));
                            
                            mergeFeature.Store();
                            mergeFeature = featCursor.NextFeature();
                        }
                        FlushCursor(false, ref featCursor);

                        // stop edit operation and session
                        CommitEdits();
                        FlushCursor(true, ref cursor);
                    }
                    catch
                    {
                        DiscardEdits();
                        throw;
                    }
                }
            }
            catch { _pipeData.Clear(); }
        }

        #endregion

        #region History

        private int[] HistorySchema(string[] historyColumns)
        {
            int ix;
            var historyFields = from c in historyColumns
                                let ordinal = (ix = MapField(c)) != -1 ? ix : FuzzyFieldOrdinal(c)
                                where ordinal != -1
                                select new
                                {
                                    FieldOrdinal = ordinal,
                                    FieldName = c.Replace(GISApp.HistoryAdditionalFieldsDelimiter, String.Empty)
                                };
            
            for (int i = 0; i < historyFields.Count(); i++)
            {
                var a = historyFields.ElementAt(i);
                _pipeData.Add(String.Format("{0}{1}{2}", a.FieldName, _pipeFieldDelimiter,
                    _hluFieldSysTypeNames[a.FieldOrdinal]));
            }

            // GeometryColumn1: Length for polygons; length for polylines; X for points
            _pipeData.Add(String.Format("{0}{1}System.Double", _historyGeometry1ColumnName, _pipeFieldDelimiter));

            // GeometryColumn2: Area for polygons; empty for polylines; Y for points
            _pipeData.Add(String.Format("{0}{1}System.Double", _historyGeometry2ColumnName, _pipeFieldDelimiter));
            
            _pipeData.Add(_pipeTransmissionInterrupt);

            return historyFields.Select(hf => hf.FieldOrdinal).ToArray();
        }

        private string History(IFeature feature, int[] historyFieldOrdinals, string[] additionalValues)
        {
            StringBuilder history = new StringBuilder();

            int j = 0;
            foreach (int i in historyFieldOrdinals)
            {
                history.Append(String.Format("{0}{1}", _pipeFieldDelimiter, i != -1 ? 
                    feature.get_Value(i) : additionalValues[j++]));
            }

            double geom1;
            double geom2;
            GetGeometryProperties(feature, out geom1, out geom2);

            history.Append(String.Format("{0}{1}{0}{2}", _pipeFieldDelimiter,
                geom1 != -1 ? geom1.ToString() : String.Empty, 
                geom2 != -1 ? geom2.ToString() : String.Empty));

            return history.Remove(0, 1).ToString();
        }

        private void GetGeometryProperties(IFeature feature, out double geom1, out double geom2)
        {
            geom1 = -1;
            geom2 = -1;
            switch (feature.Shape.GeometryType)
            {
                case esriGeometryType.esriGeometryPolygon:
                    IArea area = feature.Shape as IArea;
                    geom1 = ((IPolygon4)feature.Shape).Length;
                    geom2 = area.Area;
                    break;
                case esriGeometryType.esriGeometryPolyline:
                    IPolyline5 pline = feature.Shape as IPolyline5;
                    geom1 = pline.Length;
                    break;
                case esriGeometryType.esriGeometryPoint:
                    IPoint point = feature.Shape as IPoint;
                    geom1 = point.X;
                    geom2 = point.Y;
                    break;
            }
        }

        #endregion

        #region Update

        private void UpdateAttributes(List<int> updateFieldOrdinals, 
            List<object> updateValues, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // make sure some features are selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count == 0) return;

                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(null, false, out cursor);
                    IFeatureCursor updateCursor = (IFeatureCursor)cursor;

                    UpdateAttributesAction(updateCursor, historyColumns, updateFieldOrdinals, updateValues);

                }
            }
            catch { _pipeData.Clear(); }
        }

        private void UpdateAttributes(string scratchMdbPath, string selectionDatasetName,
            List<int> updateFieldOrdinals, List<object> updateValues, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // select features to be updated
                    SelectByJoin(scratchMdbPath, selectionDatasetName);

                    // make sure some features are selected
                    if (_hluFeatureSelection == null)
                        _hluFeatureSelection = (IFeatureSelection)_hluLayer;
                    if (_hluFeatureSelection.SelectionSet.Count == 0) return;

                    ICursor cursor;
                    _hluFeatureSelection.SelectionSet.Search(null, false, out cursor);
                    IFeatureCursor updateCursor = (IFeatureCursor)cursor;

                    UpdateAttributesAction(updateCursor, historyColumns, updateFieldOrdinals, updateValues);
                }
            }
            catch { _pipeData.Clear(); }
        }

        private void UpdateAttributes(IQueryFilter queryFilter, List<int> updateFieldOrdinals,
            List<object> updateValues, string[] historyColumns)
        {
            // make sure selection changed event handler won't intervene
            _selectFieldOrdinals = null;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    IFeatureCursor updateCursor = _hluFeatureClass.Search(queryFilter, false);

                    UpdateAttributesAction(updateCursor, historyColumns, updateFieldOrdinals, updateValues);
                }
            }
            catch { _pipeData.Clear(); }
        }

        private void UpdateAttributesAction(IFeatureCursor updateCursor,
            string[] historyColumns, List<int> updateFieldOrdinals, List<object> updateValues)
        {
            // write history field names and types to pipe
            int[] historyFieldOrdinals = HistorySchema(historyColumns);

            // start workspace edit session and operation
            StartEditing();

            try
            {
                IFeature feature = null;
                while ((feature = updateCursor.NextFeature()) != null)
                {
                    _pipeData.Add(History(feature, historyFieldOrdinals, null));

                    for (int i = 0; i < updateFieldOrdinals.Count; i++)
                        feature.set_Value(updateFieldOrdinals[i], updateValues[i]);

                    feature.Store();
                }

                Marshal.ReleaseComObject(updateCursor);

                CommitEdits();
            }
            catch
            {
                Marshal.ReleaseComObject(updateCursor);
                DiscardEdits();
                throw;
            }
        }

        #endregion

        #region Editing

        private void IsEditing()
        {
            try
            {
                lock (_pipeData)
                {
                    _pipeData.Clear();
                    _pipeData.Add((InEditingSession && HluLayerBeingEdited).ToString());
                }
            }
            catch { }
        }

        private bool InEditingSession
        {
            get
            {
                return Editor != null && Editor.EditState == esriEditState.esriStateEditing;
            }
        }

        public IEditor Editor
        {
            get
            {
                if (_editor == null)
                {
                    try
                    {
                        if (_application != null)
                        {
                            UID editorUID = new UIDClass();
                            editorUID.Value = "esriEditor.Editor";
                            _editor = _application.FindExtensionByCLSID(editorUID) as IEditor;
                        }
                    }
                    catch { _editor = null; }
                }
                return _editor;
            }
        }

        private bool HluLayerBeingEdited
        {
            get { return _hluFeatureClass != null && Editor.EditWorkspace == _hluWS; }
        }

        private IWorkspaceEdit StartEditingWorkspace(IFeatureWorkspace editWS)
        {
            IWorkspaceEdit workspaceEdit = (IWorkspaceEdit)editWS;
            IMultiuserWorkspaceEdit multiUserWSEdit = editWS as IMultiuserWorkspaceEdit;
            if (multiUserWSEdit != null)
            {
                if (multiUserWSEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMNonVersioned))
                    multiUserWSEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMNonVersioned);
                else if (multiUserWSEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMVersioned))
                    multiUserWSEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMVersioned);
                else
                    throw new Exception("Cannot start multiuser editing session");
            }
            else
            {
                workspaceEdit.StartEditing(true);
            }
            workspaceEdit.StartEditOperation();

            return workspaceEdit;
        }

        private void StartEditing()
        {
            try
            {
                if (InEditingSession)
                {
                    if (_saveEditsCmd == null)
                        _saveEditsCmd = FindSaveEditsCommand();
                    _editor.StartOperation();
                }
                else
                {
                    _hluWorkspaceEdit = StartEditingWorkspace(_hluWS);
                }
            }
            catch { }
        }

        private void CloseEditSession(bool save)
        {
            // make sure selection changed event handler won't intervene
            int[] selectFieldOrdinalsBak = _selectFieldOrdinals;
            _selectFieldOrdinals = null;

            _editor.StopEditing(save);

            if (selectFieldOrdinalsBak != null)
            {
                _selectFieldOrdinals = selectFieldOrdinalsBak;
                selectFieldOrdinalsBak = null;
            }
        }

        private void OpenEditSession()
        {
            try
            {
                if (!InEditingSession)
                {
                    Editor.StartEditing((IWorkspace)_hluWS);
                }
            }
            catch { }
        }

        private void CommitEdits()
        {
            try
            {
                if (InEditingSession)
                {
                    _editor.StopOperation(String.Empty);
                    if (_saveEditsCmd == null)
                        _saveEditsCmd = FindSaveEditsCommand();

                    if (_saveEditsCmd != null)
                    {
                        _saveEditsCmd.Execute(); // only way to save edits without closing edit session
                    }
                    else
                    {
                        _editor.StopEditing(true);
                        _editor.StartEditing((IWorkspace)_hluWS);
                    }
                }
                else
                {
                    _hluWorkspaceEdit.StopEditOperation();
                    _hluWorkspaceEdit.StopEditing(true);
                    _hluWorkspaceEdit = null;
                }
            }
            catch { DiscardEdits(); }
        }

        private ICommandItem FindSaveEditsCommand()
        {
            ICommandBars cmdBars = _application.Document.CommandBars;
            UID uid = new UIDClass();
            uid.Value = "{605F0528-1B64-4862-93B1-484A2C534EA3}"; // "esriEditor.EditingToolbarNew"
            ICommandBar mxMenuBar = cmdBars.Find(uid, false, false) as ICommandBar;
            if (mxMenuBar == null)
            {
                uid.Value = "{C671B640-83B9-11D2-850C-0000F875B9C6}"; // "esriEditor.EditorToolBar"; // "esriCore.EditorToolBar";
                mxMenuBar = cmdBars.Find(uid, false, false) as ICommandBar;
            }

            if (mxMenuBar != null)
            {
                uid.Value = "esriCore.EditorMenu";
                mxMenuBar = mxMenuBar.Find(uid, false) as ICommandBar;
                if (mxMenuBar != null)
                {
                    uid.Value = "{59D2AFD2-9EA2-11D1-9165-0080C718DF97}"; // Editor_SaveEdits
                    return mxMenuBar.Find(uid, false);
                }
            }
            return null;
        }

        private void DiscardEdits()
        {
            try
            {
                if (InEditingSession)
                {
                    _editor.AbortOperation();
                }
                else
                {
                    _hluWorkspaceEdit.AbortEditOperation();
                    _hluWorkspaceEdit.StopEditOperation();
                    _hluWorkspaceEdit.StopEditing(false);
                }
            }
            finally { _hluWorkspaceEdit = null; }
        }

        private void FlushCursor(bool noFlush, ref IFeatureCursor cursor)
        {
            try
            {
                if (!noFlush && !IsHluWorkspaceSDE && (cursor != null))
                {
                    cursor.Flush();
                }
            }
            finally
            {
                Marshal.ReleaseComObject(cursor);
                cursor = null;
            }
        }

        private void FlushCursor(bool noFlush, ref ICursor cursor)
        {
            try
            {
                if (!noFlush && !IsHluWorkspaceSDE && (cursor != null))
                {
                    cursor.Flush();
                }
            }
            finally
            {
                Marshal.ReleaseComObject(cursor);
                cursor = null;
            }
        }

        #endregion

        #region Export

        private void Export(string tempMdbPathName, string attributeDatasetName, int exportRowCount)
        {
            IRelationshipClass relClass = null;
            IFeatureClass outFeatureClass = null;

            try
            {
                SetCursor(true);

                IDataset attributeDataset;
                ITable exportAttributes;
                OpenOleDbWorkspace(tempMdbPathName, attributeDatasetName, 
                    out attributeDataset, out exportAttributes);

                IDisplayTable hluDisplayTable = (IDisplayTable)_hluLayer;
                IFeatureClass hluDisplayTableFeatureClass = (IFeatureClass)hluDisplayTable.DisplayTable;
                ITable hluLayerTable = (ITable)hluDisplayTableFeatureClass;

                // Prompt the user for where to save the export layer
                IExportOperation exportOp = new ExportOperation();
                bool saveProjection;
                esriExportTableOptions exportOptions;
                IDatasetName exportDatasetName = exportOp.GetOptions(hluDisplayTableFeatureClass,
                    _hluLayer.Name, _hluFeatureSelection != null && _hluFeatureSelection.SelectionSet.Count > 0, 
                    true, _application.hWnd, out saveProjection, out exportOptions);

                // If no export dataset name was chosen by the user then cancel the export
                if (exportDatasetName == null)
                {
                    _pipeData.Add("cancelled");
                    return;
                }

                // Open the export dataset workspace
                object outWS = ((IName)exportDatasetName.WorkspaceName).Open();

                // Get the field names to be used when joining the attribute data and the feature layer
                string originPKJoinField = _hluLayerStructure.incidColumn.ColumnName;
                string originFKJoinField =
                    _hluFeatureClass.Fields.get_Field(_hluFieldMap[_hluLayerStructure.incidColumn.Ordinal]).Name;

                // Get a list of all the fields to be used in the export layer (plus separate lists of all
                // those fields that will come from the attribute table and those that will come from the
                // feature layer).
                List<IField> attributeFields;
                List<IField> featClassFields;
                List<IField> exportFields = ExportFieldLists(originPKJoinField, originFKJoinField,
                    exportAttributes, out attributeFields, out featClassFields);

                // Add x/y, length, or area and length fields to the list of fields in the export layer
                // if the export layer is a shapefile.
                bool isShp = IsShp(outWS as IWorkspace);
                ExportAddGeometryPropertyFields(isShp, exportFields);

                // create virtual relate
                IMemoryRelationshipClassFactory memoryRelFactory = new MemoryRelationshipClassFactoryClass();
                relClass = memoryRelFactory.Open("ExportRelClass", (IObjectClass)exportAttributes,
                    originPKJoinField, (IObjectClass)hluLayerTable, originFKJoinField, "forward", "backward",
                    esriRelCardinality.esriRelCardinalityOneToMany);

                // use Relate to perform a join
                IDisplayRelationshipClass displayRelClass = (IDisplayRelationshipClass)_hluLayer;
                displayRelClass.DisplayRelationshipClass(relClass, esriJoinType.esriLeftInnerJoin);

                // create query filter for export cursor
                bool featClassFieldsQualified;
                bool attributeFieldsQualified;
                IQueryFilter exportQueryFilter = ExportQueryFilter(originPKJoinField, hluDisplayTable,
                    attributeDataset, featClassFields, attributeFields, out featClassFieldsQualified,
                    out attributeFieldsQualified);

                IGeometryDef geomDef = _hluFeatureClass.Fields.get_Field(_hluFeatureClass.FindField(
                    _hluFeatureClass.ShapeFieldName)).GeometryDef;

                // adds OID and SHAPE at beginning, possibly Shape_Length and Shape_Area at end
                // when populating new rows we loop over exportFieldOrdinals
                // if we export shp we calculate geometry props into the last two fields, which are
                // not in exportFields
                IFields outFields = CreateFieldsCollection(true, geomDef.HasZ, geomDef.HasM, outWS,
                    _hluFeatureClass.ShapeType, exportFields.Select(f => f.Length).ToArray(),
                    exportFields.Select(f => f.Name).ToArray(), exportFields.Select(f => f.Name).ToArray(),
                    exportFields.Select(f => f.Type).ToArray(), exportFields.Select(f => f.Type != 
                        esriFieldType.esriFieldTypeOID).ToArray(), geomDef.SpatialReference);

                // create output feature class
                outFeatureClass = CreateFeatureClass(exportDatasetName.Name, null, outWS,
                    outFields, esriFeatureType.esriFTSimple, _hluFeatureClass.ShapeType, null, null);

                // field map between display and output feature class, as display 
                // table always includes all fields, regardless of SubFields
                // the first two fields are always OID and SHAPE 
                // the last two Shape_Length and Shape_Area, either added automatically or here
                int[] exportFieldMap = new int[] { 0, 1 }.Concat(featClassFields
                    .Select(f => hluDisplayTable.DisplayTable.Fields.FindField(featClassFieldsQualified ?
                        _hluLayer.Name + "." + f.Name : f.Name))).Concat(attributeFields
                    .Select(f => hluDisplayTable.DisplayTable.Fields.FindField(attributeFieldsQualified ?
                        attributeDataset.Name + "." + f.Name : f.Name))).ToArray();

                // insert features into new feature class
                ExportInsertFeatures(hluDisplayTable, exportQueryFilter, exportRowCount,
                    exportFieldMap, isShp, outWS, outFeatureClass);

            }
            catch (Exception ex) { _pipeData.Add(ex.Message); }
            finally
            {
                if (relClass != null)
                    ((IDisplayRelationshipClass)_hluLayer).DisplayRelationshipClass(
                        null, esriJoinType.esriLeftInnerJoin);
                outFeatureClass = null;
                try { if (File.Exists(tempMdbPathName)) File.Delete(tempMdbPathName); }
                catch { }
                SetCursor(false);
            }
        }

        private void OpenOleDbWorkspace(string tempMdbPathName, string datasetName,
            out IDataset attributeDataset, out ITable exportAttributes)
        {
            if (!File.Exists(tempMdbPathName)) throw new IOException("File not found");

            IPropertySet propertySet = new PropertySet();
            propertySet.SetProperty("CONNECTSTRING",
                String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", tempMdbPathName));

            IWorkspaceFactory workspaceFactory = new OLEDBWorkspaceFactoryClass();
            IWorkspace exportWorkspace = workspaceFactory.Open(propertySet, 0);

            exportAttributes = null;
            attributeDataset = null;
            IEnumDataset enumExportDatasets = exportWorkspace.get_Datasets(esriDatasetType.esriDTAny);

            while ((attributeDataset = enumExportDatasets.Next()) != null)
            {
                if (attributeDataset.Name == datasetName)
                {
                    object attributeDatasetObject = attributeDataset.FullName.Open();
                    exportAttributes = attributeDataset as ITable;
                    break;
                }
            }
        }

        private List<IField> ExportFieldLists(string originPKJoinField, string originFKJoinField, 
            ITable exportAttributes, out List<IField> attributeFields, out List<IField> featClassFields)
        {
            // Build a list of all the export fields that will come from the
            // attribute table (except the field to be used as the primary key
            // when joining the attribute table to the feature layer)
            attributeFields = new List<IField>();
            List<int> attributeFieldOrdinals = new List<int>();
            for (int i = 0; i < exportAttributes.Fields.FieldCount; i++)
            {
                IField attributeField = exportAttributes.Fields.get_Field(i);
                if (attributeField.Name != originPKJoinField)
                {
                    attributeFields.Add(attributeField);
                    attributeFieldOrdinals.Add(i);
                }
            }

            // Build a list consisting of the feature layer field that is to be
            // used as the foreign key when joining to the attribute table plus
            // any other feature fields not already in the attribute table.
            featClassFields = new List<IField>();
            List<int> featClassFieldOrdinals = new List<int>();
            foreach (DataColumn c in _hluLayerStructure.Columns)
            {
                if ((c.ColumnName == originFKJoinField) ||
                    (attributeFields.Count(f => f.Name == c.ColumnName) == 0))
                {
                    int fieldOrdinal = _hluFieldMap[c.Ordinal];
                    featClassFields.Add(_hluFeatureClass.Fields.get_Field(fieldOrdinal));
                    featClassFieldOrdinals.Add(fieldOrdinal);
                }
            }

            // Append the attribute table fields to the feature layer fields
            // as a new list of fields to go into the export layer.
            List<IField> exportFields = new List<IField>(featClassFields);
            exportFields.AddRange(attributeFields);

            return exportFields;
        }

        private void ExportAddGeometryPropertyFields(bool isShp, List<IField> exportFields)
        {
            
            switch (_hluFeatureClass.ShapeType)
            {
                case esriGeometryType.esriGeometryPoint:
                    IFieldEdit geomFieldEdit = (IFieldEdit)new FieldClass();
                    geomFieldEdit.Name_2 = "x";
                    geomFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    exportFields.Add(geomFieldEdit);
                    geomFieldEdit = (IFieldEdit)new FieldClass();
                    geomFieldEdit.Name_2 = "y";
                    geomFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    exportFields.Add(geomFieldEdit);

                    break;
                case esriGeometryType.esriGeometryPolyline:
                    if (isShp)
                    {
                        IFieldEdit lengthFieldEdit = (IFieldEdit)new FieldClass();
                        lengthFieldEdit.Name_2 = "shape_length";
                        lengthFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                        exportFields.Add(lengthFieldEdit);
                    }
                    break;
                case esriGeometryType.esriGeometryPolygon:
                    if (isShp)
                    {
                        IFieldEdit lengthFieldEdit = (IFieldEdit)new FieldClass();
                        lengthFieldEdit.Name_2 = "shape_length";
                        lengthFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                        exportFields.Add(lengthFieldEdit);
                        IFieldEdit areaFieldEdit = (IFieldEdit)new FieldClass();
                        areaFieldEdit.Name_2 = "shape_area";
                        areaFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                        exportFields.Add(areaFieldEdit);
                    }
                    break;
            }
        }

        private IQueryFilter ExportQueryFilter(string originPKJoinField, IDisplayTable hluDisplayTable, 
            IDataset attributeDataset, List<IField> featClassFields, List<IField> attributeFields,
            out bool featClassFieldsQualified, out bool attributeFieldsQualified)
        {
            StringBuilder queryFilterSubFields = new StringBuilder();
            string queryFilterWhereClause = String.Empty;

            featClassFieldsQualified = hluDisplayTable.DisplayTable.Fields.get_Field(0)
                .Name.StartsWith(_hluLayer.Name + ".");
            attributeFieldsQualified = hluDisplayTable.DisplayTable.Fields.get_Field(
                _hluFeatureClass.Fields.FieldCount).Name.StartsWith(attributeDataset.Name + ".");

            if (featClassFieldsQualified)
            {
                queryFilterSubFields.Append(String.Join(",", featClassFields
                    .Select(f => _hluLayer.Name + "." + f.Name).ToArray()));
            }
            else
            {
                queryFilterSubFields.Append(String.Join(",", featClassFields.Select(f => f.Name).ToArray()));
            }

            if (attributeFieldsQualified)
            {
                queryFilterSubFields.Append(",").Append(String.Join(",",
                    attributeFields.Select(f => attributeDataset.Name + "." + f.Name).ToArray()));
                //queryFilterWhereClause = String.Format("{0} IS NOT NULL",
                  //  _hluLayer.Name + "." + originPKJoinField);
            }
            else
            {
                queryFilterSubFields.Append(",").Append(
                    String.Join(",", attributeFields.Select(f => f.Name).ToArray()));
                //queryFilterWhereClause = String.Format("{0} IS NOT NULL", originPKJoinField);
            }
            IQueryFilter exportQueryFilter = new QueryFilterClass();
            exportQueryFilter.SubFields = queryFilterSubFields.ToString();
            exportQueryFilter.WhereClause = queryFilterWhereClause;

            return exportQueryFilter;
        }

        private bool IsShp(IWorkspace workspace)
        {
            return workspace != null && workspace.WorkspaceFactory.GetClassID().Value.ToString().ToUpper() ==
                typeof(ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass).GUID.ToString("B").ToUpper();
        }

        private void ExportInsertFeatures(IDisplayTable hluDisplayTable, IQueryFilter exportQueryFilter, 
            int exportRowCount, int[] exportFieldMap, bool isShp, object outWS, IFeatureClass outFeatureClass)
        {
            IStatusBar statusBar = _application.StatusBar;
            IStepProgressor progressBar = statusBar.ProgressBar;
            progressBar.Position = 0;
            statusBar.ShowProgressBar("Exporting...", 0, exportRowCount, 1, true);

            IWorkspaceEdit workspaceEdit = null;
            IWorkspace wsOut = outWS as IWorkspace;
            bool restoreEditSession = InEditingSession;
            if (restoreEditSession) CloseEditSession(true);

            if (wsOut.WorkspaceFactory.WorkspaceType == esriWorkspaceType.esriRemoteDatabaseWorkspace)
            {
                Editor.StartEditing(wsOut);
                Editor.StartOperation();
            }

            else
            {
                workspaceEdit = (IWorkspaceEdit)outWS;
                workspaceEdit.StartEditing(true);
                workspaceEdit.StartEditOperation();
            }
            
            IFeatureCursor exportFeatureCursor =
                (IFeatureCursor)hluDisplayTable.SearchDisplayTable(exportQueryFilter, true);
            IFeature exportFeature;

            IFeatureCursor insertCursor = outFeatureClass.Insert(true);
            IFeatureBuffer featureBuffer = outFeatureClass.CreateFeatureBuffer();

            bool calcGeometry = _hluFeatureClass.ShapeType == esriGeometryType.esriGeometryPoint || isShp;
            double geom1;
            double geom2;
            int ixGeom1 = featureBuffer.Fields.FieldCount - 2;
            int ixGeom2 = featureBuffer.Fields.FieldCount - 1;

            try
            {
                object item;
                while ((exportFeature = exportFeatureCursor.NextFeature()) != null)
                {
                    featureBuffer.Shape = exportFeature.ShapeCopy;

                    for (int i = 2; i < exportFieldMap.Length; i++)
                    {
                        item = exportFeature.get_Value(exportFieldMap[i]);
                        if (item != DBNull.Value) featureBuffer.set_Value(i, item);
                    }

                    if (calcGeometry)
                    {
                        GetGeometryProperties(exportFeature, out geom1, out geom2);
                        if (geom1 != -1)
                            featureBuffer.set_Value(ixGeom1, geom1);
                        if (geom2 != -1)
                            featureBuffer.set_Value(ixGeom2, geom2);
                    }

                    try { insertCursor.InsertFeature(featureBuffer); }
                    catch { }

                    progressBar.Step();
                }
                FlushCursor(false, ref insertCursor);

                if (workspaceEdit == null)
                {
                    Editor.StopOperation(String.Empty);
                    Editor.StopEditing(true);
                }
                else
                {
                    workspaceEdit.StopEditOperation();
                    workspaceEdit.StopEditing(true);
                }
            }
            catch
            {
                if (workspaceEdit == null)
                {
                    Editor.AbortOperation();
                    Editor.StopEditing(false);
                }
                else
                {
                    workspaceEdit.AbortEditOperation();
                    workspaceEdit.StopEditOperation();
                    workspaceEdit.StopEditing(false);
                }
                throw;
            }
            finally
            {
                FlushCursor(true, ref exportFeatureCursor);
                if (restoreEditSession) OpenEditSession();
                statusBar.set_Message((int)esriStatusBarPanes.esriStatusMain, "");
                statusBar.HideProgressBar();
            }
        }

        /// <summary>
        /// Creates the field collection for a new feature class or standalone table.
        /// </summary>
        /// <param name="bSpatial">True if the new fields collection is for a feature class, false if a table.</param>
        /// <param name="bHasZ">True if the geometry field of the fields collection to be created is to be Z aware.</param>
        /// <param name="bHasM">True if the geometry field of the fields collection to be created is to be M aware.</param>
        /// <param name="oWorkspace">The workspace in which a new feature class will be created from the fields collection.</param>
        /// <param name="geomType">The geometry type of the feature class to be created from the new fields collection.</param>
        /// <param name="aFldLengths">An array of field lengths of the fields collection to be created.</param>
        /// <param name="aFldNames">An array of field names of the fields collection to be created.</param>
        /// <param name="aFldAliases">An array of field aliases of the fields collection to be created.</param>
        /// <param name="aFldTypes">An array of field types of the fields collection to be created.</param>
        /// <param name="spatialReference">The coordinate system to be assigned to the new feature class.</param>
        /// <returns>Returns the new field collection. Throws an error if no field collection can be created.</returns>
        private IFields CreateFieldsCollection(bool hasGeom, bool hasZ, bool hasM, object outWorkspace,
            esriGeometryType geometryType, int[] fieldLengths, string[] fieldNames, string[] fieldAliases,
            esriFieldType[] fieldTypes, bool[] fieldNullable, ISpatialReference spatialReference)
        {
            if ((hasGeom) && (outWorkspace is IWorkspace) && (spatialReference == null))
                throw new Exception("Spatial reference must be defined for standalone feature classes");

            IWorkspace outputWorkspace = null;
            if (outWorkspace is IWorkspace)
            {
                outputWorkspace = outWorkspace as IWorkspace;
            }
            else if (outWorkspace is IFeatureDataset)
            {
                outputWorkspace = ((IDataset)((IFeatureDataset)outWorkspace)).Workspace;
            }

            IFields fields = new FieldsClass();
            IFieldsEdit fieldsEdit = (IFieldsEdit)fields;
            if (hasGeom)
            {
                IGeometryDef geometryDef = new GeometryDefClass();
                IGeometryDefEdit geometryDefEdit = (IGeometryDefEdit)geometryDef;
                geometryDefEdit.GeometryType_2 = geometryType;
                geometryDefEdit.GridCount_2 = 1;
                geometryDefEdit.set_GridSize(0, 0); // (0, 0.5);
                geometryDefEdit.AvgNumPoints_2 = 2;
                geometryDefEdit.HasM_2 = hasM;
                geometryDefEdit.HasZ_2 = hasZ;
                if (outWorkspace is IWorkspace) geometryDefEdit.SpatialReference_2 = spatialReference;
                IField fieldOID = new FieldClass();
                IFieldEdit fieldEditOID = (IFieldEdit)fieldOID;
                fieldEditOID.Name_2 = "OBJECTID";
                fieldEditOID.AliasName_2 = "OBJECTID";
                fieldEditOID.Type_2 = esriFieldType.esriFieldTypeOID;
                fieldsEdit.AddField(fieldOID);
                IField fieldShape = new FieldClass();
                IFieldEdit fieldEditShape = (IFieldEdit)fieldShape;
                fieldEditShape.Name_2 = "SHAPE";
                fieldEditShape.AliasName_2 = "SHAPE";
                fieldEditShape.Type_2 = esriFieldType.esriFieldTypeGeometry;
                fieldEditShape.GeometryDef_2 = geometryDef;
                fieldsEdit.AddField(fieldShape);
            }
            else
            {
                IField fieldOID = new FieldClass();
                IFieldEdit fieldEditOID = (IFieldEdit)fieldOID;
                fieldEditOID.Name_2 = "OBJECTID";
                fieldEditOID.AliasName_2 = "OBJECTID";
                fieldEditOID.Type_2 = esriFieldType.esriFieldTypeOID;
                fieldsEdit.AddField(fieldOID);
            }

            if (fieldAliases == null) fieldAliases = fieldNames;

            IField fieldAtt;
            IFieldEdit fieldEditAtt;
            for (int i = 0; i < fieldTypes.Length; i++)
            {
                if (String.IsNullOrEmpty(fieldNames[i]))
                    throw new Exception(String.Format("Error creating field {0}", i));
                if (fieldTypes[i] == esriFieldType.esriFieldTypeOID) continue;
                fieldAtt = new FieldClass();
                fieldEditAtt = (IFieldEdit)fieldAtt;
                fieldEditAtt.Name_2 = fieldNames[i];
                if (!String.IsNullOrEmpty(fieldAliases[i]))
                    fieldEditAtt.AliasName_2 = fieldAliases[i];
                fieldEditAtt.Type_2 = fieldTypes[i];
                fieldEditAtt.Editable_2 = true;
                fieldEditAtt.Required_2 = false;
                fieldEditAtt.IsNullable_2 = fieldNullable[i];// true;
                if (fieldTypes[i] == esriFieldType.esriFieldTypeString)
                {
                    //if ((fieldLengths[i] == -1) || ((fieldLengths[i] > 0) && (fieldLengths[i] < 256)))
                    fieldEditAtt.Length_2 = fieldLengths[i];
                    //else
                    //    fieldEditAtt.Length_2 = 255;
                }
                fieldsEdit.AddField(fieldAtt);
            }
            IFieldChecker fldChk = new FieldCheckerClass();
            fldChk.ValidateWorkspace = outputWorkspace;

            IFields outFields;
            IEnumFieldError fieldErrors;
            fldChk.Validate(fields, out fieldErrors, out outFields);
            return outFields;
        }

        /// <summary>
        /// Creates a new feature class. Returns the feature class or null if not successful. Throws but does not handle errors. 
        /// </summary>
        /// <param name="sFeatName">Feature class name.</param>
        /// <param name="configWord">Configuration keywords for the new feature class.</param>
        /// <param name="oWorkspace">Workspace in which to create the new feature class.</param>
        /// <param name="theFields">Fields collection of the new feature class.</param>
        /// <param name="featureType">Feature type of the new feature class.</param>
        /// <param name="geometryType">geometry type of the new feature class.</param>
        /// <param name="CLSID">Class ID of the new feature class. May be null if featureType is not null.</param>
        /// <param name="CLSEXT">Class extension ID of the new feature class. Pass null if none.</param>
        /// <returns>Returns the cloned feature class. Throws an error if no feature class can be created.</returns>
        private IFeatureClass CreateFeatureClass(string featureClassName, string configWord, object outWorkspace,
            IFields fields, esriFeatureType featureType, esriGeometryType geometryType, UID CLSID, UID CLSEXT)
        {
            IFeatureClass featureClass = null;
            if (outWorkspace == null) throw (new Exception("Workspace cannot be null"));
            if (!((outWorkspace is IWorkspace) || (outWorkspace is IFeatureDataset)))
                throw (new Exception("Workspace must be IWorkspace or IFeatureDataset"));
            if (featureClassName == "") throw (new Exception("Feature class name cannot be empty string"));
            UID uidCLSID = new UIDClass();
            if (CLSID == null)
            {
                switch (featureType)
                {
                    case (esriFeatureType.esriFTSimple):
                        uidCLSID.Value = "{52353152-891A-11D0-BEC6-00805F7C4268}";
                        break;
                    case (esriFeatureType.esriFTSimpleJunction):
                        geometryType = esriGeometryType.esriGeometryPoint;
                        uidCLSID.Value = "{CEE8D6B8-55FE-11D1-AE55-0000F80372B4}";
                        break;
                    case (esriFeatureType.esriFTComplexJunction):
                        uidCLSID.Value = "{DF9D71F4-DA32-11D1-AEBA-0000F80372B4}";
                        break;
                    case (esriFeatureType.esriFTSimpleEdge):
                        geometryType = esriGeometryType.esriGeometryPolyline;
                        uidCLSID.Value = "{E7031C90-55FE-11D1-AE55-0000F80372B4}";
                        break;
                    case (esriFeatureType.esriFTComplexEdge):
                        geometryType = esriGeometryType.esriGeometryPolyline;
                        uidCLSID.Value = "{A30E8A2A-C50B-11D1-AEA9-0000F80372B4}";
                        break;
                    case (esriFeatureType.esriFTAnnotation):
                        geometryType = esriGeometryType.esriGeometryPolygon;
                        uidCLSID.Value = "{E3676993-C682-11D2-8A2A-006097AFF44E}";
                        break;
                    case (esriFeatureType.esriFTDimension):
                        geometryType = esriGeometryType.esriGeometryPolygon;
                        uidCLSID.Value = "{496764FC-E0C9-11D3-80CE-00C04F601565}";
                        break;
                }
            }
            else
            {
                uidCLSID = CLSID;
            }
            UID uidCLSEXT = new UIDClass();
            if (CLSEXT == null)
            {
                switch (featureType)
                {
                    case (esriFeatureType.esriFTAnnotation):
                        uidCLSEXT.Value = "{24429589-D711-11D2-9F41-00C04F6BC6A5}";
                        break;
                    case (esriFeatureType.esriFTDimension):
                        uidCLSEXT.Value = "{48F935E2-DA66-11D3-80CE-00C04F601565}";
                        break;
                    default:
                        uidCLSEXT = null;
                        break;
                }
            }
            else
                uidCLSEXT = CLSEXT;
            string shapeFieldName = "";
            for (int i = 0; i <= fields.FieldCount - 1; i++)
            {
                if (fields.get_Field(i).Type == esriFieldType.esriFieldTypeGeometry)
                {
                    shapeFieldName = fields.get_Field(i).Name;
                    break;
                }
            }
            if (shapeFieldName == "") throw new Exception("Cannot locate geometry field");
            if (outWorkspace is IWorkspace)
            {
                IWorkspace workspace = (IWorkspace)outWorkspace;
                IFeatureWorkspace featureWorkspace = (IFeatureWorkspace)workspace;
                featureClass = featureWorkspace.CreateFeatureClass(featureClassName, fields, uidCLSID, uidCLSEXT,
                    featureType, shapeFieldName, configWord);
            }
            else if (outWorkspace is IFeatureDataset)
            {
                IFeatureDataset featureDataset = (IFeatureDataset)outWorkspace;
                featureClass = featureDataset.CreateFeatureClass(featureClassName, fields, uidCLSID, uidCLSEXT,
                    featureType, shapeFieldName, configWord);
            }
            return featureClass;
        }

        #endregion

        private void SetCursor(bool wait)
        {
            try
            {
                IMouseCursor mouseCursor = new MouseCursorClass();
                mouseCursor.SetCursor(wait ? 2 : 0);
            }
            catch { }
        }

        #region HLU

        private bool HluDbIsRunning
        {
            get
            {
                try
                {
                    return Process.GetProcesses().Count(p => p.ProcessName.StartsWith(ArcMapApp.HluDbAppName)) > 0;
                }
                catch { return false; }
            }
        }

        public bool IsHluWorkspaceSDE
        {
            get { return _hluWSisSDE; }
        }

        /// <summary>
        /// Determines whether at least one of the layers in all the maps is an HLU layer.
        /// Called by OnOpenDocument when ArcGIS determines that a document has just been opened.
        /// </summary>
        /// <param name="maps">The maps collection (i.e. active views) of the open document.</param>
        /// <returns></returns>
        public bool IsHluWorkspace(IMaps maps)
        {
            try
            {
                IMap map = null;

                UID uid = new UIDClass();
                uid.Value = typeof(IFeatureLayer).GUID.ToString("B");

                for (int i = 0; i < maps.Count; i++)
                {
                    map = maps.get_Item(i);
                    IEnumLayer layers = map.get_Layers(uid, true);
                    ILayer layer = layers.Next();
                    while (layer != null)
                    {
                        IFeatureLayer featureLayer = layer as IFeatureLayer;
                        if (IsHluLayer(featureLayer))
                        {
                            _hluView = map as IActiveView;
                            return true;
                        }
                        layer = layers.Next();
                    }
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// Determines whether at least one of the layers in all the maps is an HLU layer
        /// and returns the field mapping of the first HLU layer found, plus a list of all
        /// the HLU layers found.
        /// Triggered from ArcMapApp after the required document has been opened.
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        public void IsHluWorkspace()
        {
            IFeatureLayer hluLayerBak = _hluLayer;
            int[] hluFieldMapBak = _hluFieldMap;
            string[] hluFieldNamesBak = _hluFieldNames;
            IFeatureClass hluFeatureClassBak = _hluFeatureClass;
            IFeatureWorkspace hluWSBak = _hluWS;
            string hluTableNameBak = _hluTableName;
            ISQLSyntax hluSqlSyntaxBak = _hluSqlSyntax;
            string quotePrefixBak = _quotePrefix;
            string quoteSuffixBak = _quoteSuffix;

            _pipeCalling = true;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    _hluLayer = null;
                    IMaps maps = ((IMxDocument)_application.Document).Maps;
                    IMap map = null;

                    UID uid = new UIDClass();
                    uid.Value = typeof(IFeatureLayer).GUID.ToString("B");

                    for (int i = 0; i < maps.Count; i++)
                    {
                        map = maps.get_Item(i);
                        int j = 0;
                        IEnumLayer layers = map.get_Layers(uid, true);
                        ILayer layer = layers.Next();
                        while (layer != null)
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR19 (Feature layer position in GIS)
                            // Only check geofeature layers in the document (to see if they are
                            // valid HLU layers) to save having to check layers (e.g. coverage
                            // annotation layers) that can't possibly be valid.
                            if (layer is IGeoFeatureLayer)
                            {
                                IFeatureLayer featureLayer = layer as IFeatureLayer;
                                if (IsHluLayer(featureLayer))
                                {
                                    _hluView = map as IActiveView;
                                    // Return details of the first valid HLU layer found (the map number,
                                    // map name, layer number, field indexes and field names).
                                    _pipeData.AddRange(new string[] { i.ToString(), map.Name, j.ToString() });
                                    _pipeData.Add(_pipeTransmissionInterrupt);
                                    _pipeData.AddRange(_hluFieldMap.Select(ix => ix.ToString()));
                                    _pipeData.Add(_pipeTransmissionInterrupt);
                                    _pipeData.AddRange(_hluFieldNames);
                                    return;
                                }
                            }
                            //---------------------------------------------------------------------
                            layer = layers.Next();
                            j++;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                _pipeCalling = false;
                _hluLayer = hluLayerBak;
                _hluFieldMap = hluFieldMapBak;
                _hluFieldNames = hluFieldNamesBak;
                _hluFeatureClass = hluFeatureClassBak;
                _hluWS = hluWSBak;
                _hluTableName = hluTableNameBak;
                _hluSqlSyntax = hluSqlSyntaxBak;
                _quotePrefix = quotePrefixBak;
                _quoteSuffix = quoteSuffixBak;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        //
        /// <summary>
        /// Determines which of the layers in all the maps are valid HLU layers
        /// and sets the list of all the HLU layers found.
        /// Triggered from ArcMapApp before switching layers.
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        public void ListHluLayers()
        {
            IFeatureLayer hluLayerBak = _hluLayer;
            int[] hluFieldMapBak = _hluFieldMap;
            string[] hluFieldNamesBak = _hluFieldNames;
            IFeatureClass hluFeatureClassBak = _hluFeatureClass;
            IFeatureWorkspace hluWSBak = _hluWS;
            string hluTableNameBak = _hluTableName;
            ISQLSyntax hluSqlSyntaxBak = _hluSqlSyntax;
            string quotePrefixBak = _quotePrefix;
            string quoteSuffixBak = _quoteSuffix;

            _pipeCalling = true;

            if (_pipeData == null) _pipeData = new List<string>();

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    //_hluLayer = null;
                    IMaps maps = ((IMxDocument)_application.Document).Maps;
                    IMap map = null;

                    UID uid = new UIDClass();
                    uid.Value = typeof(IFeatureLayer).GUID.ToString("B");

                    // Initialise the list of valid layers
                    if (_hluLayerList == null) _hluLayerList = new List<string>();

                    // Clear the valid HLU layer list
                    _hluLayerList.Clear();

                    // Loop through each map in the document
                    for (int i = 0; i < maps.Count; i++)
                    {
                        map = maps.get_Item(i);
                        int j = 0;
                        IEnumLayer layers = map.get_Layers(uid, true);
                        ILayer layer = layers.Next();
                        while (layer != null)
                        {
                            // Only check geofeature layers in the document (to see if they are
                            // valid HLU layers) to save having to check layers (e.g. coverage
                            // annotation layers) that can't possibly be valid.
                            if (layer is IGeoFeatureLayer)
                            {
                                IFeatureLayer featureLayer = layer as IFeatureLayer;
                                if (IsHluLayer(featureLayer))
                                {
                                    // Add the map number, map name, layer number and layer name of the
                                    // current layer to the list of valid layers.
                                    _hluLayerList.Add(String.Format("{0}::{1}::{2}::{3}", i.ToString(), map.Name, j.ToString(), layer.Name));
                                }
                            }
                            layer = layers.Next();
                            j++;
                        }
                    }

                    // Return the list of layers once all the layers have been checked
                    // as valid (or not).
                    _pipeData.AddRange(new string[] { _hluLayerList.Count.ToString() });
                    _pipeData.Add(_pipeTransmissionInterrupt);
                    _pipeData.AddRange(_hluLayerList);
                    return;
                }
            }
            catch { }
            finally
            {
                _pipeCalling = false;
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Determines whether the layer passed is an HLU layer, sets all related fields except _hluView 
        /// (which it cannot obtain from the layer) and sets up the selection changed event handler for the
        /// layer if it is an HLU layer.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        private bool IsHluLayer(IFeatureLayer layer)
        {
            try
            {
                //---------------------------------------------------------------------
                // CHANGED: CR31 (Switching between GIS layers)
                // Don't update the class properties unless an HLU layer (_hluLayer)
                // has not already been selected so that other feature layers can
                // be examined to see if they are also valid HLU layers (without
                // overwritting the class properties of the first layer).
                int[] hluFieldMap;
                string[] hluFieldNames;

                if (HluDbIsRunning && ArcMapAppHelperClass.IsHluLayer(layer,
                new FieldsClass(), new FieldCheckerClass(), _validWorkspaces, _typeMapSystemToSql,
                ref _hluLayerStructure, out hluFieldMap, out hluFieldNames))
                {
                    IFeatureLayer hluLayer = layer as IFeatureLayer;
                    IFeatureClass hluFeatureClass = hluLayer.FeatureClass;
                    IWorkspace hluWorkspace = ((IDataset)hluFeatureClass).Workspace;
                    IFeatureWorkspace hluWS = hluWorkspace as IFeatureWorkspace;
                    if ((!_pipeCalling) || (_hluLayer == null)) SetupSelectionChangedEvent(hluLayer);

                    int[] hluUidFieldOrdinals = new int[3];
                    hluUidFieldOrdinals[0] = hluFieldMap[_hluLayerStructure.incidColumn.Ordinal];
                    hluUidFieldOrdinals[1] = hluFieldMap[_hluLayerStructure.toidColumn.Ordinal];
                    hluUidFieldOrdinals[2] = hluFieldMap[_hluLayerStructure.toid_fragment_idColumn.Ordinal];

                    string[] hluFieldSysTypeNames = new string[hluFeatureClass.Fields.FieldCount];
                    Type sysType;
                    for (int i = 0; i < hluFeatureClass.Fields.FieldCount; i++)
                    {
                        if (_typeMapSQLToSystem.TryGetValue((int)hluFeatureClass.Fields.get_Field(i).Type, out sysType))
                            hluFieldSysTypeNames[i] = sysType.FullName;
                    }

                    string hluTableName = ((IDataset)hluFeatureClass).Name;
                    ISQLSyntax hluSqlSyntax = (ISQLSyntax)hluWS;
                    string quotePrefix =
                        hluSqlSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix);
                    string quoteSuffix =
                       hluSqlSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix);

                    if (_hluLayer == null)
                    {
                        _hluLayer = hluLayer;
                        _hluFeatureClass = hluFeatureClass;
                        _hluWS = hluWS;
                        _hluWSisSDE = hluWorkspace.WorkspaceFactory.WorkspaceType == esriWorkspaceType.esriRemoteDatabaseWorkspace;
                        _hluUidFieldOrdinals = hluUidFieldOrdinals;
                        _hluFieldSysTypeNames = hluFieldSysTypeNames;
                        _hluTableName = hluTableName;
                        _hluSqlSyntax = hluSqlSyntax;
                        _quotePrefix = quotePrefix;
                        _quoteSuffix = quoteSuffix;
                        _hluFieldMap = hluFieldMap;
                        _hluFieldNames = hluFieldNames;
                    }
                    //---------------------------------------------------------------------

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                _hluLayer = null;
                _hluFeatureClass = null;
                _hluTableName = null;
                _hluWS = null;
                _hluWSisSDE = false;
                _hluView = null;
                _hluFieldMap = null;
                _hluFieldNames = null;
                _hluUidFieldOrdinals = null;
                _selectFieldOrdinals = null;
                _hluSqlSyntax = null;
                _quotePrefix = null;
                _quoteSuffix = null;
                return false;
            }
        }

        private void IsHluLayer(int ixMap, int ixLayer)
        {
            if (_pipeData == null) _pipeData = new List<string>();

            IFeatureLayer hluLayerBak = _hluLayer;
            int[] hluFieldMapBak = _hluFieldMap;
            string[] hluFieldNamesBak = _hluFieldNames;
            IFeatureClass hluFeatureClassBak = _hluFeatureClass;
            IFeatureWorkspace hluWSBak = _hluWS;
            string hluTableNameBak = _hluTableName;
            ISQLSyntax hluSqlSyntaxBak = _hluSqlSyntax;
            string quotePrefixBak = _quotePrefix;
            string quoteSuffixBak = _quoteSuffix;

            _pipeCalling = true;

            _hluLayer = null;

            try
            {
                lock (_pipeData)
                {
                    // clear the pipe
                    _pipeData.Clear();

                    // Get the correct map based on the map number.
                    IMap map = ((IMxDocument)_application.Document).Maps.get_Item(ixMap);

                    UID uid = new UIDClass();
                    uid.Value = typeof(IFeatureLayer).GUID.ToString("B");

                    // Loop through each layer in the map looking for the correct layer
                    // by number (order).
                    int j = 0;
                    IEnumLayer layers = map.get_Layers(uid, true);
                    ILayer layer = layers.Next();
                    while (layer != null)
                    {
                        if (j == ixLayer)
                        {
                            IFeatureLayer featureLayer = layer as IFeatureLayer;
                            if (IsHluLayer(featureLayer))
                            {
                                _hluView = map as IActiveView;
                                _pipeData.Add("true");
                                _pipeData.Add(_pipeTransmissionInterrupt);
                                _pipeData.AddRange(_hluFieldMap.Select(ix => ix.ToString()));
                                _pipeData.Add(_pipeTransmissionInterrupt);
                                _pipeData.AddRange(_hluFieldNames);
                            }
                            else
                            {
                                _pipeData.Add("false");
                            }
                        }
                        layer = layers.Next();
                        j++;
                    }
                }
            }
            catch
            {
                _pipeData.Clear();
            }
            finally
            {
                _pipeCalling = false;
            }

            //---------------------------------------------------------------------
            // CHANGED: CR31 (Switching between GIS layers)
            // Reset the class properties if there has been an error or if
            // the layer is not valid.
            if (_hluLayer == null)
            {
                _hluLayer = hluLayerBak;
                _hluFieldMap = hluFieldMapBak;
                _hluFieldNames = hluFieldNamesBak;
                _hluFeatureClass = hluFeatureClassBak;
                _hluWS = hluWSBak;
                _hluTableName = hluTableNameBak;
                _hluSqlSyntax = hluSqlSyntaxBak;
                _quotePrefix = quotePrefixBak;
                _quoteSuffix = quoteSuffixBak;
            }
            //---------------------------------------------------------------------
        }

        private int MapField(string name)
        {
            name = name.Trim();
            int o;
            if ((o = FieldOrdinal(name)) != -1)
            {
                return o;
            }
            else if ((o = ColumnOrdinal(name)) != -1)
            {
                return FieldOrdinal(_hluLayerStructure.Columns[o].ColumnName);
            }
            return -1;
        }

        private int FieldOrdinal(string columnName)
        {
            int ordinal = -1;
            if ((_hluFieldMap != null) && (_hluLayerStructure != null) && !String.IsNullOrEmpty(columnName) && 
                ((ordinal = _hluLayerStructure.Columns.IndexOf(columnName.Trim())) != -1))

                return _hluFieldMap[ordinal];
            else
                return -1;
        }

        private int FieldOrdinal(int columnOrdinal)
        {
            if ((_hluFieldMap != null) && (columnOrdinal > -1) && (columnOrdinal < _hluFieldMap.Length))
                return _hluFieldMap[columnOrdinal];
            else
                return -1;
        }

        private int ColumnOrdinal(string fieldName)
        {
            if ((_hluFieldNames != null) && !String.IsNullOrEmpty((fieldName = fieldName.Trim())))
                return System.Array.IndexOf<string>(_hluFieldNames, fieldName);
            else
                return -1;
        }

        private int FuzzyFieldOrdinal(string fieldName)
        {
            return FieldOrdinal(FuzzyColumnOrdinal(fieldName));
        }

        private int FuzzyColumnOrdinal(string fieldName)
        {
            if ((_hluFieldNames != null) && !String.IsNullOrEmpty((fieldName = fieldName.Trim())))
            {
                var q = from c in _hluLayerStructure.Columns.Cast<DataColumn>()
                        join s in fieldName.Split(new String[] { GISApp.HistoryAdditionalFieldsDelimiter }, 
                            StringSplitOptions.RemoveEmptyEntries).Distinct() on c.ColumnName equals s
                        select c.Ordinal;
                if (q.Count() == 1) return q.First();
            }
            return -1;
        }

        #endregion
    }
}