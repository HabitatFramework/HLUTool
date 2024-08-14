// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
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

using System.Collections.Generic;
using System;

namespace HLU.GISApplication.MapInfo
{
    static class MapInfoConstants
    {
        #region MapBasic 10.0 System defines

        public enum Colors
        {
            BLACK = 0,
            WHITE = 16777215,
            RED = 16711680,
            GREEN = 65280,
            BLUE = 255,
            CYAN = 65535,
            MAGENTA = 16711935,
            YELLOW = 16776960
        };

        /// <summary>
        /// Maximum length for character string
        /// </summary>
        public static int MAX_STRING_LENGTH = 32767;

        public enum ButtonPadInfo
        {
            BTNPAD_INFO_FLOATING = 1,
            BTNPAD_INFO_WIDTH = 2,
            BTNPAD_INFO_NBTNS = 3,
            BTNPAD_INFO_X = 4,
            BTNPAD_INFO_Y = 5,
            BTNPAD_INFO_WINID = 6
        };

        public enum ColumnInfo
        {
            COL_INFO_NAME = 1,
            COL_INFO_NUM = 2,
            COL_INFO_TYPE = 3,
            COL_INFO_WIDTH = 4,
            COL_INFO_DECPLACES = 5,
            COL_INFO_INDEXED = 6,
            COL_INFO_EDITABLE = 7
        };

        /// <summary>
        /// Returned by ColumnInfo() for COL_INFO_TYPE
        /// </summary>
        public enum ColumnType
        {
            COL_TYPE_CHAR = 1,
            COL_TYPE_DECIMAL = 2,
            COL_TYPE_INTEGER = 3,
            COL_TYPE_SMALLINT = 4,
            COL_TYPE_DATE = 5,
            COL_TYPE_LOGICAL = 6,
            COL_TYPE_GRAPHIC = 7,
            COL_TYPE_FLOAT = 8,
            COL_TYPE_TIME = 37,
            COL_TYPE_DATETIME = 38
        };

        public enum CommandInfo
        {
            CMD_INFO_X = 1,
            CMD_INFO_Y = 2,
            CMD_INFO_SHIFT = 3,
            CMD_INFO_CTRL = 4,
            CMD_INFO_X2 = 5,
            CMD_INFO_Y2 = 6,
            CMD_INFO_TOOLBTN = 7,
            CMD_INFO_MENUITEM = 8,
            CMD_INFO_WIN = 1,
            CMD_INFO_SELTYPE = 1,
            CMD_INFO_ROWID = 2,
            CMD_INFO_INTERRUPT = 3,
            CMD_INFO_STATUS = 1,
            CMD_INFO_MSG = 1000,
            CMD_INFO_DLG_OK = 1,
            CMD_INFO_DLG_DBL = 1,
            CMD_INFO_FIND_RC = 3,
            CMD_INFO_FIND_ROWID = 4,
            CMD_INFO_XCMD = 1,
            CMD_INFO_CUSTOM_OBJ = 1,
            CMD_INFO_TASK_SWITCH = 1,
            CMD_INFO_EDIT_TABLE = 1,
            CMD_INFO_EDIT_STATUS = 2,
            CMD_INFO_EDIT_ASK = 1,
            CMD_INFO_EDIT_SAVE = 2,
            CMD_INFO_EDIT_DISCARD = 3,
            CMD_INFO_HL_WINDOW_ID = 17,
            CMD_INFO_HL_TABLE_NAME = 18,
            CMD_INFO_HL_ROWID = 19,
            CMD_INFO_HL_LAYER_ID = 20,
            CMD_INFO_HL_FILE_NAME = 21
        };

        /// <summary>
        /// Returned by CommandInfo() for CMD_INFO_TASK_SWITCH
        /// </summary>
        public enum TaskSwitches
        {
            SWITCHING_OUT_OF_MAPINFO = 0,
            SWITCHING_INTO_MAPINFO = 1
        };

        public enum DateWindow
        {
            DATE_WIN_SESSION = 1,
            DATE_WIN_CURPROG = 2
        };

        public enum FileAttr
        {
            FILE_ATTR_MODE = 1,
            FILE_ATTR_FILESIZE = 2
        };

        /// <summary>
        /// Returned by FileAttr() for FILE_ATTR_MODE
        /// </summary>
        public enum FileAccessModes
        {
            MODE_INPUT = 0,
            MODE_OUTPUT = 1,
            MODE_APPEND = 2,
            MODE_RANDOM = 3,
            MODE_BINARY = 4,
        };

        public enum GetFolderPath
        {
            FOLDER_MI_APPDATA = -1,
            FOLDER_MI_LOCAL_APPDATA = -2,
            FOLDER_MI_PREFERENCE = -3,
            FOLDER_MI_COMMON_APPDATA = -4,
            FOLDER_APPDATA = 26,
            FOLDER_LOCAL_APPDATA = 28,
            FOLDER_COMMON_APPDATA = 35,
            FOLDER_COMMON_DOCS = 46,
            FOLDER_MYDOCS = 5,
            FOLDER_MYPICS = 39
        };

        public enum IntersectNodes
        {
            INCL_CROSSINGS = 1,
            INCL_COMMON = 6,
            INCL_ALL = 7
        };

        public enum LabelInfo
        {
            LABEL_INFO_OBJECT = 1,
            LABEL_INFO_POSITION = 2,
            LABEL_INFO_ANCHORX = 3,
            LABEL_INFO_ANCHORY = 4,
            LABEL_INFO_OFFSET = 5,
            LABEL_INFO_ROWID = 6,
            LABEL_INFO_TABLE = 7,
            LABEL_INFO_EDIT = 8,
            LABEL_INFO_EDIT_VISIBILITY = 9,
            LABEL_INFO_EDIT_ANCHOR = 10,
            LABEL_INFO_EDIT_OFFSET = 11,
            LABEL_INFO_EDIT_FONT = 12,
            LABEL_INFO_EDIT_PEN = 13,
            LABEL_INFO_EDIT_TEXT = 14,
            LABEL_INFO_EDIT_TEXTARROW = 15,
            LABEL_INFO_EDIT_ANGLE = 16,
            LABEL_INFO_EDIT_POSITION = 17,
            LABEL_INFO_EDIT_TEXTLINE = 18,
            LABEL_INFO_SELECT = 19,
            LABEL_INFO_DRAWN = 20,
            LABEL_INFO_ORIENTATION = 21
        };

        /// <summary>
        /// Codes passed to the GroupLayerInfo function to get info about a group layer.
        /// </summary>
        public enum GroupLayerInfoParams
        {
            GROUPLAYER_INFO_NAME = 1,
            GROUPLAYER_INFO_LAYERLIST_ID = 2,
            GROUPLAYER_INFO_DISPLAY = 3,
            GROUPLAYER_INFO_LAYERS = 4,
            GROUPLAYER_INFO_ALL_LAYERS = 5,
            GROUPLAYER_INFO_TOPLEVEL_LAYERS = 6,
            GROUPLAYER_INFO_PARENT_GROUP_ID = 7
        };

        /// <summary>
        /// Values returned by GroupLayerInfo() for GROUPLAYER_INFO_DISPLAY. 
        /// These defines correspond to the MapBasic defines in MAPBASIC.DEF. 
        /// If you alter these you must alter MAPBASIC.DEF.
        /// </summary>
        public enum GroupLayerInfoReturn
        {
            GROUPLAYER_INFO_DISPLAY_OFF = 0,
            GROUPLAYER_INFO_DISPLAY_ON = 1
        };

        /// <summary>
        /// Codes passed to the LayerListInfo function to help enumerating all layers in a Map.
        /// </summary>
        public enum LayerListInfoParams
        {
            LAYERLIST_INFO_TYPE = 1,
            LAYERLIST_INFO_NAME = 2,
            LAYERLIST_INFO_LAYER_ID = 3,
            LAYERLIST_INFO_GROUPLAYER_ID = 4
        };

        /// <summary>
        /// Values returned by LayerListInfo() for LAYERLIST_INFO_TYPE. 
        /// These defines correspond to the MapBasic defines in MAPBASIC.DEF. 
        /// If you alter these you must alter MAPBASIC.DEF.
        /// </summary>
        public enum LayerListInfoReturn
        {
            LAYERLIST_INFO_TYPE_LAYER = 0,
            LAYERLIST_INFO_TYPE_GROUP = 1
        };

        public enum LayerInfo
        {
            LAYER_INFO_NAME = 1,
            LAYER_INFO_EDITABLE = 2,
            LAYER_INFO_SELECTABLE = 3,
            LAYER_INFO_ZOOM_LAYERED = 4,
            LAYER_INFO_ZOOM_MIN = 5,
            LAYER_INFO_ZOOM_MAX = 6,
            LAYER_INFO_COSMETIC = 7,
            LAYER_INFO_PATH = 8,
            LAYER_INFO_DISPLAY = 9,
            LAYER_INFO_OVR_LINE = 10,
            LAYER_INFO_OVR_PEN = 11,
            LAYER_INFO_OVR_BRUSH = 12,
            LAYER_INFO_OVR_SYMBOL = 13,
            LAYER_INFO_OVR_FONT = 14,
            LAYER_INFO_LBL_EXPR = 15,
            LAYER_INFO_LBL_LT = 16,
            LAYER_INFO_LBL_CURFONT = 17,
            LAYER_INFO_LBL_FONT = 18,
            LAYER_INFO_LBL_PARALLEL = 19,
            LAYER_INFO_LBL_POS = 20,
            LAYER_INFO_ARROWS = 21,
            LAYER_INFO_NODES = 22,
            LAYER_INFO_CENTROIDS = 23,
            LAYER_INFO_TYPE = 24,
            LAYER_INFO_LBL_VISIBILITY = 25,
            LAYER_INFO_LBL_ZOOM_MIN = 26,
            LAYER_INFO_LBL_ZOOM_MAX = 27,
            LAYER_INFO_LBL_AUTODISPLAY = 28,
            LAYER_INFO_LBL_OVERLAP = 29,
            LAYER_INFO_LBL_DUPLICATES = 30,
            LAYER_INFO_LBL_OFFSET = 31,
            LAYER_INFO_LBL_MAX = 32,
            LAYER_INFO_LBL_PARTIALSEGS = 33,
            LAYER_INFO_HOTLINK_EXPR = 34,
            LAYER_INFO_HOTLINK_MODE = 35,
            LAYER_INFO_HOTLINK_RELATIVE = 36,
            LAYER_INFO_HOTLINK_COUNT = 37,
            LAYER_INFO_LBL_ORIENTATION = 38,
            LAYER_INFO_LAYER_ALPHA = 39,
            LAYER_INFO_LAYER_TRANSLUCENCY = 40,
            LAYER_INFO_LABEL_ALPHA = 41,
            LAYER_INFO_LAYERLIST_ID = 42,
            LAYER_INFO_PARENT_GROUP_ID = 43
        };

        /// <summary>
        /// Values returned by LayerInfo() for LAYER_INFO_LABEL_ORIENTATION and LABEL_INFO_ORIENTATION.
        /// </summary>
        public enum LayerInfoLabelOrientation
        {
            LAYER_INFO_LABEL_ORIENT_HORIZONTAL = 0,
            LAYER_INFO_LABEL_ORIENT_PARALLEL = 1,
            LAYER_INFO_LABEL_ORIENT_CURVED = 2
        };

        /// <summary>
        /// Display Modes, returned by LayerInfo() for LAYER_INFO_DISPLAY
        /// </summary>
        public enum LayerInfoDisplayModes
        {
            LAYER_INFO_DISPLAY_OFF = 0,
            LAYER_INFO_DISPLAY_GRAPHIC = 1,
            LAYER_INFO_DISPLAY_GLOBAL = 2,
            LAYER_INFO_DISPLAY_VALUE = 3
        };

        /// <summary>
        /// Label Linetypes, returned by LayerInfo() for LAYER_INFO_LBL_LT
        /// </summary>
        public enum LayerInfoLabelLinestypes
        {
            LAYER_INFO_LBL_LT_NONE = 0,
            LAYER_INFO_LBL_LT_SIMPLE = 1,
            LAYER_INFO_LBL_LT_ARROW = 2
        };

        /// <summary>
        /// Label Positions, returned by LayerInfo() for LAYER_INFO_LBL_POS
        /// </summary>
        public enum LayerInfoLabelPositions
        {
            LAYER_INFO_LBL_POS_CC = 0,
            LAYER_INFO_LBL_POS_TL = 1,
            LAYER_INFO_LBL_POS_TC = 2,
            LAYER_INFO_LBL_POS_TR = 3,
            LAYER_INFO_LBL_POS_CL = 4,
            LAYER_INFO_LBL_POS_CR = 5,
            LAYER_INFO_LBL_POS_BL = 6,
            LAYER_INFO_LBL_POS_BC = 7,
            LAYER_INFO_LBL_POS_BR = 8
        };

        /// <summary>
        /// Layer Types, returned by LayerInfo() for LAYER_INFO_TYPE
        /// </summary>
        public enum LayerInfoLayerTypes
        {
            LAYER_INFO_TYPE_NORMAL = 0,
            LAYER_INFO_TYPE_COSMETIC = 1,
            LAYER_INFO_TYPE_IMAGE = 2,
            LAYER_INFO_TYPE_THEMATIC = 3,
            LAYER_INFO_TYPE_GRID = 4,
            LAYER_INFO_TYPE_WMS = 5
        };

        /// <summary>
        /// Label visibility modes, from LayerInfo() for LAYER_INFO_LBL_VISIBILITY
        /// </summary>
        public enum LayerInfoLabelVisibility
        {
            LAYER_INFO_LBL_VIS_OFF = 1,
            LAYER_INFO_LBL_VIS_ZOOM = 2,
            LAYER_INFO_LBL_VIS_ON = 3
        };

        public enum HotlinkInfo
        {
            HOTLINK_INFO_EXPR = 1,
            HOTLINK_INFO_MODE = 2,
            HOTLINK_INFO_RELATIVE = 3,
            HOTLINK_INFO_ENABLED = 4,
            HOTLINK_INFO_ALIAS = 5
        };

        /// <summary>
        /// Hotlink activation modes, from LayerInfo() for LAYER_INFO_HOTLINK_MODE
        /// </summary>
        public enum HotlinkInfoactivationModes
        {
            HOTLINK_MODE_LABEL = 0,
            HOTLINK_MODE_OBJ = 1,
            HOTLINK_MODE_BOTH = 2
        };

        public enum LegendInfo
        {
            LEGEND_INFO_MAP_ID = 1,
            LEGEND_INFO_ORIENTATION = 2,
            LEGEND_INFO_NUM_FRAMES = 3,
            LEGEND_INFO_STYLE_SAMPLE_SIZE = 4
        };

        /// <summary>
        /// Orientation codes, returned by LegendInfo() for LEGEND_INFO_ORIENTATION
        /// </summary>
        public enum LegenInfoOrientations
        {
            ORIENTATION_PORTRAIT = 1,
            ORIENTATION_LANDSCAPE = 2,
            ORIENTATION_CUSTOM = 3
        };

        /// <summary>
        /// Style sample codes, from LegendInfo() for LEGEND_INFO_STYLE_SAMPLE_SIZE
        /// </summary>
        public enum LegenInfoStyleSampleSizes
        {
            STYLE_SAMPLE_SIZE_SMALL = 0,
            STYLE_SAMPLE_SIZE_LARGE = 1
        };

        public enum LegendFrameInfo
        {
            FRAME_INFO_TYPE = 1,
            FRAME_INFO_MAP_LAYER_ID = 2,
            FRAME_INFO_REFRESHABLE = 3,
            FRAME_INFO_POS_X = 4,
            FRAME_INFO_POS_Y = 5,
            FRAME_INFO_WIDTH = 6,
            FRAME_INFO_HEIGHT = 7,
            FRAME_INFO_TITLE = 8,
            FRAME_INFO_TITLE_FONT = 9,
            FRAME_INFO_SUBTITLE = 10,
            FRAME_INFO_SUBTITLE_FONT = 11,
            FRAME_INFO_BORDER_PEN = 12,
            FRAME_INFO_NUM_STYLES = 13,
            FRAME_INFO_VISIBLE = 14,
            FRAME_INFO_COLUMN = 15,
            FRAME_INFO_LABEL = 16
        };

        /// <summary>
        /// Frame Types, returned by LegendFrameInfo() for FRAME_INFO_TYPE
        /// </summary>
        public enum LegendFrameInfoFrameTypes
        {
            FRAME_TYPE_STYLE = 1,
            FRAME_TYPE_THEME = 2
        };

        /// <summary>
        /// Geocode Attributes, returned by GeocodeInfo()
        /// </summary>
        public enum GeocodeInforAttributes
        {
            GEOCODE_STREET_NAME = 1,
            GEOCODE_STREET_NUMBER = 2,
            GEOCODE_MUNICIPALITY = 3,
            GEOCODE_MUNICIPALITY2 = 4,
            GEOCODE_COUNTRY_SUBDIVISION = 5,
            GEOCODE_COUNTRY_SUBDIVISION2 = 6,
            GEOCODE_POSTAL_CODE = 7,
            GEOCODE_DICTIONARY = 9,
            GEOCODE_BATCH_SIZE = 10,
            GEOCODE_FALLBACK_GEOGRAPHIC = 11,
            GEOCODE_FALLBACK_POSTAL = 12,
            GEOCODE_OFFSET_CENTER = 13,
            GEOCODE_OFFSET_CENTER_UNITS = 14,
            GEOCODE_OFFSET_END = 15,
            GEOCODE_OFFSET_END_UNITS = 16,
            GEOCODE_MIXED_CASE = 17,
            GEOCODE_RESULT_MARK_MULTIPLE = 18,
            GEOCODE_COUNT_GEOCODED = 19,
            GEOCODE_COUNT_NOTGEOCODED = 20,
            GEOCODE_UNABLE_TO_CONVERT_DATA = 21,
            GEOCODE_MAX_BATCH_SIZE = 22,
            GEOCODE_PASSTHROUGH = 100
        };

        public enum GeocodeDictionary
        {
            DICTIONARY_ALL = 1,
            DICTIONARY_ADDRESS_ONLY = 2,
            DICTIONARY_USER_ONLY = 3,
            DICTIONARY_PREFER_ADDRESS = 4,
            DICTIONARY_PREFER_USER = 5
        };

        /// <summary>
        /// ISOGRAM Attributes, returned by IsogramInfo()
        /// </summary>
        public enum IsogramInfo
        {
            ISOGRAM_BANDING = 1,
            ISOGRAM_MAJOR_ROADS_ONLY = 2,
            ISOGRAM_RETURN_HOLES = 3,
            ISOGRAM_MAJOR_POLYGON_ONLY = 4,
            ISOGRAM_MAX_OFFROAD_DIST = 5,
            ISOGRAM_MAX_OFFROAD_DIST_UNITS = 6,
            ISOGRAM_SIMPLIFICATION_FACTOR = 7,
            ISOGRAM_DEFAULT_AMBIENT_SPEED = 8,
            ISOGRAM_AMBIENT_SPEED_DIST_UNIT = 9,
            ISOGRAM_AMBIENT_SPEED_TIME_UNIT = 10,
            ISOGRAM_PROPAGATION_FACTOR = 11,
            ISOGRAM_BATCH_SIZE = 12,
            ISOGRAM_POINTS_ONLY = 13,
            ISOGRAM_RECORDS_INSERTED = 14,
            ISOGRAM_RECORDS_NOTINSERTED = 15,
            ISOGRAM_MAX_BATCH_SIZE = 16,
            ISOGRAM_MAX_BANDS = 17,
            ISOGRAM_MAX_DISTANCE = 18,
            ISOGRAM_MAX_DISTANCE_UNITS = 19,
            ISOGRAM_MAX_TIME = 20,
            ISOGRAM_MAX_TIME_UNITS = 21
        };

        public enum LegendStyleInfo
        {
            LEGEND_STYLE_INFO_TEXT = 1,
            LEGEND_STYLE_INFO_FONT = 2,
            LEGEND_STYLE_INFO_OBJ = 3
        };

        public enum LocateFile
        {
            LOCATE_PREF_FILE = 0,
            LOCATE_DEF_WOR = 1,
            LOCATE_CLR_FILE = 2,
            LOCATE_PEN_FILE = 3,
            LOCATE_FNT_FILE = 4,
            LOCATE_ABB_FILE = 5,
            LOCATE_PRJ_FILE = 6,
            LOCATE_MNU_FILE = 7,
            LOCATE_CUSTSYMB_DIR = 8,
            LOCATE_THMTMPLT_DIR = 9,
            LOCATE_GRAPH_DIR = 10,
            LOCATE_WMS_SERVERLIST = 11,
            LOCATE_WFS_SERVERLIST = 12,
            LOCATE_GEOCODE_SERVERLIST = 13,
            LOCATE_ROUTING_SERVERLIST = 14,
            LOCATE_LAYOUT_TEMPLATE_DIR = 15
        };

        public enum Map3DInfo
        {
            MAP3D_INFO_SCALE = 1,
            MAP3D_INFO_RESOLUTION_X = 2,
            MAP3D_INFO_RESOLUTION_Y = 3,
            MAP3D_INFO_BACKGROUND = 4,
            MAP3D_INFO_UNITS = 5,
            MAP3D_INFO_LIGHT_X = 6,
            MAP3D_INFO_LIGHT_Y = 7,
            MAP3D_INFO_LIGHT_Z = 8,
            MAP3D_INFO_LIGHT_COLOR = 9,
            MAP3D_INFO_CAMERA_X = 10,
            MAP3D_INFO_CAMERA_Y = 11,
            MAP3D_INFO_CAMERA_Z = 12,
            MAP3D_INFO_CAMERA_FOCAL_X = 13,
            MAP3D_INFO_CAMERA_FOCAL_Y = 14,
            MAP3D_INFO_CAMERA_FOCAL_Z = 15,
            MAP3D_INFO_CAMERA_VU_1 = 16,
            MAP3D_INFO_CAMERA_VU_2 = 17,
            MAP3D_INFO_CAMERA_VU_3 = 18,
            MAP3D_INFO_CAMERA_VPN_1 = 19,
            MAP3D_INFO_CAMERA_VPN_2 = 20,
            MAP3D_INFO_CAMERA_VPN_3 = 21,
            MAP3D_INFO_CAMERA_CLIP_NEAR = 22,
            MAP3D_INFO_CAMERA_CLIP_FAR = 23
        };

        public enum MapperInfo
        {
            MAPPER_INFO_ZOOM = 1,
            MAPPER_INFO_SCALE = 2,
            MAPPER_INFO_CENTERX = 3,
            MAPPER_INFO_CENTERY = 4,
            MAPPER_INFO_MINX = 5,
            MAPPER_INFO_MINY = 6,
            MAPPER_INFO_MAXX = 7,
            MAPPER_INFO_MAXY = 8,
            MAPPER_INFO_LAYERS = 9,
            MAPPER_INFO_EDIT_LAYER = 10,
            MAPPER_INFO_XYUNITS = 11,
            MAPPER_INFO_DISTUNITS = 12,
            MAPPER_INFO_AREAUNITS = 13,
            MAPPER_INFO_SCROLLBARS = 14,
            MAPPER_INFO_DISPLAY = 15,
            MAPPER_INFO_NUM_THEMATIC = 16,
            MAPPER_INFO_COORDSYS_CLAUSE = 17,
            MAPPER_INFO_COORDSYS_NAME = 18,
            MAPPER_INFO_MOVE_DUPLICATE_NODES = 19,
            MAPPER_INFO_DIST_CALC_TYPE = 20,
            MAPPER_INFO_DISPLAY_DMS = 21,
            MAPPER_INFO_COORDSYS_CLAUSE_WITH_BOUNDS = 22,
            MAPPER_INFO_CLIP_TYPE = 23,
            MAPPER_INFO_CLIP_REGION = 24,
            MAPPER_INFO_REPROJECTION = 25,
            MAPPER_INFO_RESAMPLING = 26,
            MAPPER_INFO_MERGE_MAP = 27,
            MAPPER_INFO_ALL_LAYERS = 28,
            MAPPER_INFO_GROUPLAYERS = 29,
            MAPPER_INFO_NUM_ADORNMENTS = 200,
            MAPPER_INFO_ADORNMENT = 200
        };

        /// <summary>
        /// Display Modes, returned by MapperInfo() for MAPPER_INFO_DISPLAY_DMS
        /// </summary>
        public enum MapperInfoDisplayDMS
        {
            MAPPER_INFO_DISPLAY_DECIMAL = 0,
            MAPPER_INFO_DISPLAY_DEGMINSEC = 1,
            MAPPER_INFO_DISPLAY_MGRS = 2,
            MAPPER_INFO_DISPLAY_USNG_WGS84 = 3,
            MAPPER_INFO_DISPLAY_USNG_NAD27 = 4
        };

        /// <summary>
        /// Display Modes, returned by MapperInfo() for MAPPER_INFO_DISPLAY
        /// </summary>
        public enum MapperInfoDisplayModes
        {
            MAPPER_INFO_DISPLAY_SCALE = 0,
            MAPPER_INFO_DISPLAY_ZOOM = 1,
            MAPPER_INFO_DISPLAY_POSITION = 2
        };

        /// <summary>
        /// Distance Calculation Types from MapperInfo() for MAPPER_INFO_DIST_CALC_TYPE
        /// </summary>
        public enum MapperInfoDistCalcType
        {
            MAPPER_INFO_DIST_SPHERICAL = 0,
            MAPPER_INFO_DIST_CARTESIAN = 1
        };

        /// <summary>
        /// Clip Types, returned by MapperInfo() for MAPPER_INFO_CLIP_TYPE
        /// </summary>
        public enum MapperInfoClipType
        {
            MAPPER_INFO_CLIP_DISPLAY_ALL = 0,
            MAPPER_INFO_CLIP_DISPLAY_POLYOBJ = 1,
            MAPPER_INFO_CLIP_OVERLAY = 2
        };

        /// <summary>
        /// MenuItemInfoByID() and MenuItemInfoByHandler() defines
        /// </summary>
        public enum MenuItemInfoByID
        {
            MENUITEM_INFO_ENABLED = 1,
            MENUITEM_INFO_CHECKED = 2,
            MENUITEM_INFO_CHECKABLE = 3,
            MENUITEM_INFO_SHOWHIDEABLE = 4,
            MENUITEM_INFO_ACCELERATOR = 5,
            MENUITEM_INFO_TEXT = 6,
            MENUITEM_INFO_HELPMSG = 7,
            MENUITEM_INFO_HANDLER = 8,
            MENUITEM_INFO_ID = 9
        };

        public enum ObjectGeography
        {
            OBJ_GEO_MINX = 1,
            OBJ_GEO_LINEBEGX = 1,
            OBJ_GEO_POINTX = 1,
            OBJ_GEO_MINY = 2,
            OBJ_GEO_LINEBEGY = 2,
            OBJ_GEO_POINTY = 2,
            OBJ_GEO_MAXX = 3,
            OBJ_GEO_LINEENDX = 3,
            OBJ_GEO_MAXY = 4,
            OBJ_GEO_LINEENDY = 4,
            OBJ_GEO_ARCBEGANGLE = 5,
            OBJ_GEO_TEXTLINEX = 5,
            OBJ_GEO_ROUNDRADIUS = 5,
            OBJ_GEO_CENTROID = 5,
            OBJ_GEO_ARCENDANGLE = 6,
            OBJ_GEO_TEXTLINEY = 6,
            OBJ_GEO_TEXTANGLE = 7,
            OBJ_GEO_POINTZ = 8,
            OBJ_GEO_POINTM = 9
        };

        public enum ObjectInfo
        {
            OBJ_INFO_TYPE = 1,
            OBJ_INFO_PEN = 2,
            OBJ_INFO_SYMBOL = 2,
            OBJ_INFO_TEXTFONT = 2,
            OBJ_INFO_BRUSH = 3,
            OBJ_INFO_NPNTS = 20,
            OBJ_INFO_TEXTSTRING = 3,
            OBJ_INFO_SMOOTH = 4,
            OBJ_INFO_FRAMEWIN = 4,
            OBJ_INFO_NPOLYGONS = 21,
            OBJ_INFO_TEXTSPACING = 4,
            OBJ_INFO_TEXTJUSTIFY = 5,
            OBJ_INFO_FRAMETITLE = 6,
            OBJ_INFO_TEXTARROW = 6,
            OBJ_INFO_FILLFRAME = 7,
            OBJ_INFO_REGION = 8,
            OBJ_INFO_PLINE = 9,
            OBJ_INFO_MPOINT = 10,
            OBJ_INFO_NONEMPTY = 11,
            OBJ_INFO_Z_UNIT_SET = 12,
            OBJ_INFO_Z_UNIT = 13,
            OBJ_INFO_HAS_Z = 14,
            OBJ_INFO_HAS_M = 15
        };

        /// <summary>
        /// Object types, returned by ObjectInfo() for OBJ_INFO_TYPE
        /// </summary>
        public enum ObjectInfoTypes
        {
            OBJ_TYPE_ARC = 1,
            OBJ_TYPE_ELLIPSE = 2,
            OBJ_TYPE_LINE = 3,
            OBJ_TYPE_PLINE = 4,
            OBJ_TYPE_POINT = 5,
            OBJ_TYPE_FRAME = 6,
            OBJ_TYPE_REGION = 7,
            OBJ_TYPE_RECT = 8,
            OBJ_TYPE_ROUNDRECT = 9,
            OBJ_TYPE_TEXT = 10,
            OBJ_TYPE_MPOINT = 11,
            OBJ_TYPE_COLLECTION = 12
        };

        /// <summary>
        ///  Constants passed to RegionInfo.
        /// </summary>
        public static int REGION_INFO_IS_CLOCKWISE = 1;

        public enum PrismMapInfo
        {
            PRISMMAP_INFO_SCALE = 1,
            PRISMMAP_INFO_BACKGROUND = 4,
            PRISMMAP_INFO_LIGHT_X = 6,
            PRISMMAP_INFO_LIGHT_Y = 7,
            PRISMMAP_INFO_LIGHT_Z = 8,
            PRISMMAP_INFO_LIGHT_COLOR = 9,
            PRISMMAP_INFO_CAMERA_X = 10,
            PRISMMAP_INFO_CAMERA_Y = 11,
            PRISMMAP_INFO_CAMERA_Z = 12,
            PRISMMAP_INFO_CAMERA_FOCAL_X = 13,
            PRISMMAP_INFO_CAMERA_FOCAL_Y = 14,
            PRISMMAP_INFO_CAMERA_FOCAL_Z = 15,
            PRISMMAP_INFO_CAMERA_VU_1 = 16,
            PRISMMAP_INFO_CAMERA_VU_2 = 17,
            PRISMMAP_INFO_CAMERA_VU_3 = 18,
            PRISMMAP_INFO_CAMERA_VPN_1 = 19,
            PRISMMAP_INFO_CAMERA_VPN_2 = 20,
            PRISMMAP_INFO_CAMERA_VPN_3 = 21,
            PRISMMAP_INFO_CAMERA_CLIP_NEAR = 22,
            PRISMMAP_INFO_CAMERA_CLIP_FAR = 23,
            PRISMMAP_INFO_INFOTIP_EXPR = 24
        };

        public enum SearchInfo
        {
            SEARCH_INFO_TABLE = 1,
            SEARCH_INFO_ROW = 2
        };

        public enum SelectionInfo
        {
            SEL_INFO_TABLENAME = 1,
            SEL_INFO_SELNAME = 2,
            SEL_INFO_NROWS = 3
        };

        #region Server statement and function defines
        
        public enum ServerReturnCodes
        {
            SRV_SUCCESS = 0,
            SRV_SUCCESS_WITH_INFO = 1,
            SRV_ERROR = -1,
            SRV_INVALID_HANDLE = -2,
            SRV_NEED_DATA = 99,
            SRV_NO_MORE_DATA = 100
        };

        /// <summary>
        /// Special values for the status associated with a fetched value
        /// </summary>
        public enum ServerFetchedValueStatus
        {
            SRV_NULL_DATA = -1,
            SRV_TRUNCATED_DATA = -2
        };

        public enum ServerColumnInfo
        {
            SRV_COL_INFO_NAME = 1,
            SRV_COL_INFO_TYPE = 2,
            SRV_COL_INFO_WIDTH = 3,
            SRV_COL_INFO_PRECISION = 4,
            SRV_COL_INFO_SCALE = 5,
            SRV_COL_INFO_VALUE = 6,
            SRV_COL_INFO_STATUS = 7,
            SRV_COL_INFO_ALIAS = 8
        };

        /// <summary>
        /// Column types, returned by Server_ColumnInfo() for SRV_COL_INFO_TYPE
        /// </summary>
        public enum ServerColumnTypes
        {
            SRV_COL_TYPE_NONE = 0,
            SRV_COL_TYPE_CHAR = 1,
            SRV_COL_TYPE_DECIMAL = 2,
            SRV_COL_TYPE_INTEGER = 3,
            SRV_COL_TYPE_SMALLINT = 4,
            SRV_COL_TYPE_DATE = 5,
            SRV_COL_TYPE_LOGICAL = 6,
            SRV_COL_TYPE_FLOAT = 8,
            SRV_COL_TYPE_FIXED_LEN_STRING = 16,
            SRV_COL_TYPE_BIN_STRING = 17
        };

        public enum ServerDriverInfo
        {
            SRV_DRV_INFO_NAME = 1,
            SRV_DRV_INFO_NAME_LIST = 2,
            SRV_DRV_DATA_SOURCE = 3
        };

        public enum ServerConnectInfo
        {
            SRV_CONNECT_INFO_DRIVER_NAME = 1,
            SRV_CONNECT_INFO_DB_NAME = 2,
            SRV_CONNECT_INFO_SQL_USER_ID = 3,
            SRV_CONNECT_INFO_DS_NAME = 4,
            SRV_CONNECT_INFO_QUOTE_CHAR = 5
        };

        /// <summary>
        /// Fetch Directions (used by ServerFetch function in some code libraries)
        /// </summary>
        public enum ServerFetchDirections
        {
            SRV_FETCH_NEXT = -1,
            SRV_FETCH_PREV = -2,
            SRV_FETCH_FIRST = -3,
            SRV_FETCH_LAST = -4
        };

        /// <summary>
        ///Oracle workspace manager
        /// <summary>
        public enum ServerOracleWSManager
        {
            SRV_WM_HIST_NONE = 0,
            SRV_WM_HIST_OVERWRITE = 1,
            SRV_WM_HIST_NO_OVERWRITE = 2
        };

        public enum SessionInfo
        {
            SESSION_INFO_COORDSYS_CLAUSE = 1,
            SESSION_INFO_DISTANCE_UNITS = 2,
            SESSION_INFO_AREA_UNITS = 3,
            SESSION_INFO_PAPER_UNITS = 4
        };
        #endregion

        public enum SetNextDocumentStyles
        {
            WIN_STYLE_STANDARD = 0,
            WIN_STYLE_CHILD = 1,
            WIN_STYLE_POPUP_FULLCAPTION = 2,
            WIN_STYLE_POPUP = 3
        };

        public enum StringCompare
        {
            STR_LT = -1,
            STR_GT = 1,
            STR_EQ = 0
        };

        public enum StyleAttr
        {
            PEN_WIDTH = 1,
            PEN_PATTERN = 2,
            PEN_COLOR = 4,
            PEN_INDEX = 5,
            PEN_INTERLEAVED = 6,
            BRUSH_PATTERN = 1,
            BRUSH_FORECOLOR = 2,
            BRUSH_BACKCOLOR = 3,
            FONT_NAME = 1,
            FONT_STYLE = 2,
            FONT_POINTSIZE = 3,
            FONT_FORECOLOR = 4,
            FONT_BACKCOLOR = 5,
            SYMBOL_CODE = 1,
            SYMBOL_COLOR = 2,
            SYMBOL_POINTSIZE = 3,
            SYMBOL_ANGLE = 4,
            SYMBOL_FONT_NAME = 5,
            SYMBOL_FONT_STYLE = 6,
            SYMBOL_KIND = 7,
            SYMBOL_CUSTOM_NAME = 8,
            SYMBOL_CUSTOM_STYLE = 9
        };

        /// <summary>
        /// Symbol kinds returned by StyleAttr() for SYMBOL_KIND
        /// </summary>
        public enum StyleAttrSymbolKinds
        {
            SYMBOL_KIND_VECTOR = 1,
            SYMBOL_KIND_FONT = 2,
            SYMBOL_KIND_CUSTOM = 3
        };

        public enum SystemInfo
        {
            SYS_INFO_PLATFORM = 1,
            SYS_INFO_APPVERSION = 2,
            SYS_INFO_MIVERSION = 3,
            SYS_INFO_RUNTIME = 4,
            SYS_INFO_CHARSET = 5,
            SYS_INFO_COPYPROTECTED = 6,
            SYS_INFO_APPLICATIONWND = 7,
            SYS_INFO_DDESTATUS = 8,
            SYS_INFO_MAPINFOWND = 9,
            SYS_INFO_NUMBER_FORMAT = 10,
            SYS_INFO_DATE_FORMAT = 11,
            SYS_INFO_DIG_INSTALLED = 12,
            SYS_INFO_DIG_MODE = 13,
            SYS_INFO_MIPLATFORM = 14,
            SYS_INFO_MDICLIENTWND = 15,
            SYS_INFO_PRODUCTLEVEL = 16,
            SYS_INFO_APPIDISPATCH = 17,
            SYS_INFO_MIBUILD_NUM = 18
        };

        /// <summary>
        /// Platform, returned by SystemInfo() for SYS_INFO_PLATFORM
        /// </summary>
        public enum SystemInfoPlatform
        {
            PLATFORM_SPECIAL = 0,
            PLATFORM_WIN = 1,
            PLATFORM_MAC = 2,
            PLATFORM_MOTIF = 3,
            PLATFORM_X11 = 4,
            PLATFORM_XOL = 5
        };

        /// <summary>
        /// Version, returned by SystemInfo() for SYS_INFO_MIPLATFORM
        /// </summary>
        public enum SystemInfoVersion
        {
            MIPLATFORM_SPECIAL = 0,
            MIPLATFORM_WIN16 = 1,
            MIPLATFORM_WIN32 = 2,
            MIPLATFORM_POWERMAC = 3,
            MIPLATFORM_MAC68K = 4,
            MIPLATFORM_HP = 5,
            MIPLATFORM_SUN = 6
        };

        public enum TableInfo
        {
            TAB_INFO_NAME = 1,
            TAB_INFO_NUM = 2,
            TAB_INFO_TYPE = 3,
            TAB_INFO_NCOLS = 4,
            TAB_INFO_MAPPABLE = 5,
            TAB_INFO_READONLY = 6,
            TAB_INFO_TEMP = 7,
            TAB_INFO_NROWS = 8,
            TAB_INFO_EDITED = 9,
            TAB_INFO_FASTEDIT = 10,
            TAB_INFO_UNDO = 11,
            TAB_INFO_MAPPABLE_TABLE = 12,
            TAB_INFO_USERMAP = 13,
            TAB_INFO_USERBROWSE = 14,
            TAB_INFO_USERCLOSE = 15,
            TAB_INFO_USEREDITABLE = 16,
            TAB_INFO_USERREMOVEMAP = 17,
            TAB_INFO_USERDISPLAYMAP = 18,
            TAB_INFO_TABFILE = 19,
            TAB_INFO_MINX = 20,
            TAB_INFO_MINY = 21,
            TAB_INFO_MAXX = 22,
            TAB_INFO_MAXY = 23,
            TAB_INFO_SEAMLESS = 24,
            TAB_INFO_COORDSYS_MINX = 25,
            TAB_INFO_COORDSYS_MINY = 26,
            TAB_INFO_COORDSYS_MAXX = 27,
            TAB_INFO_COORDSYS_MAXY = 28,
            TAB_INFO_COORDSYS_CLAUSE = 29,
            TAB_INFO_COORDSYS_NAME = 30,
            TAB_INFO_NREFS = 31,
            TAB_INFO_SUPPORT_MZ = 32,
            TAB_INFO_Z_UNIT_SET = 33,
            TAB_INFO_Z_UNIT = 34,
            TAB_INFO_BROWSER_LIST = 35,
            TAB_INFO_THEME_METADATA = 36,
            TAB_INFO_COORDSYS_CLAUSE_WITHOUT_BOUNDS = 37
        };

        /// <summary>
        /// Table type defines, returned by TableInfo() for TAB_INFO_TYPE
        /// </summary>
        public enum TableInfoType
        {
            TAB_TYPE_BASE = 1,
            TAB_TYPE_RESULT = 2,
            TAB_TYPE_VIEW = 3,
            TAB_TYPE_IMAGE = 4,
            TAB_TYPE_LINKED = 5,
            TAB_TYPE_WMS = 6,
            TAB_TYPE_WFS = 7,
            TAB_TYPE_FME = 8
        };

        public enum RasterTableInfo
        {
            RASTER_TAB_INFO_IMAGE_NAME = 1,
            RASTER_TAB_INFO_WIDTH = 2,
            RASTER_TAB_INFO_HEIGHT = 3,
            RASTER_TAB_INFO_IMAGE_TYPE = 4,
            RASTER_TAB_INFO_BITS_PER_PIXEL = 5,
            RASTER_TAB_INFO_IMAGE_CLASS = 6,
            RASTER_TAB_INFO_NUM_CONTROL_POINTS = 7,
            RASTER_TAB_INFO_BRIGHTNESS = 8,
            RASTER_TAB_INFO_CONTRAST = 9,
            RASTER_TAB_INFO_GREYSCALE = 10,
            RASTER_TAB_INFO_DISPLAY_TRANSPARENT = 11,
            RASTER_TAB_INFO_TRANSPARENT_COLOR = 12,
            RASTER_TAB_INFO_ALPHA = 13
        };

        /// <summary>
        /// Image type defines returned by RasterTableInfo() for RASTER_TAB_INFO_IMAGE_TYPE
        /// </summary>
        public enum RasterTableInfoImageTypes
        {
            IMAGE_TYPE_RASTER = 0,
            IMAGE_TYPE_GRID = 1
        };

        /// <summary>
        /// Image class defines returned by RasterTableInfo() for RASTER_TAB_INFO_IMAGE_CLASS
        /// </summary>
        public enum RasterTableInfoImageClasses
        {
            IMAGE_CLASS_BILEVEL = 0,
            IMAGE_CLASS_GREYSCALE = 1,
            IMAGE_CLASS_PALETTE = 2,
            IMAGE_CLASS_RGB = 3
        };

        public enum GridTableInfo
        {
            GRID_TAB_INFO_MIN_VALUE = 1,
            GRID_TAB_INFO_MAX_VALUE = 2,
            GRID_TAB_INFO_HAS_HILLSHADE = 3
        };

        public enum  ControlPointInfo
        {
            RASTER_CONTROL_POINT_X = 1,
            RASTER_CONTROL_POINT_Y = 2,
            GEO_CONTROL_POINT_X = 3,
            GEO_CONTROL_POINT_Y = 4,
            TAB_GEO_CONTROL_POINT_X = 5,
            TAB_GEO_CONTROL_POINT_Y = 6
        };

        public enum WindowInfo
        {
            WIN_INFO_NAME = 1,
            WIN_INFO_TYPE = 3,
            WIN_INFO_WIDTH = 4,
            WIN_INFO_HEIGHT = 5,
            WIN_INFO_X = 6,
            WIN_INFO_Y = 7,
            WIN_INFO_TOPMOST = 8,
            WIN_INFO_STATE = 9,
            WIN_INFO_TABLE = 10,
            WIN_INFO_LEGENDS_MAP = 10,
            WIN_INFO_ADORNMENTS_MAP = 10,
            WIN_INFO_OPEN = 11,
            WIN_INFO_WND = 12,
            WIN_INFO_WINDOWID = 13,
            WIN_INFO_WORKSPACE = 14,
            WIN_INFO_CLONEWINDOW = 15,
            WIN_INFO_SYSMENUCLOSE = 16,
            WIN_INFO_AUTOSCROLL = 17,
            WIN_INFO_SMARTPAN = 18,
            WIN_INFO_SNAPMODE = 19,
            WIN_INFO_SNAPTHRESHOLD = 20,
            WIN_INFO_PRINTER_NAME = 21,
            WIN_INFO_PRINTER_ORIENT = 22,
            WIN_INFO_PRINTER_COPIES = 23,
            WIN_INFO_PRINTER_PAPERSIZE = 24,
            WIN_INFO_PRINTER_LEFTMARGIN = 25,
            WIN_INFO_PRINTER_RIGHTMARGIN = 26,
            WIN_INFO_PRINTER_TOPMARGIN = 27,
            WIN_INFO_PRINTER_BOTTOMMARGIN = 28,
            WIN_INFO_PRINTER_BORDER = 29,
            WIN_INFO_PRINTER_TRUECOLOR = 30,
            WIN_INFO_PRINTER_DITHER = 31,
            WIN_INFO_PRINTER_METHOD = 32,
            WIN_INFO_PRINTER_TRANSPRASTER = 33,
            WIN_INFO_PRINTER_TRANSPVECTOR = 34,
            WIN_INFO_EXPORT_BORDER = 35,
            WIN_INFO_EXPORT_TRUECOLOR = 36,
            WIN_INFO_EXPORT_DITHER = 37,
            WIN_INFO_EXPORT_TRANSPRASTER = 38,
            WIN_INFO_EXPORT_TRANSPVECTOR = 39,
            WIN_INFO_PRINTER_SCALE_PATTERNS = 40,
            WIN_INFO_EXPORT_ANTIALIASING = 41,
            WIN_INFO_EXPORT_THRESHOLD = 42,
            WIN_INFO_EXPORT_MASKSIZE = 43,
            WIN_INFO_EXPORT_FILTER = 44,
            WIN_INFO_ENHANCED_RENDERING = 45,
            WIN_INFO_SMOOTH_TEXT = 46,
            WIN_INFO_SMOOTH_IMAGE = 47,
            WIN_INFO_SMOOTH_VECTOR = 48
        };

        /// <summary>
        /// Window types, returned by WindowInfo() for WIN_INFO_TYPE
        /// </summary>
        public enum WindowInfoWindowTypes
        {
            WIN_MAPPER = 1,
            WIN_BROWSER = 2,
            WIN_LAYOUT = 3,
            WIN_GRAPH = 4,
            WIN_BUTTONPAD = 19,
            WIN_TOOLBAR = 25,
            WIN_CART_LEGEND = 27,
            WIN_3DMAP = 28,
            WIN_ADORNMENT = 32,
            WIN_HELP = 1001,
            WIN_MAPBASIC = 1002,
            WIN_MESSAGE = 1003,
            WIN_RULER = 1007,
            WIN_INFO = 1008,
            WIN_LEGEND = 1009,
            WIN_STATISTICS = 1010,
            WIN_MAPINFO = 1011
        };

        /// <summary>
        /// Version 2 window types no longer used in version 3 or later versions
        /// </summary>
        public enum WindowInfoWindowTypesV2
        {
            WIN_TOOLPICKER = 1004,
            WIN_PENPICKER = 1005,
            WIN_SYMBOLPICKER = 1006
        };

        /// <summary>
        /// Window states, returned by WindowInfo() for WIN_INFO_STATE
        /// </summary>
        public enum WindowInfoWindowStates
        {
            WIN_STATE_NORMAL = 0,
            WIN_STATE_MINIMIZED = 1,
            WIN_STATE_MAXIMIZED = 2
        };

        /// <summary>
        /// Print orientation, returned by WindowInfo() for WIN_INFO_PRINTER_ORIENT
        /// </summary>
        public enum WindowInfoPrintOrientations
        {
            WIN_PRINTER_PORTRAIT = 1,
            WIN_PRINTER_LANDSCAPE = 2
        };

        /// <summary>
        /// Antialiasing filters, returned by WindowInfo() for WIN_INFO_EXPORT_FILTER
        /// </summary>
        public enum WindowInfoAntialiasingFilters
        {
            FILTER_VERTICALLY_AND_HORIZONTALLY = 0,
            FILTER_ALL_DIRECTIONS_1 = 1,
            FILTER_ALL_DIRECTIONS_2 = 2,
            FILTER_DIAGONALLY = 3,
            FILTER_HORIZONTALLY = 4,
            FILTER_VERTICALLY = 5
        };
        
        #endregion

        #region MapInfo version 9.5 - Menu Item Definitions
       
        public enum MenuDef95File
        {
            M_FILE_NEW = 101,
            M_FILE_OPEN = 102,
            M_FILE_OPEN_WMS = 118,
            M_FILE_OPEN_WFS = 119,
            M_FILE_OPEN_ODBC_CONN = 125,
            M_FILE_OPEN_UNIVERSAL_DATA = 126,
            M_FILE_CLOSE = 103,
            M_FILE_CLOSE_ALL = 104,
            M_FILE_CLOSE_ODBC = 124,
            M_FILE_SAVE = 105,
            M_FILE_SAVE_COPY_AS = 106,
            M_FILE_SAVE_QUERY = 117,
            M_FILE_SAVE_WORKSPACE = 109,
            M_FILE_SAVE_WINDOW_AS = 609,
            M_FILE_REVERT = 107,
            M_FILE_PAGE_SETUP = 111,
            M_FILE_PRINT = 112,
            M_FILE_EXIT = 113
        };

        public enum MenuDef95Edit
        {
            M_EDIT_UNDO = 201,
            M_EDIT_CUT = 202,
            M_EDIT_COPY = 203,
            M_EDIT_PASTE = 204,
            M_EDIT_CLEAR = 205,
            M_EDIT_CLEAROBJ = 206,
            M_EDIT_RESHAPE = 1601,
            M_EDIT_NEW_ROW = 702,
            M_EDIT_GETINFO = 207
        };

        public enum MenuDef95Tools
        {
            M_TOOLS_CRYSTAL_REPORTS_NEW = 1802,
            M_TOOLS_CRYSTAL_REPORTS_OPEN = 1803,
            M_TOOLS_RUN = 110,
            M_TOOLS_TOOL_MANAGER = 1801
        };

        public enum MenuDef95Objects
        {
            M_OBJECTS_SET_TARGET = 1610,
            M_OBJECTS_CLEAR_TARGET = 1611,
            M_OBJECTS_COMBINE = 1605,
            M_OBJECTS_DISAGG = 1621,
            M_OBJECTS_BUFFER = 1606,
            M_OBJECTS_CONVEX_HULL = 1616,
            M_OBJECTS_ENCLOSE = 1617,
            M_OBJECTS_VORONOI = 1622,
            M_OBJECTS_SPLIT = 1612,
            M_OBJECTS_ERASE = 1613,
            M_OBJECTS_ERASE_OUT = 1614,
            M_OBJECTS_POLYLINE_SPLIT = 1623,
            M_OBJECTS_POLYLINE_SPLIT_AT_NODE = 1626,
            M_OBJECTS_DRIVE_REGION = 1627,
            M_OBJECTS_OVERLAY = 1615,
            M_OBJECTS_CHECK_REGIONS = 1618,
            M_OBJECTS_CLEAN = 1619,
            M_OBJECTS_SNAP = 1620,
            M_OBJECTS_OFFSET = 1624,
            M_OBJECTS_ROTATE = 1625,
            M_OBJECTS_SMOOTH = 1602,
            M_OBJECTS_UNSMOOTH = 1603,
            M_OBJECTS_CVT_PGON = 1607,
            M_OBJECTS_CVT_PLINE = 1604
        };

        public enum MenuDef95Query
        {
            M_QUERY_SELECT = 301,
            M_QUERY_SQLQUERY = 302,
            M_QUERY_SELECTALL = 303,
            M_QUERY_INVERTSELECT = 311,
            M_QUERY_UNSELECT = 304,
            M_QUERY_FIND = 305,
            M_QUERY_FIND_SELECTION = 306,
            M_QUERY_CALC_STATISTICS = 309,
            M_QUERY_FIND_SELECTION_CURRENT_MAP = 312,
            M_QUERY_FIND_ADDRESS = 313
        };

        public enum MenuDef95TableMaintenanceRaster
        {
            M_TABLE_UPDATE_COLUMN = 405,
            M_TABLE_APPEND = 411,
            M_TABLE_GEOCODE = 407,
            M_TABLE_CREATE_POINTS = 408,
            M_TABLE_MERGE_USING_COLUMN = 406,
            M_TABLE_BUFFER = 419,
            M_TABLE_VORONOI = 420,
            M_TABLE_IMPORT = 401,
            M_TABLE_EXPORT = 402,
            M_TABLE_MODIFY_STRUCTURE = 404,
            M_TABLE_DELETE = 409,
            M_TABLE_RENAME = 410,
            M_TABLE_PACK = 403,
            M_TABLE_MAKEMAPPABLE = 415,
            M_TABLE_CHANGESYMBOL = 418,
            M_TABLE_UNLINK = 416,
            M_TABLE_REFRESH = 417,
            M_TABLE_RASTER_STYLE = 414,
            M_TABLE_RASTER_REG = 413,
            M_TOOLS_RASTER_REG = 1730,
            M_TABLE_WMS_PROPS = 421,
            M_TABLE_WFS_REFRESH = 422,
            M_TABLE_WFS_PROPS = 423,
            M_TABLE_WEB_GEOCODE = 424,
            M_TABLE_DRIVE_REGION = 425,
            M_TABLE_UNIVERSAL_DATA_REFRESH = 426,
            M_ORACLE_CREATE_WORKSPACE = 1804,
            M_ORACLE_DELETE_WORKSPACE = 1805,
            M_ORACLE_VERSION_ENABLE_ON = 1806,
            M_ORACLE_VERSION_ENABLE_OFF = 1807,
            M_ORACLE_MERGE_PARENT = 1808,
            M_ORACLE_REFRESH_FROM_PARENT = 1809
        };

        public enum MenuDef95OptionsPreferences
        {
            M_FORMAT_PICK_LINE = 501,
            M_FORMAT_PICK_FILL = 502,
            M_FORMAT_PICK_SYMBOL = 503,
            M_FORMAT_PICK_FONT = 504,
            M_WINDOW_BUTTONPAD = 605,
            M_WINDOW_LEGEND = 606,
            M_WINDOW_STATISTICS = 607,
            M_WINDOW_MAPBASIC = 608,
            M_WINDOW_STATUSBAR = 616,
            M_FORMAT_CUSTOM_COLORS = 617,
            M_EDIT_PREFERENCES = 208,
            M_EDIT_PREFERENCES_SYSTEM = 210,
            M_EDIT_PREFERENCES_MAP = 212,
            M_EDIT_PREFERENCES_LEGEND = 215,
            M_EDIT_PREFERENCES_FILE = 211,
            M_EDIT_PREFERENCES_COUNTRY = 213,
            M_EDIT_PREFERENCES_PATH = 214,
            M_EDIT_PREFERENCES_OUTPUT = 216,
            M_EDIT_PREFERENCES_PRINTER = 217,
            M_EDIT_PREFERENCES_STYLES = 218,
            M_EDIT_PREFERENCES_IMAGE_PROC = 219,
            M_EDIT_PREFERENCES_WEBSERVICES = 220,
            M_EDIT_PREFERENCES_LAYOUT = 221
        };

        /// <summary>
        /// Note: the  2nd through  80th windows can be accessed as (M_WINDOW_FIRST+i-1)
        /// </summary>
        public enum MenuDef95Window
        {
            M_WINDOW_BROWSE = 601,
            M_WINDOW_MAP = 602,
            M_WINDOW_GRAPH = 603,
            M_WINDOW_LAYOUT = 604,
            M_WINDOW_REDISTRICT = 615,
            M_WINDOW_REDRAW = 610,
            M_WINDOW_TILE = 611,
            M_WINDOW_CASCADE = 612,
            M_WINDOW_ARRANGEICONS = 613,
            M_WINDOW_MORE = 614,
            M_WINDOW_FIRST = 620
        };

        public enum MenuDef95Help
        {
            M_HELP_CONTEXTSENSITIVE = 1201,
            M_HELP_CONTENTS = 1202,
            M_HELP_SEARCH = 1203,
            M_HELP_USE_HELP = 1204,
            M_HELP_ABOUT = 1205,
            M_HELP_HELPMODE = 1206,
            M_HELP_TECHSUPPORT = 1208,
            M_HELP_CONNECT_MIFORUM = 1209,
            M_HELP_MAPINFO_WWW = 1210,
            M_HELP_MAPINFO_WWW_STORE = 1211,
            M_HELP_CHECK_FOR_UPDATE = 1212,
            M_HELP_MAPINFO_WWW_TUTORIAL = 1213,
            M_HELP_MAPINFO_3DGRAPH_HELP = 1214,
            M_HELP_MAPINFO_CONNECT_SERVICES = 1215
        };

        public enum MenuDef95Browse
        {
            M_BROWSE_PICK_FIELDS = 704,
            M_BROWSE_OPTIONS = 703
        };

        public enum MenuDef95Map
        {
            M_MAP_LAYER_CONTROL = 801,
            M_MAP_CREATE_3DMAP = 817,
            M_MAP_CREATE_PRISMMAP = 818,
            M_MAP_THEMATIC = 307,
            M_MAP_MODIFY_THEMATIC = 308,
            M_MAP_CREATE_LEGEND = 816,
            M_MAP_CHANGE_VIEW = 805,
            M_MAP_CLONE_MAPPER = 811,
            M_MAP_PREVIOUS = 806,
            M_MAP_ENTIRE_LAYER = 807,
            M_MAP_CLEAR_CUSTOM_LABELS = 814,
            M_MAP_SAVE_COSMETIC = 809,
            M_MAP_CLEAR_COSMETIC = 810,
            M_MAP_SET_CLIP_REGION = 812,
            M_MAP_CLIP_REGION_ONOFF = 813,
            M_MAP_SETUPDIGITIZER = 803,
            M_MAP_OPTIONS = 802,
            M_MAP_CREATE_SCALEBAR = 824
        };

        public enum MenuDef95Layout
        {
            M_LAYOUT_CHANGE_VIEW = 902,
            M_LAYOUT_ACTUAL = 903,
            M_LAYOUT_ENTIRE = 904,
            M_LAYOUT_PREVIOUS = 905,
            M_LAYOUT_BRING2FRONT = 906,
            M_LAYOUT_SEND2BACK = 907,
            M_LAYOUT_ALIGN = 908,
            M_LAYOUT_DROPSHADOWS = 909,
            M_LAYOUT_DISPLAYOPTIONS = 901,
            M_LAYOUT_AUTOSCROLL_ONOFF = 910
        };

        public enum MenuDef95Graph
        {
            M_GRAPH_TYPE = 1001,
            M_GRAPH_LABEL_AXIS = 1002,
            M_GRAPH_VALUE_AXIS = 1003,
            M_GRAPH_SERIES = 1004
        };

        public enum MenuDef95MGraphNew
        {
            M_GRAPH_FORMATING = 2007,
            M_GRAPH_GENERAL_OPTIONS = 2002,
            M_GRAPH_SERIES_OPTIONS = 2003,
            M_GRAPH_GRID_SCALES = 2004,
            M_GRAPH_TITLES = 2005,
            M_GRAPH_3D_VIEWING_ANGLE = 2006,
            M_GRAPH_SAVE_AS_TEMPLATE = 2008
        };

        public enum MenuDef95MapBasic
        {
            M_MAPBASIC_CLEAR = 1101,
            M_MAPBASIC_SAVECONTENTS = 1102
        };

        public enum MenuDef95Redistrict
        {
            M_REDISTRICT_ASSIGN = 705,
            M_REDISTRICT_TARGET = 706,
            M_REDISTRICT_ADD = 707,
            M_REDISTRICT_DELETE = 708,
            M_REDISTRICT_OPTIONS = 709
        };

        public enum MenuDef95Legend
        {
            M_LEGEND_PROPERTIES = 1901,
            M_LEGEND_REFRESH = 1902,
            M_LEGEND_ADD_FRAMES = 1903,
            M_LEGEND_DELETE = 1904
        };

        public enum MenuDef953DMap
        {
            M_3DMAP_VIEW_ENTIRE_GRID = 2101,
            M_3DMAP_PROPERTIES = 2102,
            M_3DMAP_REFRESH_GRID_TEXTURE = 2103,
            M_3DMAP_WIREFRAME = 2104,
            M_3DMAP_CLONE_VIEW = 2105,
            M_3DMAP_PREVIOUS_VIEW = 2106,
            M_3DMAP_VIEWPOINT_CONTROL = 2107
        };

        public enum MenuDef95ButtonpadMain
        {
            M_TOOLS_SELECTOR = 1701,
            M_TOOLS_SEARCH_RECT = 1722,
            M_TOOLS_SEARCH_RADIUS = 1703,
            M_TOOLS_SEARCH_BOUNDARY = 1704,
            M_TOOLS_EXPAND = 1705,
            M_TOOLS_SHRINK = 1706,
            M_TOOLS_RECENTER = 1702,
            M_TOOLS_PNT_QUERY = 1707,
            M_TOOLS_HOTLINK = 1736,
            M_TOOLS_LABELER = 1708,
            M_TOOLS_DRAGWINDOW = 1734,
            M_TOOLS_RULER = 1710
        };

        public enum MenuDef95ButtonpadDrawing
        {
            M_TOOLS_POINT = 1711,
            M_TOOLS_LINE = 1712,
            M_TOOLS_POLYLINE = 1713,
            M_TOOLS_ARC = 1716,
            M_TOOLS_POLYGON = 1714,
            M_TOOLS_ELLIPSE = 1715,
            M_TOOLS_RECTANGLE = 1717,
            M_TOOLS_ROUNDEDRECT = 1718,
            M_TOOLS_TEXT = 1709,
            M_TOOLS_FRAME = 1719,
            M_TOOLS_ADD_NODE = 1723
        };

        public enum MenuDef95ButtonpadDbms
        {
            M_DBMS_OPEN_ODBC = 116
        }

        /// <summary>
        /// Menu and ButtonPad items that do not appear in the standard menus
        /// </summary>
        public enum MenuDef95NoInStandardMenus
        {
            M_TOOLS_MAPBASIC = 1720,
            M_TOOLS_SEARCH_POLYGON = 1733
        };

        /// <summary>
        /// Codes used to position Adornments relative to mapper
        /// <summary>
        public enum AdornmentPositions
        {
	        ADORNMENT_INFO_MAP_POS_TL = 0,
	        ADORNMENT_INFO_MAP_POS_TC = 1,
	        ADORNMENT_INFO_MAP_POS_TR = 2,
	        ADORNMENT_INFO_MAP_POS_CL = 3,
	        ADORNMENT_INFO_MAP_POS_CC = 4,
	        ADORNMENT_INFO_MAP_POS_CR = 5,
	        ADORNMENT_INFO_MAP_POS_BL = 6,
	        ADORNMENT_INFO_MAP_POS_BC = 7,
	        ADORNMENT_INFO_MAP_POS_BR = 8,
	        SCALEBAR_INFO_BARTYPE_CHECKEDBAR = 0,
	        SCALEBAR_INFO_BARTYPE_SOLIDBAR = 1,
	        SCALEBAR_INFO_BARTYPE_LINEBAR = 2,
	        SCALEBAR_INFO_BARTYPE_TICKBAR = 3
        };

        /// <summary>
        /// Coordinate system datum id's. These match the id's from mapinfow.prj.
        /// </summary>
        public enum CoordSysDatums
        {
            DATUMID_NAD27 = 62,
            DATUMID_NAD83 = 74,
            DATUMID_WGS84 = 104
        };

        #endregion
    }
}
