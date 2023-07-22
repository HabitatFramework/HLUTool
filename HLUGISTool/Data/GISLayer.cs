// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2014, 2016 Thames Valley Environmental Records Centre
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
using System.Linq;
using System.Text;

namespace HLU.Data
{
    //---------------------------------------------------------------------
    // CHANGED: CR31 (Switching between GIS layers)
    // Enable the user to switch between different HLU layers, where
    // there is more than one valid layer in the current document.
    //
    // Contains details of each valid HLU layer (map/window number,
    // map/window name, layer number and layer name.
    //---------------------------------------------------------------------
    public class GISLayer
    {
        #region Fields

        private int _mapNum;
        private string _mapName;
        private int _layerNum;
        private string _layerName;
        private static int _mapWindowsCount;

        #endregion

        #region Constructor

        public GISLayer()
        {
        }

        public GISLayer(int mapNum, int layerNum, string layerName)
        {
            _mapNum = mapNum;
            _mapName = null;
            _layerNum = layerNum;
            _layerName = layerName;
        }

        public GISLayer(int mapNum, string mapName, int layerNum, string layerName)
        {
            _mapNum = mapNum;
            _mapName = mapName;
            _layerNum = layerNum;
            _layerName = layerName;
        }

        #endregion // Constructor

        #region Properties

        public int MapNum
        {
            get { return _mapNum; }
            set { _mapNum = value; }
        }

        public string MapName
        {
            get { return _mapName; }
            set { _mapName = value; }
        }

        public int LayerNum
        {
            get { return _layerNum; }
            set { _layerNum = value; }
        }

        public string LayerName
        {
            get { return _layerName; }
            set { _layerName = value; }
        }

        public static int MapWindowsCount
        {
            get { return _mapWindowsCount; }
            set { _mapWindowsCount = value; }
        }

        public string DisplayName
        {
            get
            {
                string mapName = null;

                // If there is no map window name (MapInfo) then set
                // the map name to the window number, otherwise (ArcGIS)
                // set the map name to the window name and number.
                if (_mapName == null)
                    mapName = string.Format(" [{0}]", _mapNum);
                else
                    mapName = string.Format(" in {0} [{1}]", _mapName, _mapNum);

                // If there is more than one map window in total then
                // include the map name/number
                if (_mapWindowsCount > 1)
                    return String.Format("{0}{1}", _layerName, mapName);
                else
                    return String.Format("{0}", _layerName);
            }
        }

        #endregion // Properties

        #region Methods

        public override string ToString()
        {
            return "Map: " + _mapName + "[" + _mapNum + "]" + " Layer: " + _layerName + "[" + _layerNum + "]";
        }

        public override int GetHashCode()
        {
            return (this.MapNum.GetHashCode() * 100) + (this.LayerNum.GetHashCode() * 10) + this.LayerName.GetHashCode();
        }

        public virtual bool Equals(GISLayer other)
        {
            if (other == null) return false;

            return ((this._mapNum == other._mapNum) && (this._mapName == other._mapName) && (this._layerNum == other._layerNum) && (this._layerName == other._layerName));
        }

        public override bool Equals(object obj)
        {
            if (this.GetType() != obj.GetType()) return false;

            return Equals(obj as GISLayer);
        }

        #endregion // Methods

    }
}