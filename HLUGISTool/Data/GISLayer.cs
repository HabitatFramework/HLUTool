// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2014 Andy Foy
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

        public string DisplayName
        {
            get
            {
                if (_mapName == null)
                {
                    return String.Format("{0} [{1}]", _layerName, _mapNum);
                }
                else
                {
                    return String.Format("{0} in {1} [{2}]", _layerName, _mapName, _mapNum);
                }
            }
        }

        #endregion // Properties

        public virtual bool Equals(GISLayer other)
        {
            if (other == null) return false;

            return ((this._mapNum == other._mapNum) && (this._layerName == other._layerName) && (this._layerNum == other._layerNum) && (this._layerName == other._layerName));
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as GISLayer);
        }

    }
}