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

using System.Collections.Generic;
using HLU.Data.Model;

namespace HLU.Data
{
    class HistoryRowEqualityComparer : IEqualityComparer<HluDataSet.historyRow>
    {
        #region IEqualityComparer<historyRow> Members

        public bool Equals(HluDataSet.historyRow x, HluDataSet.historyRow y)
        {
            if (x.incid == null)
            {
                if (y.incid != null) return false;
            }
            else
            {
                if (!x.incid.Equals(y.incid)) return false;
            }

            if (x.toid == null)
            {
                if (y.toid != null) return false;
            }
            else
            {
                if (!x.toid.Equals(y.toid)) return false;
            }

            if (x.toid_fragment_id == null)
            {
                if (y.toid_fragment_id != null) return false;
            }
            else
            {
                if (!x.toid_fragment_id.Equals(y.toid_fragment_id)) return false;
            }

            return true;
        }

        public int GetHashCode(HluDataSet.historyRow obj)
        {
            return obj.incid.GetHashCode() ^ obj.toid.GetHashCode() ^ obj.toid_fragment_id.GetHashCode();
        }

        #endregion
    }
}
