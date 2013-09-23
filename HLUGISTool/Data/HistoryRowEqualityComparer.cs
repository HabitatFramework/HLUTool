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
