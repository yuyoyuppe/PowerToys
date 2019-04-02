// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Security.Cryptography;

namespace VerificationRoutines
{
    /*
     * ObjectIdentity is used by MSAA to uniquely identify
     * a given accessible element.
     */
    public class ObjectIdentity
    {
        private List<int> _path;

        public ObjectIdentity()
        {
            _path = new List<int>();
            _path.Add(0); // root element
        }

        public static ObjectIdentity FromIntArray(object idObject)
        {
            int[] id = (int[])idObject;
            ObjectIdentity obj = new ObjectIdentity();
            for (int i = 1; i < id.Length; i++)
            {
                obj._path.Add(id[i]);
            }
            return obj;
        }

        public ObjectIdentity(ObjectIdentity parent, int idChild)
        {
            _path = new List<int>(parent._path);
            _path.Add(idChild);
        }

        public ObjectIdentity Parent
        {
            get
            {
                ObjectIdentity p = new ObjectIdentity();
                for (int i = 1; i < _path.Count - 1; i++)
                {
                    p._path.Add(_path[i]);
                }
                return p;
            }
        }

        public int[] GetFriendlyPath()
        {
            return _path.ToArray();
        }

        #region IComparable Members
        public int CompareTo(object obj)
        {
            if (obj is ObjectIdentity)
            {
                ObjectIdentity id = (ObjectIdentity)obj;
                if (id._path.Count != _path.Count)
                    return 1;
                for (int i = 0; i < this._path.Count; i++)
                {
                    if (_path[i] != id._path[i])
                        return 1;
                }
                return 0;
            }
            else
                return 1;
        }

        #endregion

        #region IEqualityComparer Members

        public override bool Equals(object x)
        {
            ObjectIdentity objX = (ObjectIdentity)x;
            //ObjectIdentity objY = (ObjectIdentity)y;
            if (objX._path.Count == _path.Count)
            {
                for (int i = 0; i < objX._path.Count; i++)
                {
                    if (objX._path[i] != _path[i])
                        return false;
                }
                return true;
            }
            else
                return false;
        }

        #endregion

        #region IHashCodeProvider Members

        public override int GetHashCode()
        {
            //if (!(obj is ObjectIdentity))
            //    return 0;
            //ObjectIdentity id = (ObjectIdentity)obj;
            StringBuilder s = new StringBuilder();
            s.Append(String.Format("{0}", _path[0]));
            for (int i = 1; i < _path.Count; i++)
            {
                s.Append(String.Format("-{0}", _path[i]));
            }
            MD5 md5 = MD5.Create();

            byte[] bytes = md5.ComputeHash(System.Text.Encoding.ASCII.GetBytes(s.ToString()));
            int res = 0;
            for (int i = 0; i < 4; i++)
            {
                res += bytes[i] * (int)Math.Pow(2, i * 4);
            }
            return res;
        }

        #endregion
    }
}
