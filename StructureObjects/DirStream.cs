using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects;

namespace VbaDiff.Decompression.StructureObjects
{
    /// <summary>
    /// 2.3.4.2
    /// </summary>
    internal class DirStream
    {
        #region Fields
        private ProjectInformation informationRecord = new ProjectInformation();                     
        #endregion

        #region Properties
        internal ProjectInformation InformationRecord
        {
            get
            {
                return this.informationRecord;
            }
        }
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream)
        {
            int position = 0;
            
            informationRecord.ParseStream(stream, ref position);                                                
        }
        #endregion

    }
}
