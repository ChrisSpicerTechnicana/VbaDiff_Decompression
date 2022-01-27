using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.1
    /// </summary>
    class ProjectSysKind
    {
        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        {
            // Check the ID.
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0001) { throw new ParseException("Failed to parse ID in ProjectSysKind"); }

            // Check the size.
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000004) { throw new ParseException("Failed to parse size in ProjectSysKind"); }

            // Check the SysKind
            uint sysKind = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (!((sysKind == 0x00000000) || (sysKind == 0x00000001) || (sysKind == 0x00000002) || (sysKind == 0x00000003))) { throw new ParseException("Failed to parse sysKind in ProjectSysKind"); }

            return;
        }
        #endregion
    }
}
