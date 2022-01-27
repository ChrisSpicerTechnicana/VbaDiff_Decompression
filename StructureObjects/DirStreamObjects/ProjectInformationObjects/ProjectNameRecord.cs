using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.5
    /// </summary>
    class ProjectNameRecord
    {
        #region Fields
        private byte[] projectName;
        #endregion

        #region Properties
        internal byte[] ProjectName
        {
            get
            {
                return this.projectName;
            }
        }
        #endregion

        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0004) { throw new ParseException("Failed to parse Id in ProjectName."); }

            uint sizeOfProjectName = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if ((sizeOfProjectName< 1)||(sizeOfProjectName > 128)) { throw new ParseException("Size of project name is not an allowed size."); }

            this.projectName = stream.SubArray(position, (int) sizeOfProjectName);
            position += (int)sizeOfProjectName;
        }
    }
}
