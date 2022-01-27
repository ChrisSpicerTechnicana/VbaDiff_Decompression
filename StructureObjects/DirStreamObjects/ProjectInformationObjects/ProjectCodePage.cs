using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    class ProjectCodePage
    {
        #region Fields
        private UInt16 codePage;
        #endregion

        #region Properties
        internal UInt16 CodePage
        {
            get
            {
                return this.codePage;
            }
        }
        #endregion

        internal void ParseStream(byte[] stream, ref int position)
        {
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if(id !=0x0003) { throw new ParseException("Failed to parse Id in ProjectCodePage.");};

            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if (size != 0x00000002) { throw new ParseException("failed to parse size in ProjectCodePage."); };

            // Parse and store the code page - we're going to need it later.
            this.codePage = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position +=2;



        }
    }
}
