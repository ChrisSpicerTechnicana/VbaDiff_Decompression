using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.5
    /// </summary>
    internal class ModuleOffset
    {
        #region Fields
        private UInt32 textOffset;  // Where to find the source code in the module stream.
        #endregion

        #region Properties
        internal UInt32 TextOffset
        {
            get
            {
                return this.textOffset;
            }
        }
        #endregion

        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x0031) { throw new ParseException("Failed to parse ID in Parsestream."); }

            // Size
            uint size = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if(size != 0x00000004){ throw new ParseException("Failed to parse size in ModuleOffset.");}

            // Text Offset
            this.textOffset = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;
        }
    }
}
