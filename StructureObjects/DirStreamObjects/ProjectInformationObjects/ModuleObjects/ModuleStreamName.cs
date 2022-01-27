using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.3
    /// </summary>
    internal class ModuleStreamName
    {
        #region Fields
        private byte[] streamName;
        private byte[] streamNameUnicode;
        #endregion

        #region Properties
        internal byte[] StreamName
        {
            get
            {
                return this.streamName;
            }
        }
        #endregion

        internal void ParseStream(byte[] stream, ref int position)
        { 
            //ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != 0x001A) { throw new ParseException("Failed to parse ID in ModuleStream name."); }

            // SizeOfStreamName
            uint sizeOfStreamName = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            //Stream Name. This is very useful!
            this.streamName = stream.SubArray(position, (int)sizeOfStreamName);
            position += (int)sizeOfStreamName;

            // Reserved
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != 0x0032) { throw new ParseException("Failed to parse reserved field in ModuleStreamName."); }

            // Size of stream in Unicode
            uint sizeOfStreamNameUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if ((sizeOfStreamNameUnicode % 2) != 0) { throw new ParseException("Faled to parse size of stream name unicode in ModuleStream name."); }

            this.streamNameUnicode = stream.SubArray(position, (int)sizeOfStreamNameUnicode);
            position += (int)sizeOfStreamNameUnicode;
        }
    }
}
