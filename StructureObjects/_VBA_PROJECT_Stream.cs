using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects
{
    /// <summary>
    /// 2.3.4.1
    /// The _VBA_PROJECT stream contains the version-dependent description of a VBA project.
    /// </summary>
    internal class _VBA_PROJECT_Stream
    {
        #region Fields
        private const UInt16 Reserved1Value = 0x61CC;
        private const UInt16 BlankVersionValue = 0xFFFF;
        private const UInt16 Reserved2Value = 0x00;
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream)
        {
            int position = 0;

            uint reserved1 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved1 != Reserved1Value) { throw new ParseException("Failed to parse reserved1 token in VBA Project Stream."); }

            uint version = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            byte reserved2 = stream[position];
            position += 1;

            if (reserved2 != 0x00) { throw new ParseException("Failed to parse reserved2 token in VBA Project Stream."); }

            uint reserved3 = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            // The rest of the stream is Performance cache.
        }

        internal byte[] Write()
        {
            byte[] byteStream = new byte[7];

            // Write reserved1
            BitConverter.GetBytes(Reserved1Value).CopyTo(byteStream, 0);

            // Now the version. When VbaDiff is writing, this will always be the default value;
            BitConverter.GetBytes(BlankVersionValue).CopyTo(byteStream, 2);

            byteStream[4] = (byte)Reserved2Value;

            return byteStream;
        }
        #endregion
    }
}
