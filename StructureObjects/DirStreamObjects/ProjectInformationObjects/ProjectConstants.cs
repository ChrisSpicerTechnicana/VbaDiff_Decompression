using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.1.11
    /// </summary>
    class ProjectConstants
    {
        #region Fields
        private byte[] constants;
        private byte[] constantsAsUnicode;

        // Constants
        private const UInt16 IdTag = 0x000C;
        private const int MaximumSizeOfConstants = 1015;
        private const UInt16 Reserved = 0x003C;
        #endregion

        #region Properties
        internal byte[] Constants
        {
            get
            {
                return this.constants;
            }
        }

        internal byte[] ConstantsAsUnicode
        {
            get
            {
                return this.constantsAsUnicode;
            }
        }

        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        { 
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != IdTag) { throw new ParseException("Failed to parse ID for ProjectConstants."); }

            // SizeOfConstants
            uint sizeOfConstants = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;
            
            if (sizeOfConstants > MaximumSizeOfConstants) { throw new ParseException("Failed to parse sizeOfConstants for ProjectConstants."); }

            // Save the constants section. We can extract the project constants from this byte array later.
            this.constants = stream.SubArray(position, (int)sizeOfConstants);
            position += (int) sizeOfConstants;

            // Reserved
            uint reserved = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (reserved != Reserved) { throw new ParseException("Failed to parse reserved field for ProjectConstants."); }

            // Size of constants in Unicode
            uint sizeOfConstantsInUnicode = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            if ((sizeOfConstantsInUnicode % 2) != 0) { throw new ParseException("Failed to pase size of constants in Unicode."); }

            // Not interested in the constants in Unicode data just yet.
            this.constantsAsUnicode = stream.SubArray(position, (int)sizeOfConstantsInUnicode);
            position += (int)sizeOfConstantsInUnicode;
        }

        /// <summary>
        /// Builds the Project Constants Record.
        /// </summary>
        /// <param name="constants">The constants in the form of a byte array.</param>
        /// <param name="constantsAsUnicode">The constants as unicode in the form of a byte array.</param>
        /// <returns></returns>
        internal byte[] BuildProjectConstantsRecord(byte[] constants, byte[] constantsAsUnicode)
        {
            List<byte> bytes = new List<byte>();

            // Add the Id Tag
            bytes.AddRange(BitConverter.GetBytes(IdTag));

            // Add the Size of the Project Constants
            UInt32 sizeOfProjectConstants = (UInt32)constants.Length;            
            bytes.AddRange(BitConverter.GetBytes(sizeOfProjectConstants));

            // Add the Constants themselves
            bytes.AddRange(constants);

            // Add the Reserved field
            bytes.AddRange(BitConverter.GetBytes(Reserved));

            // Add the Size of the Project Constants in Unicode
            UInt32 sizeOfProjectConstantsInUnicode = (UInt32)constantsAsUnicode.Length;
            bytes.AddRange(BitConverter.GetBytes(sizeOfProjectConstantsInUnicode));

            // Add the Constants as Unicode themselves
            bytes.AddRange(constantsAsUnicode);

            return bytes.ToArray();
        }
        #endregion
    }

}
