using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VbaDiff.Decompression.Exceptions;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ModuleObjects
{
    /// <summary>
    /// 2.3.4.2.3.2.1
    /// </summary>
    internal class ModuleName
    {
        #region Fields
        private byte[] moduleName;

        // Constants
        private const UInt16 IdTag = 0x0019;        
        #endregion

        #region Properties
        internal byte[] ModuleNameData
        {
            get
            {
                return this.moduleName;
            }

            set
            {
                this.moduleName = value;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Parameterless constructor.
        /// </summary>
        internal ModuleName()
        { }

        /// <summary>
        /// Constructor to create a ModuleName object around a byte stream containing the actual name of the module.
        /// </summary>
        /// <param name="moduleName"></param>
        internal ModuleName(byte[] moduleName)
        {
            this.moduleName = moduleName;
        }
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        {            
            // ID
            uint id = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);
            position += 2;

            if (id != IdTag) { throw new ParseException("Failed to Parse ID in ModuleName");}

            // SizeOfModuleName
            uint sizeOfModuleName = BitConverter.ToUInt32(stream.SubArray(position, 4), 0);
            position += 4;

            this.moduleName = stream.SubArray(position, (int)sizeOfModuleName);
            position += (int)sizeOfModuleName;
        }

        /// <summary>
        /// Converts the ModuleName to a byte array.
        /// </summary>
        internal byte[] Write()
        {
            List<byte> bytes = new List<byte>();

            // Add the ID tag.
            bytes.AddRange(BitConverter.GetBytes(IdTag));

            // Add the SizeOfModuleName as 4-byte unsigned integer.
            UInt32 sizeOfModuleName = (UInt32)this.moduleName.Length;
            bytes.AddRange(BitConverter.GetBytes(sizeOfModuleName));

            // Add the actual ModuleName.
            bytes.AddRange(this.moduleName);

            return bytes.ToArray();
        }
        #endregion
    }
}
