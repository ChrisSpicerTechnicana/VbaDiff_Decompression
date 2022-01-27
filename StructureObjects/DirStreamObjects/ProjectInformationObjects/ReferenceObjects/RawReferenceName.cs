using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// Contains the data for REFERENCE NAME 2.3.4.2.2.2
    /// </summary>
    internal class RawReferenceName
    {
        /// <summary>
        /// An array of SizeOfName bytes that specifies the name of the referenced VBA project 
        /// or Automation type library. MUST contain MBCS characters encoded using the code page
        /// specified in PROJECTCODEPAGE Record (section 2.3.4.2.1.4).
        /// </summary>
        internal byte[] NameBytes
        { get; set; }
    }
}
