using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects
{
    /// <summary>
    /// Contains the data read from a 2.3.4.2.2.1 REFERENCE Record
    /// </summary>
    internal class RawReference
    {
        internal RawReferenceName ReferenceName
        { get; set; }

        internal RawReferenceProject ReferenceProject
        { get; set; }
    }
}
