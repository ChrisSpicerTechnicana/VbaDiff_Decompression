using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using VbaDiff.Model.References;
using VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects.ReferenceObjects;

namespace VbaDiff.Decompression.StructureObjects.DirStreamObjects.ProjectInformationObjects
{
    /// <summary>
    /// 2.3.4.2.2
    /// </summary>
    class ProjectReferences
    {
        #region Fields
        private List<RawReference> rawReferences = new List<RawReference>();

        #endregion

        #region Properties
        internal List<RawReference> RawReferences
        {
            get
            {
                return this.rawReferences;
            }
        }
        #endregion

        #region Public Methods
        internal void ParseStream(byte[] stream, ref int position)
        {
            bool nextRecordIsAReference = true;

            int referenceLoopCount = 0;
            while (nextRecordIsAReference)
            {
                referenceLoopCount += 1;

                // Parse the references.                
                RawReference rawReference = ReferenceReader.ParseStream(stream, ref position);
                rawReferences.Add(rawReference);

                // If the next ID is a ProjectModules ID, you know to break.
                nextRecordIsAReference = !PeekForBeginningOfProjectModules(stream, position);

                if (referenceLoopCount > 100) { Debug.Print("Reference looping too much."); }
            }
        }
        #endregion

        #region Private Methods
        private bool PeekForBeginningOfProjectModules(byte[] stream, int position)
        {
            uint projectModulesId = BitConverter.ToUInt16(stream.SubArray(position, 2), 0);

            return (projectModulesId == 0x000F);
        }
        #endregion
    }
}
