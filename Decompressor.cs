using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Decompression;

namespace VbaDiff.Decompression
{
    /// <summary>
    /// Performs Decompression algorithm on a compressed stream.
    /// Chapter numbers refer to the MS-OVBA document.
    /// </summary>
    internal class Decompressor
    {        
        int compressedCurrent = 0;
        int compressedRecordEnd = 0;
        int compressedEnd = 0;
        int compressedChunkStart = 0;
        int decompressedChunkStart = 0;
        int decompressedCurrent = 0;

        List<byte> decompressedBuffer = new List<byte>();   //2.4.1.1.2
        byte[] compressedBuffer;

        internal byte[] DecompressByteStream(byte[] compressedBuffer)
        {
            this.compressedBuffer = compressedBuffer;
            DecompressStream();

            return this.decompressedBuffer.ToArray();
        }

        private string ConvertBytesToString(byte[] bytes)
        {
            StringBuilder bytesAsString = new StringBuilder();

            for (int i = 0; i < bytes.Count(); i++)
            {
                string byteAsString = Convert.ToString(bytes[i], 16);

                // Pad out the Number so we always have double-digits.
                if (byteAsString.Length == 1)
                {
                    byteAsString = "0" + byteAsString;
                }

                bytesAsString.Append(byteAsString);

                if (i < bytes.Count() - 1)
                {
                    bytesAsString.Append(" ");
                }
            }

            return bytesAsString.ToString();
        }

        private byte ReadCompressedByte()
        {
            return compressedBuffer[this.compressedCurrent];
        }

        private byte[] ReadCompressedBytes(int numberOfBytes)
        { 
            byte[] returnBytes = new byte[numberOfBytes];
            for(int i = 0; i < numberOfBytes; i++)
            {
                returnBytes[i] = compressedBuffer[this.compressedCurrent + i];
            }

            return returnBytes;
        }

        private void DecompressStream()
        { 
            byte firstByte = ReadCompressedByte();

            if (firstByte != 0x01)
            {
                throw new Exception("Stream didn't start right.");
            }

            // Get the length of the stream.
            compressedRecordEnd = (int) compressedBuffer.Length;

            compressedCurrent += 1;

            while (compressedCurrent < compressedRecordEnd)
            {
                compressedChunkStart = compressedCurrent;
                DecompressingACompressedChunk();
            }
        }

        /// <summary>
        /// 2.4.1.3.2
        /// </summary>
        private void DecompressingACompressedChunk()
        {
            // Reader the Header.
            int headerSize = 2;
            byte[] header = ReadCompressedBytes(headerSize);          

            // The header will tell you the ChunkSize.
            int size = ExtractCompressedChunkSize(header);
            //Debug.Print(String.Format("Chunk size is {0}", size));

            uint compressedFlag = ExtractCompressedChunkFlag(header);
            //Debug.Print(String.Format("CompressedFlag is {0}", compressedFlag));

            decompressedChunkStart = decompressedCurrent;

            compressedEnd = Math.Min(compressedRecordEnd, (compressedChunkStart + size));
            compressedCurrent = compressedChunkStart + 2;

            if (compressedFlag == 1)
            {
                while (compressedCurrent < compressedEnd)
                {
                    DecompressingATokenSequence(compressedEnd);
                }
            }
            else
            {
                DecompressingARawChunk();
            }
        }

        /// <summary>
        /// 2.4.1.3.3
        /// </summary>
        private void DecompressingARawChunk()
        {
            for (int i = 0; i < 4096; i++)
            {
                byte data = ReadCompressedByte();                
                decompressedBuffer.Add(data);
                decompressedCurrent++;
                compressedCurrent++;
            }
        }

        /// <summary>
        /// 2.4.1.3.4
        /// </summary>
        /// <param name="compressedEnd"></param>
        internal void DecompressingATokenSequence(long compressedEnd)
        {
            //Debug.Print(String.Format("DecompressingATokenSequence start with compressed current at {0}", compressedCurrent)); 
            int flagByte = ReadCompressedByte();
            compressedCurrent += 1;
            //Debug.Print(String.Format("DecompressingATokenSequence sets compressed current to {0}", compressedCurrent)); 

            if (compressedCurrent < compressedEnd)
            {
                for (int i = 0; i < 8; i++)
                {
                    if (compressedCurrent < compressedEnd)
                    {
                        DecompressingAToken(i, (byte)flagByte);
                    }
                }
            }
        }

        /// <summary>
        /// 2.4.1.3.5
        /// </summary>
        /// <param name="i"></param>
        /// <param name="flagByte"></param>
        private void DecompressingAToken(int index, byte flagByte)
        {
            int flag = ExtractFlagBit(index, flagByte);

            if (flag == 0)
            {
                byte currentByte = ReadCompressedByte();
                decompressedBuffer.Add(currentByte);

                compressedCurrent += 1;
                decompressedCurrent += 1;
            }            
            else // Packed
            {
                // Set Token to the CopyToken (2.4.1.1.8) at CompressedCurrent.
                int tokenSize = 2; // TODO: Move token sizes, etc to constancts.
                byte[] token = ReadCompressedBytes(tokenSize);
                
                UnpackCopyTokenResult unpackCopyTokenResult = UnpackCopyToken(token);

                int copySource = decompressedCurrent - unpackCopyTokenResult.offset;
                ByteCopy(copySource, decompressedCurrent, unpackCopyTokenResult.length);
                decompressedCurrent += unpackCopyTokenResult.length;
                compressedCurrent += 2;
            }
        }

        /// <summary>
        /// 2.4.1.3.11
        /// </summary>
        /// <param name="copySource"></param>
        /// <param name="destinationSouce"></param>
        /// <param name="byteCount"></param>
        private void ByteCopy(int copySource, int destinationSouce, int byteCount)
        {
            int srcCurrent = copySource;
            int dstCurrent = destinationSouce;

            for (int i = 0; i < byteCount; i++)
            {
                while (dstCurrent >= decompressedBuffer.Count)
                {
                    decompressedBuffer.Add(new byte());
                }

                if (dstCurrent >= decompressedBuffer.Count)
                {
                    //decompressedBuffer.Add(new byte());
                    //Debug.Fail("Trying to copy a byte to a point in the list that doesn't exist yet.");
                }
                byte copyByte = decompressedBuffer[srcCurrent];
                decompressedBuffer[dstCurrent] = copyByte;

                srcCurrent += 1;
                dstCurrent += 1;
            }
        }

        /// <summary>
        /// 2.4.1.3.19.1
        /// </summary>
        /// <param name="lengthMark"></param>
        /// <param name="offsetMarsk"></param>
        /// <param name="bitCount"></param>
        internal static CopyTokenHelpResult CopyTokenHelp(int decompressedCurrent, int decompressedChunkStart)
        {
            long difference = decompressedCurrent - decompressedChunkStart;

            // Set bitCount to the smallest integer thatis greaer than or equalto logarithm base2 of difference.
            
            UInt16 bitCount = (UInt16)Math.Ceiling(Math.Log(difference, 2));

            UInt16 bitCountMinimum = 4;
            bitCount = Math.Max(bitCount, bitCountMinimum);

            UInt16 lengthMask = (UInt16)(0xFFFF >> bitCount); // TODO: Check uint16 to int conversion.
            UInt16 offsetMask = (UInt16) ~((int)lengthMask);  // tilde is Bitwise NOT
            UInt16 maximumLength = (UInt16)(0xFFFF >> ((int)bitCount));
            maximumLength +=3;

            CopyTokenHelpResult copyTokenHelpResult = new CopyTokenHelpResult();
            copyTokenHelpResult.bitCount = bitCount;
            copyTokenHelpResult.lengthMask = lengthMask;
            copyTokenHelpResult.maximumLength = maximumLength;
            copyTokenHelpResult.offsetMask = offsetMask;

            return copyTokenHelpResult;
        }

        

        /// <summary>
        /// 2.4.1.3.19.2
        /// </summary>
        /// <param name="token"></param>
        /// <param name="offset"></param>
        /// <param name="length"></param>
        private UnpackCopyTokenResult UnpackCopyToken(byte[] token)
        {
            UInt16 uToken = BitConverter.ToUInt16(token, 0);

            CopyTokenHelpResult copyTokenHelpResult = Decompressor.CopyTokenHelp(this.decompressedCurrent, this.decompressedChunkStart);

            UInt16 length = (UInt16)(((int)uToken) & ((int)copyTokenHelpResult.lengthMask));
            length += 3;


            UInt16 temp1 = (UInt16)(((int)uToken) & ((int)copyTokenHelpResult.offsetMask));

            // 4. SET temp2 TO 16 MINUS BitCount.
            UInt16 sixteen = 16;
            UInt16 temp2 = ((UInt16)(sixteen - copyTokenHelpResult.bitCount));

            // 5. SET Offset TO (temp1 RIGHT SHIFT BY temp2) PLUS 1.
            UInt16 offset = (UInt16)((temp1 >> temp2) + 1);

            UnpackCopyTokenResult result = new UnpackCopyTokenResult();
            result.length = length;
            result.offset = offset;

            return result;
        }

        /// <summa
        /// 2.4.1.3.17
        /// </summary>
        /// <param name="index"></param>
        /// <param name="flagByte"></param>
        private int ExtractFlagBit(int index, byte flagByte)
        {
            int flag = ((flagByte >> index) & 1);

            Debug.Assert((flag == 0 ) || (flag == 1));

            return flag;
        }

        /// <summary>
        /// 2.4.1.3.12
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        private UInt16 ExtractCompressedChunkSize(byte[] header)
        {
            int temp = (BitConverter.ToUInt16(header, 0) & 0x0FFF);

            // Set Size to temp + 3
            int size = temp + 3;

            return (UInt16)size;
        }

        /// <summary>
        /// 2.4.1.3.15
        /// </summary>
        /// <param name="header"></param>
        private uint ExtractCompressedChunkFlag(byte[] header)
        {
            uint temp1 = (uint)(BitConverter.ToInt16(header, 0) & 0x8000);
            uint compressedFlag = temp1 >> 15;

            return compressedFlag;
        }
    }

    
}
