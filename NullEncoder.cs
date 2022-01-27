using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VbaDiff.Decompression
{
    internal static class NullEncoder
    {
        /// <summary>
        /// 2.4.4.2   Encode Nulls 
        /// The Password Hash stores Key and PasswordHash with null bytes removed. The fields are encoded by replacing 0x00 bytes with 0x01 and setting a bit on the bit-fields GrbitKey and GrbitHashNull, respectively. 
        /// </summary>
        /// <param name="inputBytes">An input array of bytes to be encoded. </param>
        /// <param name="grbitNull"> An output array of bits specifying null bytes in InputBytes. </param>
        /// <param name="encodedBytes">An output array of encoded bytes.</param>    
        internal static void EncodeNulls(byte[] inputBytes, ref BitArray grbitNull, ref byte[] encodedBytes)
        {
            grbitNull = new BitArray(inputBytes.Length);
            encodedBytes = new Byte[inputBytes.Length];

            for(int i = 0; i < inputBytes.Length; i++)
            {
                var inputByte = inputBytes[i];
                if (inputByte == 0x00)
                {
                    // This is a null.
                    encodedBytes[i] = 0x01;
                    grbitNull[i] = false;
                }
                else
                {
                    encodedBytes[i] = inputByte;
                    grbitNull[i] = true;
                }
            }
        }

        /// <summary>
        /// 2.4.4.3   Decode Nulls        
        /// The Password Hash stores Key and PasswordHash with null bytes removed as specified by Encode Nulls (section 2.4.4.2). The fields are decoded by reading bit-fields GrbitKey and GrbitHashNull, and replacing corresponding bytes in Key and PasswordHash with 0x00. 
        /// </summary>
        /// <param name="encodedBytes"></param>
        /// <param name="grBitNull"></param>
        /// <param name="decodedBytes"></param>
        internal static void DecodeNulls(byte[] encodedBytes, BitArray grBitNull, ref byte[] decodedBytes)
        {
            decodedBytes = new byte[encodedBytes.Length];

            for (int i = 0; i < encodedBytes.Length; i++)
            {
                if (!grBitNull[i])
                {
                    decodedBytes[i] = 0x00;
                }
                else
                {
                    decodedBytes[i] = encodedBytes[i];
                }
            }
        }
    }
}
