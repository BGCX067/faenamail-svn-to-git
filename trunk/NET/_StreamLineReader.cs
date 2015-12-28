using System;
using System.IO;
using System.Text;

namespace Faena.Net
{
	/// <summary>
	/// Byte[] line parser.
	/// </summary>
	internal class StreamLineReader
	{
		private Stream m_StrmSource = null;
		private long   m_StrmLen    = 0;

		/// <summary>
		/// Default constructor.
		/// </summary>
		/// <param name="strmSource"></param>
		public StreamLineReader(Stream strmSource)
		{
			m_StrmSource = strmSource;
			m_StrmLen    = strmSource.Length;
		}


		#region function ReadLine

		/// <summary>
		/// Reads byte[] line from stream.
		/// </summary>
		/// <returns>Return null if end of stream reached.</returns>
		public byte[] ReadLine()
		{
			byte[] retVal     = null;
			long strmStartPos = m_StrmSource.Position;
			long len          = 0;
			byte prevByte     = 0;
			byte currByte     = 0;

			while(true){
				// Line found
				if((prevByte == (byte)'\r' && currByte == (byte)'\n')){
					m_StrmSource.Position = strmStartPos;
					retVal = new byte[len-2];   // Remove <CRLF> 
					m_StrmSource.Read(retVal,0,retVal.Length);
					m_StrmSource.Position += 2; // Move by <CRLF>
					break;
				}

				// End of stream reached
				if(strmStartPos + len >= m_StrmLen-1){
					if(len > 0){
						m_StrmSource.Position = strmStartPos;
						retVal = new byte[len];
						m_StrmSource.Read(retVal,0,(int)len);						
					}
                    break;
				}				

				// Store byte
				prevByte = currByte;
				len++;

				// Read new byte
				currByte = (byte)m_StrmSource.ReadByte();
			}

			return retVal;
		}

		#endregion

	}
}
