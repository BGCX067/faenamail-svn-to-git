using System;
using System.IO;
using System.Net;
using System.Net.Sockets;

namespace Faena.Net
{
	/// <summary>
	/// Summary description for Core.
	/// </summary>
	public class Core
	{
	/*	/// <summary>
		/// 
		/// </summary>
		public Core()
		{			
		}*/

		#region function DoPeriodHandling

		/// <summary>
		/// Does period handling.
		/// </summary>
		/// <param name="data"></param>
		/// <param name="add_Remove">If true add periods, else removes periods.</param>
		/// <returns></returns>
		public static MemoryStream DoPeriodHandling(byte[] data,bool add_Remove)
		{
			using(MemoryStream strm = new MemoryStream(data)){
				return DoPeriodHandling(strm,add_Remove);
			}
		}

		/// <summary>
		/// Does period handling.
		/// </summary>
		/// <param name="strm">Input stream.</param>
		/// <param name="add_Remove">If true add periods, else removes periods.</param>
		/// <returns></returns>
		public static MemoryStream DoPeriodHandling(Stream strm,bool add_Remove)
		{			
			MemoryStream replyData = new MemoryStream();

			byte[] crlf = new byte[]{(byte)'\r',(byte)'\n'};

			strm.Position = 0;
			StreamLineReader r = new StreamLineReader(strm);
			byte[] line = r.ReadLine();

			// Loop through all lines
			while(line != null){
				if(line.Length > 0){
					if(line[0] == '.'){
						/* Add period Rfc 281 4.5.2
						   -  Before sending a line of mail text, the SMTP client checks the
						   first character of the line.  If it is a period, one additional
						   period is inserted at the beginning of the line.
						*/
						if(add_Remove){
							replyData.WriteByte((byte)'.');
							replyData.Write(line,0,line.Length);
						}
						/* Remove period Rfc 281 4.5.2
						 If the first character is a period , the first characteris deleted.							
						*/
						else
						{
							replyData.Write(line,1,line.Length-1);
						}
					}
					else{
						replyData.Write(line,0,line.Length);
					}
				}					

				replyData.Write(crlf,0,crlf.Length);

				// Read next line
				line = r.ReadLine();
			}

			replyData.Position = 0;

			return replyData;
		}

		#endregion


		#region function ReadLine

		/// <summary>
		/// Reads line of data from TcpClient.
		/// </summary>
		/// <param name="clnt"></param>
		/// <returns></returns>
		public static string ReadLine(TcpClient clnt)
		{
			long lastDataTime = DateTime.Now.Ticks;
			MemoryStream  mStrm = new MemoryStream();
			NetworkStream cStrm = clnt.GetStream();

			while(true){
				if(cStrm.DataAvailable){
					byte b = (byte)cStrm.ReadByte();
					mStrm.WriteByte(b);

					if(b == (byte)'\n'){
						// Check if previous is \r, then <CRLF>
						if(mStrm.Length > 1){
							mStrm.Position -= 2;
							if(mStrm.ReadByte() == (byte)'\r'){
								break;
							}

							// There isn't <CRLF>, restore pos.
							mStrm.Position += 1;
						}
					}											

					// reset last data time
					lastDataTime = DateTime.Now.Ticks;
				}
				else{
					//---- Time out stuff -----------------------//
					if(DateTime.Now.Ticks > lastDataTime + ((long)(60000)) * 10000){
						throw new Exception("Read timeout");
					}					
					System.Threading.Thread.Sleep(100);									
					//------------------------------------------//
				}
			}

			return System.Text.Encoding.ASCII.GetString(mStrm.ToArray());
		}

		#endregion

		#region function ReadData

		/// <summary>
		/// Reads byte data from stream while gets terminator or timeout.
		/// </summary>
		/// <param name="clnt"></param>
		/// <param name="terminator"></param>
		/// <returns></returns>
		public static byte[] ReadData(TcpClient clnt,string terminator)
		{
			bool ok             = false;
			long lastDataTime   = DateTime.Now.Ticks;
			MemoryStream  strm  = new MemoryStream();
			NetworkStream cStrm = clnt.GetStream();

			while(true)
			{
				if(cStrm.DataAvailable){
					byte[] repl = new byte[4024];
					int nCount = cStrm.Read(repl,0,repl.Length);
					strm.Seek(0,SeekOrigin.End);
					strm.Write(repl,0,nCount);

					//---- Check for terminator ------------------//
					if(!cStrm.DataAvailable && strm.Length >= terminator.Length){
						byte[] bufComp = new byte[terminator.Length];
						strm.Position = strm.Length - terminator.Length;
						strm.Read(bufComp,0,terminator.Length);

						string strCompare = System.Text.Encoding.ASCII.GetString(bufComp);
														               				
						// <CRLF> - means end of reply
						if(strCompare.IndexOf(terminator) > -1){
							ok = true;
							break;
						}
					}
					//---------------------------------------------//

					// reset last data time
					lastDataTime = DateTime.Now.Ticks;
				}
				else{
					//---- Time out stuff -----------------------//
					if(DateTime.Now.Ticks > lastDataTime + ((long)(20000)) * 10000){
						throw new Exception("Read timeout");
					}					
					System.Threading.Thread.Sleep(100);					
				}				
				//------------------------------------------//
			}

			if(ok){
				strm.SetLength(strm.Length - terminator.Length);

				byte[] byte_reply = strm.ToArray();
				strm.Close();

				return byte_reply;
			}
			else{
				// Never must not reach here
				throw new Exception("Terminator not found");
			}
		}

		#endregion

		#region function ReadDataS

		/// <summary>
		/// Overoladed ReadData, differs only return type.
		/// </summary>
		/// <param name="clnt"></param>
		/// <param name="terminator"></param>
		/// <returns></returns>
		public static string ReadDataS(TcpClient clnt,string terminator)
		{
			return System.Text.Encoding.ASCII.GetString(ReadData(clnt,terminator));
		}

		#endregion


		#region SendDataS

		/// <summary>
		/// Sends data to TcpClient.
		/// </summary>
		/// <param name="clnt"></param>
		/// <param name="data">Data which to send.</param>
		public static void SendDataS(TcpClient clnt,string data)
		{
			NetworkStream strm = clnt.GetStream();

			byte[] byte_data = System.Text.Encoding.ASCII.GetBytes(data);
			strm.Write(byte_data,0,byte_data.Length);
		}

		#endregion


		/// <summary>
		/// 
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		public static string ParseEmailFromPath(string path)
		{
			if(path.IndexOf("<") > -1 && path.IndexOf(">") > -1){
				if(path.IndexOf("<") < path.IndexOf(">")){
					path = path.Substring(path.IndexOf("<") + 1,path.IndexOf(">") - 1);
				}
			}

			return path;
		}


		#region function QDecode

		/// <summary>
		/// quoted-printable decoder.
		/// </summary>
		/// <param name="encoding">Input string encoding.</param>
		/// <param name="data">String which to encode.</param>
		/// <returns></returns>
		public static string QDecode(System.Text.Encoding encoding,string data)
		{
			MemoryStream strm = new MemoryStream(System.Text.Encoding.ASCII.GetBytes(data));
			int b = strm.ReadByte();

			MemoryStream dStrm = new MemoryStream();

			while(b > -1){
				// Hex eg. =E4
				if(b == '='){
					byte[] buf = new byte[2];
					strm.Read(buf,0,2);

					// <CRLF> followed by =, it's splitted line
					if(!(buf[0] == '\r' && buf[1] == '\n')){
						int val = int.Parse(System.Text.Encoding.ASCII.GetString(buf),System.Globalization.NumberStyles.HexNumber);
						string encodedChar = encoding.GetString(new byte[]{(byte)val});
						byte[] d = System.Text.Encoding.Unicode.GetBytes(encodedChar);
						dStrm.Write(d,0,d.Length);
					}
				}
				else{
					string encodedChar = encoding.GetString(new byte[]{(byte)b});
					byte[] d = System.Text.Encoding.Unicode.GetBytes(encodedChar);
					dStrm.Write(d,0,d.Length);
				}

				b = strm.ReadByte();
			}

			return System.Text.Encoding.Unicode.GetString(dStrm.ToArray());
		}

		#endregion

	}
}
