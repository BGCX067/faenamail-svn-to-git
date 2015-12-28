using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using LumiSoft.Net;

namespace LumiSoft.Net
{
	// ToDo: Partial logging, allow to specify log part size.
	//       If log part gets full, write it file.

	/// <summary>
	/// Summary description for SocketLogger.
	/// </summary>
	public class SocketLogger
	{
		private Socket  m_pSocket     = null;
		private string          m_SessionID   = "";
		private LogEventHandler m_pLogHandler = null;
		private string s = ""; //****
		
		/// <summary>
		/// Default constructor.
		/// </summary>
		/// <param name="socket"></param>
		/// <param name="logHandler"></param>
		public SocketLogger(Socket socket,LogEventHandler logHandler)
		{	
			m_pSocket     = socket;
			m_pLogHandler = logHandler;		
		}

		/// <summary>
		/// Adds data read(from remoteEndpoint) entry.
		/// </summary>
		/// <param name="text"></param>
		public void AddReadEntry(string text)
		{
			s += CreateEntry(text,">>>");
		}

		/// <summary>
		/// Adds data send(to remoteEndpoint) entry.
		/// </summary>
		/// <param name="text"></param>
		public void AddSendEntry(string text)
		{
			s += CreateEntry(text,"<<<");
		}

		/// <summary>
		/// Adds free text entry.
		/// </summary>
		/// <param name="text"></param>
		public void AddTextEntry(string text)
		{
			s += CreateEntry(text,"---");
		}

		/// <summary>
		/// 
		/// </summary>
		public void Flush()
		{
			if(m_pLogHandler != null){
				s += "//----- Sys: 'Session:'" + this.SessionID + " removed " + DateTime.Now + "\r\n";

				m_pLogHandler(this,new Log_EventArgs(s));				
		//		m_Flushed = true;
				s = "";
			}
		//	System.Windows.Forms.MessageBox.Show(s);
		}

		
		private string CreateEntry(string text,string prefix)
		{
			string retVal = "";

			if(text.EndsWith("\r\n")){
				text = text.Substring(0,text.Length - 2);
			}

			string remIP = "xxx.xxx.xxx.xxx";
			if(m_pSocket.RemoteEndPoint != null){
				remIP = ((IPEndPoint)m_pSocket.RemoteEndPoint).Address.ToString();
			}

			string[] lines = text.Replace("\r\n","\n").Split('\n');
			foreach(string line in lines){
				retVal += "SessionID: " + m_SessionID + "  RemIP: " + remIP + "  " + prefix + "  '" + line + "'\r\n";
			}

			return retVal;
		}


		/// <summary>
		/// Gets or sets session ID.
		/// </summary>
		public string SessionID
		{
			get{ return m_SessionID; }

			set{ 
				m_SessionID = value; 
				s = "//----- Sys: 'Session:'" + this.SessionID + " added " + DateTime.Now + "\r\n";
			}
		}

	}

/*	/// <summary>
	/// Log entry type.
	/// </summary>
	public enum EntryType
	{
		/// <summary>
		/// 
		/// </summary>
		ReadFromRemoteEP,

		/// <summary>
		/// 
		/// </summary>
		SendToRemoteEP,

		/// <summary>
		/// 
		/// </summary>
		FreeText,
	}*/
}
