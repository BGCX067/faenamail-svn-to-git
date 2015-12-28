using System;
using System.IO;
using System.Data;
using System.Collections;

/*
Old code, must be rewritten
*/

namespace Faena.Net.Mime
{
	/// <summary>
	/// Summary description for Mime.
	/// </summary>
	[Obsolete("Use MimeConstructor instead !")]
	public class MimeMessage
	{	
		private string      m_From       = "";
		private string[]    m_To         = null;
		private string      m_Subject    = "";
		private string      m_Date       = "";
		private string      m_Body       = "";
		private string      m_HtmlBody   = ""; 
		private Attachments m_Attachmets = null;
		private string      m_Name       = "";
		private string      m_BoundaryID = "";

		/// <summary>
		/// 
		/// </summary>
		public MimeMessage()
		{
			m_BoundaryID = Guid.NewGuid().ToString();

			m_Attachmets = new Attachments();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="from"></param>
		/// <param name="to"></param>
		/// <param name="subject"></param>
		/// <param name="body"></param>
		public MimeMessage(string from,string[] to,string subject,string body)
		{
			m_BoundaryID = Guid.NewGuid().ToString();

			m_From    = from;
			m_To      = to;
			m_Subject = subject;
			m_Body    = body;

			m_Attachmets = new Attachments();
		}


		#region function CreateMime()

		private string CreateMime()
		{
			if(m_To == null){
				return "";
			}

			string mime = "";

			mime += "From: \"" + m_Name + "\" <" + m_From + ">\n";
			mime += "To: " + Construct_To();
			mime += "Subject: " + m_Subject + "\n";
			mime += "Date: "    + m_Date    + "\n";
			mime += "MIME-Version: 1.0\n";
			mime += "Content-Type: multipart/mixed;\n";
			mime += "\tboundary=\"----=_NextPart_" + m_BoundaryID + "\"\n";

			mime += "\nThis is a multi-part message in MIME format.\n\n";

			mime += Construct_Body();
			mime += Construct_HtmlBody();

			mime += Construct_Files();

			mime += "------=_NextPart_" + m_BoundaryID + "--\r\n";

			return mime;
		}

		#endregion

		
		#region function Construct_To

		private string Construct_To()
		{
			string retVal = "";

			if(m_To != null)
			{
				if(m_To.Length == 1){
					retVal += "<" + m_To[0] + ">\n";
				}
				else
				{
					int length = m_To.Length;
					for(int i=0;i<length;i++)
					{
						retVal += "<" + m_To[0] + ">";

						if(i < length-1){
							retVal += ",\n\t";
						}
						else{
							retVal += "\n";
						}
					}
				}
			}			

			return retVal;
		}

		#endregion


		#region function Construct_Body

		private string Construct_Body()
		{
			string retVal = "";

			retVal += "------=_NextPart_" + m_BoundaryID + "\n";
			retVal += "\tcharset=\"iso-8859-1\"\n";
			retVal += "Content-Transfer-Encoding: 8bit\n\n";

			retVal += m_Body + "\n";

			retVal += "\n";

			return retVal;
		}

		#endregion

		#region function Construct_HtmlBody

		private string Construct_HtmlBody()
		{
			string retVal = "";

			if(m_HtmlBody.Length > 0 && File.Exists(m_HtmlBody))
			{
				FileStream fs = new FileStream(m_HtmlBody,FileMode.Open,FileAccess.Read);
				byte[] bData = new byte[fs.Length];
				fs.Read(bData,0,(int)fs.Length);
				fs.Close();

				string data = System.Text.Encoding.ASCII.GetString(bData);

				retVal += "------=_NextPart_" + m_BoundaryID + "\n";
				retVal += "Content-Type: text/html;\n";
				retVal += "\tcharset=\"iso-8859-1\"\n";
				retVal += "Content-Transfer-Encoding: 8bit\n\n";			
            
				retVal += data + "\n\n";
			}

			return retVal;
		}

		#endregion


		#region function Construct_Files

		private string Construct_Files()
		{
			string retVal = "";

			// Loop all attachment files
			foreach(Attachment att in m_Attachmets)
			{
				string file = att.FileName;				
				string fileName = System.IO.Path.GetFileName(file);

				string ext = System.IO.Path.GetExtension(fileName);
				string Encoding    = "";
				string ContentType = "";

				//--------- Get MIME Type ---------------------//
				switch(ext)
				{
					case "txt":
						ContentType = "text/plain";
						Encoding    = "quoted-printable";
						break;

					default:
						ContentType = "application/octet-stream";
						Encoding    = "base64";
						break;
				}
				//--------- End of MIME Type --------------------//


				if(att.AttachmentType == "file")
				{
					if(File.Exists(file)){
						using(FileStream fs = File.OpenRead(file)){
							byte[] bData = new byte[fs.Length];
							fs.Read(bData,0,(int)fs.Length);

							string data = "";					
							if(ext == "txt"){
								data = Convert.ToBase64String(bData,0,bData.Length);
							}
							else{
								data = Convert.ToBase64String(bData,0,bData.Length);
							}

							retVal += Construct_File(ContentType,fileName,Encoding,data);
						}						
					}
				}
				else
				{
					byte[] bData = att.FileData;

					string data = "";
					if(ext == "txt"){
						data = Convert.ToBase64String(bData,0,bData.Length);
					}
					else{
						data = Convert.ToBase64String(bData,0,bData.Length);
					}

					retVal += Construct_File(ContentType,fileName,Encoding,data);
				}
			}		

			return retVal;
		}

		#endregion

		#region function Construct_File

		private string Construct_File(string ContentType,string FileName,string Encoding,string Data)
		{
			string retVal = "";
			
			retVal += "------=_NextPart_" + m_BoundaryID + "\n";
			retVal += "Content-Type: " + ContentType + ";\n";
			retVal +=		"\tname=\"" + FileName + "\"\n";
			retVal += "Content-Transfer-Encoding: " + Encoding + "\n";
			retVal += "Content-Disposition: attachment;\n";
			retVal +=		"\tfilename=\"" + FileName + "\"\n\n";

			retVal += Data + "\n\n";

			return retVal;
		}

		#endregion


		//------- Static Functions ---//

		#region Static  Functions

		/// <summary>
		/// 
		/// </summary>
		/// <param name="from"></param>
		/// <param name="to"></param>
		/// <param name="subject"></param>
		/// <param name="body"></param>
		/// <param name="attachments"></param>
		/// <returns></returns>
		public static MimeMessage CreateMime(string from,string[] to,string subject,string body,string[] attachments)
		{
			MimeMessage mime = new MimeMessage(from,to,subject,body);
	//		mime.Attachments = attachments;
			return mime;
		}

		#endregion

		//----------------------------//
	

		#region Propeties Implementation

		/// <summary>
		/// 
		/// </summary>
		public string From
		{
			get{
				return m_From;
			}

			set{
				m_From = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string[] To
		{
			get{
				return m_To;
			}

			set{
				m_To = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string Subject
		{
			get{
				return m_Subject;
			}

			set{
				m_Subject = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string Date
		{
			get{
				return m_Date;
			}

			set{
				m_Date = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string Body
		{
			get{
				return m_Body;
			}

			set{
				m_Body = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string HtmlBody
		{
			get{
				return m_HtmlBody;
			}

			set{
				m_HtmlBody = value;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public Attachments Attachments
		{
			get{
				return m_Attachmets;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public string Name
		{
			get{
				return m_Name;
			}

			set{
				m_Name = value;
			}
		}


		/// <summary>
		/// 
		/// </summary>
		public string Mime
		{
			get{
				return CreateMime();
			}			
		}

		#endregion
		
	}
}
