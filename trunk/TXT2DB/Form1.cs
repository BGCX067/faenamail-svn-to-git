using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace TXT2DB
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>

	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnGo;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private ValidatorControl.Validator validator1;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Button btnShowODBC;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.ComponentModel.IContainer components;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.label1 = new System.Windows.Forms.Label();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.btnBrowse = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.btnGo = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.validator1 = new ValidatorControl.Validator(this.components);
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.btnShowODBC = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label1.Location = new System.Drawing.Point(16, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(65, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Source File:";
			// 
			// textBox1
			// 
			this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.validator1.SetDataType(this.textBox1, ValidatorControl.DataTypeConstants.StringType);
			this.validator1.SetIfValidate(this.textBox1, true);
			this.textBox1.Location = new System.Drawing.Point(124, 28);
			this.validator1.SetMaxValue(this.textBox1, "");
			this.validator1.SetMinValue(this.textBox1, "");
			this.textBox1.Name = "textBox1";
			this.validator1.SetRegularExpression(this.textBox1, "");
			this.validator1.SetRequired(this.textBox1, true);
			this.textBox1.Size = new System.Drawing.Size(344, 20);
			this.textBox1.TabIndex = 1;
			this.textBox1.Text = "";
			this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// btnBrowse
			// 
			this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnBrowse.Location = new System.Drawing.Point(472, 28);
			this.btnBrowse.Name = "btnBrowse";
			this.btnBrowse.TabIndex = 2;
			this.btnBrowse.Text = "Browse";
			this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label2.Location = new System.Drawing.Point(16, 68);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(103, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Target ODBC DSN:";
			// 
			// btnGo
			// 
			this.btnGo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGo.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnGo.Location = new System.Drawing.Point(300, 236);
			this.btnGo.Name = "btnGo";
			this.btnGo.TabIndex = 5;
			this.btnGo.Text = "Go";
			this.btnGo.Click += new System.EventHandler(this.btnGo_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnCancel.Location = new System.Drawing.Point(436, 236);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.DefaultExt = "*.txt";
			this.openFileDialog1.Filter = "Text files|*.txt|All files|*.*";
			this.openFileDialog1.Title = "OpenText File";
			// 
			// validator1
			// 
			this.validator1.DisplayControl = null;
			// 
			// comboBox1
			// 
			this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.validator1.SetDataType(this.comboBox1, ValidatorControl.DataTypeConstants.StringType);
			this.validator1.SetIfValidate(this.comboBox1, false);
			this.comboBox1.Location = new System.Drawing.Point(124, 64);
			this.validator1.SetMaxValue(this.comboBox1, "");
			this.validator1.SetMinValue(this.comboBox1, "");
			this.comboBox1.Name = "comboBox1";
			this.validator1.SetRegularExpression(this.comboBox1, "");
			this.validator1.SetRequired(this.comboBox1, false);
			this.comboBox1.Size = new System.Drawing.Size(196, 21);
			this.comboBox1.TabIndex = 7;
			this.comboBox1.Text = "comboBox1";
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
			// 
			// btnShowODBC
			// 
			this.btnShowODBC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnShowODBC.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnShowODBC.Location = new System.Drawing.Point(336, 64);
			this.btnShowODBC.Name = "btnShowODBC";
			this.btnShowODBC.Size = new System.Drawing.Size(108, 23);
			this.btnShowODBC.TabIndex = 8;
			this.btnShowODBC.Text = "Show ODBC";
			this.btnShowODBC.Click += new System.EventHandler(this.btnShowODBC_Click);
			// 
			// listView1
			// 
			this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader1,
																						this.columnHeader2});
			this.listView1.Location = new System.Drawing.Point(8, 104);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(528, 120);
			this.listView1.TabIndex = 9;
			this.listView1.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Contact";
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "Status";
			this.columnHeader2.Width = 461;
			// 
			// Form1
			// 
			this.AcceptButton = this.btnGo;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(548, 270);
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.btnShowODBC);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnGo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.textBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnBrowse);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "TXT2DB";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		/// 
		FaenaMailLibrary.ClassRegistry reg;

		[STAThread]
		static void Main() 
		{
			Application.EnableVisualStyles();
			LibExcept.UnhandledExceptionManager.AddHandler(false);
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			reg = new FaenaMailLibrary.ClassRegistry();
			textBox1.Text = reg.GetSetting(@"Faena Mail\TXT2DB", "SourceFile","");
			FaenaMailLibrary.tools tools = new FaenaMailLibrary.tools();
			string [] a = tools.LoadDSNList();
			comboBox1.Items.Clear();
			if (a!=null) comboBox1.Items.AddRange(a);
			comboBox1.Text = reg.GetSetting(@"Faena Mail\TXT2DB", "ODBCDSN","");
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnGo_Click(object sender, System.EventArgs e)
		{
			btnGo.Enabled = false;
			btnCancel.Text = "&Cancel";
			ListViewItem itm;
			ListViewItem.ListViewSubItem sitm;
			string companyName="";
			string address="";
			string phone="";
			string fax="";
			string email="";
			string website="";
			string contact="";
			string title="";
			int cnt = 0;
			if (validator1.ValidateAll())
			{
				FaenaMailLibrary.Functions db = new FaenaMailLibrary.Functions(comboBox1.Text);
				if (File.Exists(textBox1.Text))
				{
					reg.SaveSetting(@"Faena Mail\TXT2DB", "SourceFile", textBox1.Text);
					reg.SaveSetting(@"Faena Mail\TXT2DB", "ODBCDSN", comboBox1.Text);
					System.IO.FileStream fs = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read);
					// Create an instance of StreamReader to read from a file.
					// The using statement also closes the StreamReader.
					using (StreamReader sr = new StreamReader(fs)) 
					{
						String line;
						// Read and display lines from the file until the end of 
						// the file is reached.
						line = sr.ReadLine();
						while (line != null) 
						{
							if (line!="")
							{
								cnt++;
								itm = listView1.Items.Add(cnt.ToString());
								sitm =  itm.SubItems.Add ("importing...");
								listView1.Refresh();
								try 
								{		
									companyName = line.Substring(3).Trim();

									while ((line = sr.ReadLine()) != "") 
									{
										if (line == null) break;
										if (line.StartsWith("Address:\t"))
										{
											address = line.Replace("Address:\t","").Trim();
										}
										if (line.StartsWith("Phone:\t"))
										{
											phone = line.Replace("Phone:\t","").Trim();
										}
										if (line.StartsWith("Fax:\t"))
										{
											fax = line.Replace("Fax:\t","").Trim();
										}
										if (line.StartsWith("E-mail:\t"))
										{
											email = line.Replace("E-mail:\t","").Trim();
										}
										if (line.StartsWith("Website:\t"))
										{
											website = line.Replace("Website:\t","").Trim();
										}
										if (line.StartsWith("Contact:\t"))
										{
											contact = line.Replace("Contact:\t","").Trim();
											if (contact.IndexOf(",")>0)
											{
												title = contact.Substring(contact.IndexOf(",") + 1).Trim();
												contact = contact.Substring(0, contact.IndexOf(","));
											}
											else
												title = "";
											if (contact.StartsWith("Mr. ")) contact = contact.Substring(3).Trim();
											if (contact.StartsWith("Mr ")) contact = contact.Substring(3).Trim();
											if (contact.StartsWith("Ms. ")) contact = contact.Substring(3).Trim();
											if (contact.StartsWith("Ms ")) contact = contact.Substring(3).Trim();
										}
									}
									string cmd = "INSERT INTO users (CompanyName,Address,Phone,Fax,Email,WebSite,Name,Title,EmailList)" 
										+ " VALUES ('" + companyName
										+ "','" + db.EscapeString(address)
										+ "','" + db.EscapeString(phone)
										+ "','" + db.EscapeString(fax)
										+ "','" + db.EscapeString(email)
										+ "','" + db.EscapeString(website)
										+ "','" + db.EscapeString(contact)
										+ "','" + db.EscapeString(title)
										+ "', Yes)";
									db.GetCommandObject(cmd, true);
									sitm.Text = " imported";
								}
								catch (Exception ex)
								{
									sitm.Text = ex.Message;
								}
								itm.EnsureVisible();
								listView1.Refresh();
							}
							else
								line = sr.ReadLine();	
						}
						sr.Close();
					}
					fs.Close();
					MessageBox.Show(cnt + " record(s) imported.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
					MessageBox.Show(textBox1.Text + " not found", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			btnCancel.Text= "&Close";
		}

		private void btnBrowse_Click(object sender, System.EventArgs e)
		{
			if (openFileDialog1.ShowDialog()==DialogResult.OK)
			{
				textBox1.Text = openFileDialog1.FileName;
			}
		}

		private void btnShowODBC_Click(object sender, System.EventArgs e)
		{
			FaenaMailLibrary.tools tools = new FaenaMailLibrary.tools();
			tools.Shell( "Rundll32.exe shell32.dll,Control_RunDLL Odbccp32.cpl");
			string [] a = tools.LoadDSNList();
			comboBox1.Items.Clear();
			if (a!=null) comboBox1.Items.AddRange(a);
			comboBox1.Text = reg.GetSetting(@"Faena Mail\TXT2DB", "ODBCDSN","");
		}

		private void textBox1_TextChanged(object sender, System.EventArgs e)
		{
			btnGo.Enabled=true;
		}

		private void comboBox1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			btnGo.Enabled=true;
		}
	}
}
