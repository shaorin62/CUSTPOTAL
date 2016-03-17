using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using DEXTUpload.NET;

namespace Dextx_dextnet
{
	/// <summary>
	/// PostScript�� ���� ��� �����Դϴ�.
	/// </summary>
	public class PostScript : System.Web.UI.Page
	{
		private void Page_Load(object sender, System.EventArgs e)
		{
			using(DEXTUpload.NET.FileUpload fileUpload = new DEXTUpload.NET.FileUpload())
			{
				string UploadedPath;

				for (int i=0; i<fileUpload["DextuploadX"].Count; i++)
				{
					//<input type="file" ...> ������̰� ���ε��� ������ ������ ��츸 ȭ�鿡 ������.
					if (fileUpload["DextuploadX"][i].IsFileElement && fileUpload["DextuploadX"][i].Value != "")
					{
						UploadedPath = fileUpload["DextuploadX"][i].Save(false);
						
						Response.Write("LastSavedFileName(Server) :" + fileUpload["DextuploadX"][i].LastSavedFileName + "<br>");
						Response.Write("Original Path (Client) :" + fileUpload["DextuploadX"][i].FilePath + "<br>");
						Response.Write("Uploaded Path (Server) :" + UploadedPath + "<br>");
						Response.Write("File Length :" + fileUpload["DextuploadX"][i].FileLength + " byte(s)<br>");
						Response.Write("File MimeType : </td><td>" + fileUpload["DextuploadX"][i].MimeType + "<br><br>");		
					}
				}
			}
		}

		#region Web Form �����̳ʿ��� ������ �ڵ�
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: �� ȣ���� ASP.NET Web Form �����̳ʿ� �ʿ��մϴ�.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// �����̳� ������ �ʿ��� �޼����Դϴ�.
		/// �� �޼����� ������ �ڵ� ������� �������� ���ʽÿ�.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
	}
}
