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
	/// PostScript에 대한 요약 설명입니다.
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
					//<input type="file" ...> 폼요소이고 업로드할 파일을 지정한 경우만 화면에 보여줌.
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

		#region Web Form 디자이너에서 생성한 코드
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 이 호출은 ASP.NET Web Form 디자이너에 필요합니다.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 디자이너 지원에 필요한 메서드입니다.
		/// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
	}
}
