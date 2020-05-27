using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;

namespace HuiQianDan2.Controllers
{
    public class HuiQianController : Controller
    {

        private string connString;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public HuiQianController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "appData\\" + "huiqiandan2.db";
            connString = "Data Source=" + dataPath;
        }

        public IActionResult Index()
        {
            string advice_wangwu = "";
            string advice_lisi = "";
            string advice_zhangsan = "";
            string strZSSignature = "";
            string id = Request.Query["id"];
            string user = Request.Query["user"];


            user = "zhangsan";
            string userName = "张三";
            if ("zhangsan".Equals(user))
            {
                userName = "张三";
            }
            else if ("lisi".Equals(user))
            {
                userName = "李四";
            }
            else if ("wangwu".Equals(user))
            {
                userName = "王五";
            }
            string sql = "select * from contract where id = " + id;

            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                advice_zhangsan = dr["Advice_zs"].ToString();
                advice_lisi = dr["Advice_ls"].ToString();
                advice_wangwu = dr["Advice_ww"].ToString();
                strZSSignature = dr["ZSSIGNATURE"].ToString();
                if (null == advice_zhangsan) advice_zhangsan = "";
                if (null == advice_lisi) advice_lisi = "";
                if (null == advice_wangwu) advice_wangwu = "";
                if (null == strZSSignature) strZSSignature = "";
            }
            dr.Close();
            conn.Close();

            PageOfficeNetCore.HtmlSignCtrl htmlSignCtrl = new PageOfficeNetCore.HtmlSignCtrl(Request);
            htmlSignCtrl.ServerPage = "../PageOffice/POServer";
            htmlSignCtrl.LoadToPage(strZSSignature, PageOfficeNetCore.HtmlSignMode.Signer, userName);
            ViewBag.htCtrl = htmlSignCtrl.GetHtmlCode("HtmlSignCtrl1");
            ViewBag.user = user;
            return View();
        }

        public IActionResult Word()
        {
            string user = Request.Query["user"];
            user = "wangwu";
            string userName = "";
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            pageofficeCtrl.Caption = "钢材采购合同";
            if (user == "zhangsan")
            {
                userName = "张三";
                pageofficeCtrl.CustomToolbar = false;
            }
            else if (user == "lisi")
            {
                userName = "李四";
                pageofficeCtrl.CustomToolbar = false;
            }
            else if (user == "wangwu")
            {
                userName = "王五";
                pageofficeCtrl.SaveFilePage = "SaveDoc";
                pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
                pageofficeCtrl.AddCustomToolButton("盖章", "AddSeal", 2);
                pageofficeCtrl.AddCustomToolButton("返回", "GoBack", 21);
            }
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/test.doc", PageOfficeNetCore.OpenModeType.docReadOnly, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.userName = userName;
            return View();
        }
        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }

    }
}