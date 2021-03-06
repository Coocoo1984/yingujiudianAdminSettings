﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using BasicSettingsMVC.Models;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Data;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;

namespace BasicSettingsMVC.Controllers
{
    public class HomeController : Controller
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private IHostingEnvironment _hostingEnvironment;
        private MyDbContext _context;

        public HomeController(IHostingEnvironment hostingEnvironment, MyDbContext context)
        {
            _hostingEnvironment = hostingEnvironment;
            _context = context;
        }

        [HttpGet]
        [Authorize]
        public IActionResult Index()
        {
            if (ModelState.IsValid) { }
            return View();
        }

        public class UsrPermissions {
            public string WechatID;

        }

        [HttpGet]
        [Authorize]
        public IActionResult Export()
        {
            FileStream fsReuslt = null;
            try
            {
                #region ///读取模板
                XSSFWorkbook wk = null;
                FileInfo strFileTemplate = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath, "Template", "BasicSettingsTemplate.xlsx"));
                using (FileStream excelTemplate = new FileStream(strFileTemplate.ToString(), FileMode.Open))
                {
                    //把xls文件读入workbook变量里后就关闭
                    wk = new XSSFWorkbook(excelTemplate);
                    excelTemplate.Close();
                }
                #endregion
                Console.WriteLine(wk.ToString());//
                #region ///读取数据 从db
                DataSet ds = new DataSet();
                
                List<BizType> listBizType = _context.BizType.ToList();
                DataTable dtBizType = DbModel.ToDataTableKeyValue(listBizType, ExcelUtil.BizTypeDataTableName, ExcelUtil.BizTypeDictionary);
                ds.Tables.Add(dtBizType);
                

                List<GoodsUnit> listGoodsUnit = _context.GoodsUnit
                    .Include(i=>i.Goods)
                    .ToList();
                DataTable dtGoodsUnit = DbModel.ToDataTableKeyValue(listGoodsUnit, ExcelUtil.GoodsUnitDataTableName, ExcelUtil.GoodsUnitDictionary);
                ds.Tables.Add(dtGoodsUnit);

                List<GoodsClass> listGoodsClass = _context.GoodsClass
                    .Include(i=>i.BizType)
                    .ToList();
                DataTable dtGoodsClass = DbModel.ToDataTableKeyValue(listGoodsClass, ExcelUtil.GoodsClassDataTableName, ExcelUtil.GoodsClassDictionary);
                ds.Tables.Add(dtGoodsClass);

                List<Goods> listGoods = _context.Goods
                    .Include(i=>i.GoodsClass)
                    .Include(i=>i.GoodsUnit)
                    .ToList();
                DataTable dtGoods = DbModel.ToDataTableKeyValue(listGoods, ExcelUtil.GoodsDataTableName, ExcelUtil.GoodsDictionary);
                ds.Tables.Add(dtGoods);

                List<RsPermission> listRsPermission = _context.RsPermission
                    .Include(i => i.Permission)
                    .ToList();

                var wechatids = listRsPermission.Select(s => s.UsrWechatId).Distinct().ToList();

                List<Usr> listUsrs = //new List<Usr>();
                    _context.Usr.Include(i => i.Role).ToList();
                foreach (string wechatid in wechatids)
                {
                    Usr usr = listUsrs.SingleOrDefault(s => s.WechatID.Equals(wechatid));//微信ID匹配用户权限 暂不将权限与角色关联 角色仅存在于发送消息
                    if (usr?.ID > 0)
                    {
                        var wechatList = listRsPermission.Where(w => w.UsrWechatId.Equals(wechatid)).ToList();
                        foreach (var item in wechatList)
                        {
                            switch (item.PermissionId)
                            {
                                case 1: usr.QuoteDetailRead = item.Disable ? "否" : "是"; break;
                                case 2: usr.QuoteAudit = item.Disable ? "否" : "是"; break;
                                case 3: usr.QuoteAudit2 = item.Disable ? "否" : "是"; break;
                                case 4: usr.PurchaceAudit = item.Disable ? "否" : "是"; break;
                                case 5: usr.PurchaceAudit2 = item.Disable ? "否" : "是"; break;
                                case 6: usr.PurchaceAudit3 = item.Disable ? "否" : "是"; break;
                                case 7: usr.ChargeBackAudit = item.Disable ? "否" : "是"; break;
                                case 8: usr.DepotAdmin = item.Disable ? "否" : "是"; break;
                                case 9: usr.ReportExport = item.Disable ? "否" : "是"; break;
                                case 10: usr.QuoteCommit = item.Disable ? "否" : "是"; break;
                                case 11: usr.PurchaceCommit = item.Disable ? "否" : "是"; break;
                                case 12: usr.ChargeBackCommit = item.Disable ? "否" : "是"; break;
                            }
                        }
                    }
                }
                DataTable dtUsrs = DbModel.ToDataTableKeyValue(listUsrs, ExcelUtil.RsPermissionDataTableName, ExcelUtil.RsPermissionDictionary);
                ds.Tables.Add(dtUsrs);


                #endregion

                //写入新对象
                int result = ExcelUtil.SetDataSet2Workbook(ds, wk);
                Console.WriteLine($"ExcelUtil.SetDataSet2Workbook(ds, wk):{result}");//

                //生成新excel
                FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath, "Download", DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".xlsx"));
                using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
                { 
                    wk.Write(fs);
                }

                fsReuslt = System.IO.File.OpenRead(file.ToString());

                return File(fsReuslt, XlsxContentType, DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".xlsx");
            }
            catch (Exception ex)
            {
                // excel sheet中DataTtable默认选中范围太小引发 NPOI对象越界
                return new NotFoundObjectResult(ex.Message);
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
