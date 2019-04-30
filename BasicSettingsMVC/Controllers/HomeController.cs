using System;
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

        [Authorize]
        public IActionResult Index()
        {
            if (ModelState.IsValid) { }
                return View();
        }

        public IActionResult Export()
        {
            FileStream fsReuslt = null;
            try
            {
                #region ///读取模板
                XSSFWorkbook wk = null;

                FileInfo strFileTemplate = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath + @"\\Template", "BasicSettingsTemplate.xlsx"));
                using (FileStream excelTemplate = new FileStream(strFileTemplate.ToString(), FileMode.Open))
                {
                    //把xls文件读入workbook变量里后就关闭
                    wk = new XSSFWorkbook(excelTemplate);
                    excelTemplate.Close();
                }
                #endregion

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
                #endregion

                //写入新对象
                int result = ExcelUtil.SetDataSet2Workbook(ds, wk);


                //生成新excel
                FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath + @"\\DownLoad", DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".xlsx"));
                using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
                { 
                    wk.Write(fs);
                }

                fsReuslt = System.IO.File.OpenRead(file.ToString());

                return File(fsReuslt, XlsxContentType, DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".xlsx");
            }
            catch (Exception ex)
            {
                return NotFound();
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
