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
                //读取模板
                //string strTemplateFileName = @"C:\\Users\\Juan\\source\\repos\\WebApplicationTestNPOI\\WebApplicationTestNPOI\\BasicSettingsTemplate.xlsx";
                XSSFWorkbook wk = null;

                FileInfo strFileTemplate = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath + @"\\Template", "BasicSettingsTemplate.xlsx"));
                using (FileStream excelTemplate = new FileStream(strFileTemplate.ToString(), FileMode.Open))
                {
                    //把xls文件读入workbook变量里后就关闭
                    wk = new XSSFWorkbook(excelTemplate);
                    excelTemplate.Close();
                }

                //读取数据 从db
                DataSet ds = new DataSet();

                List<BizType> listBizType = _context.BizType.ToList();
                DataTable dtBizType = DbModel.ToDataTableKeyValue(listBizType, ExcelUtil.BizTypeDictionary);
                dtBizType.TableName = ExcelUtil.BizTypeDataTableName;
                ds.Tables.Add(dtBizType);


                List<GoodsUnit> listGoodsUnit = _context.GoodsUnit.ToList();
                DataTable dtGoodsUnit = DbModel.ToDataTableKeyValue(listGoodsUnit, ExcelUtil.GoodsUnitDictionary);
                dtGoodsUnit.TableName = ExcelUtil.GoodsUnitDataTableName;
                ds.Tables.Add(dtGoodsUnit);

                List<GoodsClass> listGoodsClass = _context.GoodsClass.ToList();
                DataTable dtGoodsClass = DbModel.ToDataTableKeyValue(listGoodsClass, ExcelUtil.GoodsClassDictionary);
                dtGoodsClass.TableName = ExcelUtil.GoodsClassDataTableName;
                ds.Tables.Add(dtGoodsClass);

                List<Goods> listGoods = _context.Goods.ToList();
                DataTable dtGoods = DbModel.ToDataTable(listGoods, ExcelUtil.GoodsModelPropertyArray);
                dtGoods.TableName = ExcelUtil.GoodsDataTableName;
                ds.Tables.Add(dtGoods);


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

        [HttpPost]
        public async Task<IActionResult> Import(List<IFormFile> files)
        {

            //文件上传保存
            string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + Path.GetExtension(files[0].FileName);
            var filePath = Path.Combine(_hostingEnvironment.ContentRootPath + @"\\Upload", fileName);
            FileInfo strFileData = new FileInfo(filePath);
            using (FileStream excelData = new FileStream(strFileData.ToString(), FileMode.Open))
            {
                await files[0].CopyToAsync(excelData);
                excelData.Close();
            }
            //读取数据文件构造DataSet
            DataSet ds;
            using (FileStream excelData = new FileStream(filePath, FileMode.Open))
            {
                ds = ExcelUtil.GetDataSet(excelData);
                excelData.Close();
            }
            //批量写入数据库
            List<BizType> listBizType = DbModel.ToListKeyValue<BizType>(ds.Tables[ExcelUtil.BizTypeDataTableName], ExcelUtil.BizTypeModelPropertyArray);

            return Ok("200");
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
