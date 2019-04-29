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
using System.Collections.Generic;
using System;

namespace BasicSettingsMVC.Controllers
{
    public class ImportController : Controller
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private IHostingEnvironment _hostingEnvironment;
        private MyDbContext _context;

        public ImportController(IHostingEnvironment hostingEnvironment, MyDbContext context)
        {
            _hostingEnvironment = hostingEnvironment;
            _context = context;
        }
        [HttpGet]
        [Authorize]
        public IActionResult Index()
        {

            FileInfo strDataFile = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath + @"\\Template", "BasicSettingsWithData.xlsx"));

            //读取数据文件构造DataSet
            DataSet ds;
            using (FileStream excelData = new FileStream(strDataFile.ToString(), FileMode.Open))
            {
                ds = ExcelUtil.GetDataSet(excelData);
                excelData.Close();
            }

            //转换为模型对象集合
            List<BizType> listBizType = DbModel.ToListKeyValue<BizType>(ds.Tables[ExcelUtil.BizTypeDataTableName], ExcelUtil.BizTypeDictionary);
            List<GoodsClass> listGoodsClass = DbModel.ToListKeyValue<GoodsClass>(ds.Tables[ExcelUtil.GoodsClassDataTableName], ExcelUtil.GoodsClassDictionary);
            List<GoodsUnit> listGoodsUnit = DbModel.ToListKeyValue<GoodsUnit>(ds.Tables[ExcelUtil.GoodsUnitDataTableName], ExcelUtil.GoodsUnitDictionary);
            List<Goods> listGoods = DbModel.ToListKeyValue<Goods>(ds.Tables[ExcelUtil.GoodsDataTableName], ExcelUtil.GoodsDictionary);


            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Index(IFormFileCollection files)
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
            //转换为模型对象集合
            List<BizType> listBizType = DbModel.ToListKeyValue<BizType>(ds.Tables[ExcelUtil.BizTypeDataTableName], ExcelUtil.BizTypeDictionary);
            List<GoodsClass> listGoodsClass = DbModel.ToListKeyValue<GoodsClass>(ds.Tables[ExcelUtil.GoodsClassDataTableName], ExcelUtil.GoodsClassDictionary);
            List<GoodsUnit> listGoodsUnit = DbModel.ToListKeyValue<GoodsUnit>(ds.Tables[ExcelUtil.GoodsUnitDataTableName], ExcelUtil.GoodsUnitDictionary);
            List<Goods> listGoods = DbModel.ToListKeyValue<Goods>(ds.Tables[ExcelUtil.GoodsDataTableName], ExcelUtil.GoodsDictionary);

            return View();
        }
    }
}