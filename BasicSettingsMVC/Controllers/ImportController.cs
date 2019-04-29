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

            //更新新增(不作删除操作)
            #region BizType
            //待更新
            List<BizType> entityBizTypes4update = _context.BizType.Where(w => listBizType.Select(s=>s.Name).Contains(w.Name)).ToList<BizType>();
            List<BizType> listBizTypeRemove = new List<BizType>();
            foreach (BizType bt in entityBizTypes4update)
            {
                foreach(BizType b in listBizType)
                {
                    if(b.Name == bt.Name)
                    {
                        bt.Disable = b.Disable;
                        bt.Desc = b.Desc;
                        listBizTypeRemove.Add(b);
                    }
                }
            }
            _context.BizType.UpdateRange(entityBizTypes4update);

            //新增
            IEnumerable<BizType> listBizTypeInsert = listBizType.Except(listBizTypeRemove);
            List<BizType> entityBizTypes4Add = new List<BizType>();
            if (listBizTypeInsert?.Count() > 0)
            {
                foreach (BizType newBizType in listBizTypeInsert)
                {
                    entityBizTypes4Add.Add(new BizType
                    {
                        Code = newBizType.Name,
                        Name = newBizType.Name,
                        Desc = newBizType.Desc
                    });
                }
            }
            _context.BizType.AddRange(entityBizTypes4Add);

            #endregion

            #region GoodsClass
            //待更新
            List<GoodsClass> entityGoodsClass4update = _context.GoodsClass.Where(w => listGoodsClass.Select(s => s.Name).Contains(w.Name)).ToList<GoodsClass>();
            List<GoodsClass> listGoodsClassRemove = new List<GoodsClass>();
            foreach (GoodsClass gc in entityGoodsClass4update)
            {
                foreach (GoodsClass g in listGoodsClass)
                {
                    if (g.Name == gc.Name)
                    {
                        gc.Disable = g.Disable;
                        gc.Desc = g.Desc;

                        listGoodsClassRemove.Add(g);
                    }
                }
            }
            _context.GoodsClass.UpdateRange(entityGoodsClass4update);

            //新增
            IEnumerable<GoodsClass> listGoodsClassInsert = listGoodsClass.Except(listGoodsClassRemove);
            List<GoodsClass> entityGoodsClass4Add = new List<GoodsClass>();
            if (listGoodsClassInsert?.Count() > 0)
            {
                foreach (GoodsClass newGoodsClass in listGoodsClassInsert)
                {
                    entityGoodsClass4Add.Add(new GoodsClass
                    {
                        Code = newGoodsClass.Name,
                        Name = newGoodsClass.Name,
                        Desc = newGoodsClass.Desc
                    });
                }
            }
            _context.BizType.AddRange(entityBizTypes4Add);

            #endregion

            _context.SaveChanges();


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