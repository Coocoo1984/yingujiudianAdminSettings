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
using Microsoft.EntityFrameworkCore;

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
            #region
            ////FileInfo strDataFile = new FileInfo(Path.Combine(_hostingEnvironment.ContentRootPath,"Template", "BasicSettingsWithData.xlsx"));

            //////读取数据文件构造DataSet
            ////DataSet ds;
            ////using (FileStream excelData = new FileStream(strDataFile.ToString(), FileMode.Open))
            ////{
            ////    ds = ExcelUtil.GetDataSet(excelData);
            ////    excelData.Close();
            ////}

            //////转换为模型对象集合
            ////List<BizType> listBizType = DbModel.ToListKeyValue<BizType>(ds.Tables[ExcelUtil.BizTypeDataTableName], ExcelUtil.BizTypeDictionary);
            ////List<GoodsClass> listGoodsClass = DbModel.ToListKeyValue<GoodsClass>(ds.Tables[ExcelUtil.GoodsClassDataTableName], ExcelUtil.GoodsClassDictionary);
            ////List<GoodsUnit> listGoodsUnit = DbModel.ToListKeyValue<GoodsUnit>(ds.Tables[ExcelUtil.GoodsUnitDataTableName], ExcelUtil.GoodsUnitDictionary);
            ////List<Goods> listGoods = DbModel.ToListKeyValue<Goods>(ds.Tables[ExcelUtil.GoodsDataTableName], ExcelUtil.GoodsDictionary);
            ////List<RsPermission> listRsPermission = DbModel.ToListKeyValue<RsPermission>(ds.Tables[ExcelUtil.GoodsDataTableName], ExcelUtil.GoodsDictionary);

            //////更新新增(不作删除操作)

            ////#region BizType
            //////待更新
            ////List<BizType> entityBizTypes4update = _context.BizType.Where(w => listBizType.Select(s=>s.Name).Contains(w.Name)).ToList<BizType>();
            ////List<BizType> listBizTypeRemove = new List<BizType>();
            ////foreach (BizType bt in entityBizTypes4update)
            ////{
            ////    foreach(BizType b in listBizType)
            ////    {
            ////        if(b.Name == bt.Name)
            ////        {
            ////            bt.Disable = b.Disable;
            ////            bt.Desc = b.Desc;
            ////            listBizTypeRemove.Add(b);
            ////        }
            ////    }
            ////}
            ////_context.BizType.UpdateRange(entityBizTypes4update);

            //////新增
            ////IEnumerable<BizType> listBizTypeInsert = listBizType.Except(listBizTypeRemove);
            ////if (listBizTypeInsert?.Count() > 0)
            ////{
            ////    foreach (BizType newBizType in listBizTypeInsert)
            ////    {
            ////        if (newBizType.Code == null)
            ////        {
            ////            newBizType.Code = newBizType.Name;
            ////        }
            ////        if (newBizType.Desc == null)
            ////        {
            ////            newBizType.Desc = newBizType.Desc;
            ////        }
            ////    }
            ////}
            ////_context.BizType.AddRange(listBizTypeInsert);

            ////#endregion

            ////#region GoodsClass
            //////待更新
            ////List<GoodsClass> entityGoodsClass4update = _context.GoodsClass.Include(i=>i.BizType).Where(w => listGoodsClass.Select(s => s.Name).Contains(w.Name)).ToList<GoodsClass>();
            ////List<GoodsClass> listGoodsClassRemove = new List<GoodsClass>();
            ////foreach (GoodsClass gc in entityGoodsClass4update)
            ////{
            ////    foreach (GoodsClass g in listGoodsClass)
            ////    {
            ////        if (g.Name == gc.Name)
            ////        {
            ////            gc.Disable = g.Disable;
            ////            gc.Specification = g.Specification;
            ////            gc.Desc = g.Desc;
            ////            if(g.BizTypeName != gc.BizType.Name)
            ////            {
            ////                g.BizTypeId = entityBizTypes4update.Where(w => g.BizTypeName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////            }

            ////            listGoodsClassRemove.Add(g);
            ////        }
            ////    }
            ////}
            ////_context.GoodsClass.UpdateRange(entityGoodsClass4update);

            //////新增
            ////IEnumerable<GoodsClass> listGoodsClassInsert = listGoodsClass.Except(listGoodsClassRemove);
            ////if (listGoodsClassInsert?.Count() > 0)
            ////{
            ////    foreach (GoodsClass newGoodsClass in listGoodsClassInsert)
            ////    {
            ////        if(newGoodsClass.Specification == null)
            ////        {
            ////            newGoodsClass.Specification = newGoodsClass.Name;
            ////        }
            ////        if (newGoodsClass.Code == null)
            ////        {
            ////            newGoodsClass.Code = newGoodsClass.Name;
            ////        }
            ////        if (newGoodsClass.Desc == null)
            ////        {
            ////            newGoodsClass.Desc = newGoodsClass.Name;
            ////        }
            ////        newGoodsClass.BizTypeId = entityBizTypes4update.Where(w => newGoodsClass.BizTypeName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////    }
            ////}
            ////_context.GoodsClass.AddRange(listGoodsClassInsert);

            ////#endregion

            ////#region GoodsUnit
            //////待更新
            ////List<GoodsUnit> entityGoodsUnit4update = _context.GoodsUnit.Where(w => listGoodsUnit.Select(s => s.Name).Contains(w.Name)).ToList<GoodsUnit>();
            ////List<GoodsUnit> listGoodsUnitRemove = new List<GoodsUnit>();
            ////foreach (GoodsUnit gc in entityGoodsUnit4update)
            ////{
            ////    foreach (GoodsUnit g in listGoodsUnit)
            ////    {
            ////        if (g.Name == gc.Name)
            ////        {
            ////            gc.Name = g.Name;
            ////            gc.Desc = g.Desc;


            ////            listGoodsUnitRemove.Add(g);
            ////        }
            ////    }
            ////}
            ////_context.GoodsUnit.UpdateRange(entityGoodsUnit4update);

            //////新增
            ////IEnumerable<GoodsUnit> listGoodsUnitInsert = listGoodsUnit.Except(listGoodsUnitRemove);
            ////if (listGoodsUnitInsert?.Count() > 0)
            ////{
            ////    foreach (GoodsUnit newGoodsUnit in listGoodsUnitInsert)
            ////    {
            ////        if (newGoodsUnit.Code == null)
            ////        {
            ////            newGoodsUnit.Code = newGoodsUnit.Name;
            ////        }
            ////        if (newGoodsUnit.Desc == null)
            ////        {
            ////            newGoodsUnit.Desc = newGoodsUnit.Name;
            ////        }
            ////    }
            ////}
            ////_context.GoodsUnit.AddRange(listGoodsUnitInsert);

            ////#endregion

            ////#region Goods
            //////待更新
            ////List<Goods> entityGoods4update = _context.Goods

            ////    .Include(i => i.GoodsUnit)
            ////    .Include(i => i.GoodsClass)
            ////    .Where(w => listGoods.Select(s => s.Name).Contains(w.Name))
            ////    .ToList<Goods>();
            ////List<Goods> listGoodsRemove = new List<Goods>();
            ////foreach (Goods gc in entityGoods4update)
            ////{
            ////    foreach (Goods g in listGoods)
            ////    {
            ////        if (g.Name == gc.Name)
            ////        {
            ////            gc.Disable = g.Disable;
            ////            gc.Specification = g.Specification;
            ////            gc.Desc = g.Desc;
            ////            if (g.GoodsClassName != gc.GoodsClass.Name)
            ////            {
            ////                g.GoodsClassId = entityGoodsClass4update.Where(w => g.GoodsClassName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////            }
            ////            if (g.GoodsUnitName != gc.GoodsUnit.Name)
            ////            {
            ////                g.GoodsUnitId = entityGoodsUnit4update.Where(w => g.GoodsUnitName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////            }
            ////            listGoodsRemove.Add(g);
            ////        }
            ////    }
            ////}
            ////_context.Goods.UpdateRange(entityGoods4update);

            //////新增
            ////IEnumerable<Goods> listGoodsInsert = listGoods.Except(listGoodsRemove);
            ////if (listGoodsInsert?.Count() > 0)
            ////{
            ////    foreach (Goods newGoods in listGoodsInsert)
            ////    {
            ////        if (newGoods.Specification == null)
            ////        {
            ////            newGoods.Specification = newGoods.Name;
            ////        }
            ////        if (newGoods.Code == null)
            ////        {
            ////            newGoods.Code = newGoods.Name;
            ////        }
            ////        if (newGoods.Desc == null)
            ////        {
            ////            newGoods.Desc = newGoods.Name;
            ////        }
            ////        newGoods.GoodsClassId = entityGoodsClass4update.Where(w => newGoods.GoodsClassName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////        newGoods.GoodsUnitId = entityGoodsUnit4update.Where(w => newGoods.GoodsUnitName.Equals(w.Name)).SingleOrDefault()?.Id;
            ////    }
            ////}
            ////_context.Goods.AddRange(listGoodsInsert);

            ////#endregion

            ////#region Permission




            ////#endregion

            ////_context.SaveChanges();

            #endregion
            return View();
        }

        [HttpPost]
        [Authorize]
        public async Task<IActionResult> Index(IFormFileCollection files)
        {
            //文件上传保存
            string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + Path.GetExtension(files[0].FileName);
            var filePath = Path.Combine(_hostingEnvironment.ContentRootPath,"Upload", fileName);
            FileInfo strFileData = new FileInfo(filePath);
            using (FileStream excelData = new FileStream(strFileData.ToString(), FileMode.Create))
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
            List<Usr> listUsr = DbModel.ToListKeyValue<Usr>(ds.Tables[ExcelUtil.RsPermissionDataTableName], ExcelUtil.RsPermissionDictionary);

#region BizType
            //待更新
            List<BizType> entityBizTypes4update = _context.BizType.Where(w => listBizType.Select(s => s.Name).Contains(w.Name)).ToList<BizType>();
            List<BizType> listBizTypeRemove = new List<BizType>();
            foreach (BizType bt in entityBizTypes4update)
            {
                foreach (BizType b in listBizType)
                {
                    if (b.Name == bt.Name)
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
            if (listBizTypeInsert?.Count() > 0)
            {
                foreach (BizType newBizType in listBizTypeInsert)
                {
                    if (newBizType.Code == null)
                    {
                        newBizType.Code = newBizType.Name;
                    }
                    if (newBizType.Desc == null)
                    {
                        newBizType.Desc = newBizType.Desc;
                    }
                }
                _context.BizType.AddRange(listBizTypeInsert);
            }
            

#endregion

#region GoodsClass
            //待更新
            List<GoodsClass> entityGoodsClass4update = _context.GoodsClass.Include(i => i.BizType).Where(w => listGoodsClass.Select(s => s.Name).Contains(w.Name)).ToList<GoodsClass>();
            List<GoodsClass> listGoodsClassRemove = new List<GoodsClass>();
            foreach (GoodsClass gc in entityGoodsClass4update)
            {
                foreach (GoodsClass g in listGoodsClass)
                {
                    if (g.Name == gc.Name)
                    {
                        gc.Disable = g.Disable;
                        gc.Specification = g.Specification;
                        gc.Desc = g.Desc;
                        if (g.BizTypeName != gc.BizType.Name)
                        {
                            g.BizTypeId = entityBizTypes4update.Where(w => g.BizTypeName.Equals(w.Name)).SingleOrDefault()?.Id;
                        }

                        listGoodsClassRemove.Add(g);
                    }
                }
            }
            _context.GoodsClass.UpdateRange(entityGoodsClass4update);

            //新增
            IEnumerable<GoodsClass> listGoodsClassInsert = listGoodsClass.Except(listGoodsClassRemove);
            if (listGoodsClassInsert?.Count() > 0)
            {
                foreach (GoodsClass newGoodsClass in listGoodsClassInsert)
                {
                    if (newGoodsClass.Specification == null)
                    {
                        newGoodsClass.Specification = newGoodsClass.Name;
                    }
                    if (newGoodsClass.Code == null)
                    {
                        newGoodsClass.Code = newGoodsClass.Name;
                    }
                    if (newGoodsClass.Desc == null)
                    {
                        newGoodsClass.Desc = newGoodsClass.Name;
                    }
                    newGoodsClass.BizTypeId = entityBizTypes4update.Where(w => newGoodsClass.BizTypeName.Equals(w.Name)).SingleOrDefault()?.Id;
                }
            }
            _context.GoodsClass.AddRange(listGoodsClassInsert);

#endregion

#region GoodsUnit
            //待更新
            List<GoodsUnit> entityGoodsUnit4update = _context.GoodsUnit.Where(w => listGoodsUnit.Select(s => s.Name).Contains(w.Name)).ToList<GoodsUnit>();
            List<GoodsUnit> listGoodsUnitRemove = new List<GoodsUnit>();
            foreach (GoodsUnit gc in entityGoodsUnit4update)
            {
                foreach (GoodsUnit g in listGoodsUnit)
                {
                    if (g.Name == gc.Name)
                    {
                        gc.Name = g.Name;
                        gc.Desc = g.Desc;


                        listGoodsUnitRemove.Add(g);
                    }
                }
            }
            _context.GoodsUnit.UpdateRange(entityGoodsUnit4update);

            //新增
            IEnumerable<GoodsUnit> listGoodsUnitInsert = listGoodsUnit.Except(listGoodsUnitRemove);
            if (listGoodsUnitInsert?.Count() > 0)
            {
                foreach (GoodsUnit newGoodsUnit in listGoodsUnitInsert)
                {
                    if (newGoodsUnit.Code == null)
                    {
                        newGoodsUnit.Code = newGoodsUnit.Name;
                    }
                    if (newGoodsUnit.Desc == null)
                    {
                        newGoodsUnit.Desc = newGoodsUnit.Name;
                    }
                }
            }
            _context.GoodsUnit.AddRange(listGoodsUnitInsert);

#endregion

#region Goods
            //待更新
            List<Goods> entityGoods4update = _context.Goods

                .Include(i => i.GoodsUnit)
                .Include(i => i.GoodsClass)
                .Where(w => listGoods.Select(s => s.Name).Contains(w.Name))
                .ToList<Goods>();
            List<Goods> listGoodsRemove = new List<Goods>();
            foreach (Goods gc in entityGoods4update)
            {
                foreach (Goods g in listGoods)
                {
                    if (g.Name == gc.Name)
                    {
                        gc.Disable = g.Disable;
                        gc.Specification = g.Specification;
                        gc.Desc = g.Desc;
                        if (g.GoodsClassName != gc.GoodsClass.Name)
                        {
                            g.GoodsClassId = entityGoodsClass4update.Where(w => g.GoodsClassName.Equals(w.Name)).SingleOrDefault()?.Id;
                        }
                        if (g.GoodsUnitName != gc.GoodsUnit.Name)
                        {
                            g.GoodsUnitId = entityGoodsUnit4update.Where(w => g.GoodsUnitName.Equals(w.Name)).SingleOrDefault()?.Id;
                        }
                        listGoodsRemove.Add(g);
                    }
                }
            }
            _context.Goods.UpdateRange(entityGoods4update);

            //新增
            IEnumerable<Goods> listGoodsInsert = listGoods.Except(listGoodsRemove);
            if (listGoodsInsert?.Count() > 0)
            {
                foreach (Goods newGoods in listGoodsInsert)
                {
                    if (newGoods.Specification == null)
                    {
                        newGoods.Specification = newGoods.Name;
                    }
                    if (newGoods.Code == null)
                    {
                        newGoods.Code = newGoods.Name;
                    }
                    if (newGoods.Desc == null)
                    {
                        newGoods.Desc = newGoods.Name;
                    }
                    newGoods.GoodsClassId = entityGoodsClass4update.Where(w => newGoods.GoodsClassName.Equals(w.Name)).SingleOrDefault()?.Id;
                    newGoods.GoodsUnitId = entityGoodsUnit4update.Where(w => newGoods.GoodsUnitName.Equals(w.Name)).SingleOrDefault()?.Id;
                }
            }
            _context.Goods.AddRange(listGoodsInsert);

            #endregion

            #region Permission

            //待更新
            List<RsPermission> listRsPermission = new List<RsPermission>();
            foreach(Usr usr  in listUsr)
            {
                if (!string.IsNullOrWhiteSpace(usr.QuoteDetailRead))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 1,
                        Disable = usr.QuoteDetailRead == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.QuoteAudit))
                {
                    RsPermission obj = new RsPermission {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 2,
                        Disable = usr.QuoteAudit == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.QuoteAudit2))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 3,
                        Disable = usr.QuoteAudit2 == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.PurchaceAudit))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 4,
                        Disable = usr.PurchaceAudit == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.PurchaceAudit2))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 5,
                        Disable = usr.PurchaceAudit2 == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.PurchaceAudit3))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 6,
                        Disable = usr.PurchaceAudit3 == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.ChargeBackAudit))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 7,
                        Disable = usr.ChargeBackAudit == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.DepotAdmin))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 8,
                        Disable = usr.DepotAdmin == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
                if (!string.IsNullOrWhiteSpace(usr.ReportExport))
                {
                    RsPermission obj = new RsPermission
                    {
                        UsrWechatId = usr.WechatID,
                        PermissionId = 9,
                        Disable = usr.ReportExport == "是" ? true : false
                    };
                    listRsPermission.Add(obj);
                }
            }



            //待更新
            List<RsPermission> entityRsPermission4update = _context.RsPermission.Where(w => listRsPermission.Select(s => s.UsrWechatId).Contains(w.UsrWechatId)).ToList<RsPermission>();
            List<RsPermission> listRsPermissionRemove = new List<RsPermission>();
            foreach (RsPermission rp in entityRsPermission4update)
            {
                foreach (RsPermission b in listRsPermission)
                {
                    if (b.UsrWechatId == rp.UsrWechatId && b.PermissionId == rp.PermissionId)
                    {
                        rp.Disable = b.Disable;
                        listRsPermissionRemove.Add(b);
                    }
                }
            }
            _context.RsPermission.UpdateRange(entityRsPermission4update);

            //新增
            IEnumerable<RsPermission> listRsPermissionInsert = listRsPermission.Except(listRsPermissionRemove);
            if (listRsPermissionInsert?.Count() > 0)
            {
                _context.RsPermission.AddRange(listRsPermissionInsert);
            }
            

            #endregion

            _context.SaveChanges();

            return RedirectToAction("Index", "Home");
        }
    }
}