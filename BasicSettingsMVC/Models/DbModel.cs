using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace BasicSettingsMVC.Models
{
    public static class DbModel
    {
        public static bool ImportAll(DataSet ds)
        {
            bool result = false;

            DataTable dtBizTypes = ds.Tables[ExcelUtil.BizTypeModelName];
            DataTable dtGoodsClass = ds.Tables[ExcelUtil.GoodsClassModelName];
            DataTable dtGoodsUnit = ds.Tables[ExcelUtil.GoodsUnitModelName];
            DataTable dtGoods = ds.Tables[ExcelUtil.GoodsModelName];

            List<BizType> listBizTypes = ToList<BizType>(dtBizTypes);
            List<GoodsClass> listGoodsClass = ToList<GoodsClass>(dtGoodsClass);
            List<GoodsUnit> listGoodsUnit = ToList<GoodsUnit>(dtGoodsUnit);
            List<Goods> listGoods = ToList<Goods>(dtGoods);



            return result;
        }


        public static DataSet ExportAll()
        {
            DataSet result = new DataSet();

            return result;
        }

        /// <summary>    
        /// DataTable 转换为List 集合    
        /// </summary>    
        /// <typeparam name="TResult">类型</typeparam>    
        /// <param name="dt">DataTable</param>    
        /// <returns></returns>    
        public static List<T> ToList<T>(this DataTable dt) where T : class, new()
        {
            //创建一个属性的列表    
            List<PropertyInfo> prlist = new List<PropertyInfo>();
            //获取TResult的类型实例  反射的入口    

            Type t = typeof(T);

            //获得TResult 的所有的Public 属性 并找出TResult属性和DataTable的列名称相同的属性(PropertyInfo) 并加入到属性列表     
            Array.ForEach<PropertyInfo>(t.GetProperties(), p => { if (dt.Columns.IndexOf(p.Name) != -1) prlist.Add(p); });

            //创建返回的集合    

            List<T> oblist = new List<T>();

            foreach (DataRow row in dt.Rows)
            {
                //创建TResult的实例    
                T ob = new T();
                //找到对应的数据  并赋值    
                prlist.ForEach(p => { if (row[p.Name] != DBNull.Value) p.SetValue(ob, row[p.Name], null); });
                //放入到返回的集合中.    
                oblist.Add(ob);
            }
            return oblist;
        }

        /// <summary>    
        /// DataTable 转换为List 集合    
        /// </summary>    
        /// <typeparam name="TResult">类型</typeparam>    
        /// <param name="dt">DataTable</param>    
        /// <returns></returns>    
        public static List<T> ToListKeyValue<T>(this DataTable dt, string[] propertyArray) where T : class, new()
        {
            //创建一个属性的列表    
            List<PropertyInfo> prlist = new List<PropertyInfo>();
            //获取TResult的类型实例  反射的入口    

            Type t = typeof(T);

            //获得TResult 的所有的Public 属性 并找出TResult属性和DataTable的列名称相同的属性(PropertyInfo) 并加入到属性列表     
            Array.ForEach<PropertyInfo>(
                t.GetProperties(), 
                p => {
                    if (dt.Columns.IndexOf(p.Name) != -1)
                        prlist.Add(p);
                });

            //创建返回的集合    

            List<T> result = new List<T>();

            foreach (DataRow row in dt.Rows)
            {
                //创建TResult的实例    
                T ob = new T();
                //找到对应的数据  并赋值    
                prlist.ForEach(
                    p => {
                        if (propertyArray.Contains(p.Name))
                        {
                            if(row[p.Name] != DBNull.Value)
                                p.SetValue(ob, row[p.Name], null);
                        }
                    });
                //放入到返回的集合中.    
                result.Add(ob);
            }
            return result;
        }

        /// <summary>     
        /// 转化一个DataTable    
        /// </summary>      
        /// <typeparam name="T"></typeparam>    
        /// <param name="list"></param>    
        /// <returns></returns>    
        public static DataTable ToDataTable<T>(this IEnumerable<T> list)
        {
            //创建属性的集合    
            List<PropertyInfo> pList = new List<PropertyInfo>();
            //获得反射的入口    
            Type type = typeof(T);
            DataTable dt = new DataTable();
            //把所有的public属性加入到集合 并添加DataTable的列    
            Array.ForEach<PropertyInfo>(type.GetProperties(), p => { pList.Add(p); dt.Columns.Add(p.Name, p.PropertyType); });
            foreach (var item in list)
            {
                //创建一个DataRow实例    
                DataRow row = dt.NewRow();
                //给row 赋值    
                pList.ForEach(p => row[p.Name] = p.GetValue(item, null));
                //加入到DataTable    
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>    
        /// 将泛型集合类转换成DataTable    
        /// </summary>    
        /// <typeparam name="T">集合项类型</typeparam>    

        /// <param name="list">集合</param>    
        /// <returns>数据集(表)</returns>    
        public static DataTable ToDataTable<T>(IList<T> list)
        {
            return ToDataTable<T>(list, null);

        }




        /// <summary>    
        /// 将泛型集合类转换成DataTable    
        /// </summary>    
        /// <typeparam name="T">集合项类型</typeparam>    
        /// <param name="list">集合</param>    
        /// <param name="propertyName">需要返回的列的列名</param>    
        /// <returns>数据集(表)</returns>    
        public static DataTable ToDataTable<T>(IList<T> list, params string[] propertyName)
        {
            DataTable result = new DataTable();
            
            foreach (string name in propertyName)
            {
                result.Columns.Add(name, typeof(string));
            }

            if (list.Count > 0)
            {
                PropertyInfo[] propertys = list[0].GetType().GetProperties();
                for (int i = 0; i < list.Count; i++)
                {
                    DataRow dr = result.NewRow();
                    foreach (string name in propertyName)
                    {
                        PropertyInfo pi = propertys.ToList().Single(s => s.Name.Equals(name));
                        dr[name] = (pi.GetValue(list[i], null)?.ToString() ?? string.Empty);
                    }
                    result.Rows.Add(dr);
                }
            }
            return result;
        }

        /// <summary>    
        /// 将泛型集合类转换成DataTable    
        /// </summary>    
        /// <typeparam name="T">集合项类型</typeparam>    
        /// <param name="list">集合</param>    
        /// <param name="keyValuePairs">key:T字段名;value:Datatable列名</param>    
        /// <returns>数据集(表)</returns>   
        public static DataTable ToDataTableKeyValue<T>(IList<T> list, Dictionary<string, Tuple<string, string>> keyValuePairs)
        {
            DataTable result = new DataTable();
            
            if(list != null && list.Count > 0 && keyValuePairs != null && keyValuePairs.Count > 0)
            {
                foreach (var kv in keyValuePairs)
                {
                    //循环构造列
                    result.Columns.Add(kv.Value.Item2, typeof(string));
                }

                PropertyInfo[] propertys = list[0].GetType().GetProperties();

                if (list.Count > 0)
                {
                    for (int i = 0; i < list.Count; i++)
                    {
                        //循环填充数据
                        DataRow dr = result.NewRow();
                        foreach (PropertyInfo pi in propertys)
                        {
                            //遍历T的每个属性
                            if(keyValuePairs.ContainsKey(pi.Name))
                            {
                                //定义的需要导出到excel的列名

                                    //需要显示的字段名
                                    dr[keyValuePairs[pi.Name].Item2] = pi.GetValue(list[i], null)?.ToString() ?? string.Empty;
                            }
                        }
                        result.Rows.Add(dr);
                    }
                }
            }
            
            return result;
        }
    }
}
