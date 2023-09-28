using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoPipelines
{
    public static class DictTools
    {
        /// <summary>
        /// 添加对象扩展字典项
        /// </summary>
        /// <param name="db"></param>
        /// <param name="searchKey"></param>
        /// <returns></returns>
        public static ObjectId AddNamedDictionary(this Database db, string searchKey)
        {
            ObjectId id = ObjectId.Null;
            DBDictionary dicts = (DBDictionary)db.NamedObjectsDictionaryId.GetObject(OpenMode.ForRead);
            if (!dicts.Contains(searchKey))
            {
                DBDictionary dict = new DBDictionary();
                dicts.UpgradeOpen();
                id = dicts.SetAt(searchKey, dict);
                dicts.DowngradeOpen();
                db.TransactionManager.AddNewlyCreatedDBObject(dict, true);
            }
            return id;
        }

        /// <summary>
        /// 添加一条扩展记录
        /// </summary>
        /// <param name="id"></param>
        /// <param name="searchKey"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ObjectId AddXrecord(this ObjectId id, string searchKey, ResultBuffer values)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            if (obj.ExtensionDictionary.IsNull)
            {
                obj.UpgradeOpen();
                obj.CreateExtensionDictionary();
                obj.DowngradeOpen();
            }
            DBDictionary dict = (DBDictionary)obj.ExtensionDictionary.GetObject(OpenMode.ForRead);
            if (dict.Contains(searchKey)) return ObjectId.Null;
            Xrecord xrec = new Xrecord();
            xrec.Data = values;
            dict.UpgradeOpen();
            ObjectId idXrec = dict.SetAt(searchKey, xrec);
            id.Database.TransactionManager.AddNewlyCreatedDBObject(xrec, true);
            dict.DowngradeOpen();
            return idXrec;
        }

        //  获取扩展记录
        public static ResultBuffer GetXrecord(this ObjectId id, string searchKey)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return null;
            DBDictionary dict = (DBDictionary)dictId.GetObject(OpenMode.ForRead);
            if (!dict.Contains(searchKey)) return null;
            ObjectId xrecordId = dict.GetAt(searchKey);
            Xrecord xrecord = (Xrecord)xrecordId.GetObject(OpenMode.ForRead);
            return xrecord.Data;
        }

        // 重载获取扩展记录（在没有提供searchKey的情况下默认以Handle值作为key来查找）
        public static ResultBuffer GetXrecord(this ObjectId id)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            string searchKey = id.Handle.Value.ToString();
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return null;
            DBDictionary dict = (DBDictionary)dictId.GetObject(OpenMode.ForRead);
            // if (!dict.Contains(searchKey)) return null; 由于图元必有句柄因此这句判断无效
            ObjectId xrecordId = dict.GetAt(searchKey);
            Xrecord xrecord = (Xrecord)xrecordId.GetObject(OpenMode.ForRead);
            return xrecord.Data;
        }

        //  替换扩展记录
        public static TypedValueList ModXrecord(this ObjectId id, string searchKey, TypedValueList values)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return null;
            DBDictionary dict = dictId.GetObject(OpenMode.ForRead) as DBDictionary;
            if (!dict.Contains(searchKey)) return null;
            ObjectId xrecordId = dict.GetAt(searchKey);
            Xrecord xrecord = xrecordId.GetObject(OpenMode.ForWrite) as Xrecord;
            xrecord.Data = values;
            xrecord.DowngradeOpen();
            return xrecord.Data;
        }

        /// <summary>
        /// 删除对象扩展字典中指定的一条扩展记录
        /// </summary>
        /// <param name="id"></param>
        /// <param name="searchKey"></param>
        public static void DelXrecord(this ObjectId id, string searchKey)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return;
            DBDictionary dict = dictId.GetObject(OpenMode.ForRead) as DBDictionary;
            if (dict.Contains(searchKey))
            {
                dict.UpgradeOpen();
                dict.Remove(searchKey);
                dict.DowngradeOpen();
            }
        }

        /// <summary>
        /// 删除对象扩展字典中所有的扩展记录
        /// </summary>
        /// <param name="id"></param>
        public static void DelAllXrecord(this ObjectId id)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return;
            DBDictionary dict = dictId.GetObject(OpenMode.ForRead) as DBDictionary;
            List<string> searchKeys = new List<string>();
            foreach (var item in dict)
                searchKeys.Add(item.Key);
            foreach (var key in searchKeys)
                dict.Remove(key);
            dict.DowngradeOpen();
        }

        /// <summary>
        /// 删除对象的有名扩展字典
        /// </summary>
        /// <param name="id"></param>
        public static void DelNamedDictionary(this ObjectId id)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return;
            obj.UpgradeOpen();
            id.DelAllXrecord();
            obj.ReleaseExtensionDictionary();
            obj.DowngradeOpen();
        }

        /// <summary>
        /// 添加对象的扩展数据
        /// </summary>
        /// <param name="id"></param>
        /// <param name="regAppName"></param>
        /// <param name="values"></param>
        public static void AddXData(this ObjectId id, string regAppName, TypedValueList values)
        {
            Database db = id.Database;
            RegAppTable regTable = db.RegAppTableId.GetObject(OpenMode.ForWrite) as RegAppTable;
            // 先检查注册应用程序表里面是否有注册有该应用程序，若没有，则添加：
            if (!regTable.Has(regAppName))
            {
                RegAppTableRecord regRecord = new RegAppTableRecord
                {
                    Name = regAppName
                };
                regTable.Add(regRecord);
                db.TransactionManager.AddNewlyCreatedDBObject(regRecord, true);
            }

            DBObject obj = id.GetObject(OpenMode.ForWrite);
            // 任何扩展数据在写入之前都必须添加它的注册应用程序名（必须在第一条添加，否则会报错）：
            values.Insert(0, new TypedValue((int)DxfCode.ExtendedDataRegAppName, regAppName));
            // 添加完之后把TypedValue附加到实体的扩展数据上
            obj.XData = values;
            regTable.DowngradeOpen();
        }

        /// <summary>
        /// 获取对象的扩展数据
        /// </summary>
        /// <param name="id"></param>
        /// <param name="regAppName"></param>
        /// <returns></returns>
        public static ResultBuffer GetXData(this ObjectId id, string regAppName)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            ResultBuffer values = obj.GetXDataForApplication(regAppName);
            return values;
        }

        /// <summary>
        /// 删除对象的扩展数据
        /// 设置对象的XData属性为null是不能删除扩展数据的，只能对XData属性重新赋值为只包含
        /// 应用程序名的扩展数据项，才可以达到删除扩展数据的目的
        /// </summary>
        /// <param name="id"></param>
        /// <param name="regAppName"></param>
        public static void RemoveXData(this ObjectId id, string regAppName)
        {
            DBObject obj = id.GetObject(OpenMode.ForWrite);
            // 获取注册应用程序regAppName名下的扩展数据列表
            TypedValueList xdata = obj.GetXDataForApplication(regAppName);
            if (!(xdata is null)) // 如果有扩展数据
            {
                // 新建一个TypedValue列表，并只添加注册应用程序名扩展数据项
                TypedValueList newValues = new TypedValueList
                {
                    { DxfCode.ExtendedDataRegAppName, regAppName }
                };
                // 为对象的XData属性重新赋值，从而删除扩展数据
                obj.XData = newValues;
            }
            obj.DowngradeOpen();
        }

        /// <summary>
        /// 修改对象的扩展数据
        /// </summary>
        /// <param name="id"></param>
        /// <param name="regAppName"></param>
        /// <param name="code"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        public static void ModXData(this ObjectId id, string regAppName, DxfCode code,
                                    object oldValue, object newValue)
        {
            DBObject obj = id.GetObject(OpenMode.ForWrite);
            TypedValueList xdata = obj.GetXDataForApplication(regAppName);
            for (int i = 0; i < xdata.Count; i++)
            {
                TypedValue tv = xdata[i];
                if (tv.TypeCode == (short)code && tv.Value.Equals(oldValue))
                {
                    xdata[i] = new TypedValue(tv.TypeCode, newValue);
                    break;
                }
            }
            obj.XData = xdata;
            obj.DowngradeOpen();
        }
    }
}
