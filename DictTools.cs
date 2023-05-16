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

        public static ResultBuffer GetXrecord(this ObjectId id)
        {
            DBObject obj = id.GetObject(OpenMode.ForRead);
            string searchKey = id.ToString();
            ObjectId dictId = obj.ExtensionDictionary;
            if (dictId.IsNull) return null;
            DBDictionary dict = (DBDictionary)dictId.GetObject(OpenMode.ForRead);
            if (!dict.Contains(searchKey)) return null;
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
        public static  void DelAllXrecord(this ObjectId id)
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
    }
}
