using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoPipelines
{
    public class TypedValueList:List<TypedValue>
    {
        public TypedValueList(params TypedValue[] args)
        {
            AddRange(args);
        }
        
        public void Add(int typecode, object value)
        {
            base.Add(new TypedValue(typecode, value));
        }

        public void Add(DxfCode dxfcode, object value)
        {
            base.Add(new TypedValue((int)dxfcode, value));
        }

        public void Add(Type entityType)
        {
            base.Add(new TypedValue(0, RXClass.GetClass(entityType).DxfName));
        }

        public static implicit operator TypedValue[](TypedValueList src)
        {
            return src != null ? src.ToArray() : null;
        }

        public static implicit operator ResultBuffer(TypedValueList values)
        {
            if (values.Count > 0)
            {
                ResultBuffer buffer = new ResultBuffer();
                foreach (var value in values)
                    buffer.Add(value);
                return buffer;
            }
            else
                return null;
        }

        public static implicit operator TypedValueList(ResultBuffer buffer)
        {
            if (buffer != null)
            {
                return new TypedValueList(buffer.AsArray());
            }
            else
                return null;
        }

}
}
