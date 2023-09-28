﻿using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoPipelines
{
    public static class PipeXRecord
    {
        public static TypedValueList ToTypedValueList(this PipeLineProperty plp, string method = "all")
        {
            TypedValueList typedValues;
            switch (method)
            {
                case "all":
                    typedValues = new TypedValueList
                {
                    {DxfCode.Int32, plp.RowInd}, // 0
                    {DxfCode.Text, plp.Name},
                    {DxfCode.Text, plp.WTName},
                    {DxfCode.Int32, (int)plp.PipeLineType},
                    {DxfCode.Text, plp.Connect},
                    {DxfCode.Text, plp.Attribute}, // 5
                    {DxfCode.Text, plp.Attachment},
                    {DxfCode.Real, plp.X},
                    {DxfCode.Real, plp.Y},
                    {DxfCode.Real, plp.H},
                    {DxfCode.Real, plp.SPH}, // 10
                    {DxfCode.Real, plp.EPH},
                    {DxfCode.Real, plp.WellDepth},
                    {DxfCode.Real, plp.SPDepth},
                    {DxfCode.Real, plp.EPDepth},
                    {DxfCode.Text, plp.Size}, // 15
                    {DxfCode.Text, plp.Material},
                    {DxfCode.Text, plp.Pressure},
                    {DxfCode.Text, plp.Voltage},
                    {DxfCode.Int32, plp.TotalBHNum},
                    {DxfCode.Int32, plp.UsedBHNum}, // 20
                    {DxfCode.Int32, plp.CableNum},
                    {DxfCode.Text, plp.Company},
                    {DxfCode.Text, plp.BuryMethod},
                    {DxfCode.Text, plp.BuryDate},
                    {DxfCode.Text, plp.RoadName}, // 25
                    {DxfCode.Text, plp.Comment},
                    {DxfCode.Text, plp.Tag},
                };
                    return typedValues;
                case "line":
                    typedValues = new TypedValueList
                    {
                    {DxfCode.Int32, (int)plp.PipeLineType},
                    {DxfCode.Real, plp.SPDepth},
                    {DxfCode.Real, plp.EPDepth},
                    {DxfCode.Text, plp.Size},
                    {DxfCode.Text, plp.Material},
                    {DxfCode.Text, plp.Pressure},
                    {DxfCode.Text, plp.Voltage},
                    {DxfCode.Text, plp.BuryMethod},
                    {DxfCode.Text, plp.Company},
                    };
                    return typedValues;
                default:
                    return null;
            }
        }

        public static PipeLineProperty ToPipeLineProperty(this TypedValueList tv)
        {
            PipeLineProperty pipeLine = new PipeLineProperty()
            {
                RowInd = (ushort)(int)tv[0].Value,
                Name = tv[1].Value.ToString(),
                WTName = tv[2].Value.ToString(),
                PipeLineType = (PipeLineType)tv[3].Value,
                Connect = tv[4].Value.ToString(),
                Attribute = tv[5].Value.ToString(),
                Attachment = tv[6].Value.ToString(),
                X = Convert.ToDouble(tv[7].Value),
                Y = Convert.ToDouble(tv[8].Value),
                H = Convert.ToDouble(tv[9].Value),
                SPH = Convert.ToDouble(tv[10].Value),
                EPH = Convert.ToDouble(tv[11].Value),
                WellDepth = Convert.ToDouble(tv[12].Value),
                SPDepth = Convert.ToDouble(tv[13].Value),
                EPDepth = Convert.ToDouble(tv[14].Value),
                Size = tv[15].Value.ToString(),
                Material = tv[16].Value.ToString(),
                Pressure = tv[17].Value.ToString(),
                Voltage = tv[18].Value.ToString(),
                TotalBHNum = (ushort)(int)tv[19].Value,
                UsedBHNum = (ushort)(int)tv[20].Value,
                CableNum = (ushort)(int)tv[21].Value,
                Company = tv[22].Value.ToString(),
                BuryMethod = tv[23].Value.ToString(),
                BuryDate = tv[24].Value.ToString(),
                RoadName = tv[25].Value.ToString(),
                Comment = tv[26].Value.ToString(),
                Tag = tv[27].Value.ToString(),
            };
            return pipeLine;
        }

    }

}
