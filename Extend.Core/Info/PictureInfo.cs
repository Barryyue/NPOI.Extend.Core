﻿/*
 类：PictureInfo
 描述：图片信息
 编 码 人：韩兆新 日期：
 修改记录：

*/

using System;

namespace NPOI.Extend.Core.Info
{
    public class PictureInfo
    {
        public PictureInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData, PictureStyle pictureStyle)
        {
            MinRow = minRow;
            MaxRow = maxRow;
            MinCol = minCol;
            MaxCol = maxCol;
            PictureData = pictureData;
            PicturesStyle = pictureStyle;
        }

        public int MinRow { get; set; }
        public int MaxRow { get; set; }
        public int MinCol { get; set; }
        public int MaxCol { get; set; }
        public Byte[] PictureData { get; set; }
        public PictureStyle PicturesStyle { get; set; }
    }
}