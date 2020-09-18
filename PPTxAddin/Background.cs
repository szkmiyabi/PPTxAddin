using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace PPTxAddin
{
    partial class Ribbon1
    {
        //RGBコードを取得
        private int getRGB(int r, int g, int b)
        {
            Color c = Color.FromArgb(r, g, b);
            var cint = (c.R + 0x100 * c.G + 0x10000 * c.B);
            return (int)cint;
        }
    }
}
