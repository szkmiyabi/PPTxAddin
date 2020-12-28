using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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

        //TXTファイル保存先を取得
        private string _get_txt_save_path()
        {
            string path = "";
            SaveFileDialog fda = new SaveFileDialog();
            fda.Filter = "Textファイル(*.txt)|*.txt";
            fda.Title = "名前を付けて保存";
            if (fda.ShowDialog() == DialogResult.OK)
            {
                path = fda.FileName;
            }
            return path;
        }

        //デフォルトの作業ディレクトリを取得
        private string getDefaultWorkDir()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\";
        }

        //ユーザのホームフォルダパス
        private string getUserHomePath()
        {
            return System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

    }
}
