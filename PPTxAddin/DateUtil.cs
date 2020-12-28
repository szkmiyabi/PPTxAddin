using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PPTxAddin
{
    class DateUtil
    {

        //ファイル用時刻文字列を生成
        public static string fetch_filename_logtime()
        {
            DateTime ld = DateTime.Now;
            return ld.ToString("yyyy-MM-dd_HH-mm-ss");
        }

        //ログ出力の時刻文字列を生成
        public static string get_logtime()
        {
            DateTime ld = DateTime.Now;
            return ld.ToString("yyyy-MM-dd HH:mm:ss");
        }

    }
}
