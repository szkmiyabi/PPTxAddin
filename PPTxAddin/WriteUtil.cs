using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using PPT = Microsoft.Office.Interop.PowerPoint;


namespace PPTxAddin
{
    partial class Ribbon1
    {

        private string br_sp = "<bkmk:br>";

        //カレントスライドページを取得
        private PPT.Slide getCurrentSlide()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.Slides[
                Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex
            ];
        }

        //指定番号のスライドページを取得
        private PPT.Slide getSlide(int idx)
        {
            return Globals.ThisAddIn.Application.ActivePresentation.Slides[idx];
        }

        //プレイスホルダーの数を取得
        private int getPlaceholdersCount(PPT.Slide cs)
        {
            return cs.Shapes.Placeholders.Count;
        }

        //Shapesの数を取得
        private int getShapesCount(PPT.Slide cs)
        {
            return cs.Shapes.Count;
        }

        //スライドタイトル判定
        Boolean checkHasTitle(PPT.Slide cs)
        {
            var flg = cs.Shapes.HasTitle;
            if (flg == MsoTriState.msoTrue) return true;
            else return false;
        }

        //選択範囲を取得
        PPT.Selection getSelection()
        {
            return Globals.ThisAddIn.Application.ActiveWindow.Selection;
        }

        //コンボで選択した文字列を挿入する
        private void write_comment()
        {
            string src = writeCommentCombo.Text;
            src = src.Replace(br_sp, "\r\n");
            var sa = getSelection();
            var cr_page = getCurrentSlide();
            string search = sa.TextRange.Text;
            cr_page.Shapes.Placeholders[2].TextFrame.TextRange.Replace(search, search + src);
        }

        //ドロップダウンに値を追加する
        private void do_add_comment()
        {
            var sa = getSelection();
            string buff = sa.TextRange.Text;
            buff = buff.Replace("\r\n", br_sp);
            buff = buff.Replace("\r", br_sp);

            if(buff != "")
            {
                //check onなら全クリア
                if (addCommentPreClearCheck.Checked == true) do_clear_combo_comment_all();
                RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                item.Label = buff;
                writeCommentCombo.Items.Add(item);
                MessageBox.Show("値の追加に成功しました");
            }
        }

        //テキストファイルからドロップダウンに値を追加する
        private void do_add_comment_from_file()
        {
            string filename = "";
            string body = "";
            List<string> arr = new List<string>();

            //check onなら全クリア
            if (addCommentPreClearCheck.Checked == true) do_clear_combo_comment_all();

            OpenFileDialog f = new OpenFileDialog();
            f.Filter = "テキストファイル(*.txt)|*.txt";
            if (f.ShowDialog() == DialogResult.OK)
            {
                filename = f.FileName;
            }
            if (filename == "") return;
            StreamReader sr = new StreamReader(filename, System.Text.Encoding.GetEncoding("shift_jis"));
            while (sr.Peek() > -1)
            {
                string line = sr.ReadLine();
                arr.Add(line);
            }
            sr.Close();

            for (int i = 0; i < arr.Count; i++)
            {
                RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                item.Label = arr[i].ToString();
                writeCommentCombo.Items.Add(item);
            }
            MessageBox.Show("値の追加に成功しました");

        }

        //ドロップダウン選択項目削除
        private void do_clear_combo_comment_single()
        {
            int idx = 0;
            string cr = writeCommentCombo.Text;

            for (int i = 0; i < writeCommentCombo.Items.Count; i++)
            {
                RibbonDropDownItem opt = writeCommentCombo.Items[i];
                if (opt.Label.Equals(cr))
                {
                    writeCommentCombo.Items.RemoveAt(idx);
                    writeCommentCombo.Text = "";
                    break;
                }
                idx++;
            }
        }

        //ドロップダウン項目全削除
        private void do_clear_combo_comment_all()
        {
            writeCommentCombo.Items.Clear();
            writeCommentCombo.Text = "";
        }

        //ドロップダウンの値を保存
        private void do_save_val_comment()
        {
            int cnt = writeCommentCombo.Items.Count;
            string body = "";
            for (int i = 0; i < cnt; i++)
            {
                string val = writeCommentCombo.Items[i].Label;
                body += val;
                if (i != (cnt - 1)) body += "\r\n";
            }
            string path = _get_txt_save_path();
            Encoding enc = Encoding.GetEncoding("Shift_JIS");
            StreamWriter sw = new StreamWriter(path, false, enc);
            sw.WriteLine(body);
            sw.Close();
            MessageBox.Show("保存できました!");
        }

        //角丸赤枠を挿入
        private void insert_rounded_rect()
        {
            var cs = getCurrentSlide();
            float[] size = { 120, 90 };
            var rectangle = cs.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 100, 100,  size[0], size[1]);
            rectangle.Fill.Visible = MsoTriState.msoFalse;
            rectangle.Line.ForeColor.RGB = getRGB(255, 0, 0);
            rectangle.Line.Weight = 2.5F;
        }

        //図形矢印を挿入
        private void insert_arrow()
        {
            var cs = getCurrentSlide();
            float[] size = { 120, 90 };
            var arrow = cs.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, 110, 110, size[0], size[1]);
            arrow.Fill.ForeColor.RGB = getRGB(255, 153, 0);
            arrow.Line.Visible = MsoTriState.msoFalse;
        }

        //文字枠を挿入
        private void insert_textbox()
        {
            var cs = getCurrentSlide();
            float[] size = { 180, 90 };
            var textbox = cs.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 90, 90, size[0], size[1]);
            textbox.Line.Visible = MsoTriState.msoTrue;
            textbox.Line.ForeColor.RGB = getRGB(166, 166, 166);
            textbox.Line.Weight = 2.5F;
            textbox.Fill.Visible = MsoTriState.msoTrue;
            textbox.Fill.ForeColor.RGB = getRGB(255, 255, 255);
            textbox.TextFrame.AutoSize = PPT.PpAutoSize.ppAutoSizeNone;
            textbox.TextFrame.TextRange.Font.Name = "HGPｺﾞｼｯｸM";
            textbox.TextFrame.TextRange.Text = "Textbox";
        }

    }
}
