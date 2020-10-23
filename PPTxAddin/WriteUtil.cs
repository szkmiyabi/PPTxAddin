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
            //sa.TextRange.InsertAfter(src);
            sa.TextRange.Text = src;

        }

        //改行挿入
        private void insert_br()
        {
            var sa = getSelection();
            sa.TextRange.Text = "\r\n";
        }

        //選択範囲の文字を複製
        private void duplicate_selection()
        {
            var sa = getSelection();
            var old_txt = sa.TextRange.Text;
            sa.TextRange.Text = old_txt + old_txt;
        }

        //選択範囲の文字を削除
        private void delete_selection()
        {
            var sa = getSelection();
            sa.TextRange.Text = "";
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

        //記号で前後を挟む
        private string mark_ham(string mark, string body)
        {
            Regex pt = new Regex(@"(\[|【|「|『|""|\')( )(\]|】|」|』|""|\')", RegexOptions.Compiled);
            if (!pt.IsMatch(mark))
                return body;
            Match mt = pt.Match(mark);
            return mt.Groups[1].Value + body + mt.Groups[3].Value;
        }

        //コンボで選択した記号を挿入する
        private void write_mark()
        {
            var sa = getSelection();
            string src = writeMarkCombo.Text;

            if (writeMarkHamCheck.Checked)
            {
                string body = sa.TextRange.Text;
                body = mark_ham(src, body);
                sa.TextRange.Text = body;
            }
            else
            {
                sa.TextRange.Text = src;
            }

        }

        //角丸赤枠を挿入
        private void insert_rounded_rect()
        {
            var cs = getCurrentSlide();
            float[] size = { 120, 29 };
            var rectangle = cs.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 80, 80,  size[0], size[1]);
            rectangle.Fill.Visible = MsoTriState.msoFalse;
            rectangle.Line.ForeColor.RGB = getRGB(255, 0, 0);
            rectangle.Line.Weight = 2.5F;
            rectangle.Shadow.Visible = MsoTriState.msoTrue;
            rectangle.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            rectangle.Shadow.OffsetX = 1;
            rectangle.Shadow.OffsetY = 1;
            rectangle.Shadow.Transparency = 0.5F;
        }

        //図形矢印を挿入
        private void insert_arrow()
        {
            var cs = getCurrentSlide();
            float[] size = { 200, 75 };
            var arrow = cs.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, 90, 90, size[0], size[1]);
            arrow.Fill.ForeColor.RGB = getRGB(255, 153, 0);
            arrow.Line.Visible = MsoTriState.msoFalse;
            arrow.Shadow.Visible = MsoTriState.msoTrue;
            arrow.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
            arrow.Shadow.OffsetX = 1;
            arrow.Shadow.OffsetY = 1;
            arrow.Shadow.Transparency = 0.5F;
        }

        //文字枠を挿入
        private void insert_textbox()
        {
            var cs = getCurrentSlide();
            float[] size = { 250, 65 };
            var textbox = cs.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, size[0], size[1]);
            textbox.Line.Visible = MsoTriState.msoTrue;
            textbox.Line.ForeColor.RGB = getRGB(166, 166, 166);
            textbox.Line.Weight = 2.5F;
            textbox.Fill.Visible = MsoTriState.msoTrue;
            textbox.Fill.ForeColor.RGB = getRGB(255, 255, 255);
            textbox.TextFrame.AutoSize = PPT.PpAutoSize.ppAutoSizeNone;
            textbox.TextFrame.TextRange.Font.Name = "HGPｺﾞｼｯｸM";
            textbox.TextFrame.TextRange.Text = "説明を入力";
        }

        //吹出を挿入
        private void insert_callout()
        {
            var cs = getCurrentSlide();
            float[] size = { 120, 60 };
            var callout = cs.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangularCallout, 110, 110, size[0], size[1]);
            callout.Fill.Visible = MsoTriState.msoTrue;
            callout.Fill.ForeColor.RGB = getRGB(255, 255, 255);
            callout.Line.ForeColor.RGB = getRGB(255, 153, 0);
            callout.Line.Weight = 2.5F;
            callout.TextFrame.TextRange.Font.Name = "HGPｺﾞｼｯｸM";
            callout.TextFrame.TextRange.Text = "説明を入力";
            callout.TextFrame.TextRange.Font.Color.RGB = getRGB(0, 0, 0);
            callout.TextFrame.TextRange.Font.Size = 12;
        }

        //スライド複製
        private void duplicate_slide()
        {
            var cs = getCurrentSlide();
            cs.Duplicate();
        }

        //スライド挿入
        private void insert_slide()
        {
            var crSlides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            crSlides.AddSlide(Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex + 1, getCurrentSlide().CustomLayout);
        }

        //スライドコピー
        private void copy_slide()
        {
            var cs = getCurrentSlide();
            cs.Copy();
        }

        //スライド貼り付け
        private void paste_slide()
        {
            var crSlides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            crSlides.Paste(Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex + 1);
        }

        //赤字
        private void paint_text_red()
        {
            var sa = getSelection();
            sa.TextRange.Font.Color.RGB = getRGB(255, 0, 0);
        }

        //青字
        private void paint_text_blue()
        {
            var sa = getSelection();
            sa.TextRange.Font.Color.RGB = getRGB(0, 0, 255);
        }

        //黒字
        private void paint_text_black()
        {
            var sa = getSelection();
            sa.TextRange.Font.Color.RGB = getRGB(0, 0, 0);
        }

        //太字
        private void bold_text()
        {
            var sa = getSelection();
            sa.TextRange.Font.Bold = MsoTriState.msoTrue;
        }

        //細字
        private void narrow_text()
        {
            var sa = getSelection();
            sa.TextRange.Font.Bold = MsoTriState.msoFalse;
        }
    }
}
