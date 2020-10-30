using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace PPTxAddin
{
    public partial class Ribbon1
    {

        //Formオブジェクト
        private static InputForm _inpfrmObj;
        private static ComboEditForm _cmbefrmObj;

        //コンストラクタ
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //InputFormインスタンスの取得
        public static InputForm inpfrmObj
        {
            get
            {
                if (_inpfrmObj == null || _inpfrmObj.IsDisposed)
                {
                    _inpfrmObj = new InputForm();
                }
                return _inpfrmObj;
            }
        }

        //ComboEditFormインスタンスの取得
        public static ComboEditForm cmbefrmObj
        {
            get
            {
                if (_cmbefrmObj == null || _cmbefrmObj.IsDisposed)
                {
                    _cmbefrmObj = new ComboEditForm();
                }
                return _cmbefrmObj;
            }
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
        //コンボで選択した文字列を挿入する
        private void writeCommentInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            write_comment();
        }

        //ドロップダウンに値を追加する
        private void writeCommentAddButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_add_comment();
        }
        //テキストファイルからドロップダウンに値を追加する
        private void writeCommentAddFromFileButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_add_comment_from_file();
        }
        //ドロップダウン選択項目削除
        private void delCommentSingleButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_clear_combo_comment_single();
        }
        //ドロップダウン項目全削除
        private void delCommentAllButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_clear_combo_comment_all();
        }
        //ドロップダウンの値を保存
        private void writeCommentComboSaveButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_save_val_comment();
        }
        //フォームから値追加
        private void writeCommentAddFromFormButton_Click(object sender, RibbonControlEventArgs e)
        {
            inpfrmObj.Show();
        }
        //値編集
        private void doEditComboButton_Click(object sender, RibbonControlEventArgs e)
        {
            cmbefrmObj.Show();
        }

        //赤枠
        private void insertRectangleButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_rounded_rect();
        }

        //図矢印
        private void insertArrowButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_arrow();
        }

        //文字枠
        private void insertTextboxButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_textbox();
        }

        //吹出
        private void insertCalloutButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_callout();
        }

        //記号挿入
        private void writeMarkInputButton_Click(object sender, RibbonControlEventArgs e)
        {
            write_mark();
        }

        //改行
        private void insertBrButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_br();
        }

        //文字複製
        private void duplicateTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            duplicate_selection();
        }

        //スライド複製
        private void duplicateSlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            duplicate_slide();
        }

        //赤字
        private void fontRedButton_Click(object sender, RibbonControlEventArgs e)
        {
            paint_text_red();
        }

        //青字
        private void fontBlueButton_Click(object sender, RibbonControlEventArgs e)
        {
            paint_text_blue();
        }

        //黒字
        private void fontBlackButton_Click(object sender, RibbonControlEventArgs e)
        {
            paint_text_black();
        }

        //削除
        private void textDeleteButton_Click(object sender, RibbonControlEventArgs e)
        {
            delete_selection();
        }

        //新規挿入
        private void insertSlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_slide();
        }

        //太字
        private void fontBoldButton_Click(object sender, RibbonControlEventArgs e)
        {
            bold_text();
        }

        //細字
        private void fontNarrowButton_Click(object sender, RibbonControlEventArgs e)
        {
            narrow_text();
        }

        //コピー
        private void slideCopyButton_Click(object sender, RibbonControlEventArgs e)
        {
            copy_slide();
        }

        //貼付
        private void slidePasteButton_Click(object sender, RibbonControlEventArgs e)
        {
            paste_slide();
        }

        //書式無
        private void resetShapeStyleButton_Click(object sender, RibbonControlEventArgs e)
        {
            reset_shape_style();
        }

        private void bringFrontButton_Click(object sender, RibbonControlEventArgs e)
        {
            bring_front();
        }

        private void textCutButton_Click(object sender, RibbonControlEventArgs e)
        {
            text_cut();
        }

        private void textPasteButton_Click(object sender, RibbonControlEventArgs e)
        {
            text_paste();
        }

        private void textCopyButton_Click(object sender, RibbonControlEventArgs e)
        {
            text_copy();
        }

        private void insertLineArrowButton_Click(object sender, RibbonControlEventArgs e)
        {
            insert_line_arrow();
        }
    }
}
