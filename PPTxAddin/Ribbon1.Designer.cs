namespace PPTxAddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl15 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl16 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl17 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl18 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl19 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.writeCommentCombo = this.Factory.CreateRibbonComboBox();
            this.writeCommentInputButton = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.writeCommentAddButton = this.Factory.CreateRibbonButton();
            this.writeCommentAddFromFormButton = this.Factory.CreateRibbonButton();
            this.writeCommentAddFromFileButton = this.Factory.CreateRibbonButton();
            this.addCommentPreClearCheck = this.Factory.CreateRibbonCheckBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.delCommentSingleButton = this.Factory.CreateRibbonButton();
            this.delCommentAllButton = this.Factory.CreateRibbonButton();
            this.doEditComboButton = this.Factory.CreateRibbonButton();
            this.writeCommentComboSaveButton = this.Factory.CreateRibbonButton();
            this.textCopyButton = this.Factory.CreateRibbonButton();
            this.textPasteButton = this.Factory.CreateRibbonButton();
            this.box5 = this.Factory.CreateRibbonBox();
            this.writeMarkCombo = this.Factory.CreateRibbonComboBox();
            this.writeMarkInputButton = this.Factory.CreateRibbonButton();
            this.writeMarkHamCheck = this.Factory.CreateRibbonCheckBox();
            this.box6 = this.Factory.CreateRibbonBox();
            this.insertBrButton = this.Factory.CreateRibbonButton();
            this.duplicateTextButton = this.Factory.CreateRibbonButton();
            this.textCutButton = this.Factory.CreateRibbonButton();
            this.textDeleteButton = this.Factory.CreateRibbonButton();
            this.box4 = this.Factory.CreateRibbonBox();
            this.fontRedButton = this.Factory.CreateRibbonButton();
            this.fontBlueButton = this.Factory.CreateRibbonButton();
            this.fontBlackButton = this.Factory.CreateRibbonButton();
            this.fontBoldButton = this.Factory.CreateRibbonButton();
            this.fontNarrowButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.box8 = this.Factory.CreateRibbonBox();
            this.insertRectangleButton = this.Factory.CreateRibbonButton();
            this.insertLineArrowButton = this.Factory.CreateRibbonButton();
            this.insertArrowButton = this.Factory.CreateRibbonButton();
            this.insertCalloutButton = this.Factory.CreateRibbonButton();
            this.insertTextboxButton = this.Factory.CreateRibbonButton();
            this.box9 = this.Factory.CreateRibbonBox();
            this.resetShapeStyleButton = this.Factory.CreateRibbonButton();
            this.bringFrontButton = this.Factory.CreateRibbonButton();
            this.selectObjectButton = this.Factory.CreateRibbonButton();
            this.flipHorizontalButton = this.Factory.CreateRibbonButton();
            this.flipVerticalButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.box7 = this.Factory.CreateRibbonBox();
            this.duplicateSlideButton = this.Factory.CreateRibbonButton();
            this.insertSlideButton = this.Factory.CreateRibbonButton();
            this.box10 = this.Factory.CreateRibbonBox();
            this.slideCopyButton = this.Factory.CreateRibbonButton();
            this.slidePasteButton = this.Factory.CreateRibbonButton();
            this.box11 = this.Factory.CreateRibbonBox();
            this.exportCurrentSlideAsPNGButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.box5.SuspendLayout();
            this.box6.SuspendLayout();
            this.box4.SuspendLayout();
            this.group3.SuspendLayout();
            this.box8.SuspendLayout();
            this.box9.SuspendLayout();
            this.group2.SuspendLayout();
            this.box7.SuspendLayout();
            this.box10.SuspendLayout();
            this.box11.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "PPTxA";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.box3);
            this.group1.Items.Add(this.box5);
            this.group1.Items.Add(this.box6);
            this.group1.Items.Add(this.box4);
            this.group1.Label = "文章編集";
            this.group1.Name = "group1";
            // 
            // box1
            // 
            this.box1.Items.Add(this.writeCommentCombo);
            this.box1.Items.Add(this.writeCommentInputButton);
            this.box1.Name = "box1";
            // 
            // writeCommentCombo
            // 
            this.writeCommentCombo.Label = "comboBox1";
            this.writeCommentCombo.Name = "writeCommentCombo";
            this.writeCommentCombo.ShowLabel = false;
            this.writeCommentCombo.SizeString = "AAAAAAAAAAAAAAAAA";
            this.writeCommentCombo.Text = null;
            // 
            // writeCommentInputButton
            // 
            this.writeCommentInputButton.Label = "挿入";
            this.writeCommentInputButton.Name = "writeCommentInputButton";
            this.writeCommentInputButton.OfficeImageId = "BrowseNext";
            this.writeCommentInputButton.ScreenTip = "語句挿入";
            this.writeCommentInputButton.ShowImage = true;
            this.writeCommentInputButton.ShowLabel = false;
            this.writeCommentInputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentInputButton_Click);
            // 
            // box2
            // 
            this.box2.Items.Add(this.writeCommentAddButton);
            this.box2.Items.Add(this.writeCommentAddFromFormButton);
            this.box2.Items.Add(this.writeCommentAddFromFileButton);
            this.box2.Items.Add(this.addCommentPreClearCheck);
            this.box2.Name = "box2";
            // 
            // writeCommentAddButton
            // 
            this.writeCommentAddButton.Label = "選択範囲";
            this.writeCommentAddButton.Name = "writeCommentAddButton";
            this.writeCommentAddButton.OfficeImageId = "SectionRename";
            this.writeCommentAddButton.ScreenTip = "選択範囲";
            this.writeCommentAddButton.ShowImage = true;
            this.writeCommentAddButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentAddButton_Click);
            // 
            // writeCommentAddFromFormButton
            // 
            this.writeCommentAddFromFormButton.Label = "フォーム";
            this.writeCommentAddFromFormButton.Name = "writeCommentAddFromFormButton";
            this.writeCommentAddFromFormButton.OfficeImageId = "FormControlInsertMenu";
            this.writeCommentAddFromFormButton.ScreenTip = "フォーム";
            this.writeCommentAddFromFormButton.ShowImage = true;
            this.writeCommentAddFromFormButton.ShowLabel = false;
            this.writeCommentAddFromFormButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentAddFromFormButton_Click);
            // 
            // writeCommentAddFromFileButton
            // 
            this.writeCommentAddFromFileButton.Label = "読込";
            this.writeCommentAddFromFileButton.Name = "writeCommentAddFromFileButton";
            this.writeCommentAddFromFileButton.OfficeImageId = "CreateDocumentLibrary";
            this.writeCommentAddFromFileButton.ScreenTip = "ファイル読込";
            this.writeCommentAddFromFileButton.ShowImage = true;
            this.writeCommentAddFromFileButton.ShowLabel = false;
            this.writeCommentAddFromFileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentAddFromFileButton_Click);
            // 
            // addCommentPreClearCheck
            // 
            this.addCommentPreClearCheck.Label = "全削除追加";
            this.addCommentPreClearCheck.Name = "addCommentPreClearCheck";
            // 
            // box3
            // 
            this.box3.Items.Add(this.delCommentSingleButton);
            this.box3.Items.Add(this.delCommentAllButton);
            this.box3.Items.Add(this.doEditComboButton);
            this.box3.Items.Add(this.writeCommentComboSaveButton);
            this.box3.Items.Add(this.textCopyButton);
            this.box3.Items.Add(this.textPasteButton);
            this.box3.Name = "box3";
            // 
            // delCommentSingleButton
            // 
            this.delCommentSingleButton.Label = "1件削除";
            this.delCommentSingleButton.Name = "delCommentSingleButton";
            this.delCommentSingleButton.OfficeImageId = "SectionMergeWithPrevious";
            this.delCommentSingleButton.ScreenTip = "1件削除";
            this.delCommentSingleButton.ShowImage = true;
            this.delCommentSingleButton.ShowLabel = false;
            this.delCommentSingleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delCommentSingleButton_Click);
            // 
            // delCommentAllButton
            // 
            this.delCommentAllButton.Label = "全件削除";
            this.delCommentAllButton.Name = "delCommentAllButton";
            this.delCommentAllButton.OfficeImageId = "SectionRemoveAll";
            this.delCommentAllButton.ScreenTip = "全件削除";
            this.delCommentAllButton.ShowImage = true;
            this.delCommentAllButton.ShowLabel = false;
            this.delCommentAllButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delCommentAllButton_Click);
            // 
            // doEditComboButton
            // 
            this.doEditComboButton.Label = "値編集";
            this.doEditComboButton.Name = "doEditComboButton";
            this.doEditComboButton.OfficeImageId = "SearchTools";
            this.doEditComboButton.ScreenTip = "値編集";
            this.doEditComboButton.ShowImage = true;
            this.doEditComboButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.doEditComboButton_Click);
            // 
            // writeCommentComboSaveButton
            // 
            this.writeCommentComboSaveButton.Label = "値保存";
            this.writeCommentComboSaveButton.Name = "writeCommentComboSaveButton";
            this.writeCommentComboSaveButton.OfficeImageId = "SaveHollow";
            this.writeCommentComboSaveButton.ScreenTip = "ファイル保存";
            this.writeCommentComboSaveButton.ShowImage = true;
            this.writeCommentComboSaveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentComboSaveButton_Click);
            // 
            // textCopyButton
            // 
            this.textCopyButton.Label = "文字コピー";
            this.textCopyButton.Name = "textCopyButton";
            this.textCopyButton.OfficeImageId = "ContactCardCopy";
            this.textCopyButton.ScreenTip = "文字コピー";
            this.textCopyButton.ShowImage = true;
            this.textCopyButton.ShowLabel = false;
            this.textCopyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.textCopyButton_Click);
            // 
            // textPasteButton
            // 
            this.textPasteButton.Label = "文字貼付";
            this.textPasteButton.Name = "textPasteButton";
            this.textPasteButton.OfficeImageId = "ContactCardPaste";
            this.textPasteButton.ScreenTip = "文字貼付";
            this.textPasteButton.ShowImage = true;
            this.textPasteButton.ShowLabel = false;
            this.textPasteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.textPasteButton_Click);
            // 
            // box5
            // 
            this.box5.Items.Add(this.writeMarkCombo);
            this.box5.Items.Add(this.writeMarkInputButton);
            this.box5.Items.Add(this.writeMarkHamCheck);
            this.box5.Name = "box5";
            // 
            // writeMarkCombo
            // 
            ribbonDropDownItemImpl1.Label = "※";
            ribbonDropDownItemImpl2.Label = "●";
            ribbonDropDownItemImpl3.Label = "○";
            ribbonDropDownItemImpl4.Label = "×";
            ribbonDropDownItemImpl5.Label = "■";
            ribbonDropDownItemImpl6.Label = "←";
            ribbonDropDownItemImpl7.Label = "→";
            ribbonDropDownItemImpl8.Label = "↑";
            ribbonDropDownItemImpl9.Label = "↓";
            ribbonDropDownItemImpl10.Label = "\" \"";
            ribbonDropDownItemImpl11.Label = "\' \'";
            ribbonDropDownItemImpl12.Label = "( )";
            ribbonDropDownItemImpl13.Label = "（ ）";
            ribbonDropDownItemImpl14.Label = "【 】";
            ribbonDropDownItemImpl15.Label = "[ ]";
            ribbonDropDownItemImpl16.Label = "「 」";
            ribbonDropDownItemImpl17.Label = "『 』";
            ribbonDropDownItemImpl18.Label = "< >";
            ribbonDropDownItemImpl19.Label = "＜ ＞";
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl1);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl2);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl3);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl4);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl5);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl6);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl7);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl8);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl9);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl10);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl11);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl12);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl13);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl14);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl15);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl16);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl17);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl18);
            this.writeMarkCombo.Items.Add(ribbonDropDownItemImpl19);
            this.writeMarkCombo.Label = "comboBox1";
            this.writeMarkCombo.Name = "writeMarkCombo";
            this.writeMarkCombo.ShowLabel = false;
            this.writeMarkCombo.SizeString = "AAAA";
            this.writeMarkCombo.Text = null;
            // 
            // writeMarkInputButton
            // 
            this.writeMarkInputButton.Label = "記号挿入";
            this.writeMarkInputButton.Name = "writeMarkInputButton";
            this.writeMarkInputButton.OfficeImageId = "BrowseNext";
            this.writeMarkInputButton.ScreenTip = "挿入";
            this.writeMarkInputButton.ShowImage = true;
            this.writeMarkInputButton.ShowLabel = false;
            this.writeMarkInputButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeMarkInputButton_Click);
            // 
            // writeMarkHamCheck
            // 
            this.writeMarkHamCheck.Label = "はさみこみ";
            this.writeMarkHamCheck.Name = "writeMarkHamCheck";
            // 
            // box6
            // 
            this.box6.Items.Add(this.insertBrButton);
            this.box6.Items.Add(this.duplicateTextButton);
            this.box6.Items.Add(this.textCutButton);
            this.box6.Items.Add(this.textDeleteButton);
            this.box6.Name = "box6";
            // 
            // insertBrButton
            // 
            this.insertBrButton.Label = "改行";
            this.insertBrButton.Name = "insertBrButton";
            this.insertBrButton.OfficeImageId = "ParagraphDialog";
            this.insertBrButton.ScreenTip = "改行";
            this.insertBrButton.ShowImage = true;
            this.insertBrButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertBrButton_Click);
            // 
            // duplicateTextButton
            // 
            this.duplicateTextButton.Label = "複製";
            this.duplicateTextButton.Name = "duplicateTextButton";
            this.duplicateTextButton.OfficeImageId = "DelegateAccess";
            this.duplicateTextButton.ScreenTip = "複製";
            this.duplicateTextButton.ShowImage = true;
            this.duplicateTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateTextButton_Click);
            // 
            // textCutButton
            // 
            this.textCutButton.Label = "切取";
            this.textCutButton.Name = "textCutButton";
            this.textCutButton.OfficeImageId = "ContactCardCut";
            this.textCutButton.ScreenTip = "切取";
            this.textCutButton.ShowImage = true;
            this.textCutButton.ShowLabel = false;
            this.textCutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.textCutButton_Click);
            // 
            // textDeleteButton
            // 
            this.textDeleteButton.Label = "削除";
            this.textDeleteButton.Name = "textDeleteButton";
            this.textDeleteButton.OfficeImageId = "ContactDelete";
            this.textDeleteButton.ScreenTip = "削除";
            this.textDeleteButton.ShowImage = true;
            this.textDeleteButton.ShowLabel = false;
            this.textDeleteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.textDeleteButton_Click);
            // 
            // box4
            // 
            this.box4.Items.Add(this.fontRedButton);
            this.box4.Items.Add(this.fontBlueButton);
            this.box4.Items.Add(this.fontBlackButton);
            this.box4.Items.Add(this.fontBoldButton);
            this.box4.Items.Add(this.fontNarrowButton);
            this.box4.Name = "box4";
            // 
            // fontRedButton
            // 
            this.fontRedButton.Label = "赤";
            this.fontRedButton.Name = "fontRedButton";
            this.fontRedButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontRedButton_Click);
            // 
            // fontBlueButton
            // 
            this.fontBlueButton.Label = "青";
            this.fontBlueButton.Name = "fontBlueButton";
            this.fontBlueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontBlueButton_Click);
            // 
            // fontBlackButton
            // 
            this.fontBlackButton.Label = "黒";
            this.fontBlackButton.Name = "fontBlackButton";
            this.fontBlackButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontBlackButton_Click);
            // 
            // fontBoldButton
            // 
            this.fontBoldButton.Label = "太字";
            this.fontBoldButton.Name = "fontBoldButton";
            this.fontBoldButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontBoldButton_Click);
            // 
            // fontNarrowButton
            // 
            this.fontNarrowButton.Label = "細字";
            this.fontNarrowButton.Name = "fontNarrowButton";
            this.fontNarrowButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fontNarrowButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.box8);
            this.group3.Items.Add(this.box9);
            this.group3.Label = "図形処理";
            this.group3.Name = "group3";
            // 
            // box8
            // 
            this.box8.Items.Add(this.insertRectangleButton);
            this.box8.Items.Add(this.insertLineArrowButton);
            this.box8.Items.Add(this.insertArrowButton);
            this.box8.Items.Add(this.insertCalloutButton);
            this.box8.Items.Add(this.insertTextboxButton);
            this.box8.Name = "box8";
            // 
            // insertRectangleButton
            // 
            this.insertRectangleButton.Label = "赤枠";
            this.insertRectangleButton.Name = "insertRectangleButton";
            this.insertRectangleButton.OfficeImageId = "MsoInkStyle5";
            this.insertRectangleButton.ScreenTip = "赤枠";
            this.insertRectangleButton.ShowImage = true;
            this.insertRectangleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertRectangleButton_Click);
            // 
            // insertLineArrowButton
            // 
            this.insertLineArrowButton.Label = "矢印";
            this.insertLineArrowButton.Name = "insertLineArrowButton";
            this.insertLineArrowButton.OfficeImageId = "Arrow";
            this.insertLineArrowButton.ShowImage = true;
            this.insertLineArrowButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertLineArrowButton_Click);
            // 
            // insertArrowButton
            // 
            this.insertArrowButton.Label = "図矢印";
            this.insertArrowButton.Name = "insertArrowButton";
            this.insertArrowButton.OfficeImageId = "MarginsAdjust";
            this.insertArrowButton.ShowImage = true;
            this.insertArrowButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertArrowButton_Click);
            // 
            // insertCalloutButton
            // 
            this.insertCalloutButton.Label = "吹出";
            this.insertCalloutButton.Name = "insertCalloutButton";
            this.insertCalloutButton.OfficeImageId = "Callout";
            this.insertCalloutButton.ShowImage = true;
            this.insertCalloutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertCalloutButton_Click);
            // 
            // insertTextboxButton
            // 
            this.insertTextboxButton.Label = "文字枠";
            this.insertTextboxButton.Name = "insertTextboxButton";
            this.insertTextboxButton.OfficeImageId = "CharacterBorder";
            this.insertTextboxButton.ShowImage = true;
            this.insertTextboxButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertTextboxButton_Click);
            // 
            // box9
            // 
            this.box9.Items.Add(this.resetShapeStyleButton);
            this.box9.Items.Add(this.bringFrontButton);
            this.box9.Items.Add(this.selectObjectButton);
            this.box9.Items.Add(this.flipHorizontalButton);
            this.box9.Items.Add(this.flipVerticalButton);
            this.box9.Name = "box9";
            // 
            // resetShapeStyleButton
            // 
            this.resetShapeStyleButton.Label = "書式無";
            this.resetShapeStyleButton.Name = "resetShapeStyleButton";
            this.resetShapeStyleButton.OfficeImageId = "Clear";
            this.resetShapeStyleButton.ShowImage = true;
            this.resetShapeStyleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resetShapeStyleButton_Click);
            // 
            // bringFrontButton
            // 
            this.bringFrontButton.Label = "最前面";
            this.bringFrontButton.Name = "bringFrontButton";
            this.bringFrontButton.OfficeImageId = "CircularReferences";
            this.bringFrontButton.ShowImage = true;
            this.bringFrontButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bringFrontButton_Click);
            // 
            // selectObjectButton
            // 
            this.selectObjectButton.Label = "全選択";
            this.selectObjectButton.Name = "selectObjectButton";
            this.selectObjectButton.OfficeImageId = "ReadingModePageColorMenu";
            this.selectObjectButton.ShowImage = true;
            this.selectObjectButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectObjectButton_Click);
            // 
            // flipHorizontalButton
            // 
            this.flipHorizontalButton.Label = "横反転";
            this.flipHorizontalButton.Name = "flipHorizontalButton";
            this.flipHorizontalButton.OfficeImageId = "ReviewCompareTwoVersions";
            this.flipHorizontalButton.ShowImage = true;
            this.flipHorizontalButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.flipHorizontalButton_Click);
            // 
            // flipVerticalButton
            // 
            this.flipVerticalButton.Label = "縦反転";
            this.flipVerticalButton.Name = "flipVerticalButton";
            this.flipVerticalButton.OfficeImageId = "RowHeight";
            this.flipVerticalButton.ShowImage = true;
            this.flipVerticalButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.flipVerticalButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.box7);
            this.group2.Items.Add(this.box10);
            this.group2.Items.Add(this.box11);
            this.group2.Label = "スライド編集";
            this.group2.Name = "group2";
            // 
            // box7
            // 
            this.box7.Items.Add(this.duplicateSlideButton);
            this.box7.Items.Add(this.insertSlideButton);
            this.box7.Name = "box7";
            // 
            // duplicateSlideButton
            // 
            this.duplicateSlideButton.Label = "複製";
            this.duplicateSlideButton.Name = "duplicateSlideButton";
            this.duplicateSlideButton.OfficeImageId = "DelegateAccess";
            this.duplicateSlideButton.ScreenTip = "複製";
            this.duplicateSlideButton.ShowImage = true;
            this.duplicateSlideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateSlideButton_Click);
            // 
            // insertSlideButton
            // 
            this.insertSlideButton.Label = "新規";
            this.insertSlideButton.Name = "insertSlideButton";
            this.insertSlideButton.OfficeImageId = "NewNotebookMenu";
            this.insertSlideButton.ScreenTip = "新規";
            this.insertSlideButton.ShowImage = true;
            this.insertSlideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertSlideButton_Click);
            // 
            // box10
            // 
            this.box10.Items.Add(this.slideCopyButton);
            this.box10.Items.Add(this.slidePasteButton);
            this.box10.Name = "box10";
            // 
            // slideCopyButton
            // 
            this.slideCopyButton.Label = "コピー";
            this.slideCopyButton.Name = "slideCopyButton";
            this.slideCopyButton.OfficeImageId = "ContactCardCopy";
            this.slideCopyButton.ScreenTip = "コピー";
            this.slideCopyButton.ShowImage = true;
            this.slideCopyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.slideCopyButton_Click);
            // 
            // slidePasteButton
            // 
            this.slidePasteButton.Label = "貼付";
            this.slidePasteButton.Name = "slidePasteButton";
            this.slidePasteButton.OfficeImageId = "ContactCardPaste";
            this.slidePasteButton.ScreenTip = "貼付";
            this.slidePasteButton.ShowImage = true;
            this.slidePasteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.slidePasteButton_Click);
            // 
            // box11
            // 
            this.box11.Items.Add(this.exportCurrentSlideAsPNGButton);
            this.box11.Name = "box11";
            // 
            // exportCurrentSlideAsPNGButton
            // 
            this.exportCurrentSlideAsPNGButton.Label = "頁画像";
            this.exportCurrentSlideAsPNGButton.Name = "exportCurrentSlideAsPNGButton";
            this.exportCurrentSlideAsPNGButton.OfficeImageId = "PhotoAlbumInsert";
            this.exportCurrentSlideAsPNGButton.ShowImage = true;
            this.exportCurrentSlideAsPNGButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportCurrentSlideAsPNGButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box8.ResumeLayout(false);
            this.box8.PerformLayout();
            this.box9.ResumeLayout(false);
            this.box9.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box7.ResumeLayout(false);
            this.box7.PerformLayout();
            this.box10.ResumeLayout(false);
            this.box10.PerformLayout();
            this.box11.ResumeLayout(false);
            this.box11.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox writeCommentCombo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeCommentInputButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeCommentAddButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeCommentAddFromFormButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeCommentAddFromFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delCommentSingleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delCommentAllButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton doEditComboButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox addCommentPreClearCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeCommentComboSaveButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertRectangleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertArrowButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertTextboxButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertCalloutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox writeMarkCombo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton writeMarkInputButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox writeMarkHamCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertBrButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateSlideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fontRedButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fontBlueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fontBlackButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton textDeleteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertSlideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fontBoldButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fontNarrowButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton slideCopyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton slidePasteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resetShapeStyleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bringFrontButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box8;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton textCutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton textPasteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton textCopyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertLineArrowButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton flipHorizontalButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton flipVerticalButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectObjectButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportCurrentSlideAsPNGButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
