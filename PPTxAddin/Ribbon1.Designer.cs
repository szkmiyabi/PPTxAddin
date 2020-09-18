﻿namespace PPTxAddin
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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.box4 = this.Factory.CreateRibbonBox();
            this.insertRectangleButton = this.Factory.CreateRibbonButton();
            this.insertTextboxButton = this.Factory.CreateRibbonButton();
            this.insertArrowButton = this.Factory.CreateRibbonButton();
            this.insertCalloutButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.group2.SuspendLayout();
            this.box4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "PPTxA";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.box3);
            this.group1.Label = "文字編集";
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
            this.writeCommentCombo.SizeString = "AAAAAAAAAAA";
            this.writeCommentCombo.Text = null;
            // 
            // writeCommentInputButton
            // 
            this.writeCommentInputButton.Label = "語句挿入";
            this.writeCommentInputButton.Name = "writeCommentInputButton";
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
            this.writeCommentAddButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentAddButton_Click);
            // 
            // writeCommentAddFromFormButton
            // 
            this.writeCommentAddFromFormButton.Label = "フォーム";
            this.writeCommentAddFromFormButton.Name = "writeCommentAddFromFormButton";
            this.writeCommentAddFromFormButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentAddFromFormButton_Click);
            // 
            // writeCommentAddFromFileButton
            // 
            this.writeCommentAddFromFileButton.Label = "ファイルから";
            this.writeCommentAddFromFileButton.Name = "writeCommentAddFromFileButton";
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
            this.box3.Name = "box3";
            // 
            // delCommentSingleButton
            // 
            this.delCommentSingleButton.Label = "1件削除";
            this.delCommentSingleButton.Name = "delCommentSingleButton";
            this.delCommentSingleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delCommentSingleButton_Click);
            // 
            // delCommentAllButton
            // 
            this.delCommentAllButton.Label = "全件削除";
            this.delCommentAllButton.Name = "delCommentAllButton";
            this.delCommentAllButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delCommentAllButton_Click);
            // 
            // doEditComboButton
            // 
            this.doEditComboButton.Label = "値編集";
            this.doEditComboButton.Name = "doEditComboButton";
            this.doEditComboButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.doEditComboButton_Click);
            // 
            // writeCommentComboSaveButton
            // 
            this.writeCommentComboSaveButton.Label = "ファイル保存";
            this.writeCommentComboSaveButton.Name = "writeCommentComboSaveButton";
            this.writeCommentComboSaveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.writeCommentComboSaveButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.box4);
            this.group2.Label = "描画";
            this.group2.Name = "group2";
            // 
            // box4
            // 
            this.box4.Items.Add(this.insertRectangleButton);
            this.box4.Items.Add(this.insertArrowButton);
            this.box4.Items.Add(this.insertTextboxButton);
            this.box4.Items.Add(this.insertCalloutButton);
            this.box4.Name = "box4";
            // 
            // insertRectangleButton
            // 
            this.insertRectangleButton.Label = "赤枠";
            this.insertRectangleButton.Name = "insertRectangleButton";
            this.insertRectangleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertRectangleButton_Click);
            // 
            // insertTextboxButton
            // 
            this.insertTextboxButton.Label = "文字枠";
            this.insertTextboxButton.Name = "insertTextboxButton";
            this.insertTextboxButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertTextboxButton_Click);
            // 
            // insertArrowButton
            // 
            this.insertArrowButton.Label = "図矢印";
            this.insertArrowButton.Name = "insertArrowButton";
            this.insertArrowButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertArrowButton_Click);
            // 
            // insertCalloutButton
            // 
            this.insertCalloutButton.Label = "吹出";
            this.insertCalloutButton.Name = "insertCalloutButton";
            this.insertCalloutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertCalloutButton_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertRectangleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertArrowButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertTextboxButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertCalloutButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
