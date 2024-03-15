namespace CY_UnrealHeaderAddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.group2 = this.Factory.CreateRibbonGroup();
            this.RegistButton = this.Factory.CreateRibbonButton();
            this.AddBothButton = this.Factory.CreateRibbonButton();
            this.AddCsvButton = this.Factory.CreateRibbonButton();
            this.AddHeaderButton = this.Factory.CreateRibbonButton();
            this.Generate = this.Factory.CreateRibbonButton();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.IncludeCsvCheckBox = this.Factory.CreateRibbonCheckBox();
            this.group2.SuspendLayout();
            this.tab2.SuspendLayout();
            this.SuspendLayout();
            // 
            // group2
            // 
            this.group2.Items.Add(this.RegistButton);
            this.group2.Items.Add(this.AddBothButton);
            this.group2.Items.Add(this.AddCsvButton);
            this.group2.Items.Add(this.AddHeaderButton);
            this.group2.Items.Add(this.Generate);
            this.group2.Label = "Add Module";
            this.group2.Name = "group2";
            // 
            // RegistButton
            // 
            this.RegistButton.Label = "Regist Direction";
            this.RegistButton.Name = "RegistButton";
            this.RegistButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnClick_RegistPathButton);
            // 
            // AddBothButton
            // 
            this.AddBothButton.Label = "Add Both";
            this.AddBothButton.Name = "AddBothButton";
            this.AddBothButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnClick_AddBothButton);
            // 
            // AddCsvButton
            // 
            this.AddCsvButton.Label = "Add Csv";
            this.AddCsvButton.Name = "AddCsvButton";
            this.AddCsvButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnClick_AddCsvButton);
            // 
            // AddHeaderButton
            // 
            this.AddHeaderButton.Label = "Add Header";
            this.AddHeaderButton.Name = "AddHeaderButton";
            this.AddHeaderButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnClick_AddHeaderButton);
            // 
            // Generate
            // 
            this.Generate.Label = "Generate Unreal Code";
            this.Generate.Name = "Generate";
            this.Generate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Generate_Click);
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "Unreal Header AddIn";
            this.tab2.Name = "tab2";
            // 
            // IncludeCsvCheckBox
            // 
            this.IncludeCsvCheckBox.Label = "Include Csv";
            this.IncludeCsvCheckBox.Name = "IncludeCsvCheckBox";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddCsvButton;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RegistButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddHeaderButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddBothButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox IncludeCsvCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Generate;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
