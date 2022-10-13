using EA;
using MiniAddins.Models;
using MiniAddins.View;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;

namespace MiniAddinsFacade
{
    /// <summary>
    /// Summary description for frmModel.
    /// </summary>
    public class frmModel : System.Windows.Forms.Form
	{

        #region Private variable and Public Interface Define
        private EA.ModelWatcher modelWatcher;
        private bool isOpen =false;
        private int displayToolCatsIndex;
        private EA.Repository m_Repository;
        private int openModal;
		private System.Windows.Forms.Integration.ElementHost elementHost;
        private MiniAddins.View.ContentView ContentView;

        private ConcurrentDictionary<EA.IDualPackage, string> packageDict = new ConcurrentDictionary<IDualPackage,string>();
        private ConcurrentDictionary<string, EA.IDualDiagram> diagramDict = new ConcurrentDictionary<string, EA.IDualDiagram>();
        
        private List<CustomSetting> customSettings = new List<CustomSetting>();
        private List<EaElement> eaElements = new List<EaElement>();
        private System.Windows.Forms.Timer timerModelWatcher;
        private String[] eaModelInfo;

        /// <summary>
        /// Set EA Repository Object
        /// </summary>
        /// <param name="m_Repository"></param>
        public void setEARepository(EA.Repository m_Repository)
        {
            this.m_Repository = m_Repository;

        }

        /// <summary>
        /// Set EA modelWatcher
        /// </summary>
        /// <param name="modelWatcher"></param>
        public void setEAModelWatcher(EA.ModelWatcher modelWatcher)
        {
            this.modelWatcher = modelWatcher;

        }
        
        /// <summary>
        /// Set ToolCats Index From Menu Item
        /// </summary>
        /// <param name="displayToolCatsIndex"></param>
        public void SetDisplayToolCatsIndex(int displayToolCatsIndex)
        {
            this.displayToolCatsIndex = displayToolCatsIndex;

        }
        
        /// <summary>
        /// Set Form Open Modal
        /// </summary>
        /// <param name="modal"></param>
        public void SetModal(int modal)
        {
            this.openModal = modal;

        }

        /// <summary>
        /// Open Form
        /// </summary>
        /// <param name="modal"></param>
        public void OpenForm(int modal)
        {
            // set modal
            SetModal(modal);

            if (!this.isOpen)
            {
                // 初始化数据
                InitilizeData();
            }
            
            // open form with modal
            if (this.openModal == 0)
            {
                this.ShowDialog();
            }
            else
            {
                this.ShowInTaskbar = true;
                this.TopMost = true;
                this.timerModelWatcher.Enabled = true;
                this.Show();

            }

            this.isOpen = true;
        }


        #endregion


        #region Form Define
        private IContainer components;
        public frmModel()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
                        
            //
            // Bind Event Handler
            //
            this.ContentView = new MiniAddins.View.ContentView(this.displayToolCatsIndex);
            this.elementHost.Child = this.ContentView;

            this.ContentView.OnExit += new MiniAddins.View.ContentView.ElementHostEixtEventHandler(ContentView_OnExit);
            this.ContentView.OnDragMove += new MiniAddins.View.ContentView.ElementHostMouseEventHandler(ContentView_OnDragMove);
            this.ContentView.OnImageExport += new MiniAddins.View.ContentView.ElementHostImageExportEventHandler(ContentView_OnImageExport);
            this.ContentView.OnColButtonClick += new MiniAddins.View.ContentView.ElementHostColButtonClickEventHandler(ContentView_OnColButtonClick);
            this.ContentView.OnAllTypesClick += new MiniAddins.View.ContentView.ElementHostAllTypesClickEventHandler(ContentView_OnAllTypesClick);

        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

        #endregion


        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.elementHost = new System.Windows.Forms.Integration.ElementHost();
            this.timerModelWatcher = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // elementHost
            // 
            this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
            this.elementHost.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.elementHost.Location = new System.Drawing.Point(0, 0);
            this.elementHost.Name = "elementHost";
            this.elementHost.Size = new System.Drawing.Size(938, 552);
            this.elementHost.TabIndex = 3;
            this.elementHost.Text = "elementHost1";
            this.elementHost.Child = null;
            // 
            // timerModelWatcher
            // 
            this.timerModelWatcher.Interval = 5000;
            this.timerModelWatcher.Tick += new System.EventHandler(this.timerModelWatcher_Tick);
            // 
            // frmModel
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
            this.ClientSize = new System.Drawing.Size(938, 552);
            this.Controls.Add(this.elementHost);
            this.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmModel";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Model Tree";
            this.Load += new System.EventHandler(this.frmModel_Load);
            this.ResumeLayout(false);

        }
        #endregion

        
        #region Window Form Load
        /// <summary>
        /// windows form load事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmModel_Load(object sender, System.EventArgs e)
		{
            // Adjust From and Font Size
            int sizeScale = GetWindowsScaling();

            if (sizeScale != 100)
            {
                // comment out the fellow codes because AutoScale=Font
                // Font Size
                //float newFontSize = this.Font.Size * 100 / sizeScale;
                //Font newFont = new Font(this.Font.FontFamily, newFontSize);
                //this.Font = newFont;
                //this.elementHost.Font = newFont;

                // Form Size
                this.Width = (int)this.Width * sizeScale / 100;
                this.Height = (int)this.Height * sizeScale / 100;
            }

            // Send Data to View
            SetDataToContentView();

        }

        /// <summary>
        /// Get OS Font Scaling Rate
        /// </summary>
        /// <returns></returns>
        public  int GetWindowsScaling()
        {
            return (int)(100 * Screen.PrimaryScreen.Bounds.Width / SystemParameters.PrimaryScreenWidth);
        }
        #endregion


        #region Search & Scan EA Repository

        /// <summary>
        /// 初始化数据值
        /// </summary>
        private void InitilizeData()
        {
            
            // clear data dict
            this.diagramDict.Clear();
            this.packageDict.Clear();
            this.customSettings.Clear();
            this.eaElements.Clear();

            // Parse EA model fiie
            eaModelInfo = ParseRootPathFromModel(this.m_Repository.ConnectionString);

            // Scan All Packages
            foreach (EA.IDualPackage p in m_Repository.Models)
            {
                ScanEaPackage(p);
            }

            // Read Diagram and Elements from EA Repository
            ReadDiagramFromEaPackage();

        }
        /// <summary>
        /// Initilize Element Host User Control
        /// </summary>
        /// <param name="IsLocal"></param>
        private void SetDataToContentView()
        {
            // Initilize Element Host User Control
            this.ContentView.SetOpenModelName(eaModelInfo[1]);

            // set image export data to view
            this.ContentView.SetImageExportViewDisplay(eaModelInfo[0], customSettings, null);

            // set statistics data to view
            this.ContentView.SetStatisticsResult(eaElements);

            this.ContentView.SetToolCatsIndex(this.displayToolCatsIndex);

            if (this.openModal == 0)
            {
                this.ContentView.SetStatusBarVisibility();
            }

            this.ContentView.ResetChangelog();
        }

        /// <summary>
        /// 解析出根目录
        /// 根目录为模型文件的上两层目录
        /// 例如：C：\workspace\projectname\model\proj_main.eapx，保存根目录为C：\workspace\projectname\images
        /// </summary>
        /// <param name="modelConnectionString"></param>
        /// <returns></returns>
        private string[] ParseRootPathFromModel(string modelConnectionString)
        {
            string[] modelString = modelConnectionString.Split(@"\".ToCharArray());
            string rootPath = String.Join(@"\", modelString, 0, modelString.Length - 2);

            return new string[] { $"{rootPath}\\images", modelString[modelString.Length - 1] };

        }
                
        /// <summary>
        /// Scan EA Packages
        /// </summary>
        /// <param name="parentPackage"></param>
        private void ScanEaPackage(EA.IDualPackage parentPackage)
        {
            
            this.packageDict.AddOrUpdate(parentPackage, parentPackage.PackageGUID, (key,value)=> parentPackage.PackageGUID);
            
            EA.IDualCollection c = parentPackage.Packages;
            foreach(EA.IDualPackage package in c)
            {
                ScanEaPackage(package);
            }

        }

        /// <summary>
        /// Get Diagram Infomation,return CustomSetting List
        /// </summary>        
        /// <returns></returns>
        private void ReadDiagramFromEaPackage()
        {
            foreach(EA.IDualPackage Kid in this.packageDict.Keys)
            {
                // read Diagram
                ReadDiagramFromByPackage(Kid, this.diagramDict, this.customSettings);
                
            }
           
            // Count Diagram as Element 
            GetElementFromDiagram(this.eaElements);

        }

        /// <summary>
        /// Get Elements Infomation,return CustomSetting List
        /// </summary>        
        /// <returns></returns>
        private void ReadElementFromEaPackage()
        {
            foreach (EA.IDualPackage Kid in this.packageDict.Keys)
            {
              
                // read element
                ReadElementFromByPackage(Kid, this.eaElements);
            }


        }

        /// <summary>
        /// Diagram as Element Stastistics
        /// </summary>
        /// <param name="eaElements"></param>
        private void GetElementFromDiagram(List<EaElement> eaElements)
        {

            foreach(EA.IDualDiagram diagram in this.diagramDict.Values)
            {
                // Elements
                var setting = eaElements.FirstOrDefault(item => item.ElementGUID.Equals(diagram.DiagramGUID));
                if (setting != null)
                {
                    eaElements.Remove(setting);
                }

                eaElements.Add(new EaElement()
                {
                    ElementGUID = diagram.DiagramGUID,
                    ElementType = "Diagram",
                    Author = diagram.Author,
                    Name = diagram.Name
                });
            }
            
            
        }

        /// <summary>
        /// Read Diagram Infomation By Package
        /// </summary>
        /// <param name="dualPackage"></param>
        /// <param name="diagramDict"></param>
        /// <param name="customSettings"></param>
        /// <param name="eaElements"></param>
        private void ReadDiagramFromByPackage(EA.IDualPackage dualPackage, ConcurrentDictionary<string, EA.IDualDiagram> diagramDict, List<CustomSetting> customSettings)
        {

            foreach( EA.IDualDiagram diagram in dualPackage.Diagrams)
            {
                // 原始Diagram
                diagramDict.AddOrUpdate(diagram.DiagramGUID, diagram, (key, value) => diagram);

                var setting = customSettings.FirstOrDefault(item => item.DiagramGUID.Equals(diagram.DiagramGUID));
                if (setting != null)
                {
                    customSettings.Remove(setting);
                }

                customSettings.Add(new CustomSetting()
                {
                    DiagramGUID = diagram.DiagramGUID,
                    Diagram = diagram.Name,
                    Selected = true,
                    PackageName = dualPackage.Name,
                    Created = diagram.CreatedDate,
                    Modified = diagram.ModifiedDate,
                    PackageGUID = dualPackage.PackageGUID,
                    ExportState = ExportState.NoExport,
                    PackageFullName = GetPackageFullPath(dualPackage.PackageID)

                });

            }
        }
        
        /// <summary>
        /// 获得包的全路径
        /// </summary>
        /// <param name="packageID"></param>
        /// <returns></returns>
        private string GetPackageFullPath(int packageID)
        {
            EA.IDualPackage package = this.m_Repository.GetPackageByID(packageID);
            List<string> pathList = new List<string>();
            StringBuilder stringBuilder = new StringBuilder();
            pathList.Add(package.Name);

            while (package.ParentID > 0)
            {   
                packageID = package.ParentID;
                package = this.m_Repository.GetPackageByID(packageID);
                pathList.Add(package.Name);
            }

            // remove root package 
            for(int i = pathList.Count - 1; i >= 0; i--)
            {
                stringBuilder.Append($"{pathList[i]}\\");
            }
            string result = stringBuilder.ToString();
            return result.Substring(0, result.Length-1);
        }

        /// <summary>
        /// Read Element From a Package
        /// </summary>
        /// <param name="dualPackage"></param>
        /// <param name="eaElements"></param>
        private void ReadElementFromByPackage(EA.IDualPackage dualPackage,  List<EaElement> eaElements)
        {
            foreach(EA.IDualElement el in dualPackage.Elements)
            {
                var setting = eaElements.FirstOrDefault(item => item.ElementGUID.Equals(el.ElementGUID));
                if (setting != null)
                {
                    eaElements.Remove(setting);
                }

                eaElements.Add(new EaElement()
                {
                    ElementGUID = el.ElementGUID,
                    ElementType = el.Type,
                    Author = el.Author,
                    Name = el.Name
                });
            }

        }

        #endregion


        #region Event Handler from Element Host User Control Component

        /// <summary>
        /// 画面拖动事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void ContentView_OnDragMove(object sender, MiniAddins.View.ContentView.LocationEventArgs args)
        {
            this.Location = new System.Drawing.Point(this.Location.X + (int)args.X, this.Location.Y + (int)args.Y);     
            
            this.Update();
        }

        /// <summary>
        /// 关闭退出事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void ContentView_OnExit(object sender, EventArgs args)
        {
            if (this.openModal == 0)
            {
                this.Close();
            }
            else
            {
                this.Hide();
            }
            
        }

        /// <summary>
        /// 执行图片导出事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void ContentView_OnImageExport(object sender, MiniAddins.View.ContentView.ImageExportEventArgs args)
        {

            // ****************** 一些commnets **************

            // 图像输出方式
            // 方式1、IDualProject接口中的GetAllDiagramImageAndMap/GetDiagramImageAndMap
            //    条件：The 'Auto Create Diagram Image and Image Map' option must be selected in the model-specific options for this function to save the image and image-map
            //    不足：输出图像中没有title及边框
            // 
            // 方式2、Diagram类中的SaveImagePage 按指定页方式来输出图片
            //    不足：如果模型跨多页时不能在一页中显示完全;输出图像中没有title及边框
            //
            // 方式3、IDualProject接口中的PutDiagramImageToFile
            //    不足：每导出一个图像时，EA中也将打开相应Diagram.

            // **********************************************
            if (args.customSettings == null || args.customSettings.Count == 0)
            {
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            EA.IDualProject dualProject = this.m_Repository.GetProjectInterface();

            try
            {

                args.customSettings.ForEach(item =>
                {

                    Debug.WriteLine($"ManagedThreadId={System.Threading.Thread.CurrentThread.ManagedThreadId}");

                    bool success = false;
                    string imagePath = $"{args.basicSetting.RootPath}\\{item.RelativePath}";
                    if (!Directory.Exists(imagePath))
                    {
                        Directory.CreateDirectory(imagePath);
                    }

                    success = dualProject.PutDiagramImageToFile(item.DiagramGUID, $"{imagePath}\\{item.Diagram}.png", 1);
                    if (!success)
                    {
                        Debug.WriteLine(dualProject.GetLastError());
                    }
                    else
                    {
                        // 关闭EA中打开的Diagram
                        this.m_Repository.CloseDiagram(this.diagramDict[item.DiagramGUID].DiagramID);
                    }

                });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        /// <summary>
        /// Column Button Click Event Handler
        /// 1. ButtonType=image，copy image to clipboard
        /// 2. ButtonType=open，open diagramGUID in EA
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void ContentView_OnColButtonClick(object sender, ContentView.ColButtonClickEventArgs args)
        {
            if (args.ButtonType.Equals("image"))
            {
                EA.IDualProject dualProject = this.m_Repository.GetProjectInterface();                
                bool success = false;
                success = dualProject.PutDiagramImageOnClipboard(args.DiagramGUID, 1);
                if (!success)
                {
                    Debug.WriteLine(dualProject.GetLastError());
                }
                else
                {                    
                    this.m_Repository.CloseDiagram(this.diagramDict[args.DiagramGUID].DiagramID);
                }

            }
            else if (args.ButtonType.Equals("open"))
            {
                this.m_Repository.OpenDiagram(this.diagramDict[args.DiagramGUID].DiagramID);

            }
        }

        /// <summary>
        /// All Types Button Click Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void ContentView_OnAllTypesClick(object sender, EventArgs args)
        {
            Cursor.Current = Cursors.WaitCursor;

            // read elements
            ReadElementFromEaPackage();

            // set statistics data to view
            this.ContentView.SetStatisticsResult(this.eaElements);

            Cursor.Current = Cursors.Default;

        }
        #endregion


        #region EA Watcher Change Handler
        /// <summary>
        /// 主动扫描EA模型的变更点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timerModelWatcher_Tick(object sender, EventArgs e)
        {
            object reloadItem;
            ReloadType reloadType = ReloadType.rtNone;
            int packageID =-1;
            EA.IDualPackage containPackage;

            reloadType = this.modelWatcher.GetReloadItem(out reloadItem);

            if (reloadType == ReloadType.rtNone) return;

            // 重新打开了模型文件
            if(reloadType == ReloadType.rtEntireModel)
            {
                // 初始化数据
                InitilizeData();

                // 把数据反映到视图上
                SetDataToContentView();

                return;
            }

            // when package is changed
            if (reloadType == ReloadType.rtPackage && reloadItem !=null)
            {
                EA.IDualPackage dualPackage = reloadItem as EA.IDualPackage;
                packageID = dualPackage.PackageID;
                // Debug.WriteLine($"changed package name ={dualPackage.Name},package id={packageID}");
                
            }

            // 
            if (reloadType == ReloadType.rtElement && reloadItem != null)
            {
                EA.IDualElement dualElement = reloadItem as EA.IDualElement;
                packageID = dualElement.PackageID;
                // Debug.WriteLine($"changed element name ={dualElement.Name},Object type={dualElement.ObjectType},package id={packageID}");

            }

            if (packageID == -1) return;

            containPackage = this.m_Repository.GetPackageByID(packageID);
            if (containPackage.Diagrams.Count == 0) return;

            // read modified diagram
            List<CustomSetting> newCustomSettings = new List<CustomSetting>();

            ReadDiagramFromByPackage(containPackage, this.diagramDict, newCustomSettings);
            this.ContentView.SetImageExportViewDisplay(eaModelInfo[0], this.customSettings, newCustomSettings);
            

        }
        #endregion

    }
}
