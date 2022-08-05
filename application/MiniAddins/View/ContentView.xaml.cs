using MiniAddins.Models;
using MiniAddins.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;


namespace MiniAddins.View
{
    /// <summary>
    /// ElementHost UserControl 的交互逻辑Entry
    /// </summary>
    public partial class ContentView : UserControl
    {

        /// <summary>
        /// Exit Event Handler Define  From ElementHost Control 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ElementHostEixtEventHandler(object sender, EventArgs args);
        public event ElementHostEixtEventHandler OnExit;


        /// <summary>
        /// Mouse Drag Move Event Handler Define  From ElementHost Control 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ElementHostMouseEventHandler(object sender, LocationEventArgs args);
        public event ElementHostMouseEventHandler OnDragMove;


        /// <summary>
        /// Image Export Event Handler Define  From ElementHost Control,
        /// Because Use EA Repository
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ElementHostImageExportEventHandler(object sender, ImageExportEventArgs args);
        public event ElementHostImageExportEventHandler OnImageExport;


        /// <summary>
        /// Column Button Click Event Handler Defination From ElementHost Control.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ElementHostColButtonClickEventHandler(object sender, ColButtonClickEventArgs args);
        public event ElementHostColButtonClickEventHandler OnColButtonClick;


        /// <summary>
        /// Get All Elements  Click Event Handler Defination From ElementHost Control.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public delegate void ElementHostAllTypesClickEventHandler(object sender, EventArgs args);
        public event ElementHostAllTypesClickEventHandler OnAllTypesClick;

        /// <summary>
        /// Element Host Control Project Main Enrty ViewModel
        /// </summary>
        private ContentViewModel contentViewModel;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="toolCatsIndex"></param>
        public ContentView(int toolCatsIndex)
        {
            if (null == System.Windows.Application.Current)
            {
                // User Control Dll中不存在Application全局对象，为此预先自行创建，用于保存全局入口UserControl实体
                new System.Windows.Application();
            }
            Application.Current.Resources.MergedDictionaries.Clear();
            Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri("/MiniAddins;component/Resource/resource.xaml", UriKind.Relative) });
            Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri("/MiniAddins;component/Resource/en.xaml", UriKind.Relative) });

            // 把自身埋入全局容器中，以备在其它类中进行访问
            string keyMain = "mainUserControl";
            if (Application.Current.Properties.Contains(keyMain))
            {
                Application.Current.Properties.Remove(keyMain);                
            }
            Application.Current.Properties.Add(keyMain, this);

            // 作成主ViewModel
            contentViewModel = new ContentViewModel(this);
            this.DataContext = contentViewModel;
            InitializeComponent();
            
        }
       
        /// <summary>
        /// Set Image Export View Display Value
        /// </summary>
        /// <param name="rootPath"></param>
        /// <param name="customSettings"></param>
        /// <param name="newCustomSettings"></param>
        public void SetImageExportViewDisplay(string rootPath, List<CustomSetting> customSettings, List<CustomSetting> newCustomSettings)
        {
            // 模型中什么也没时，退出
            if(customSettings==null || customSettings.Count==0)
            {
                return;
            }

            if (newCustomSettings == null)
            {
                ImageExportViewModel savedImageExportViewModel;

                // 从保存历史文件中取值
                savedImageExportViewModel = this.contentViewModel.DeserializeViewModel();
                if(savedImageExportViewModel !=null )
                {

                    rootPath = savedImageExportViewModel.BasicSetting.RootPath;
                    // 从保存历史文件中取值
                    DeserializeImageExportData(customSettings, savedImageExportViewModel);
                }
                
            }
            else
            {

                // 与改变的模型进行比较
                MergeImageExportData(customSettings, newCustomSettings);

            }

            this.contentViewModel.ImageExportViewModel.SetRootPath(rootPath);

            // 进行排序
            customSettings = (customSettings.AsParallel().OrderByDescending(item => item.Modified)).ToList();
            this.contentViewModel.ImageExportViewModel.SetCustomSetting(customSettings);

        }

        /// <summary>
        /// 合并有改变的模型
        /// </summary>
        /// <param name="customSettings"></param>
        /// <param name="newCustomSettings"></param>
        private void MergeImageExportData(List<CustomSetting> customSettings, List<CustomSetting> newCustomSettings)
        {
            Dictionary<String, CustomSetting> modelGuidList = new Dictionary<string, CustomSetting>();
            customSettings.ForEach(item =>
            {
                modelGuidList.Add(item.DiagramGUID, item);
            });


            // ****** new change merge ****** 

            // 1. 模型中存在，新清单中无(同一Package)          -> 需要删除
            // 2. 模型中存在，新清单中有且modifed time相异时   -> 状况=has change
            // 3. 模型中存在，新清单中有且modifed time相同时   -> 状况=保持原值
            // 4. 模型中无， 新清单有                          -> 需要增加
            Dictionary<String, CustomSetting> newGuidList = new Dictionary<string, CustomSetting>();
            newCustomSettings.ForEach(item =>
            {
                newGuidList.Add(item.DiagramGUID, item);
            });

            //
            // Do 1.
            //
            CustomSetting firstCustomSetting = newGuidList.Values.FirstOrDefault();

            var deletedObject = (from item in modelGuidList
                                 where !newGuidList.Keys.Contains(item.Key) && firstCustomSetting.PackageGUID.Equals(item.Value.PackageGUID)
                                 select item).ToList();

            deletedObject.ForEach(item =>
            {
                customSettings.Remove(item.Value);

                this.contentViewModel.AddChangeLog(item.Value, "removed");

            });

            //
            // Do 2.~3. 
            //
            var existingObject = (from item in newGuidList
                                  where modelGuidList.Keys.Contains(item.Key)
                                  select item).ToList();


            existingObject.ForEach(item =>
            {

                // modifed time相异时
                if (!item.Value.Modified.Equals(modelGuidList[item.Key].Modified))
                {
                    if (modelGuidList[item.Key].ExportState == ExportState.Already)
                    {
                        modelGuidList[item.Key].ExportState = ExportState.HasChange;
                    }
                    modelGuidList[item.Key].Diagram = item.Value.Diagram;
                    modelGuidList[item.Key].Modified = item.Value.Modified;
                    modelGuidList[item.Key].PackageName = item.Value.PackageName;
                    modelGuidList[item.Key].PackageFullName = item.Value.PackageFullName;

                    // 记录
                    this.contentViewModel.AddChangeLog(item.Value, "modified");
                }

            });

            //
            // Do 4.
            //
            var addingObject = (from item in newGuidList
                                where !modelGuidList.Keys.Contains(item.Key)
                                select item).ToList();

            if (addingObject != null && addingObject.Count > 0)
            {
                addingObject.ForEach(item =>
                {
                    customSettings.Add(item.Value);

                    // 记录
                    this.contentViewModel.AddChangeLog(item.Value, "added");

                });

            }
        }

        /// <summary>
        /// 从上次保存文件中读取数据
        /// </summary>
        /// <param name="customSettings"></param>
        /// <param name="savedImageExportViewModel"></param>
        private void DeserializeImageExportData(List<CustomSetting> customSettings, ImageExportViewModel savedImageExportViewModel)
        {
            Dictionary<String, CustomSetting> modelGuidList = new Dictionary<string, CustomSetting>();
            List<CustomSetting> savedCustomSettings;

            if (savedImageExportViewModel == null) return;

            customSettings.ForEach(item =>
            {
                modelGuidList.Add(item.DiagramGUID, item);
            });

            savedCustomSettings = savedImageExportViewModel.CustomSettings.ToList();

            // ****** history merge ****** 

            // 1. 模型中存在，历史清单中无                     -> 状况=保持原值
            // 2. 模型中存在，历史清单中有且modifed time相异时 -> 状况=has change
            // 3. 模型中存在，历史清单中有且modifed time相同时 -> 状况=历史结果
            // 4. 模型中存在，无历史清单                       -> 状况=no export[default]

            Dictionary<String, CustomSetting> savedGuidList = new Dictionary<string, CustomSetting>();
            savedCustomSettings.ForEach(item =>
            {
                savedGuidList.Add(item.DiagramGUID, item);
            });

            //
            // Do 1.
            //
            var noExportObject = (from item in modelGuidList
                                    where !savedGuidList.Keys.Contains(item.Key)
                                    select item.Value).ToList();

            noExportObject.ForEach(item =>
            {
                item.ExportState = ExportState.NoExport;
            });


            //
            // Do 2.~3.
            //
            var existingObject = (from item in modelGuidList
                                    where savedGuidList.Keys.Contains(item.Key)
                                    select item).ToList();


            existingObject.ForEach(item =>
            {

                // modifed time相异时
                if (!item.Value.Modified.Equals(savedGuidList[item.Key].Modified))
                {
                    if (savedGuidList[item.Key].ExportState == ExportState.Already)
                    {
                        item.Value.ExportState = ExportState.HasChange;
                    }
                    
                }
                else
                {
                    // modifed time相同时
                    item.Value.ExportState = savedGuidList[item.Key].ExportState;

                }
                item.Value.RelativePath = savedGuidList[item.Key].RelativePath;
                item.Value.Selected = savedGuidList[item.Key].Selected;
            });
            
        }

        /// <summary>
        /// Set Tool Cats SelectedIndex.
        /// </summary>
        /// <param name="selectedIndex"></param>
        public void SetToolCatsIndex(int selectedIndex)
        {
            this.contentViewModel.SelectedIndex = selectedIndex;
        }

        /// <summary>
        /// Reset Change log
        /// </summary>
        public void ResetChangelog()
        {
            this.contentViewModel.ChangeLog = null;
        }

        /// <summary>
        /// Set Status Bar Show
        /// </summary>
        public void SetStatusBarVisibility()
        {
            this.contentViewModel.StatusVisibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Set Statistics Result
        /// </summary>
        /// <param name="eaElements"></param>
        public void SetStatisticsResult(List<EaElement> eaElements)
        {

            // 汇总类型
            this.contentViewModel.StatisticsViewModel.ElementTypes.Clear();
            var elementTypes = (from item in eaElements
                               select item.ElementType).Distinct();

            
            foreach (string type in elementTypes)
            {
                this.contentViewModel.StatisticsViewModel.ElementTypes.Add(new CheckItem() { Name = type, IsChecked = true});
            }

            // 汇总作者
            this.contentViewModel.StatisticsViewModel.Authors.Clear();
            var authors = (from item in eaElements
                                select item.Author).Distinct();

            
            foreach (string author in authors)
            {
                this.contentViewModel.StatisticsViewModel.Authors.Add(new CheckItem() { Name = author, IsChecked = true });
            }


            // 汇总结果
            var st = from item in eaElements
                     group item by new { item.ElementType, item.Author }
                     into g
                     select new StatisticsResult {
                         Author=g.Key.Author,
                         ElementType=g.Key.ElementType,
                         Count=g.Count()
                     };

            this.contentViewModel.StatisticsViewModel.OrgStatisticsResults.Clear();
            this.contentViewModel.StatisticsViewModel.StatisticsResults.Clear();
            foreach (StatisticsResult statistics in st)
            {
                this.contentViewModel.StatisticsViewModel.OrgStatisticsResults.Add(statistics);
                this.contentViewModel.StatisticsViewModel.StatisticsResults.Add(statistics);

            }

        }

        /// <summary>
        /// Set Open EA Model Name
        /// </summary>
        /// <param name="eaModel"></param>
        public void SetOpenModelName(string eaModel)
        {
            this.contentViewModel.OpenModelName = eaModel;
        }
        
        /// <summary>
        /// 把事件传递到Hosting Form
        /// </summary>
        public void FireExitEvent()
        {
            if(this.OnExit !=null)
            {
                this.OnExit(this, null);
            }
        }

        /// <summary>
        /// 激发DragMove事件到Windows Form
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void FireDragMoveEvent(double x,double y)
        {
            if (this.OnDragMove != null)
            {
                this.OnDragMove(this, new LocationEventArgs() { X = x, Y = y });
            }
        }

        /// <summary>
        /// 激发ImageExport事件到Windows Form
        /// </summary>
        /// <param name="basicSetting"></param>
        /// <param name="customSettings"></param>
        public void FireImageExportEvent(BasicSetting basicSetting, List<CustomSetting> customSettings)
        {
            if (this.OnImageExport != null)
            {
                this.OnImageExport(this, new ImageExportEventArgs() { basicSetting = basicSetting,customSettings=customSettings} );
            }
        }

        /// <summary>
        /// 激发ColButtonClick事件到Windows Form
        /// </summary>
        /// <param name="type"></param>
        /// <param name="diagramGUID"></param>
        public void FireColButtonClickEvent(string type,string diagramGUID)
        {
            if (this.OnColButtonClick != null)
            {
                this.OnColButtonClick(this, new ColButtonClickEventArgs() { ButtonType = type, DiagramGUID = diagramGUID });
            }

        }

        /// <summary>
        /// 激发AllTypesClick事件到Windows Form
        /// </summary>
        public void FireAllTypesClickEvent()
        {
            if (this.OnAllTypesClick != null)
            {
                this.OnAllTypesClick(this, new EventArgs());
            }

        }


        public class LocationEventArgs : EventArgs
        {
            public double Y { get; set; }
            public double X { get; set; }
        }

        public class ImageExportEventArgs : EventArgs
        {
            public BasicSetting basicSetting { get; set; }
            public List<CustomSetting> customSettings { get; set; }
            
        }

        public class ColButtonClickEventArgs : EventArgs
        {
            public string ButtonType { get; set; }
            public string DiagramGUID { get; set; }

        }

    }

    
}
