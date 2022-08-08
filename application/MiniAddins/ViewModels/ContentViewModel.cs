using MiniAddins.Command;
using MiniAddins.Models;
using MiniAddins.View;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Xml.Serialization;

namespace MiniAddins.ViewModels
{
    /// <summary>
    /// 主ViewModel
    /// </summary>
    public class ContentViewModel : ViewModel
    {

        #region 构建函数及变量定义
        public ContentView userControl;
        private List<ChangeLog> changeLogs;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="mainview"></param>
        public ContentViewModel(ContentView mainview)
        {
            this.userControl = mainview;
            this.toolCats = new ObservableCollection<ToolCat>();
            this.changeLogs = new List<ChangeLog>();
            UpdateToolCats(this.toolCats);

        }

        /// <summary>
        /// Add Change Log
        /// </summary>
        /// <param name="customSetting"></param>
        public void AddChangeLog(CustomSetting customSetting,string type)
        {
            this.ChangeLog = new Models.ChangeLog() { LogContent = customSetting, LogType = type,AtTime=customSetting.Modified };
            this.changeLogs.Add(this.ChangeLog);
        }

        /// <summary>
        /// Open Model Name ,for example:test.eapx
        /// </summary>
        public string OpenModelName { set; get; }

        #endregion


        #region 状况栏的表示可否
        private Visibility statusVisibility = Visibility.Visible;
        public Visibility StatusVisibility
        {
            get
            {

                return this.statusVisibility;
            }
            set
            {
                if (Equals(value, this.statusVisibility))
                {
                    return;
                }

                this.statusVisibility = value;
                OnPropertyChanged(nameof(StatusVisibility));

            }
        }
        #endregion


        #region ToolCatCollection
        private ObservableCollection<ToolCat>  toolCats = null;
        public ObservableCollection<ToolCat> ToolCats
        {
            get
            {
               
                return this.toolCats;
            }
            set
            {
                if (Equals(value, this.toolCats))
                {
                    return;
                }

                this.toolCats = value;
                OnPropertyChanged(nameof(ToolCats));

            }

        }
        /// <summary>
        /// 动态增加左侧菜单项
        /// </summary>
        /// <param name="toolCats"></param>
        private void UpdateToolCats(ObservableCollection<ToolCat> toolCats)
        {
            toolCats.Clear();
            toolCats.Add(new ToolCat() { Key = "export_image", Icon = @"..\Icons\file-export.png", View = @"ImageExportView", DisplayNameResourceKey = "multilang_01" });
            toolCats.Add(new ToolCat() { Key = "export_image", Icon = @"..\Icons\sigma.png", View = @"StatisticsView", DisplayNameResourceKey = "multilang_02" });

        }
        #endregion


        #region ImageExportViewModel
        ImageExportViewModel imageExportViewModel = null;
        public ImageExportViewModel ImageExportViewModel
        {
            get
            {
                if (imageExportViewModel == null)
                    imageExportViewModel = new ImageExportViewModel(this);

                return imageExportViewModel;
            }
        }

        #endregion


        #region StatisticsViewModel

        StatisticsViewModel statisticsViewModel = null;
        public StatisticsViewModel StatisticsViewModel
        {
            get
            {
                if (statisticsViewModel == null)
                    statisticsViewModel = new StatisticsViewModel(this);

                return statisticsViewModel;
            }
        }
        #endregion


        #region Selected Index
        int _selectedIndex = 0;
        public int SelectedIndex
        {
            get { return this._selectedIndex; }
            set
            {
                if (Equals(value, _selectedIndex))
                {
                    return;
                }

                _selectedIndex = value;
                OnPropertyChanged(nameof(SelectedIndex));
            }
        }
        #endregion


        #region Language
        string _language = "cn";
        public string Language
        {
            get { return this._language; }
            set
            {
                if (Equals(value, _language))
                {
                    return;
                }

                _language = value;
                OnPropertyChanged(nameof(Language));
            }
        }
        #endregion


        #region Language Change
        private RelayCommand changeLanguageCommand;
        /// <summary>
        /// Change Language
        /// </summary>
        public ICommand ChangeLanguageCommand
        {
            get
            {
                if (this.changeLanguageCommand == null)
                {
                    this.changeLanguageCommand = new RelayCommand<string>((p) => ChangeLanguage(p));
                }

                return this.changeLanguageCommand;
            }
        }

        private void ChangeLanguage(string p)
        {

            // change language resource
            ResourceDictionary rd = this.userControl.Resources.MergedDictionaries.LastOrDefault<ResourceDictionary>();
            rd.Source = new System.Uri(string.Format("/MiniAddins;component/Resource/{0}.xaml", p), System.UriKind.Relative);

            // refresh tool cats listview
            int oldIndex = SelectedIndex;
            ObservableCollection<ToolCat> toolCats = new ObservableCollection<ToolCat>();
            UpdateToolCats(this.ToolCats);
            SelectedIndex = oldIndex;
            

            // set change language
            if (p.Equals("cn"))
            {
                this.Language = "en";
            }
            if (p.Equals("en"))
            {
                this.Language = "cn";
            }

            // refresh image export view
            ObservableCollection<CustomSetting> customSettings = this.ImageExportViewModel.DisplayCustomSettings;
            this.ImageExportViewModel.DisplayCustomSettings = null;
            this.ImageExportViewModel.DisplayCustomSettings = customSettings;
        }
        #endregion


        #region Exit
        private RelayCommand exitCommand;
        /// <summary>
        /// Exit from the application
        /// </summary>
        public ICommand ExitCommand
        {
            get
            {
                if (this.exitCommand == null)
                {
                    this.exitCommand = new RelayCommand(()=> {
                        this.userControl.FireExitEvent();
                    });
                }

                return this.exitCommand;
            }
        }

        #endregion


        #region Drag
        private bool mouseDown;
        private Point lastLocation;
        private RelayCommand dragCommand;
        private RelayCommand dragCommandUp;
        private RelayCommand dragCommandMove;
        public ICommand DragCommand
        {
            get
            {
                if (this.dragCommand == null)
                {
                    this.dragCommand = new RelayCommand<object>((p)=> {
                        mouseDown = true;
                        lastLocation = Mouse.GetPosition(this.userControl);
                        this.userControl.MouseContainer.CaptureMouse();

                    },(p)=>true);
                }

                return this.dragCommand;
            }
        }

        public ICommand DragCommandUp
        {
            get
            {
                if (this.dragCommandUp == null)
                {
                    this.dragCommandUp = new RelayCommand<object>((p) => {

                        mouseDown = false;
                        this.userControl.MouseContainer.ReleaseMouseCapture();

                    }, (p) => true);
                }

                return this.dragCommandUp;
            }
        }

        public ICommand DragCommandMove
        {
            get
            {
                if (this.dragCommandMove == null)
                {
                    this.dragCommandMove = new RelayCommand<MouseEventArgs>((p) => {

                        if (mouseDown)
                        {
                            Point location = Mouse.GetPosition(this.userControl);
                            double x;
                            double y;

                            x = location.X - lastLocation.X;
                            y = location.Y - lastLocation.Y;

                            Debug.WriteLine($"dx={x},dy={y}") ;
                            this.userControl.FireDragMoveEvent(x, y);
                        };

                    }, (p) => true);
                }

                return this.dragCommandMove;
            }
        }

        #endregion


        #region Change Log
        ChangeLog _changeLog = null;
        public ChangeLog ChangeLog
        {
            get { return this._changeLog; }
            set
            {
                if (Equals(value, _changeLog))
                {
                    return;
                }

                _changeLog = value;
                OnPropertyChanged(nameof(ChangeLog));
            }
        }
        #endregion


        #region Save Configuration
        private RelayCommand saveConfigurationCommand;
        /// <summary>
        /// Save View Display Value
        /// </summary>
        public ICommand SaveConfigurationCommand
        {
            get
            {
                if (this.saveConfigurationCommand == null)
                {
                    this.saveConfigurationCommand = new RelayCommand(()=> {
                        SerializeViewModel();
                        this.userControl.FireExitEvent();
                    });
                }

                return this.saveConfigurationCommand;
            }
        }

        #endregion


        #region 序列化，反序列化方法
        /// <summary>
        /// Get Xml Config File Path
        /// </summary>
        /// <returns></returns>
        private string GetConfigFilePath()
        {

            string userTempPath = System.IO.Path.GetTempPath();
            FileInfo fileInfo = new FileInfo(userTempPath);
            string addinsForlder = "Addins-config";
            string configDirectory = $"{fileInfo.DirectoryName}\\{addinsForlder}";
            string xmlFile = $"{configDirectory}\\{this.OpenModelName}.xml";

            if (!Directory.Exists(configDirectory))
            {
                fileInfo.Directory.CreateSubdirectory(addinsForlder);
            }
            return xmlFile;
        }

        /// <summary>
        /// 序列化ImageExportViewModel
        /// </summary>
        public void SerializeViewModel()
        {
            string xmlFile = GetConfigFilePath();
            XmlSerializer serializer = new XmlSerializer(typeof(ImageExportViewModel));

            // 序列化
            TextWriter writer = new StreamWriter(xmlFile);
            serializer.Serialize(writer, this.ImageExportViewModel);
            writer.Close();
        }

        /// <summary>
        /// 反序列化
        /// </summary>
        /// <returns></returns>
        public ImageExportViewModel DeserializeViewModel()
        {
            string xmlFile = GetConfigFilePath();
            if (!File.Exists(xmlFile))
            {
                return null;
            }
            XmlSerializer serializer = new XmlSerializer(typeof(ImageExportViewModel));
            FileStream fs = new FileStream(xmlFile, FileMode.Open);

            return serializer.Deserialize(fs) as ImageExportViewModel;

        }
        #endregion


        #region Copy Change Log
        private RelayCommand copyChangeLogCommand;
        /// <summary>
        /// Copy Change Log
        /// </summary>
        public ICommand CopyChangeLogCommand
        {
            get
            {
                if (this.copyChangeLogCommand == null)
                {
                    this.copyChangeLogCommand = new RelayCommand<string>((p) => CopyChangeLog(p),(p) => {
                        if ("one".Equals(p))
                        {
                            return this.ChangeLog != null;
                        }
                        else
                        {
                            return this.changeLogs.Count > 0;
                        }
                    });
                }

                return this.copyChangeLogCommand;
            }
        }

        private void CopyChangeLog(string p)
        {
            if ("one".Equals(p))
            {
                Clipboard.SetText(this.ChangeLog.DisplayString);
            }
            else if ("multi".Equals(p))
            {
                StringBuilder stringBuilder = new StringBuilder();
                foreach(ChangeLog changeLog in this.changeLogs)
                {
                    stringBuilder.AppendLine(changeLog.DisplayString);
                }
                Clipboard.SetText(stringBuilder.ToString());

            }

        }
        #endregion
    }
}
