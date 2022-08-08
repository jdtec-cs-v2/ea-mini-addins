using MiniAddins.Command;
using MiniAddins.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Serialization;
using MSAPI = Microsoft.WindowsAPICodePack;

namespace MiniAddins.ViewModels
{
    [XmlRoot("ImageExport",Namespace="http://wwww.csadd-ins.com",IsNullable =false)]
    public class ImageExportViewModel : ViewModel
    {
        private ContentViewModel _mainvm;

        public ImageExportViewModel()
        {

        }

        public ImageExportViewModel(ContentViewModel mainvm)
        {
            this._mainvm = mainvm;

        }
        #region Set Value
        public void SetRootPath(string rootPath)
        {
            BasicSetting setting =  this.basicSetting.Clone();
            setting.RootPath =rootPath;
            this.BasicSetting = setting;
            return;
        }

        public void SetCustomSetting(List<CustomSetting> customSettings)
        {
            this.CustomSettings.Clear();
            this.DisplayCustomSettings.Clear();
            customSettings.ForEach(item => {

                this.CustomSettings.Add(item);
                this.DisplayCustomSettings.Add(item);

            });
        }
        #endregion


        #region filter value
        private string searchValue ="";
        public string SearchValue
        {
            get { return this.searchValue; }
            set
            {
                if (Equals(value, this.searchValue))
                {
                    return;
                }
                this.searchValue = value;
                OnPropertyChanged(nameof(SearchValue));               

            }
        }
        #endregion


        #region All Select CheckBox
        private bool allSelected = false;
        public bool AllSelected
        {
            get { return this.allSelected; }
            set
            {
                if (Equals(value, this.allSelected))
                {
                    return;
                }
                this.allSelected = value;
                OnPropertyChanged(nameof(AllSelected));
                //
                // 设置DataGrid中的值
                //
                foreach (CustomSetting customSetting in this.DisplayCustomSettings)
                {
                    customSetting.Selected = this.AllSelected;
                }

            }
        }
        #endregion


        #region not export CheckBox
        private bool notExportSelected = false;
        public bool NotExportSelected
        {
            get { return this.notExportSelected; }
            set
            {
                if (Equals(value, this.notExportSelected))
                {
                    return;
                }
                this.notExportSelected = value;
                OnPropertyChanged(nameof(NotExportSelected));
                //
                // 设置DataGrid中的值
                //
                foreach (CustomSetting customSetting in this.DisplayCustomSettings)
                {
                    if (customSetting.ExportState == ExportState.NoExport)
                    {
                        customSetting.Selected = this.NotExportSelected;
                    }
                    
                }

            }
        }
        #endregion


        #region Already CheckBox
        private bool alreadySelected = false;
        public bool AlreadySelected
        {
            get { return this.alreadySelected; }
            set
            {
                if (Equals(value, this.alreadySelected))
                {
                    return;
                }
                this.alreadySelected = value;
                OnPropertyChanged(nameof(AlreadySelected));
                //
                // 设置DataGrid中的值
                //
                foreach (CustomSetting customSetting in this.DisplayCustomSettings)
                {
                    if (customSetting.ExportState == ExportState.Already)
                    {
                        customSetting.Selected = this.alreadySelected;
                    }

                }

            }
        }
        #endregion


        #region Already CheckBox
        private bool hasChangeSelected = false;
        public bool HasChangeSelected
        {
            get { return this.hasChangeSelected; }
            set
            {
                if (Equals(value, this.hasChangeSelected))
                {
                    return;
                }
                this.hasChangeSelected = value;
                OnPropertyChanged(nameof(HasChangeSelected));
                //
                // 设置DataGrid中的值
                //
                foreach (CustomSetting customSetting in this.DisplayCustomSettings)
                {
                    if (customSetting.ExportState == ExportState.HasChange)
                    {
                        customSetting.Selected = this.hasChangeSelected;
                    }

                }

            }
        }
        #endregion


        #region BasicSetting
        private BasicSetting basicSetting = new BasicSetting();
        
        public BasicSetting BasicSetting
        {
            get { return this.basicSetting; }
            set
            {
                if (Equals(value, this.basicSetting))
                {
                    return;
                }
                this.basicSetting = value;
                OnPropertyChanged(nameof(BasicSetting));
            }
        }

        #endregion


        #region CustomSettings
        public ObservableCollection<CustomSetting> CustomSettings = new ObservableCollection<CustomSetting>();
        
        private ObservableCollection<CustomSetting> displayCustomSettings = new ObservableCollection<CustomSetting>();
        public ObservableCollection<CustomSetting> DisplayCustomSettings
        {
            get { return this.displayCustomSettings; }
            set
            {
                if (Equals(value, this.displayCustomSettings))
                {
                    return;
                }
                this.displayCustomSettings = value;
                OnPropertyChanged(nameof(DisplayCustomSettings));
            }

        }

        #endregion


        #region Export Image
        RelayCommand _ExportImage = null;
        public ICommand ExportImage
        {
            get
            {
                if (_ExportImage == null)
                {
                    _ExportImage = new RelayCommand(() => OnExportImage(), () => {

                        return this.DisplayCustomSettings.Count(item=>item.Selected)>0;
                    });
                }

                return _ExportImage;
            }
        }
        
        private void OnExportImage()
        {
            // select
            var selectedList = from item in this.DisplayCustomSettings
                               where item.Selected
                               select item;

            foreach(CustomSetting setting in selectedList)
            {
                setting.ExportState = ExportState.Already;
            }

            this._mainvm.userControl.FireImageExportEvent(this.BasicSetting, selectedList.ToList());

        }
        #endregion


        #region Open
        RelayCommand _openCommand = null;
        public ICommand OpenCommand
        {
            get
            {
                if (_openCommand == null)
                {
                    _openCommand = new RelayCommand<string>((p) => OnOpen(p), (p) => true);


                }

                return _openCommand;
            }
        }
        
        private void OnOpen(string p)
        {
            string sResult = string.Empty;
            var dlg = new MSAPI::Dialogs.CommonOpenFileDialog();
            dlg.IsFolderPicker = true;
            dlg.Title = (string)(Application.Current.Properties["mainUserControl"] as UserControl).Resources["multilang_message_01"];
            dlg.InitialDirectory = this.basicSetting.RootPath;

            if (dlg.ShowDialog() == MSAPI::Dialogs.CommonFileDialogResult.Ok)
            {
                sResult = dlg.FileName;
                BasicSetting newBasicSetting = BasicSetting.Clone();
                newBasicSetting.RootPath = sResult;
                BasicSetting = newBasicSetting;
            }

        }
        #endregion


        #region Set Default Sub Forlder
        RelayCommand _setDefaultForlder = null;
        public ICommand SetDefaultForlder
        {
            get
            {
                if (_setDefaultForlder == null)
                {
                    _setDefaultForlder = new RelayCommand(() => {

                        var selectedList = from item in this.DisplayCustomSettings
                                           where item.Selected
                                           select item;

                        foreach (CustomSetting setting in selectedList)
                        {
                            bool possiblePath = false;
                            possiblePath= (setting.PackageFullName.Replace("\\","").IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) == -1);
                            if (possiblePath)
                            {
                                setting.RelativePath = setting.PackageFullName;
                            }
                            

                        }
                    }, () => {

                        return this.DisplayCustomSettings.Count(item => item.Selected) > 0;
                    });
                }

                return _setDefaultForlder;
            }
        }
        #endregion


        #region Copy Diagram Image Path
        RelayCommand _copyImagePathCommand = null;
        public ICommand CopyImagePathCommand
        {
            get
            {
                if (_copyImagePathCommand == null)
                {
                    _copyImagePathCommand = new RelayCommand<string>((p) => {

                        Clipboard.SetText($"{this.BasicSetting.RootPath}\\{p}");

                    }, (p) => true);


                }

                return _copyImagePathCommand;
            }
        }
        #endregion


        #region Open Diagram 
        RelayCommand _openDiagramCommand = null;
        public ICommand OpenDiagramCommand
        {
            get
            {
                if (_openDiagramCommand == null)
                {
                    _openDiagramCommand = new RelayCommand<string>((p) => {
                        this._mainvm.userControl.FireColButtonClickEvent("open",p);
                    }, (p) => true);


                }

                return _openDiagramCommand;
            }
        }
        #endregion


        #region copy image to clipboard
        RelayCommand _copyImageCommand = null;
        public ICommand CopyImageCommand
        {
            get
            {
                if (_copyImageCommand == null)
                {
                    _copyImageCommand = new RelayCommand<string>((p) => {
                        this._mainvm.userControl.FireColButtonClickEvent("image", p);
                    }, (p) => true);


                }

                return _copyImageCommand;
            }
        }
        #endregion

        
        #region Search Diagram
        RelayCommand _searchCommand = null;
        public ICommand SearchCommand
        {
            get
            {
                if (_searchCommand == null)
                {
                    _searchCommand = new RelayCommand(() => {

                        // Refresh List
                        this.DisplayCustomSettings.Clear();

                        var list = from item in this.CustomSettings
                                   where item.Diagram.Contains(this.SearchValue) 
                                   || (item.RelativePath!=null && item.RelativePath.Contains(this.SearchValue)) 
                                   || item.PackageFullName.Contains(this.SearchValue)
                                   select item;

                        foreach (CustomSetting item in list)
                        {
                            this.DisplayCustomSettings.Add(item);
                        }

                    }, () => true);


                }

                return _searchCommand;
            }
        }
        #endregion

    }
}
