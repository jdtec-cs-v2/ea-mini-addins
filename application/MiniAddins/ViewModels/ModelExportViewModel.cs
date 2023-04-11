using Microsoft.WindowsAPICodePack.Dialogs;
using MiniAddins.Command;
using MiniAddins.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Serialization;
using MSAPI = Microsoft.WindowsAPICodePack;

namespace MiniAddins.ViewModels
{
    [XmlRoot("ModelExport", Namespace = "http://wwww.csadd-ins.com", IsNullable = false)]
    public class ModelExportViewModel : ViewModel
    {
        private ContentViewModel _mainvm;

        public ModelExportViewModel()
        {

        }

        public ModelExportViewModel(ContentViewModel mainvm)
        {
            this._mainvm = mainvm;

        }
        #region Set Value
        public void SetModelSetting(List<ModelSetting> modelSettings)
        {
            this.ModelSettings.Clear();
            this.DisplayModelSetting.Clear();
            modelSettings.ForEach(item => {

                this.ModelSettings.Add(item);
                this.DisplayModelSetting.Add(item);

            });
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
                foreach (ModelSetting modelSetting in this.DisplayModelSetting)
                {
                    modelSetting.Selected = this.AllSelected;
                }

            }
        }
        #endregion
        

        #region ModelOptionSetting
        private ModelOptionSetting modelOptionSetting = new ModelOptionSetting();
        public ModelOptionSetting ModelOptionSetting
        {
            get { return this.modelOptionSetting; }
            set
            {
                if (Equals(value, this.modelOptionSetting))
                {
                    return;
                }
                this.modelOptionSetting = value;
                OnPropertyChanged(nameof(ModelOptionSetting));
            }
        }
        #endregion


        #region ModelSettings
        [XmlIgnore]
        public ObservableCollection<ModelSetting> ModelSettings = new ObservableCollection<ModelSetting>();
        private ObservableCollection<ModelSetting> displayModelSetting = new ObservableCollection<ModelSetting>();

        [XmlIgnore]
        public ObservableCollection<ModelSetting> DisplayModelSetting
        {
            get { return this.displayModelSetting; }
            set
            {
                if (Equals(value, this.displayModelSetting))
                {
                    return;
                }
                this.displayModelSetting = value;
                OnPropertyChanged(nameof(DisplayModelSetting));
            }

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
            dlg.IsFolderPicker = false;
            CommonFileDialogFilter commonFileDialogFilter = new CommonFileDialogFilter("Excel file", ".xlsx");
            CommonFileDialogFilter commonFileDialogFilterOther = new CommonFileDialogFilter("Excel file", ".xls");
            dlg.Filters.Add(commonFileDialogFilter);
            dlg.Filters.Add(commonFileDialogFilterOther);

            dlg.Title = (string)(Application.Current.Properties["mainUserControl"] as UserControl).Resources["multilang_message_03"];
            
            if (dlg.ShowDialog() == MSAPI::Dialogs.CommonFileDialogResult.Ok)
            {
                sResult = dlg.FileName;
                ModelOptionSetting newModelOptionSetting = this.ModelOptionSetting.Clone();
                newModelOptionSetting.ExcelTemplate = sResult;
                this.ModelOptionSetting = newModelOptionSetting;
            }

        }
        #endregion

        #region Copy LDM to clipboard
        RelayCommand _CopyLdmToClipboard = null;
        public ICommand CopyLdmToClipboard
        {
            get
            {
                if (_CopyLdmToClipboard == null)
                {
                    _CopyLdmToClipboard = new RelayCommand(
                        () => OnCopyLdmToClipboard(), 
                        () => {

                            return this.DisplayModelSetting.Count(item => item.Selected) > 0;
                        });
                }

                return _CopyLdmToClipboard;
            }
        }

        private void OnCopyLdmToClipboard()
        {
            StringBuilder stringBuilder = new StringBuilder();


            stringBuilder.AppendLine($"\"实体编码\",\"实体名称\",\"属性名称\",\"属性代码\",\"数据类型\",\"是否主键\",\"是否非空\"");
            // select
            var selectedList = from item in this.DisplayModelSetting
                               where item.Selected
                               select item;

            foreach (ModelSetting setting in selectedList)
            {

                stringBuilder.AppendLine($"===== {setting.Diagram} =====");

                setting.ModelReportRecord.ForEach(item=> {

                    stringBuilder.AppendLine(item.ToCsvString(this.ModelOptionSetting.IsUpperCase));
                });
                
                
            }
            Clipboard.SetText(stringBuilder.ToString());
        }
        #endregion

        #region Export LDM Excel
        RelayCommand _exportLdmExcelCommand = null;
        public ICommand ExportLdmExcelCommand
        {
            get
            {
                if (_exportLdmExcelCommand == null)
                {
                    _exportLdmExcelCommand = new RelayCommand<string>(
                        (p) => {

                            var selectedList = from item in this.DisplayModelSetting
                                               where item.Selected
                                               select item;

                            // 重名处理，重名时加上No.
                            foreach(ModelSetting item in selectedList)
                            {
                                int count = selectedList.Count(value => value.DiagramUnique.Equals(item.DiagramUnique,System.StringComparison.OrdinalIgnoreCase));
                                if (count  > 1)
                                {
                                    item.DiagramUnique = item.DiagramUnique + $"_{count}";
                                }
                                else
                                {
                                    item.DiagramUnique = item.DiagramUnique;
                                }

                                int countPackage = selectedList.Count(value => value.PackageNameUnique.Equals(item.PackageNameUnique, System.StringComparison.OrdinalIgnoreCase));
                                if (countPackage > 1)
                                {
                                    item.PackageNameUnique = item.PackageNameUnique + $"_{countPackage}";
                                }
                                else
                                {
                                    item.PackageNameUnique = item.PackageNameUnique;
                                }
                            }


                            this._mainvm.userControl.FireModelExportEvent(this.modelOptionSetting,selectedList.ToList());
                        }, 
                        (p) => {
                            return System.IO.File.Exists(this.ModelOptionSetting.ExcelTemplate);                            
                        });
                }

                return _exportLdmExcelCommand;
            }
        }
        #endregion
    }
}
