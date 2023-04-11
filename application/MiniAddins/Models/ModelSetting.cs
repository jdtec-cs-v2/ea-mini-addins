using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml.Serialization;

namespace MiniAddins.Models
{
    public class ModelOptionSetting
    {
        public String ExcelTemplate { get; set; }
        public bool IsUpperCase { get; set; }
        public bool IsDivdedSheet { get; set; }
        public bool OriginalSize { get; set; }
        public bool CustomizeSize { get; set; }
        public int RowMultiple { get; set; }

        public ModelOptionSetting Clone()
        {
            return new ModelOptionSetting() {
                ExcelTemplate = ExcelTemplate,
                IsUpperCase =IsUpperCase,
                IsDivdedSheet=IsDivdedSheet,
                OriginalSize = OriginalSize,
                CustomizeSize = CustomizeSize,
                RowMultiple =RowMultiple
            };
        }

        public ModelOptionSetting()
        {
            IsUpperCase = true;
            IsDivdedSheet = false;
            OriginalSize = true;
            CustomizeSize = !OriginalSize;
            RowMultiple = 10;
        }
    }

    /// <summary>
    /// Model report record.
    /// </summary>
    public class ModelReportRecordDto
    {
        public string TableGUID { get; set; }
        public string TableNote { get; set; }
        public string TableName { get; set; }
        public string FieldName { get; set; }
        public string FieldType { get; set; }
        public string FieldNote { get; set; }
        public string FieldIsId { get; set; }

        public string ToCsvString(bool Uppercase)
        {
            string sRtn =  $"\"{TableNote}\",\"{TableName}\",\"{FieldName}\",\"{FieldNote}\",\"{FieldType}\",\"{FieldIsId}\",\"{FieldIsId}\"";
            if (Uppercase)
            {
                return sRtn.ToUpper();
            }
            else
            {
                return sRtn;
            }            
        }
    }

    public class ModelSetting : INotifyPropertyChanged
    {
        private bool _selected;
        public string Diagram { get; set; }
        public string DiagramUnique { get; set; }
        public string DiagramGUID { get; set; }
        public string PackageName { get; set; }
        public string PackageNameUnique { get; set; }
        public string PackageFullName { get; set; }
        public string PackageGUID { get; set; }

        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public List<ModelReportRecordDto> ModelReportRecord { get; set; }


        public bool Selected
        {
            get { return this._selected; }
            set
            {
                if (Equals(value, this._selected))
                {
                    return;

                }
                this._selected = value;
                OnPropertyChanged(nameof(Selected));
            }
        }

        // Declare the event
        public event PropertyChangedEventHandler PropertyChanged;
        // Create the OnPropertyChanged method to raise the event
        protected void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            handler?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
          
    
}
