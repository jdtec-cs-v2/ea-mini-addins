using System;
using System.ComponentModel;

namespace MiniAddins.Models
{
    public enum ExportState
    {
        NoExport=0,
        Already,
        HasChange
    }

    public class ChangeLog
    {
        public CustomSetting LogContent { get; set; }
        public string LogType { get; set; }
        public DateTime AtTime { get; set; }

        public string DisplayString {
            get
            {
                return $"{LogType} [{LogContent.Diagram}] diagram in [{LogContent.PackageName}] package  at {AtTime.ToString("yyyy-MM-dd HH:mm:ss")}";

            }
        }

    }

    public class CustomSetting : INotifyPropertyChanged
    {
        private bool _selected;
        private string _relativePath;
        private ExportState _exportState = ExportState.NoExport;

        public string Diagram { get; set; }
        public string DiagramGUID { get; set; }
        public string PackageName { get; set; }
        public string PackageFullName { get; set; }
        public string PackageGUID { get; set; }

        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        
        public ExportState ExportState
        {
            get { return this._exportState; }
            set
            {
                if (Equals(value, this._exportState))
                {
                    return;

                }
                this._exportState = value;
                OnPropertyChanged(nameof(ExportState));
            }
        }


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

        public string RelativePath {
            get { return this._relativePath; }
            set
            {
                if (Equals(value, this._relativePath))
                {
                    return;

                }
                this._relativePath = value;
                OnPropertyChanged(nameof(RelativePath));
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
