using MiniAddins.Command;
using MiniAddins.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


namespace MiniAddins.ViewModels
{
    public class StatisticsViewModel : ViewModel
    {
        private ContentViewModel _mainvm;

        public StatisticsViewModel(ContentViewModel mainvm)
        {
            _mainvm = mainvm;

#if DEBUG

#endif
        }

        #region ElementType Combobox
        private ObservableCollection<CheckItem> elementTypes = new ObservableCollection<CheckItem>();
        public ObservableCollection<CheckItem> ElementTypes
        {
            get { return this.elementTypes; }

        }
        #endregion


        #region Author Combobox
        private ObservableCollection<CheckItem> authors = new ObservableCollection<CheckItem>();
        public ObservableCollection<CheckItem> Authors
        {
            get { return this.authors; }

        }
        #endregion


        #region Result List
        private ObservableCollection<StatisticsResult>  statisticsResults = new ObservableCollection<StatisticsResult>();
        public ObservableCollection<StatisticsResult>  StatisticsResults 
        {
            get { return this.statisticsResults; }            

        }
        public List<StatisticsResult> OrgStatisticsResults = new List<StatisticsResult>();
        #endregion


        #region Save to Clipboard Command
        RelayCommand _saveClipboardCommand = null;
        public ICommand SaveClipboardCommand
        {
            get
            {
                if (_saveClipboardCommand == null)
                {
                    _saveClipboardCommand = new RelayCommand(() => OnSaveClipboardCommand(), () => true);


                }

                return _saveClipboardCommand;
            }
        }

        private void OnSaveClipboardCommand()
        {
            
            Clipboard.SetText(GetStatisticsResultCSV());
            return;
        }
        #endregion


        #region Save to File Command
        RelayCommand _saveToFileCommand = null;
        public ICommand SaveToFileCommand
        {
            get
            {
                if (_saveToFileCommand == null)
                {
                    _saveToFileCommand = new RelayCommand(() => OnSaveToFileCommand(), () => true);


                }

                return _saveToFileCommand;
            }
        }
        
        private void OnSaveToFileCommand()
        {
            string sResult = string.Empty;
            
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "ea report";
            dlg.DefaultExt = ".txt"; 
            dlg.Filter = "ea report (.txt)|*.txt";
            
            dlg.Title = (string)(Application.Current.Properties["mainUserControl"] as UserControl).Resources["multilang_message_02"];

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;
                StreamWriter sw = new StreamWriter(filename);
                sw.Write(GetStatisticsResultCSV());
                sw.Close();
            }

        }

        /// <summary>
        /// output CSV format file 
        /// </summary>
        /// <returns></returns>
        private string GetStatisticsResultCSV()
        {

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"{nameof(StatisticsResult.ElementType)},{nameof(StatisticsResult.Author)},{nameof(StatisticsResult.Count)}");
            foreach (StatisticsResult item in this.StatisticsResults)
            {
                stringBuilder.AppendLine($"{item.ElementType},{item.Author},{item.Count}");
            }

            return stringBuilder.ToString();
        }
        #endregion


        #region All Type Click Command
        RelayCommand _allTypeCommand = null;
        public ICommand AllTypeCommand
        {
            get
            {
                if (_allTypeCommand == null)
                {
                    _allTypeCommand = new RelayCommand(() => {

                        this._mainvm.userControl.FireAllTypesClickEvent();

                    }, () => true);
                }

                return _allTypeCommand;
            }
        }


        #endregion


        #region  Check Click Command
        RelayCommand _checkClickCommand = null;
        public ICommand CheckClickCommand
        {
            get
            {
                if (_checkClickCommand == null)
                {
                    _checkClickCommand = new RelayCommand(() => RefreshResultList(), () => true);
                }

                return _checkClickCommand;
            }
        }

        /// <summary>
        /// Reset Result list
        /// </summary>
        private void RefreshResultList()
        {
            // Refresh List
            this.statisticsResults.Clear();

            var types = from item in this.ElementTypes
                        where item.IsChecked
                        select item.Name;

            var authors = from item in this.Authors
                          where item.IsChecked
                          select item.Name;


            var list = from item in this.OrgStatisticsResults
                       where types.Contains(item.ElementType)
                       where authors.Contains(item.Author)
                       select item;

            foreach (StatisticsResult item in list)
            {
                this.statisticsResults.Add(item);
            }
        }
        #endregion
               
    }
}
