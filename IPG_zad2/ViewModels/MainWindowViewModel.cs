using IPG_zad2.Infrastructure;
using IPG_zad2.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace IPG_zad2.ViewModels
{
    public class MainWindowViewModel: INotifyPropertyChanged
    {
        public ExcelDataModel excelDataModel = new ExcelDataModel();

        public ICommand ExitCommand => new CommandHandler(() => App.Current.Shutdown(), true);
        public ICommand OpenFileCommand => new CommandHandler(() =>
        {
            OpenFileDialog openFileDIalog = new OpenFileDialog();
            openFileDIalog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDIalog.ShowDialog() == true)
            {
                excelDataModel.ReadFile(openFileDIalog.FileName);
                foreach (SheetModel sheet in excelDataModel.fileDataModel.SheetList)
                    sheet.CalcAveragePricePerLevel();
            }
            
        }, true);

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
