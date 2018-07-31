using IPG_zad2.Infrastructure;
using IPG_zad2.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace IPG_zad2.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public ExcelDataModel excelDataModel = new ExcelDataModel();

        public DataGrid averagePriceDT;
        public DataGrid excelDT;

        private int currentSheetNumber = 0;
        private string _currentSheetName = "Załaduj plik";

        public string CurrentSheetName { get => _currentSheetName; set {
                _currentSheetName = value;
                PropertyChanged(this, new PropertyChangedEventArgs("currentSheetName")); } }
        public List<LevelAveragePrice> levelAveragePriceListVM;

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
                currentSheetNumber = 0;
                CurrentSheetName = excelDataModel.fileDataModel.SheetList[0].SheetName;
                BindAveragePriceLevelDG(currentSheetNumber);
                BindExcelDG(currentSheetNumber);
            }

        }, true);

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindowViewModel(DataGrid avg, DataGrid ex)
        {
            averagePriceDT = avg;
            excelDT = ex;
            currentSheetNumber = 0;
        }

        private void BindAveragePriceLevelDG (int sheetNum)
        {
            averagePriceDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Poziom",
                Binding = new Binding("Level")
            });
            averagePriceDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Średnia cena na dzień",
                Binding = new Binding("AveragePrice")
            });
            averagePriceDT.ItemsSource = excelDataModel.fileDataModel.SheetList[sheetNum].AveragePricePerLevel;
        }

        private void BindExcelDG (int sheetNum)
        {
            //TO DO
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header= "Nazwa",
                Binding= new Binding("Name")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "ID",
                Binding = new Binding("Id")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Cena",
                Binding = new Binding("Price")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Pozycja",
                Binding = new Binding("Position")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Poziom",
                Binding = new Binding("Level")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Opis",
                Binding = new Binding("Description")
            });
            excelDT.Columns.Add(new DataGridTextColumn
            {
                Header = "Nr Zamówienia",
                Binding = new Binding("Order")
            });
        }
    }
}
