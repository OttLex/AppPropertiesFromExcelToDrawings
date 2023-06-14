using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using Tekla.Structures.Model;

namespace AppPropertiesFromExcelToDrawings
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Model _model;
        private List<WorkDockRow> _workDockRows;

        public string FilePath { get; set; } = "";


        public Model Model { get => _model; set => _model = value; }



        public MainWindow()
        {
            _model = new Model();
            InitializeComponent();
        }

        #region На будующее открыть и сохранить файл через диалог.


        public bool OpenFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Exel (*.xlsx;)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            //C:\Users\oav\source\repos\CalculateBridgeService\CalculateBridgeService\template.xlsx


            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                return true;
            }
            return false;
        }
        public bool SaveFileDialog()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() == true)
            {
                FilePath = saveFileDialog.FileName;
                return true;
            }
            return false;
        }
        #endregion


        private string GetFilesNamesInFolder()
        {
            DirectoryInfo dir = new DirectoryInfo(_model.GetInfo().ModelPath); //Assuming Test is your Folder

            List<FileInfo> Files = dir.GetFiles("*.xlsx").ToList(); //Getting Text files

            string rowFileName = "";
            string actualFileName = "";
            string path = "";
            try
            {
                rowFileName = Files.Where(f => f.Name.ToLower().Contains("состав рабочей документации")).First().Name;
                actualFileName = rowFileName.Trim(new Char[] { '~', '$' });
                path = dir.FullName + @"\" + actualFileName;
            }
            catch
            {

            }
            return path;
        }

        private void openExelButton_Click(object sender, RoutedEventArgs e)
        {
            string pathToFile = GetFilesNamesInFolder();

            if (pathToFile == "")
            {
                MessageBox.Show("Ошибка. Файл \"Состав рабочей документации.xlsx\" не найден!");
                return;
            }




            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(pathToFile))
            {
                var sheet = package.Workbook.Worksheets[0];
                _workDockRows = GetList(sheet);
            }


        }

        private List<WorkDockRow> GetList(ExcelWorksheet sheet)
        {
            int startRow = 2;
            int startColumn = 1;
            List<WorkDockRow> list = new List<WorkDockRow>();

            for (int rowIndex = startRow; rowIndex < sheet.Dimension.Rows + 1; rowIndex++)
            {
                var row = sheet.Cells[rowIndex, startColumn, rowIndex, sheet.Dimension.Columns].Value as Array;

                WorkDockRow Row = new WorkDockRow(row, rowIndex);


                list.Add(Row);
            }

            return list;
        }
    }
}
