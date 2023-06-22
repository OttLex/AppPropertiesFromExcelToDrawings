using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Drawing = Tekla.Structures.Drawing.Drawing;


namespace AppPropertiesFromExcelToDrawings
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Model _model;
        private List<WorkDockRow> _workDockRows;
        private DrawingHandler _CurrentDrawingHandler;
        private List<Drawing> _drawings;
        private string _filePath;
        private Stopwatch _stopwatch;
        List<string> _currentDrawings;

        private List<Exception> _excelException;
        private List<Exception> _drawingException;

        public string FilePath { get => _filePath; set => _filePath = value; }
        public Model Model { get => _model; set => _model = value; }
        public DrawingHandler CurrentDrawingHandler { get => _CurrentDrawingHandler; set => _CurrentDrawingHandler = value; }
        public List<Drawing> Drawings { get => _drawings; set => _drawings = value; }

        public MainWindow()
        {
            Model = new Model();
            CurrentDrawingHandler = new DrawingHandler();
            _stopwatch = new Stopwatch();
            InitializeComponent();


        }

        private void ExecuteOperations()
        {
            GetDataFromExcel();
            GetDrawingsFromTekla();
            UpdateDrawings();
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

        private bool ValidateExcel()
        {
            FilePath = GetFilesNamesInFolder();

            if (FilePath == "")
            {
                MessageBox.Show("Ошибка. Файл \"Состав рабочей документации.xlsx\" не найден!");
                return false;
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            try
            {
                using (ExcelPackage package = new ExcelPackage(FilePath))
                {
                    package.Save();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }


        private void openExelButton_Click(object sender, RoutedEventArgs e)
        {
            GetDataFromExcel();
        }
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
        private void GetDataFromExcel()
        {
            _stopwatch = Stopwatch.StartNew();
            FilePath = GetFilesNamesInFolder();

            if (FilePath == "")
            {
                MessageBox.Show("Ошибка. Файл \"Состав рабочей документации.xlsx\" не найден!");
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(FilePath))
                {
                    var sheet = package.Workbook.Worksheets[0];
                    _workDockRows = GetList(sheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }

            Debug.WriteLine("Данные из экселя выгружены: " + _stopwatch.Elapsed);
        }
        private List<WorkDockRow> GetList(ExcelWorksheet sheet)
        {
            int startRow = 3;
            int startColumn = 1;
            List<WorkDockRow> list = new List<WorkDockRow>();

            for (int rowIndex = startRow; rowIndex <= sheet.Dimension.Rows + 1; rowIndex++)
            {
                var row = sheet.Cells[rowIndex, startColumn, rowIndex, sheet.Dimension.Columns].Value as Array;

                WorkDockRow Row = new WorkDockRow(row, rowIndex);

                list.Add(Row);

            }

            return list;
        }



        private void openDrawingsButton_Click(object sender, RoutedEventArgs e)
        {
            GetDrawingsFromTekla();
        }
        private void GetDrawingsFromTekla()
        {
            try
            {
                DrawingEnumerator drawingsEnum = CurrentDrawingHandler.GetDrawings();

                if (drawingsEnum != null)
                {
                    Drawings = new List<Drawing>();

                    foreach (Drawing Drawing in drawingsEnum)
                    {
                        Drawings.Add(Drawing);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }

            Debug.WriteLine("Чертежи выгружены: " + _stopwatch.Elapsed);
        }



        private void updateDrawingsButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateDrawings();
        }

        private void UpdateDrawings()
        {
            _currentDrawings = GetDrawingsContainsExcelRow().Select(d=>d.Title2).ToList();
            List<Drawing> currentDrawings= GetDrawingsContainsExcelRow();
            List<Drawing> updatedDrawings;
            RewriteDrawingsData(currentDrawings, out updatedDrawings);

            UpdateDrawingsTekla(ref updatedDrawings);
        }

        private List<Drawing> GetDrawingsContainsExcelRow()
        {

            var validDataRows = Drawings.Where(d =>
                                                    _workDockRows
                                                    .Where(w => w.IsValidName == true)
                                                    .Select(w => w.Id)
                                                    .ToList()
                                                    .Contains(d.Title2)).ToList();

            Debug.WriteLine("Сопоставлены: "+_stopwatch.Elapsed);
            return validDataRows;
        }

        private void RewriteDrawingsData(List<Drawing> currentDrawings, out List<Drawing> updatedDrawings)
        {
            foreach (Drawing drawing in currentDrawings)
            {
                WorkDockRow curRow = _workDockRows.Where(r => r.Id == drawing.Title2).FirstOrDefault();

                drawing.Title1 = curRow.KitCode;
                drawing.Title2 = curRow.KitName;
                string dateDeveloped = curRow.Date;
                drawing.SetUserProperty("DR_ASSIGNED_TO", dateDeveloped);
            }

            updatedDrawings = currentDrawings;
        }

        private void UpdateDrawingsTekla(ref List<Drawing> drawings)
        {
            foreach (Drawing drawing in drawings)
            {
                drawing.Modify();
            }

            Debug.WriteLine("Чертежи обновлены: " + _stopwatch.Elapsed);
            UpdateExcel();
        }

        private void UpdateExcel()
        {
            using (ExcelPackage package = new ExcelPackage(FilePath))
            {
                var sheet = package.Workbook.Worksheets[0];

                int startRow = 3;

                for (int rowIndex = startRow; rowIndex <= sheet.Dimension.Rows + 1; rowIndex++)
                {
                    string id = sheet.GetValue(rowIndex, 1).ToString();
                    WorkDockRow currentRow = _workDockRows.Where(r => r.Id == id 
                                                                   && r.IsValidName==true).FirstOrDefault();

                    if (currentRow != null)
                    {
                        if (_currentDrawings.Where(d => d == currentRow.Id).Count() > 0)
                        {
                            sheet.Cells[rowIndex, sheet.Dimension.Columns].Clear();
                            sheet.Cells[rowIndex, sheet.Dimension.Columns].Value = "+";
                        }
                    }
                }

                package.Save();
                Debug.WriteLine("Эксель обновлён: " + _stopwatch.Elapsed);
            }
        }

        private void ExecuteOperationsButton_Click(object sender, RoutedEventArgs e)
        {

            if (ValidateExcel())
            {
                ExecuteOperations();
            }
            else
            {
                MessageBox.Show("Excel файл должен быть закрыт!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error); ExecuteOperations();
            }
        }
    }
}
