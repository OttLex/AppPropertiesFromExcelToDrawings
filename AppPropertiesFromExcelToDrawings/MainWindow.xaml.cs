﻿using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

        public string FilePath { get; set; } = "";


        public Model Model { get => _model; set => _model = value; }
        public DrawingHandler CurrentDrawingHandler
        {
            get { return this._CurrentDrawingHandler; }
            set { this._CurrentDrawingHandler = value; }
        }

        public List<Drawing> Drawings { get => _drawings; set => _drawings = value; }

        public MainWindow()
        {
            Model = new Model();
            CurrentDrawingHandler = new DrawingHandler();
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



        private void openDrawingsButton_Click(object sender, RoutedEventArgs e)
        {
            GetDrawingsFromTekla();
        }
        private void GetDrawingsFromTekla()
        {
            DrawingEnumerator drawingsEnum = CurrentDrawingHandler.GetDrawings();

            if (drawingsEnum != null)
            {
                Drawings = new List<Drawing>();

                foreach (Drawing Drawing in drawingsEnum)
                {
                    Drawings.Add(Drawing);
                }

                foreach (Drawing drawing in Drawings)
                {
                    string dateDeveloped = "";
                    drawing.GetUserProperty("DR_ASSIGNED_TO", ref dateDeveloped);


                    //CurrentDrawingHandler.UpdateDrawing(drawing);
                }
            }
        }

        private void updateDrawingsButton_Click(object sender, RoutedEventArgs e)
        {
            ValidateExcelData();
        }

        private List<WorkDockRow> ValidateExcelData()
        {

            var validDataRows = (List<WorkDockRow>)Drawings.Where(d =>                                                
                                                                        _workDockRows
                                                                        .Where(w => w.IsValidName == true)
                                                                        .Select(w=>w.Id)
                                                                        .ToList()
                                                                        .Contains(d.Title2));

            return validDataRows;
        }

    }
}
