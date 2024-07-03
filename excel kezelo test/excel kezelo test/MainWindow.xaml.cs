using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace excel_kezelo_test
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTable dataTable;
        private string currentFilePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenTableButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls";
            if (openFileDialog.ShowDialog() == true)
            {
                currentFilePath = openFileDialog.FileName;
                LoadExcel(currentFilePath);
            }
        }

        private void LoadExcel(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var range = worksheet.RangeUsed();
                dataTable = new DataTable();

                foreach (var headerCell in range.FirstRow().Cells())
                {
                    dataTable.Columns.Add(NormalizeColumnName(headerCell.GetValue<string>()));
                }

                foreach (var dataRow in range.RowsUsed().Skip(1))
                {
                    var newRow = dataTable.NewRow();
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        newRow[i] = dataRow.Cell(i + 1).GetValue<string>();
                    }
                    dataTable.Rows.Add(newRow);
                }

                DataGrid.ItemsSource = dataTable.DefaultView;
            }
        }

        private string NormalizeColumnName(string columnName)
        {
            string normalized = columnName.Normalize(System.Text.NormalizationForm.FormD);
            var chars = normalized.Where(c => System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark).ToArray();
            normalized = new string(chars).Normalize(System.Text.NormalizationForm.FormC);
            return normalized.Replace(' ', '_');
        }

        public void SaveToExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = dataTable.Rows[i][j].ToString();
                    }
                }

                workbook.SaveAs(currentFilePath);
            }
        }

        private void SearchBox_KeyUp(object sender, KeyEventArgs e)
        {
            var filterText = SearchBox.Text;
            var dv = dataTable.DefaultView;
            dv.RowFilter = string.Format("Column1 LIKE '%{0}%'", filterText);
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var newRow = dataTable.NewRow();
            dataTable.Rows.Add(newRow);
            var modifyWindow = new ModifyWindow(newRow, this);
            modifyWindow.ShowDialog();
            DataGrid.Items.Refresh();
        }

        private void ModifyButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid.SelectedItem != null)
            {
                var selectedRow = ((DataRowView)DataGrid.SelectedItem).Row;
                var modifyWindow = new ModifyWindow(selectedRow, this);
                modifyWindow.ShowDialog();
                DataGrid.Items.Refresh();
            }
        }

        private void DeactivateButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataGrid.SelectedItem != null)
            {
                var selectedRow = ((DataRowView)DataGrid.SelectedItem).Row;
                DataGrid.Items.Refresh();
                SaveToExcel();
            }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
