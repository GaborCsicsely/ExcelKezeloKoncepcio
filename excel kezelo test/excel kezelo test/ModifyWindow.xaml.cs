using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;

namespace excel_kezelo_test
{
    /// <summary>
    /// Interaction logic for ModifyWindow.xaml
    /// </summary>
    public partial class ModifyWindow : Window
    {
        private DataRow dataRow;
        private MainWindow mainWindow;

        public ModifyWindow(DataRow row, MainWindow mainWindow)
        {
            InitializeComponent();
            dataRow = row;
            this.mainWindow = mainWindow;
            CreateFields();
        }

        private void CreateFields()
        {
            foreach (DataColumn column in dataRow.Table.Columns)
            {
                StackPanel stackPanel = new StackPanel { Orientation = Orientation.Horizontal };
                TextBlock textBlock = new TextBlock
                {
                    Text = column.ColumnName,
                    Width = 100,
                    Margin = new Thickness(5)
                };
                TextBox textBox = new TextBox
                {
                    Name = column.ColumnName,
                    Text = dataRow[column.ColumnName].ToString(),
                    Width = 200,
                    Margin = new Thickness(5)
                };

                stackPanel.Children.Add(textBlock);
                stackPanel.Children.Add(textBox);
                FieldsPanel.Children.Add(stackPanel);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var child in FieldsPanel.Children)
            {
                if (child is StackPanel stackPanel)
                {
                    foreach (var innerChild in stackPanel.Children)
                    {
                        if (innerChild is TextBox textBox)
                        {
                            dataRow[textBox.Name] = textBox.Text;
                        }
                    }
                }
            }
            mainWindow.SaveToExcel();
            this.DialogResult = true;
            this.Close();
        }
    }
}
