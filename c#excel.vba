using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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


namespace PreToPostMadeEasy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            openExcel();
        }

        private void openExcel()
        {
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"%USERPROFILE%\Desktop";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            fdlg.ShowDialog();
            fname = fdlg.FileName;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            /*
             * 
             * DataGridColumn column = new DataGridTextColumn();

                column.Header = "lol";

                excelViewer.Columns.Add(column);
             * 
             * */
            for (int i = 1; i <= colCount; i++)
            {
                DataGridColumn column = new DataGridTextColumn();
                column.Header = "lol";
                excelViewer.Columns.Add(column);
            }

            DataGridRow gf = new DataGridRow();

            excelViewer.Items.Add("MOOOOOOOOOOOOOOOOOO");



            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }
    }
}





 <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="100"/>
            <ColumnDefinition Width="100"/>
            <ColumnDefinition Width="100"/>
            <ColumnDefinition Width="100"/>
            <ColumnDefinition/>
            <ColumnDefinition/>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>

        <Grid.RowDefinitions>
            <RowDefinition Height="25"/>
            <RowDefinition/>
            <RowDefinition/>
        </Grid.RowDefinitions>

        <Label Content="Status" />
        <Button x:Name="btnOpen" Content="Open" Grid.Column="2" Click="btnOpen_Click" />
        <Button Content="Generate" Grid.Column="3" />
        <Button Content="Send" Grid.Column="4" />
        <ListView Grid.RowSpan="2" Grid.Row="1">
            <ListView.View>
                <GridView>
                    <GridViewColumn/>
                </GridView>
            </ListView.View>
        </ListView>
        <DataGrid x:Name="excelViewer" Grid.Column="1" Grid.Row="1" Grid.ColumnSpan="7" Grid.RowSpan="2"/>

    </Grid>