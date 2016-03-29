using System;
using System.Collections.Generic;
using System.IO;
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

namespace Login_Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int Counter = 1;
        public List<Person> ListOfPeople;
        bool click;
        Brush brush;
        List<CheckBox> lstCHK;
        public MainWindow()
        {

            click = true;
            InitializeComponent();
            ListOfPeople = new List<Person>();
            lstListPeople.ItemsSource = ListOfPeople;
         //commented out for testing
         //   WindowState = System.Windows.WindowState.Maximized;
            WindowStyle = WindowStyle.None;
            brush = txtLastName.BorderBrush;
            lstCHK = new List<CheckBox>();
            lstCHK.Add(chkInfo);
            lstCHK.Add(chkAFCOOL);
            lstCHK.Add(chkCCAF);
            lstCHK.Add(chkCommission);
            lstCHK.Add(chkEdLevel);
            lstCHK.Add(chkInOut);
            lstCHK.Add(chkPME);
            lstCHK.Add(chkTA);
            lstCHK.Add(chkVA);
            lstCHK.Add(chkWithdraw);

        }


        public void Method(Object sender, ExecutedRoutedEventArgs e)
        {
            if (click)
            {
                Login.Visibility = System.Windows.Visibility.Hidden;
                Admin.Visibility = System.Windows.Visibility.Visible;
                click = !click;
            }
            else
            {
                Admin.Visibility = System.Windows.Visibility.Hidden;
                Login.Visibility = System.Windows.Visibility.Visible;
                click = !click;
            }

        }

        private void Submit(object sender, RoutedEventArgs e)
        {
            bool blnProblems = false;
            if (txtFirstName.Text == "")
            {

                txtFirstName.BorderBrush = new SolidColorBrush(Colors.Red);
                blnProblems = true;
            }

            if (txtLastName.Text == "")
            {

                txtLastName.BorderBrush = new SolidColorBrush(Colors.Red);
                blnProblems = true;
            }

            if (blnProblems == false)
            {

                string sst = "First Name: " + txtFirstName.Text + Environment.NewLine +
                            "Last Name: " + txtLastName.Text + Environment.NewLine +
                            "Grade: " + cboGrade.Text + Environment.NewLine;


                if (MessageBoxResult.Yes == MessageBox.Show(sst, "Is this correct?", MessageBoxButton.YesNo))
                {

                    List<string> lst = new List<string>();
                    foreach (CheckBox chk in lstCHK)
                    {
                        if (chk.IsChecked == true) { lst.Add(chk.Content.ToString()); }

                    }
                    Person newpep = new Person { id = Counter, Appt = cboAppt.Text, Helped = "", FirstName = txtFirstName.Text, LastName = txtLastName.Text, Email = txtEmail.Text, Branch = cboBranch.Text, Grade = cboGrade.Text, Status = cboStatus.Text, ReasonsForVisit = lst };
                    ListOfPeople.Add(newpep);
                    lstListPeople.ItemsSource = null;
                    Counter++;
                    ListOfPeople = ListOfPeople.OrderBy(a => a.id).ToList();
                    lstListPeople.ItemsSource = ListOfPeople;
                    txtLastName.BorderBrush = brush;
                    txtFirstName.BorderBrush = brush;
                    //Reorders list

                    //Reset form
                    foreach (CheckBox chk in lstCHK)
                    {
                        chk.IsChecked = false;
                    }
                    txtFirstName.Text = "";
                    txtLastName.Text = "";
                    txtEmail.Text = "";
                    cboBranch.Text = "";
                    cboGrade.Text = "";
                    cboStatus.Text = "";
                    cboAppt.Text = "";
                }
            }
        }

        private void SaveToExcelandClose(object sender, RoutedEventArgs e)
        {

            if (MessageBoxResult.Yes == MessageBox.Show("Program will close and generate Excel document.", "Are you sure?", MessageBoxButton.YesNo))
            {

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Fatal error: no Excel installed on system.");
                    return;
                }
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                //Column Labeling
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[2, 1] = "Staff";
                xlWorkSheet.Cells[2, 2] = "Name, L";
                xlWorkSheet.Cells[2, 3] = "Name, F";
                xlWorkSheet.Cells[2, 4] = "Grade";
                xlWorkSheet.Cells[2, 5] = "Branch";
                xlWorkSheet.Cells[2, 6] = "Status";
                xlWorkSheet.Cells[2, 7] = "Email";
                xlWorkSheet.Cells[2, 8] = "Reasons";

                //Column sizing
                Microsoft.Office.Interop.Excel.Range FormatRange = xlWorkSheet.UsedRange.Columns;
                Microsoft.Office.Interop.Excel.Range s = FormatRange.Columns[1];
                FormatRange.Columns[1].ColumnWidth = 5.71;
                FormatRange.Columns[2].ColumnWidth = 25;
                FormatRange.Columns[3].ColumnWidth = 25;
                FormatRange.Columns[4].ColumnWidth = 8;
                FormatRange.Columns[5].ColumnWidth = 11.14;
                FormatRange.Columns[6].ColumnWidth = 19.29;
                FormatRange.Columns[7].ColumnWidth = 41;
                FormatRange.Columns[8].ColumnWidth = 40;

                int Counter = 3;
                foreach (Person p in ListOfPeople)
                {
                    xlWorkSheet.Cells[Counter, 1] = p.Helped;
                    xlWorkSheet.Cells[Counter, 2] = p.LastName;
                    xlWorkSheet.Cells[Counter, 3] = p.FirstName;
                    xlWorkSheet.Cells[Counter, 4] = p.Grade;
                    xlWorkSheet.Cells[Counter, 5] = p.Branch;
                    xlWorkSheet.Cells[Counter, 6] = p.Status;
                    xlWorkSheet.Cells[Counter, 7] = p.Email;
                    xlWorkSheet.Cells[Counter, 8] = p.Reasons;
                    Counter++;
                }
                string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\EduLog_" + String.Format("{0:MM-dd-yy}", DateTime.Now) + ".xls";
                xlWorkBook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);

                this.Close();



            }



        }

        private void Help(object sender, RoutedEventArgs e)
        {
            if (txtIntial.Text == "")

            {txtIntial.BorderBrush = new SolidColorBrush(Colors.Red);
                txtIntial.Focus();
            }
            else {
                txtIntial.BorderBrush = brush;

                var sen = (Button)sender;
                sen.Visibility = Visibility.Hidden;
                Person per = (Person)sen.DataContext;
                Person found = ListOfPeople.Find(a => a.id == per.id);
                found.id += 1000;
                found.Helped = txtIntial.Text;
                ListOfPeople = ListOfPeople.OrderBy(a => a.id).ToList();

                lstListPeople.ItemsSource = null;
                lstListPeople.ItemsSource = ListOfPeople;
            }
        

        }




    }
    public class Person
    {
        public int id;
        public string Helped { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Grade { get; set; }
        public string Branch { get; set; }
        public string Status { get; set; }
        public string Appt { get; set; }
        public Visibility Vis { get { if (id > 1000) { return Visibility.Hidden; } return Visibility.Visible; } }
        public Visibility RVis { get { if (Vis == Visibility.Hidden) { return Visibility.Visible; } return Visibility.Hidden; } }
        public string Reasons { get { string str = ""; foreach (string s in ReasonsForVisit) { str += s + ", "; } if (str.Count() < 3) { return ""; } else { return str.Substring(0, str.Count() - 2); } } }
        public List<string> ReasonsForVisit { get; set; }
    }
}
