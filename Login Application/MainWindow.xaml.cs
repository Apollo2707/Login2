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

        double thisDpiWidthFactor;
          double  thisDpiHeightFactor;


        public MainWindow()
        {

            click = true;
            InitializeComponent();
            ListOfPeople = new List<Person>();
            lstListPeople.ItemsSource = ListOfPeople;
         //commented out for testing
       // WindowState = System.Windows.WindowState.Maximized;
            WindowStyle = WindowStyle.None;
            brush = txtLastName.BorderBrush;
            lstCHK = new List<CheckBox>();


            Left = 0;
            Top = 0;
            ResizeMode = ResizeMode.NoResize;
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

            CalculateDpiFactors(this);
            double ScreenHeight = SystemParameters.PrimaryScreenHeight * thisDpiHeightFactor;
            double ScreenWidth = SystemParameters.PrimaryScreenWidth * thisDpiWidthFactor;
            Height = ScreenHeight;
            Width = ScreenWidth * 2;
            Login.Width = ScreenWidth;



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
                    if ((bool)chkCCAF.IsChecked) { newpep.ReasonCCAF = true; }
                    if ((bool)chkCommission.IsChecked) { newpep.ReasonCommission = true; }
                    if ((bool)chkEdLevel.IsChecked) { newpep.ReasonEdLevel = true; }
                    if ((bool)chkInfo.IsChecked) { newpep.ReasonGeneralInfo = true; }
                    if ((bool)chkInOut.IsChecked) { newpep.ReasonInOut = true; }
                    if ((bool)chkPME.IsChecked) { newpep.ReasonPME = true; }
                    if ((bool)chkTA.IsChecked) { newpep.ReasonTA = true; }
                    if ((bool)chkVA.IsChecked) { newpep.ReasonVA = true; }
                    if ((bool)chkAFCOOL.IsChecked) { newpep.ReasonAFCOOL = true; }
                    if ((bool)chkWithdraw.IsChecked) { newpep.ReasonWithdrawReimburse = true; }
                    if ((bool)chkOther.IsChecked) { newpep.ReasonOther = true; }


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

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
          //  if (MessageBoxResult.Yes == MessageBox.Show("Program will close and generate Excel document.", "Are you sure?", MessageBoxButton.YesNo))
          //  {

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                
                if (xlApp == null)
                {
                    MessageBox.Show("Fatal error: no Excel installed on system.");
                    return;
                }


                var newfile = Login_Application.Properties.Resources.Sign_In3;
                //save resource to disk
                
                string strPathToResource = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + @"\Sign_In3.xls";
                if (!File.Exists(strPathToResource))
                {
                    using (FileStream cFileStream = new FileStream(strPathToResource, FileMode.CreateNew))
                    {
                        cFileStream.Write(Properties.Resources.Sign_In3, 0, Properties.Resources.Sign_In3.Length);
                    }
                }

                //open workbook
                xlWorkBook = xlApp.Workbooks.Open(strPathToResource);
                //end of code
               object misValue = System.Reflection.Missing.Value;
                ////Column Labeling
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[5, 7] = String.Format("{0:dd-MMM-yy}", DateTime.Now);

                int Counter = 6;

                foreach (Person p in ListOfPeople)
                {
                    xlWorkSheet.Cells[Counter, 1] = p.Helped;
                    xlWorkSheet.Cells[Counter, 2] = p.LastName;
                    xlWorkSheet.Cells[Counter, 3] = p.FirstName;
                    xlWorkSheet.Cells[Counter, 4] = p.Grade;
                    xlWorkSheet.Cells[Counter, 5] = p.Branch;
                    xlWorkSheet.Cells[Counter, 6] = p.Status;
                    xlWorkSheet.Cells[Counter, 7] = p.Email;
                    if (p.Appt == "Yes") { xlWorkSheet.Cells[Counter, 8] = "X"; }
                    else { xlWorkSheet.Cells[Counter, 9] = "X"; }
                    if (p.ReasonCCAF == true) { xlWorkSheet.Cells[Counter, 10] = "X"; }
                    if (p.ReasonCommission == true) { xlWorkSheet.Cells[Counter, 11] = "X"; }
                    if (p.ReasonEdLevel == true) { xlWorkSheet.Cells[Counter, 12] = "X"; }
                    if (p.ReasonGeneralInfo == true) { xlWorkSheet.Cells[Counter, 13] = "X"; }
                    if (p.ReasonInOut == true) { xlWorkSheet.Cells[Counter, 14] = "X"; }
                    if (p.ReasonPME == true) { xlWorkSheet.Cells[Counter, 15] = "X"; }
                    if (p.ReasonTA == true) { xlWorkSheet.Cells[Counter, 16] = "X"; }
                    if (p.ReasonVA == true) { xlWorkSheet.Cells[Counter, 17] = "X"; }
                    if (p.ReasonAFCOOL == true) { xlWorkSheet.Cells[Counter, 18] = "X"; }
                    if (p.ReasonWithdrawReimburse == true) { xlWorkSheet.Cells[Counter, 19] = "X"; }
                    if (p.ReasonOther == true) { xlWorkSheet.Cells[Counter, 20] = "X"; }
                    Counter++;
                }
                string filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\EduLog_" + String.Format("{0:MM-dd-yy}", DateTime.Now) + ".xls";
                xlWorkBook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);

                this.Close();



           // }



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

        private void CalculateDpiFactors(Window t)
        {
            Window MainWindow = t;
            Window ms = new MessageBox();
            PresentationSource MainWindowPresentationSource = PresentationSource.FromVisual(MainWindow);
            Matrix m = MainWindowPresentationSource.CompositionTarget.TransformToDevice;
           thisDpiWidthFactor = m.M11;
            thisDpiHeightFactor = m.M22;
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
        public bool ReasonOther { get; set; }
        public bool ReasonCCAF { get; set; }
        public bool ReasonCommission { get; set; }
        public bool ReasonEdLevel { get; set; }
        public bool ReasonGeneralInfo { get; set; }
        public bool ReasonInOut { get; set; }
        public bool ReasonPME { get; set; }
        public bool ReasonTA { get; set; }
        public bool ReasonVA { get; set; }
        public bool ReasonAFCOOL { get; set; }
        public bool ReasonWithdrawReimburse { get; set; }
        public Visibility Vis { get { if (id > 1000) { return Visibility.Hidden; } return Visibility.Visible; } }
        public Visibility RVis { get { if (Vis == Visibility.Hidden) { return Visibility.Visible; } return Visibility.Hidden; } }
        public string Reasons { get { string str = ""; foreach (string s in ReasonsForVisit) { str += s + ", "; } if (str.Count() < 3) { return ""; } else { return str.Substring(0, str.Count() - 2); } } }
        public List<string> ReasonsForVisit { get; set; }
    }



}
