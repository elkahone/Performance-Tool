using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Collections.Generic;
using Window = System.Windows.Window;

namespace Performance_Tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string constring = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\kasia\source\repos\Performance_Tool\Performance_Tool\DatabaseTickets.mdf;Integrated Security = True; Integrated Security=True";
        //creatig array list for combobox
        string[] dllist = new string[]{
                                "(DL) Server monitoring",
                                "(DL) Adding/removing mailbox permissions",
                                "(DL) Distribution list amendment",
                                "(DL) Shared mailbox creation / deletion",
                                "(DL) SMTP contact amendment",
                                "(DL) Calendar permissions",
                                "(DL) Mailbox extension",
                                "(DL) User Mailbox Disabling",
                                "(DL) Non-Standard request",
                                "(DL) Other" };
        string[] isrlist = new string[]{
                                "(DL) ISR Wintel AD"};
        string[] vwlist = new string[]{
                                "(VW) S4B iPhone/iPad device activation",
                                "(VW) Ticket update",
                                "(VW) Troubleshooting",
                                "(VW) Voice activation",
                                "(VW) Lync/S4B Activation",
                                "(VW) Other"};
        string[] crhlist = new string[]{
                                "(CRH) Call",
                                "(CRH) Major incident",
                                "(CRH) LAPS password"};

        string[] byklist = new string[]{
                                "(BYK) User creation",
                                "(BYK) Details change in AD",
                                "(BYK) Disabling account",
                                "(BYK) Enabling account",
                                "(BYK) Changing ownerships",
                                "(BYK) Adding/removing from/to groups in AD",
                                "(BYK) Renaming account",
                                "(BYK) Double-checking of ticket",
                                "(BYK) Non-Standard request",
                                "(BYK) Other"};
        public MainWindow()
        {
            InitializeComponent();
            MainMethod();

            //first check if Date exist in DataBase if not add new object
            SqlConnection con = new SqlConnection(constring);
            con.Open();
            SqlCommand check_date = new SqlCommand("SELECT COUNT(*) FROM [Tickets] WHERE ([Date] ='" + DateBox.Text + "')", con);
            int dateexist = (int)check_date.ExecuteScalar();
            if (dateexist > 0)
            {
                con.Close();
            }
            else
            {
                con.Close();
                UpdateSQLBase();
            }

        }

        //update DateBox, ProjectBox and ShiftBox
        private void MainMethod()
        {
            DateBox.Text = DateTime.Today.ToString("dd.MM.yyyy");
            string[] projectlist = new string[] { "BYK", "DL", "ISR", "VW", "CRH" };
            ProjectBox.ItemsSource = projectlist;
            TimeSpan time = DateTime.Now.TimeOfDay;
            int hour = time.Hours;

            if (hour >= 19)
            {
                ShiftBox.Text = "Night/WE";
            }
            ShiftBox.Text = "Day";
        }

        //uploading array list  to comnobox depending of the selected project
        void ProjectBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectBox.SelectedItem.Equals("BYK"))
            {
                SubBox.ItemsSource = byklist;
            }
            if (ProjectBox.SelectedItem.Equals("DL"))
            {
                SubBox.ItemsSource = dllist;
            }
            if (ProjectBox.SelectedItem.Equals("VW"))
            {
                SubBox.ItemsSource = vwlist;
            }
            if (ProjectBox.SelectedItem.Equals("CRH"))
            {
                SubBox.ItemsSource = crhlist;
            }
            if (ProjectBox.SelectedItem.Equals("ISR"))
            {
                SubBox.ItemsSource = isrlist;
            }
        }

        // performance file update
        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            ExcelUpdate(DateTime.Today.ToString("yyyy.MM.dd"), SubBox.Text, RefBox.Text, ShiftBox.Text, ComentBox.Text);
            ComentBox.Clear();
            TicketUpdate(ProjectBox.Text);
            Log(DateBox.Text, SubBox.Text, RefBox.Text, ComentBox.Text);
            //short break before chars will updated
            System.Threading.Thread.Sleep(2000);
            showChart(DateBox.Text);

        }

                                                         

        //upload every new ticket(in according to the projects) to database
        private void TicketUpdate(string data1)
        {
            SqlConnection con = new SqlConnection(constring);
            SqlCommand cmd;
           
            if (ProjectBox.Text.Equals(data1))
            {
                try
                {
                    con.Open();
                    string selectsql = "select * from Tickets where Date ='"+DateBox.Text+"'";
                    cmd = new SqlCommand(selectsql, con);
                    SqlDataReader rd = cmd.ExecuteReader();
                    if (rd.Read())
                    {
                        int value = int.Parse(rd[data1].ToString());
                        con.Close();
                        rd.Close();
                        value = value + 1;
                        con.Open();
                        string sql = "update Tickets SET " + data1 + "=@val1 where Date=@val2";
                        cmd = new SqlCommand(sql, con);
                        cmd.Parameters.AddWithValue("@val1", value);
                        cmd.Parameters.AddWithValue("@val2", DateBox.Text);
                        cmd.ExecuteNonQuery();

                    }
                    con.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        //upload toady date to DataBase
        private void UpdateSQLBase()
        {
            SqlConnection con = new SqlConnection(constring);
            SqlCommand cmd;

                try
                {
                    con.Open();
                    string sql = "insert into Tickets (Date, DL, ISR, VW, BYK, CRH) values (@val, @val2, @val2, @val2, @val2, @val2)";
                    cmd = new SqlCommand(sql, con);
                    int nul = 0;
                    cmd.Parameters.AddWithValue("@val", DateTime.Today.ToString("dd.MM.yyyy"));
                    cmd.Parameters.AddWithValue("@val2", nul);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

        }


        //update excel file method
        private static void ExcelUpdate(string data1, string data2, string data3, string data4, string data5)
        {
            int row = 5000;
            string password = "1234";
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Workbook oWB = null;
            Worksheet oSheet = null;
            Range range = null;
            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(@"\\g02.fujitsu.local\DFS\LDZT\Groups\LDZ_SharedDesk_PL\1. SD Agent\1.5 Trainings\Amelia Latacz\Team Performance\Lukasz Kucharski Performance Tasks DL VW BYK CRH.xlsm", 0, false, 5, password);
                oSheet = oWB.ActiveSheet;


                range = oSheet.Cells[row, 1].EntireRow;


                while (row > 0)
                {
                    row++;
                    if (range.Cells[row, 2].Value == null)
                    {
                        oSheet.Cells[row, 2] = "Kucharski Łukasz";
                        oSheet.Cells[row, 3] = data1;
                        oSheet.Cells[row, 4] = data2;
                        oSheet.Cells[row, 5] = data3;
                        oSheet.Cells[row, 6] = data4;
                        oSheet.Cells[row, 8] = data5;
                        break;
                    }

                }


                oWB.Save();
                MessageBox.Show("Done!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                    oWB.Close();
            }
        }

        private void showChart(string date1)
        {
            SqlConnection con = new SqlConnection(constring);
            SqlCommand cmd;
            try
            {
                con.Open();
                string selectsql = "select * from Tickets where Date ='" + date1 + "'";
                cmd = new SqlCommand(selectsql, con);
                SqlDataReader rd = cmd.ExecuteReader();
                if (!rd.Read())
                {
                    List<KeyValuePair<string, int>> MyValue = new List<KeyValuePair<string, int>>();
                    MyValue.Add(new KeyValuePair<string, int>("No Data", 0));
                    ColumnChart1.DataContext = MyValue;
                }
                else
                {
                    int value1 = int.Parse(rd["DL"].ToString());
                    int value2 = int.Parse(rd["ISR"].ToString());
                    int value3 = int.Parse(rd["CRH"].ToString());
                    int value4 = int.Parse(rd["BYK"].ToString());
                    int value5 = int.Parse(rd["VW"].ToString());
                    con.Close();
                    List<KeyValuePair<string, int>> MyValue = new List<KeyValuePair<string, int>>();
                    MyValue.Add(new KeyValuePair<string, int>("DL - "+value1, value1));
                    MyValue.Add(new KeyValuePair<string, int>("ISR - "+value2, value2));
                    MyValue.Add(new KeyValuePair<string, int>("CRH - "+value3, value3));
                    MyValue.Add(new KeyValuePair<string, int>("BYK - " + value4, value4));
                    MyValue.Add(new KeyValuePair<string, int>("VW - "+value5, value5));
                    ColumnChart1.DataContext = MyValue;
                }
                rd.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            showChart(DatePickerBox.Text);
        }

        private void Log(string date1, string sub1, string ref1, string com1)
        {
            List<Ticket> items = new List<Ticket>();
            items.Add(new Ticket() { Date = date1, SubType = sub1, Ref = ref1, Coment = com1 });
            LogBox.Items.Add(items);
        }

        public class Ticket
        {
            public string Date { get; set; }

            public string SubType { get; set; }

            public string Ref { get; set; }

            public string Coment { get; set; }
        }

    }
}
