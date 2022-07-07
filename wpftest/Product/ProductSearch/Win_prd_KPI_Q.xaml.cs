using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WooJung.PopUP;

namespace WizMes_WooJung
{
    /// <summary>
    /// Win_prd_KPI_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_KPI_Q : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        int rowNum = 0;

        public Win_prd_KPI_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            DatePickerStartDateSearch.SelectedDate = DateTime.Today;
            DatePickerEndDateSearch.SelectedDate = DateTime.Today;
        }

        #region 상단 검색조건
        //전년
        private void ButtonLastYear_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (DatePickerStartDateSearch.SelectedDate != null)
                {
                    DatePickerStartDateSearch.SelectedDate = DatePickerStartDateSearch.SelectedDate.Value.AddYears(-1);
                }
                else
                {
                    DatePickerStartDateSearch.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //전월
        private void ButtonLastMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (DatePickerStartDateSearch.SelectedDate != null)
                {
                    DateTime FirstDayOfMonth = DatePickerStartDateSearch.SelectedDate.Value.AddDays(-(DatePickerStartDateSearch.SelectedDate.Value.Day - 1));
                    DateTime FirstDayOfLastMonth = FirstDayOfMonth.AddMonths(-1);

                    DatePickerStartDateSearch.SelectedDate = FirstDayOfLastMonth;
                }
                else
                {
                    DateTime FirstDayOfMonth = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));

                    DatePickerStartDateSearch.SelectedDate = FirstDayOfMonth;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //금년
        private void ButtonThisYear_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (DatePickerStartDateSearch.SelectedDate != null)
                {
                    DatePickerStartDateSearch.SelectedDate = lib.BringThisYearDatetimeFormat()[0];
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //금월
        private void ButtonThisMonth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DatePickerStartDateSearch.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #endregion

        #region Re_Search
        private void re_Search(int selectedIndex)
        {
            try
            {
                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }

                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

                FillGrid();

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        #endregion

        #region 공수조회
        private void FillGrid()
        {
            try
            {
                if (dgdOut.Items.Count > 0)
                {
                    dgdOut.Items.Clear();
                }
                if (dgdGonsu.Items.Count > 0)
                {
                    dgdGonsu.Items.Clear();
                }

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("FromDate", DatePickerStartDateSearch.SelectedDate == null ? "" : DatePickerStartDateSearch.SelectedDate.Value.ToString().Replace("-", ""));
                sqlParameter.Add("ToDate", DatePickerEndDateSearch.SelectedDate == null ? "" : DatePickerEndDateSearch.SelectedDate.Value.ToString().Replace("-", ""));
                sqlParameter.Add("ArticleNo", chkArticleNo.IsChecked == true && txtArticleNoSearch.Tag != null ? txtArticleNoSearch.Tag.ToString() : ""); //품번
                ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sKPI_KPI", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WPKQC = new Win_prd_KPI_Q_CodeView()
                            {
                                Num = i + 1,

                                GbnName = dr["GbnName"].ToString(),
                                ArticleNo = dr["ARTICLENO"].ToString(),
                                Article = dr["article"].ToString(),
                                WorkQty = Convert.ToDouble(dr["WorkQty"].ToString()),
                                WorkTime = lib.returnNumStringZero(dr["WorkTime"].ToString()),
                                WorkQtyPerHour = Convert.ToDouble(dr["WorkQtyPerHour"].ToString()),
                                WorkMan = dr["WorkMan"].ToString(),
                                Gonsu = dr["Gonsu"].ToString(),
                                OrderQty = dr["OrderQty"].ToString(),
                                DiffOutDate = dr["DiffOutDate"].ToString(),
                                DiffOutDayPerQty = dr["DiffOutDayPerQty"].ToString(),
                                DefectQty = Convert.ToDouble(dr["DefectQty"].ToString()),
                                DefectWorkQty = Convert.ToDouble(dr["DefectWorkQty"].ToString()),
                                DefectRate = dr["DefectRate"].ToString(),
                                gbn = dr["gbn"].ToString(),
                                Sort = dr["Sort"].ToString(),
                            };

                            WPKQC.Gonsu = lib.returnNumStringZero(WPKQC.Gonsu);
                            WPKQC.OrderQty = lib.returnNumStringZero(WPKQC.OrderQty);
                            WPKQC.DiffOutDayPerQty = lib.returnNumStringZero(WPKQC.DiffOutDayPerQty);
                            //WPKQC.DefectRate = lib.returnNumStringOne(WPKQC.DefectRate);
                            WPKQC.WorkTime = lib.returnNumString(WPKQC.WorkTime);

                            if (WPKQC.gbn == "Q")
                            {
                                dgdOut.Items.Add(WPKQC);
                            }

                            if (WPKQC.gbn == "P")
                            {
                                dgdGonsu.Items.Add(WPKQC);
                            }

                            i++;
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                rowNum = 0;
                re_Search(rowNum);

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btiClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                lib.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void btiExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //if(dgdOut.Items.Count == 0 && dgdGonsu.Items.Count == 0)
                //{
                //    MessageBox.Show("먼저 검색해 주세요.");
                //    return;
                //}

                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = "KPI작업공수";
                lst[1] = "KPI납기";
                lst[2] = dgdGonsu.Name;
                lst[3] = dgdOut.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdGonsu.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdGonsu);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdGonsu);

                        Name = dgdGonsu.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(dgdOut.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdOut);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdOut);

                        Name = dgdOut.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else
                    {
                        if (dt != null)
                        {
                            dt.Clear();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        private void lblArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                chkArticleNo.IsChecked = false;
            }
            else
            {
                chkArticleNo.IsChecked = true;
            }
        }
        // 품번 체크박스 이벤트
        private void chkArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleNo.IsChecked = true;
            txtArticleNoSearch.IsEnabled = true;
            btnArticleNoSearch.IsEnabled = true;
        }
        private void chkArticleNo_UnChecked(object sender, RoutedEventArgs e)
        {
            chkArticleNo.IsChecked = false;
            txtArticleNoSearch.IsEnabled = false;
            btnArticleNoSearch.IsEnabled = false;
        }
        // 품번 텍스트박스 엔터 → 플러스파인더
        private void txtArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleNoSearch, 76, "");
            }
        }
        // 품번 플러스파인더 이벤트
        private void btnArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            // 거래처 : 0
            MainWindow.pf.ReturnCode(txtArticleNoSearch, 76, "");
        }

        //품명 라벨 클릭
        private void LabelArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if(CheckBoxArticleSearch.IsChecked == true)
            {
                CheckBoxArticleSearch.IsChecked = false;
            }
            else
            {
                CheckBoxArticleSearch.IsChecked = true;
            }
        }

        private void CheckBoxArticleSearch_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxArticleSearch.IsEnabled = true;
            ButtonArticleSearch.IsEnabled = true;
        }

        private void CheckBoxArticleSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            TextBoxArticleSearch.IsEnabled = false;
            ButtonArticleSearch.IsEnabled = false;
        }

        private void TextBoxArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if(e.Key == Key.Enter)
                {
                    pf.ReturnCode(TextBoxArticleSearch, 77, TextBoxArticleSearch.Text);
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void ButtonArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(TextBoxArticleSearch, 77, TextBoxArticleSearch.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        
    }

    #region CodeView
    class Win_prd_KPI_Q_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public string GbnName { get; set; }
        public string ArticleNo { get; internal set; }
        public string Article { get; internal set; }
        public double WorkQty { get; internal set; }
        public string WorkTime { get; internal set; }
        public double WorkQtyPerHour { get; internal set; }
        public string WorkMan { get; set; }
        public string Gonsu { get; set; }
        public string OrderQty { get; set; }
        public string DiffOutDate { get; set; }
        public string DiffOutDayPerQty { get; set; }
        public double DefectQty { get; set; }
        public double DefectWorkQty { get; set; }
        public string DefectRate { get; set; }
        public string gbn { get; set; }
        public string Sort { get; set; }


    }

    #endregion

}