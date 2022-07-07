using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
    /// Win_Qul_DateProd_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DateProd_Q : UserControl
    {
        int rowNum = 0;
        Lib lib = new Lib(); 
        PlusFinder pf = new PlusFinder();

        public Win_Qul_DateProd_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);

            //콤보박스 셋팅       
            SetComboBox();

            //입고일자 체크
            chkDate.IsChecked = true;

            //공정 체크 해제
            chkProcess.IsChecked = false;

            //품명 체크 해제
            chkArticle.IsChecked = false;

            //데이트피커 오늘 날짜
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;

            //콤보박스 기본값 '전체'
            cboProcess.SelectedIndex = 0;

            //조건 박스 false
            cboProcess.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }

        //콤보박스 셋팅
        private void SetComboBox()
        {
            //공정
            ObservableCollection<CodeView> cboProcessGroup = ComboBoxUtil.Instance.GetWorkProcess(0, "");
            this.cboProcess.ItemsSource = cboProcessGroup;
            this.cboProcess.DisplayMemberPath = "code_name";
            this.cboProcess.SelectedValuePath = "code_id";
        }

        #region 클릭 이벤트

        //입고일자 라벨 클릭 이벤트
        private void LblchkDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true)
            {
                chkDate.IsChecked = false;
                dtpSDate.IsEnabled = false;
                dtpEDate.IsEnabled = false;
            }
            else
            {
                chkDate.IsChecked = true;
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //입고일자 체크 이벤트
        private void ChkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //입고일자 체크해제 이벤트
        private void ChkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }



        //공정 라벨 클릭 이벤트
        private void LblProcess_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProcess.IsChecked == true)
            {
                chkProcess.IsChecked = false;
                cboProcess.IsEnabled = false;
            }
            else
            {
                chkProcess.IsChecked = true;
                cboProcess.IsEnabled = true;
            }
        }

        //공정 체크 이벤트
        private void ChkProcess_Checked(object sender, RoutedEventArgs e)
        {
            cboProcess.IsEnabled = true;
        }

        //공정 체크 해제 이벤트
        private void ChkProcess_Unchecked(object sender, RoutedEventArgs e)
        {
            cboProcess.IsEnabled = false;
        }

        //품명
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
        }
        //품명
        private void chkArticle_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
                txtArticle.Focus();
            }
        }

        //품명 체크 이벤트
        private void ChkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }

        //품명 체크 해제 이벤트
        private void ChkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }


        #region 플러스파인더

        // 플러스파인더 _ 품명 찾기
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 77, txtArticle.Text);
        }

        // 품명 키다운 _ 품명 찾기
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 77, txtArticle.Text);
            }
        }
      
        // 플러스파인더 _ 품번 찾기
        private void btnArticleNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
        }

        // 품번 키다운 
        private void TxtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
            }
        }

        //품번
        private void chkArticleNo_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                txtArticleNo.IsEnabled = true;
                txtArticleNo.Focus();
                btnArticleNo.IsEnabled = true;
            }
            else
            {
                txtArticleNo.IsEnabled = false;
                btnArticleNo.IsEnabled = false;
            }
        }

        //품번
        private void chkArticleNo_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                chkArticleNo.IsChecked = false;
                txtArticleNo.IsEnabled = false;
                btnArticleNo.IsEnabled = false;
            }
            else
            {
                chkArticleNo.IsChecked = true;
                txtArticleNo.IsEnabled = true;
                btnArticleNo.IsEnabled = true;
                txtArticleNo.Focus();
            }
        }
        #endregion

        #endregion 클릭이벤트, 날짜

        #region CRUD 버튼

        //검색(조회)
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                if (CheckData())
                {
                    re_Search(rowNum);
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "공정불량현황";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.Check.Equals("Y"))
                    dt = lib.DataGridToDTinHidden(dgdMain);
                else
                    dt = lib.DataGirdToDataTable(dgdMain);

                Name = dgdMain.Name;

                if (lib.GenerateExcel(dt, Name))
                {
                    lib.excel.Visible = true;
                    lib.ReleaseExcelObject(lib.excel);
                }
                else
                    return;
            }
            else
            {
                if (dt != null)
                {
                    dt.Clear();
                }
            }
            lib = null;
        }

        #endregion CRUD 버튼


        #region 데이터그리드 이벤트

        //데이터그리드 셀렉션체인지드
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //조회만 하는 화면이라 이 친구는 필요가 없지요.
        }

        #endregion 데이터그리드 이벤트

        #region 조회관련(Fillgrid)

        //재조회
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
        }

        //조회
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                int chkArticleID = 0;
                string ArticleID = "";

                if (chkArticle.IsChecked == true)
                {
                    chkArticleID = 1;
                    ArticleID = txtArticle.Tag.ToString();
                }
                if (chkArticleNo.IsChecked == true)
                {
                    chkArticleID = 1;
                    ArticleID = txtArticleNo.Tag.ToString();
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("chkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("eDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkProcess", chkProcess.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProcessID", chkProcess.IsChecked == true ? cboProcess.SelectedValue.ToString() : "");

                sqlParameter.Add("ChkArticle", chkArticleID);
                sqlParameter.Add("ArticleID", ArticleID);
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true && !txtArticle.Text.Trim().Equals("") ?  txtArticle.Text : "");
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true && !txtArticleNo.Text.Trim().Equals("") ? txtArticleNo.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sInspectDefect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var DefectInfo = new Win_Qul_DateProd_Q_CodeView()
                            {
                                Num = i,
                                cls = dr["cls"].ToString(),
                                ScanDate = dr["ScanDate"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                DefectID = dr["DefectID"].ToString(),
                                KDefect = dr["KDefect"].ToString(),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                WorkPersonName = dr["WorkPersonName"].ToString(),
                                MCNAME = dr["MCNAME"].ToString()
                                //LabelID = dr["LabelID"].ToString(),
                                //ChildLabelID = dr["ChildLabelID"].ToString()
                            };

                            if ((DefectInfo.ScanDate != "" && DefectInfo.ScanDate != null))
                            {
                                  DefectInfo.ScanDate = DefectInfo.ScanDate.ToString().Substring(0, 4) + "-"
                                + DefectInfo.ScanDate.ToString().Substring(4, 2) + "-"
                                + DefectInfo.ScanDate.ToString().Substring(6, 2);
                            }

                            

                            if (DefectInfo.DefectQty.Equals("") && DefectInfo.cls.Equals("9"))
                            {
                                MessageBox.Show("조회결과가 없습니다.");
                                return;
                            }

                            if (DefectInfo.cls.Equals("9"))
                            {
                                dgdTotal.Items.Clear();
                                DefectInfo.ScanDate = "총 발생수량";
                                //DefectInfo.ColorLightLightGray = "false";
                                DefectInfo.ColorGold = "true";
                                dgdTotal.Items.Add(DefectInfo);
                            }
                            else
                            {
                                dgdMain.Items.Add(DefectInfo);
                            }
                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //검색 조건 Check
        private bool CheckData()
        {
            bool flag = true;

            if (chkArticle.IsChecked == true)
            {
                if(txtArticle.Text == "")
                {
                    MessageBox.Show("품번 선택이 되지 않았습니다. 체크를 해제하거나 품번을 선택하고 검색해 주세요.");
                    flag = false;
                    return flag;
                }
            }

            return flag;
        }


        #endregion 조회관련(Fillgrid)
        
        #region 기타 메소드 
        //특수문자 포함 검색
        private string Escape(string str)
        {
            string result = "";

            for (int i = 0; i < str.Length; i++)
            {
                string txt = str.Substring(i, 1);

                bool isSpecial = Regex.IsMatch(txt, @"[^a-zA-Z0-9가-힣]");

                if (isSpecial == true)
                {
                    result += (@"/" + txt);
                }
                else
                {
                    result += txt;
                }
            }
            return result;
        }

        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
        #endregion

        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }
    }

    #region 생성자들(CodeView)

    class Win_Qul_DateProd_Q_CodeView : BaseView
    {
        public int Num { get; set; }
        public string cls { get; set; }
        public string ScanDate { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string BuyerModelID { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string DefectID { get; set; }
        public string KDefect { get; set; }
        public string DefectQty { get; set; }
        public string WorkPersonID { get; set; }
        public string WorkPersonName { get; set; }
        public string MCNAME { get; set; }

        //public string LabelID { get; set; }
        //public string ChildLabelID { get; set; }

        public string ColorLightLightGray { get; set; }

        public string ColorGold { get; set; }
    }

    #endregion 생성자들(CodeView)
}