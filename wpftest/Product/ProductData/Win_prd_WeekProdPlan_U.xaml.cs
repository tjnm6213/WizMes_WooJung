/**
 * 
 * @details 주간생산계획 작성
 * @author 정승학
 * @date 2019-07-30
 * @version 1.0
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * - 2000-01-01    - 정승학       : -----
 * 
 * 
 * */

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
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
    /// Win_prd_WeekProdPlan_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_WeekProdPlan_U : UserControl
    {
        Win_prd_WeekProdPlan_U_CodeView WinWeekProdPlan = new Win_prd_WeekProdPlan_U_CodeView();
        Lib lib = new Lib();

        public Win_prd_WeekProdPlan_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpYearMonth.SelectedDate = DateTime.Today;
        }

                
        #region 일자변경
        //기간

        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            //if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            //else { chkDate.IsChecked = true; }
        }

        //기간
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            //if (dtpYearMonth != null)
            //{
            //    dtpYearMonth.IsEnabled = true;
            //}
        }

        //기간
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            //dtpYearMonth.IsEnabled = false;
        }
        #endregion

        #region 상단 레이아웃 활성화 & 비활성화
        //품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }
        }

        //품명
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnPfArticle.IsEnabled = true;
            txtArticle.Focus();
        }

        //품명
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnPfArticle.IsEnabled = false;
        }

        //품명 → 품번으로 변경(GLS)
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }

        //품명 → 품번으로 변경(GLS)
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
        }

        //거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //거래처
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
            txtCustom.Focus();
        }

        //거래처
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        #endregion

        #region 주간생산계획생성 버튼
        //주간생산계획생성
        private void btnWeekPlan_Click(object sender, RoutedEventArgs e)
        {
            if(MessageBox.Show("주간 생산 계획을 생성하시겠습니까?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
                MessageBox.Show("생성 취소하였습니다.");
            }
            else
            {
                //MessageBox.Show("YES일때");
                if (AutoWeekPlan(dtpYearMonth.SelectedDate != null ? dtpYearMonth.SelectedDate.Value.ToString("yyyyMM") : DateTime.Today.ToString("yyyyMM")))
                {
                    FillGrid();
                }
            }



        }
        #endregion

        #region 주간생산계획생성
        //
        private bool AutoWeekPlan(string strYearMonth)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("YYYYMM", strYearMonth);
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_WeekPlan_AutoiWeeklyPlan", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
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

            return flag;
        }
        #endregion

        #region 우측 상단 버튼
        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                Lib.Instance.UiButtonEnableChange_SCControl(this);
                btnWeekPlan.IsEnabled = false;
            }
            else
            {
                MessageBox.Show("수정할 자료가 없습니다.");
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData())
            {
                Lib.Instance.UiButtonEnableChange_IUControl(this);
                btnWeekPlan.IsEnabled = true;
                FillGrid();
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            btnWeekPlan.IsEnabled = true;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "작업지시 주간계획";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;

                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
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
            }
        }
        #endregion

        #region 조회
        //
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("YYYYMM", dtpYearMonth.SelectedDate != null ? dtpYearMonth.SelectedDate.Value.ToString("yyyyMM") : DateTime.Today.ToString("yyyyMM"));
                sqlParameter.Add("nchkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticle.IsChecked == true && txtArticle.Tag != null ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("nchkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustom.IsChecked == true && txtCustom.Tag != null ? txtCustom.Tag.ToString() : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_WeekPlan_sWeeklyPlan", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 1)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinWeekPlan = new Win_prd_WeekProdPlan_U_CodeView()
                            {
                                cls = dr["cls"].ToString(),
                                Num = i,
                                KCustom = dr["KCustom"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                maxChasu = dr["maxChasu"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                PlanQty1 = dr["PlanQty1"].ToString(),
                                PlanQty2 = dr["PlanQty2"].ToString(),
                                PlanQty3 = dr["PlanQty3"].ToString(),
                                PlanQty4 = dr["PlanQty4"].ToString(),
                                PlanQty5 = dr["PlanQty5"].ToString(),
                                PlanQtySum = dr["PlanQtySum"].ToString(),
                                flagUpdate = false,
                            };

                            WinWeekPlan.PlanQty1 = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQty1);
                            WinWeekPlan.PlanQty2 = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQty2);
                            WinWeekPlan.PlanQty3 = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQty3);
                            WinWeekPlan.PlanQty4 = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQty4);
                            WinWeekPlan.PlanQty5 = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQty5);
                            WinWeekPlan.PlanQtySum = Lib.Instance.returnNumStringZero(WinWeekPlan.PlanQtySum);

                            if (WinWeekPlan.maxChasu.Equals("4"))
                            {
                                dgdtpe5Week.Visibility = Visibility.Collapsed;
                            }
                            else if (WinWeekPlan.maxChasu.Equals("5"))
                            {
                                dgdtpe5Week.Visibility = Visibility.Visible;
                            }

                            dgdMain.Items.Add(WinWeekPlan);
                        }

                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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
        #endregion

        #region 저장
        /// <summary>
        /// 저장
        /// </summary>
        /// <returns></returns>
        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (dtpYearMonth.SelectedDate == null)
                {
                    MessageBox.Show("상단의 기간 날짜를 선택해주세요.");
                    return false;
                }

                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    WinWeekProdPlan = dgdMain.Items[i] as Win_prd_WeekProdPlan_U_CodeView;

                    if (WinWeekProdPlan.flagUpdate)
                    {
                        int count = 0;
                        count = int.Parse(WinWeekProdPlan.maxChasu);

                        for (int j = 0; j < count; j++)
                        {
                            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("YYYYMM", dtpYearMonth.SelectedDate.Value.ToString("yyyyMM"));
                            sqlParameter.Add("CustomID", WinWeekProdPlan.CustomID);
                            sqlParameter.Add("ArticleID", WinWeekProdPlan.ArticleID);
                            sqlParameter.Add("UserID", MainWindow.CurrentUser);

                            if (j == 0)
                            {
                                sqlParameter.Add("WeekChaSu", 1);
                                sqlParameter.Add("PlanQty", Lib.Instance.CheckNullZero(WinWeekProdPlan.PlanQty1.Replace(",", "")));
                            }
                            if (j == 1)
                            {
                                sqlParameter.Add("WeekChaSu", 2);
                                sqlParameter.Add("PlanQty", Lib.Instance.CheckNullZero(WinWeekProdPlan.PlanQty2.Replace(",", "")));
                            }
                            if (j == 2)
                            {
                                sqlParameter.Add("WeekChaSu", 3);
                                sqlParameter.Add("PlanQty", Lib.Instance.CheckNullZero(WinWeekProdPlan.PlanQty3.Replace(",", "")));
                            }
                            if (j == 3)
                            {
                                sqlParameter.Add("WeekChaSu", 4);
                                sqlParameter.Add("PlanQty", Lib.Instance.CheckNullZero(WinWeekProdPlan.PlanQty4.Replace(",", "")));
                            }
                            if (j == 4)
                            {
                                sqlParameter.Add("WeekChaSu", 5);
                                sqlParameter.Add("PlanQty", Lib.Instance.CheckNullZero(WinWeekProdPlan.PlanQty5.Replace(",", "")));
                            }

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_WeekPlan_uWeeklyPlan";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sMCID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);
                        }
                    }
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
                }
                else
                {
                    flag = true;
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

            return flag;
        }
        #endregion

        #region 데이터 그리드 키보드 이동관련
        //
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
            int startColCount = dgdMain.Columns.IndexOf(dgdtpe1Week);
            int endColCount = dgdMain.Columns.IndexOf(dgdtpe5Week);
            int colCount = dgdMain.Columns.IndexOf(dgdMain.CurrentCell.Column);
            int rowCount = dgdMain.Items.IndexOf(dgdMain.CurrentItem);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endColCount == colCount && dgdMain.Items.Count - 1 > rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount + 1], dgdMain.Columns[startColCount]);
                }
                else if (endColCount > colCount && dgdMain.Items.Count - 1 > rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                }
                else if (endColCount == colCount && dgdMain.Items.Count - 1 == rowCount)
                {
                    btnSave.Focus();
                }
                else if (endColCount > colCount && dgdMain.Items.Count - 1 == rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                }
                else
                {
                    MessageBox.Show("있으면 찾아보자...");
                }
            }
        }

        //
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }
        #endregion

        #region 데이터 그리드 셀 데이터 변경
        //
        private void Text_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumericOnly((TextBox)sender, e);
        }

        //1주차
        private void dgdtpetxtPlanQty1_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
                WinWeekProdPlan.flagUpdate = true;

                if (WinWeekProdPlan != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinWeekProdPlan.PlanQty1 = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                    }

                    sender = tb1;
                }
            }
        }

        //2주차
        private void dgdtpetxtPlanQty2_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
                WinWeekProdPlan.flagUpdate = true;

                if (WinWeekProdPlan != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinWeekProdPlan.PlanQty2 = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                    }

                    sender = tb1;
                }
            }
        }

        //3주차
        private void dgdtpetxtPlanQty3_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
                WinWeekProdPlan.flagUpdate = true;

                if (WinWeekProdPlan != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinWeekProdPlan.PlanQty3 = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                    }

                    sender = tb1;
                }
            }
        }

        //4주차
        private void dgdtpetxtPlanQty4_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
                WinWeekProdPlan.flagUpdate = true;

                if (WinWeekProdPlan != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinWeekProdPlan.PlanQty4 = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                    }

                    sender = tb1;
                }
            }
        }

        //5주차
        private void dgdtpetxtPlanQty5_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                WinWeekProdPlan = dgdMain.CurrentItem as Win_prd_WeekProdPlan_U_CodeView;
                WinWeekProdPlan.flagUpdate = true;

                if (WinWeekProdPlan != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        WinWeekProdPlan.PlanQty5 = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                    }

                    sender = tb1;
                }
            }
        }

        #endregion

    }

    #region CodeView
    class Win_prd_WeekProdPlan_U_CodeView
    {
        public int Num { get; set; }
        public string cls { get; set; }
        public string CustomID { get; set; }
        public string ArticleID { get; set; }
        public string maxChasu { get; set; }

        public string KCustom { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string PlanQty1 { get; set; }
        public string PlanQty2 { get; set; }
        public string PlanQty3 { get; set; }

        public string PlanQty4 { get; set; }
        public string PlanQty5 { get; set; }
        public string PlanQtySum { get; set; }
        public bool flagUpdate { get; set; }
    }

    #endregion
}
