using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using WizMes_WooJung.PopUP;
using WPF.MDI;
using System.ComponentModel;
using System.Windows.Input;
using System.Threading;

namespace WizMes_WooJung
{
    /// <summary>
    /// 
    /// </summary>
    public partial class Win_Qul_DateBox_Q : UserControl
    {
        Win_Qul_DateBox_QView WIDV = new Win_Qul_DateBox_QView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        // 엑셀 활용 용도 (프린트)


        WizMes_WooJung.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        //(기다림 알림 메시지창)

        System.Data.DataTable DT;

        public Win_Qul_DateBox_Q()
        {
            InitializeComponent();
            this.DataContext = WIDV;
        }

        private void Window_InsDateBox_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }

        #region 시작 첫단계 // 콤보박스 세팅 // 조회용 각종 체크박스 활성화 // 일자버튼

        // 시작 첫 단계.
        private void First_Step()
        {
            chkInspectDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;
            rbnManageNumber.IsChecked = true;

            
            // no check > no use.
            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
            txtOrderID.IsEnabled = false;
            cboFaultyGBN.IsEnabled = false;
            txtBoxID.IsEnabled = false;
            txtCID.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }




        //출고일자(날짜) 체크
        private void chkInspectDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
            {
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
            else
            {
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
        }
        //출고일자(날짜) 체크
        private void chkInspectDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
            {
                chkInspectDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkInspectDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //거래처
        private void chkCustomer_Click(object sender, RoutedEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
            else
            {
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
        }
        //거래처
        private void chkCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
        }        
        // 품명
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
        // 품명
        private void chkArticle_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
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
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
        }
        //관리번호 + order NO.
        private void chkOrderID_Click(object sender, RoutedEventArgs e)
        {
            if (chkOrderID.IsChecked == true)
            {
                txtOrderID.IsEnabled = true;
                txtOrderID.Focus();
            }
            else { txtOrderID.IsEnabled = false; }
        }
        //관리번호 + order NO.
        private void chkOrderID_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOrderID.IsChecked == true)
            {
                chkOrderID.IsChecked = false;
                txtOrderID.IsEnabled = false;
            }
            else
            {
                chkOrderID.IsChecked = true;
                txtOrderID.IsEnabled = true;
                txtOrderID.Focus();
            }
        }

        //불량구분
        private void chkFaultyGBN_Click(object sender, RoutedEventArgs e)
        {
            if (chkFaultyGBN.IsChecked == true)
            {
                cboFaultyGBN.IsEnabled = true;
                cboFaultyGBN.Focus();
            }
            else { cboFaultyGBN.IsEnabled = false; }
        }
        //불량구분
        private void chkFaultyGBN_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkFaultyGBN.IsChecked == true)
            {
                chkFaultyGBN.IsChecked = false;
                cboFaultyGBN.IsEnabled = false;
            }
            else
            {
                chkFaultyGBN.IsChecked = true;
                cboFaultyGBN.IsEnabled = true;
                cboFaultyGBN.Focus();
            }
        }
        //박스번호
        private void chkBoxID_Click(object sender, RoutedEventArgs e)
        {
            if (chkBoxID.IsChecked == true)
            {
                txtBoxID.IsEnabled = true;
                txtBoxID.Focus();
            }
            else { txtBoxID.IsEnabled = false; }
        }
        //박스번호
        private void chkBoxID_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkBoxID.IsChecked == true)
            {
                chkBoxID.IsChecked = false;
                txtBoxID.IsEnabled = false;
            }
            else
            {
                chkBoxID.IsChecked = true;
                txtBoxID.IsEnabled = true;
                txtBoxID.Focus();
            }
        }
        //박스번호
        private void chkCID_Click(object sender, RoutedEventArgs e)
        {
            if (chkCID.IsChecked == true)
            {
                txtCID.IsEnabled = true;
                txtCID.Focus();
            }
            else { txtCID.IsEnabled = false; }
        }
        //박스번호
        private void chkCID_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCID.IsChecked == true)
            {
                chkCID.IsChecked = false;
                txtCID.IsEnabled = false;
            }
            else
            {
                chkCID.IsChecked = true;
                txtCID.IsEnabled = true;
                txtCID.Focus();
            }
        }
        private void rbnOrderNO_Click(object sender, RoutedEventArgs e)
        {
            txbOrderID.Text = "Order NO";
        }

        private void rbnManageNumber_Click(object sender, RoutedEventArgs e)
        {
            txbOrderID.Text = "관리번호";
        }

        // 콤보박스 목록 불러오기.
        private void ComboBoxSetting()
        {
            cboFaultyGBN.Items.Clear();
            
            ObservableCollection<CodeView> cbFaultyGBN = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DFGRP", "Y", "", "");

            this.cboFaultyGBN.ItemsSource = cbFaultyGBN;
            this.cboFaultyGBN.DisplayMemberPath = "code_name";
            this.cboFaultyGBN.SelectedValuePath = "code_id";
            this.cboFaultyGBN.SelectedIndex = 0;
        }

        #endregion


        #region 플러스 파인더
        //플러스 파인더

        // 거래처
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");
        }

        // 품명(품번 검색으로 변경요청, 2020.03.23, 장가빈)
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 77, txtArticle.Text);
        }

        #endregion


        #region 조회 // 조회 프로시저

        // 검색. 조회버튼 클릭.
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

        private void FillGrid()
        {
                                        //날짜.
            string sDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string eDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");
            if (chkInspectDay.IsChecked == false)
            {
                sDate = "20000101";
                eDate = "29991230";
            }
            int nClss = 0;              //불량구분.
            if (chkFaultyGBN.IsChecked == true)
            {
                nClss = Convert.ToInt32(cboFaultyGBN.SelectedValue.ToString());
            }
            int nChkCustom = 0;        //거래처.
            string sCustomID = string.Empty;
            if (chkCustomer.IsChecked == true)
            {
                nChkCustom = 1;
                sCustomID = txtCustomer.Tag.ToString();                
            }
            int nChkOrder = 0;       //관리번호.
            string sOrder = string.Empty;
            if (chkOrderID.IsChecked == true)
            {
                nChkOrder = 1;
                sOrder = txtOrderID.Text;
            }
            int nChkBoxID = 0;       //박스번호.
            string sBoxID = string.Empty;
            if (chkBoxID.IsChecked == true)
            {
                nChkBoxID = 1;
                sBoxID = txtBoxID.Text;
            }
            int nChkCID = 0;       //박스번호.
            string sCID = string.Empty;
            if (chkCID.IsChecked == true)
            {
                nChkCID = 1;
                sCID = txtCID.Text;
            }
            
            //품명 또는 품번
            int chkArticleID = 0;
            string ArticleID = "";

            if (chkArticle.IsChecked == true && txtArticle.Text != "")
            {
                chkArticleID = 1;
                ArticleID = txtArticle.Tag.ToString();
            }
            if (chkArticleNo.IsChecked == true && txtArticleNo.Text != "") 
            {
                chkArticleID = 1;
                ArticleID = txtArticleNo.Tag.ToString();
            }

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sDate", sDate);
            sqlParameter.Add("eDate", eDate);
            sqlParameter.Add("nClss", nClss);
            sqlParameter.Add("nChkCustom", nChkCustom);
            sqlParameter.Add("sCustomID", sCustomID);

            sqlParameter.Add("nChkArticle", chkArticleID); // nChkArticle);
            sqlParameter.Add("sArticleID", ArticleID); // sArticleID);
            sqlParameter.Add("nChkOrder", nChkOrder);
            sqlParameter.Add("sOrder",  sOrder);
            sqlParameter.Add("nChkCID", nChkCID);
            sqlParameter.Add("sCID", sCID); 
            sqlParameter.Add("nChkBoxID", nChkBoxID);
            sqlParameter.Add("sBoxID", sBoxID);
            sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");
            sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Inspect_sInspectByBox", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];
                DT = null;
                DT = dt;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회결과가 없습니다.");
                    return;
                }
                else
                {
                    try
                    {
                        dgdInspect.Items.Clear();
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            if (item["cls"].ToString() == "1")
                            {
                                var Window_Ins_DateBox_ViewInsert = new Win_Qul_DateBox_QView()
                                {
                                    ExamDate = item["ExamDate"].ToString(),
                                    KCustom = item["KCustom"].ToString(),
                                    OrderNo = item["OrderNo"].ToString(),
                                    OrderID = item["OrderID"].ToString(),
                                    Article = item["Article"].ToString(),

                                    spec = item["spec"].ToString(),
                                    BuyerModel = item["BuyerModel"].ToString(),
                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    PackBoxID = item["PackBoxID"].ToString(),

                                    BoxID = item["BoxID"].ToString(), //2021-06-25
                                    PackID = item["PackID"].ToString(),      //2021-06-25

                                    CtrlQty = item["CtrlQty"].ToString(),
                                    PassRoll = item["PassRoll"].ToString(),
                                    PassQty = item["PassQty"].ToString(),
                                    DefectRoll = item["DefectRoll"].ToString(),
                                    DefectQty = item["DefectQty"].ToString(),

                                    ExamNo = item["ExamNo"].ToString(),
                                    UnitClss = item["UnitClss"].ToString(),

                                    ColorGreen = "false",
                                    ColorRed = "false"
                                };
                                Window_Ins_DateBox_ViewInsert.CtrlQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.CtrlQty);
                                Window_Ins_DateBox_ViewInsert.PassRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassRoll);
                                Window_Ins_DateBox_ViewInsert.PassQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassQty);
                                Window_Ins_DateBox_ViewInsert.DefectRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectRoll);
                                Window_Ins_DateBox_ViewInsert.DefectQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectQty);
                                dgdInspect.Items.Add(Window_Ins_DateBox_ViewInsert);
                            }
                            if (item["cls"].ToString() == "2")        // 오더계
                            {
                                var Window_Ins_DateBox_ViewInsert = new Win_Qul_DateBox_QView()
                                {
                                    ExamDate = item["ExamDate"].ToString(),
                                    KCustom = item["KCustom"].ToString(),
                                    OrderNo = item["OrderNO"].ToString(),
                                    OrderID = item["OrderID"].ToString(),
                                    Article = item["Article"].ToString(),

                                    spec = "",
                                    BuyerModel = "",
                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    PackID = "오더계",

                                    CtrlQty = item["CtrlQty"].ToString(),
                                    PassRoll = item["PassRoll"].ToString(),
                                    PassQty = item["PassQty"].ToString(),
                                    DefectRoll = item["DefectRoll"].ToString(),
                                    DefectQty = item["DefectQty"].ToString(),

                                    ExamNo = item["ExamNo"].ToString(),
                                    UnitClss = item["UnitClss"].ToString(),

                                    ColorGreen = "true",
                                    ColorRed = "false"
                                };
                                Window_Ins_DateBox_ViewInsert.CtrlQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.CtrlQty);
                                Window_Ins_DateBox_ViewInsert.PassRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassRoll);
                                Window_Ins_DateBox_ViewInsert.PassQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassQty);
                                Window_Ins_DateBox_ViewInsert.DefectRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectRoll);
                                Window_Ins_DateBox_ViewInsert.DefectQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectQty);
                                dgdInspect.Items.Add(Window_Ins_DateBox_ViewInsert);
                            }
                            if (item["cls"].ToString() == "3")        // 일계
                            {
                                var Window_Ins_DateBox_ViewInsert = new Win_Qul_DateBox_QView()
                                {
                                    ExamDate = "일계",
                                    KCustom = "",
                                    OrderNo = "",
                                    OrderID = "",
                                    Article = "",

                                    spec = "",
                                    BuyerModel = "",
                                    BuyerArticleNo = "",
                                    OrderQty = "",
                                    BoxID = "",

                                    CtrlQty = item["CtrlQty"].ToString(),
                                    PassRoll = item["PassRoll"].ToString(),
                                    PassQty = item["PassQty"].ToString(),
                                    DefectRoll = item["DefectRoll"].ToString(),
                                    DefectQty = item["DefectQty"].ToString(),

                                    ExamNo = item["ExamNo"].ToString(),
                                    UnitClss = item["UnitClss"].ToString(),

                                    ColorGreen = "true",
                                    ColorRed = "false"
                                };
                                Window_Ins_DateBox_ViewInsert.CtrlQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.CtrlQty);
                                Window_Ins_DateBox_ViewInsert.PassRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassRoll);
                                Window_Ins_DateBox_ViewInsert.PassQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassQty);
                                Window_Ins_DateBox_ViewInsert.DefectRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectRoll);
                                Window_Ins_DateBox_ViewInsert.DefectQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectQty);
                                dgdInspect.Items.Add(Window_Ins_DateBox_ViewInsert);
                            }
                            if (item["cls"].ToString() == "4")        // 총계
                            {
                                dgdTotal.Items.Clear();
                                var Window_Ins_DateBox_ViewInsert = new Win_Qul_DateBox_QView()
                                {
                                    ExamDate = "총계",
                                    KCustom = "",
                                    OrderNo = "",
                                    OrderID = "",
                                    Article = "",

                                    spec = "",
                                    BuyerModel = "",
                                    BuyerArticleNo = "",
                                    OrderQty = "",
                                    BoxID = "",

                                    CtrlQty = item["CtrlQty"].ToString(),
                                    PassRoll = item["PassRoll"].ToString(),
                                    PassQty = item["PassQty"].ToString(),
                                    DefectRoll = item["DefectRoll"].ToString(),
                                    DefectQty = item["DefectQty"].ToString(),

                                    ExamNo = item["ExamNo"].ToString(),
                                    UnitClss = "",

                                    ColorGreen = "false",
                                    ColorRed = "false"
                                };
                                Window_Ins_DateBox_ViewInsert.CtrlQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.CtrlQty);
                                Window_Ins_DateBox_ViewInsert.PassRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassRoll);
                                Window_Ins_DateBox_ViewInsert.PassQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.PassQty);
                                Window_Ins_DateBox_ViewInsert.DefectRoll = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectRoll);
                                Window_Ins_DateBox_ViewInsert.DefectQty = Lib.Instance.returnNumString(Window_Ins_DateBox_ViewInsert.DefectQty);

                                dgdTotal.Items.Add(Window_Ins_DateBox_ViewInsert);
                            }    
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        DataStore.Instance.CloseConnection();
                    }
                }
            }
        }

        #endregion


        #region 엑셀

        // 엑셀버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdInspect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib = new Lib();
            System.Data.DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdInspect.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdInspect.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdInspect);
                    else
                        dt = lib.DataGirdToDataTable(dgdInspect);

                    Name = dgdInspect.Name;
                    if(lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
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



        //닫기 기능.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.subProgramID.ToString().Contains("MDI"))
                {
                    if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                    }
                }
                i++;
            }
        }


        //정렬 버튼이벤트.
        private void btnMultiSort_Click(object sender, RoutedEventArgs e)
        {
            PopUp.MultiLevelSort MLS = new PopUp.MultiLevelSort(dgdInspect);
            MLS.ShowDialog();

            if (MLS.DialogResult.HasValue)
            {
                string targetSortProperty = string.Empty;
                int targetColIndex;
                dgdInspect.Items.SortDescriptions.Clear();

                for (int x = 0; x < MLS.ColName.Count; x++)
                {
                    targetSortProperty = MLS.SortingProperty[x];
                    targetColIndex = MLS.ColIndex[x];
                    var targetCol = dgdInspect.Columns[targetColIndex];

                    if (targetSortProperty == "UP")
                    {
                        dgdInspect.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Ascending));
                        targetCol.SortDirection = ListSortDirection.Ascending;
                    }
                    else
                    {
                        dgdInspect.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Descending));
                        targetCol.SortDirection = ListSortDirection.Descending;
                    }
                }
                dgdInspect.Refresh();
            }
        }

        private void TxtArticle_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
             if (e.Key == Key.Enter)
            {
                btnArticle_Click(null, null);
            }
        }

        private void txtCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnCustomer_Click(null, null);
            }
        }

        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

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

        // 플러스파인더 _ 품번 찾기
        private void btnArticleNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
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

        private void TxtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
            }
        }

        //품번
        private void chkArticleNo_Click(object sender, MouseButtonEventArgs e)
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
    }

    class Win_Qul_DateBox_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }

        public string ExamDate { get; set; }
        public string OrderID { get; set; }
        public string OrderNo { get; set; }
        public string OrderSeq { get; set; }
        public string BoxID { get; set; }

        public string OrderQty { get; set; }
        public string CustomID { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerModelID { get; set; }

        public string BuyerModel { get; set; }
        public string UnitClss { get; set; }
        public string KCustom { get; set; }
        public string spec { get; set; }
        public string BuyerArticleNo { get; set; }

        public string DefectClss { get; set; }
        public string CtrlQty { get; set; }
        public string lossQty { get; set; }
        public string PassRoll { get; set; }
        public string PassQty { get; set; }

        public string DefectRoll { get; set; }
        public string DefectQty { get; set; }
        public string ExamNo { get; set; }

        public string PackBoxID { get; set; } 
        public string PackID { get; set; }
        public string ColorGreen { get; set; }
        public string ColorRed { get; set; }
    }

}
