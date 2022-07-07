using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WooJung.PopUp;
using WizMes_WooJung.PopUP;
using WPF.MDI;

namespace WizMes_WooJung
{
    /// <summary>
    /// Win_ord_OrderClose_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OrderClose_U : UserControl
    {
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;

        Win_ord_OrderClose_U_CodeView WinOrderClose = new Win_ord_OrderClose_U_CodeView();
        Lib lib = new Lib();
        string rowHeaderNum = string.Empty;
        int rowNum = 0;

        WizMes_WooJung.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        System.Data.DataTable DT;
        System.Data.DataTable DTotal;


        public Win_ord_OrderClose_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            btnToday_Click(null, null);
            SetComboBox();
            Check_bdrOrder();
            cboOrderStatus.SelectedIndex = 0;
        }

        //콤보박스 세팅
        private void SetComboBox()
        {
            List<string> strValue = new List<string>();
            strValue.Add("전체");
            strValue.Add("진행건");
            strValue.Add("마감건");

            ObservableCollection<CodeView> cbOrderStatus = ComboBoxUtil.Instance.Direct_SetComboBox(strValue);
            ObservableCollection<CodeView> cbWork = ComboBoxUtil.Instance.GetCode_SetComboBox("Work", null);

            this.cboOrderStatus.ItemsSource = cbOrderStatus;
            this.cboOrderStatus.DisplayMemberPath = "code_name";
            this.cboOrderStatus.SelectedValuePath = "code_id";

            this.cboWork.ItemsSource = cbWork;
            this.cboWork.DisplayMemberPath = "code_name";
            this.cboWork.SelectedValuePath = "code_id";
        }

        #region 라벨 체크박스 이벤트 관련

        //일자
        private void lblOrderDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderDay.IsChecked == true) { chkOrderDay.IsChecked = false; }
            else { chkOrderDay.IsChecked = true; }
        }

        //일자
        private void chkOrderDay_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //일자
        private void chkOrderDay_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        //전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpSDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpEDate.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

            if (dtpSDate.SelectedDate != null)
            {
                DateTime ThatMonth1 = dtpSDate.SelectedDate.Value.AddDays(-(dtpSDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                dtpSDate.SelectedDate = LastMonth1;
                dtpEDate.SelectedDate = LastMonth31;
            }
            else
            {
                DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                dtpSDate.SelectedDate = LastMonth1;
                dtpEDate.SelectedDate = LastMonth31;
            }
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        ////금년
        //private void btnThisYear_Click(object sender, RoutedEventArgs e)
        //{
        //    dtpSDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[0];
        //    dtpEDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[1];
        //}

        //전일
        private void BtnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            if (dtpSDate.SelectedDate != null)
            {
                dtpSDate.SelectedDate = dtpSDate.SelectedDate.Value.AddDays(-1);
                dtpEDate.SelectedDate = dtpSDate.SelectedDate;
            }
            else
            {
                dtpSDate.SelectedDate = DateTime.Today.AddDays(-1);
                dtpEDate.SelectedDate = DateTime.Today.AddDays(-1);
            }
        }

        //수주상태
        private void cboOrderStatus_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboOrderStatus.SelectedIndex == 0)
            {
                btnFinal.IsEnabled = false;
                rowNum = 0;
                re_Search();
            }
            else if (cboOrderStatus.SelectedIndex == 1)
            {
                btnFinal.IsEnabled = true;
                btnFinal.Content = "마감처리";
                rowNum = 0;
                re_Search();
            }
            else
            {
                btnFinal.IsEnabled = true;
                btnFinal.Content = "진행처리";
                rowNum = 0;
                re_Search();
            }
        }

        //수주 진행 건은 마감처리 / 마감 건은 진행처리로 변경하는 버튼
        private void BtnFinal_Click(object sender, RoutedEventArgs e)
        {
            //string OrderID = string.Empty;

            // 다중선택 했을 때 각각 OrderID 들어가도록 설정했으므로 이건 안써도 돼
            //var Order = dgdMain.SelectedItem as Win_ord_OrderClose_U_CodeView;
            //if (Order != null)
            //{
            //    OrderID = Order.OrderID;
            //}

            string CloseFlag = string.Empty;
            string CloseClss = string.Empty;

            if (btnFinal.Content.ToString().Equals("마감처리"))
            {
                CloseFlag = "1";
                CloseClss = "1";

                if (MessageBox.Show("해당 건을 마감처리 하시겠습니까?", "처리 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }
                }
            }
            else if (btnFinal.Content.ToString().Equals("진행처리"))
            {
                CloseFlag = "2";
                CloseClss = "";

                if (MessageBox.Show("해당 건을 진행처리 하시겠습니까?", "처리 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }
                }
            }

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            try
            {
                //일괄처리할 때 쓰는 변수
                int CheckCount = 0;

                //데이터그리드의 체크박스 true된 수 많음 CheckCount 수 늘리기
                foreach (Win_ord_OrderClose_U_CodeView OrderCloseU in dgdMain.Items)
                {
                    if (OrderCloseU.IsCheck == true)
                    {
                        CheckCount++;
                    }
                }

                //체크된 그리드가 하나 이상일 경우(1개라도 체크가 되어 있을 경우)
                if (CheckCount > 0)
                {
                    foreach (Win_ord_OrderClose_U_CodeView OrderCloseU in dgdMain.Items)
                    {
                        if (OrderCloseU != null)
                        {
                            if (OrderCloseU.IsCheck == true)
                            {
                                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("CloseFlag", CloseFlag);
                                sqlParameter.Add("OrderID", OrderCloseU.OrderID);
                                sqlParameter.Add("CloseClss", CloseClss);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_OrderClose_uCloseClss";     //마감처리 누르면 CloseClss에 1 저장, 진행처리 누르면 '' 저장 Order테이블에.
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "OrderID";
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
                    }
                }
                else
                {
                    MessageBox.Show("[저장실패]\r\n 처리할 체크항목이 없습니다.");
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

            dgdMain.Items.Clear();
            FillGrid();
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

        //품명(품번 검색으로 수정, 2020.03.17, 장가빈)
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 81, txtArticle.Text);
            }
        }

        //품명(품번 검색으로 수정, 2020.03.17, 장가빈)
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 81, txtArticle.Text);
        }

        //차종
        private void lblModel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkModel.IsChecked == true) { chkModel.IsChecked = false; }
            else { chkModel.IsChecked = true; }
        }

        //차종
        private void chkModel_Checked(object sender, RoutedEventArgs e)
        {
            txtModel.IsEnabled = true;
            btnPfModel.IsEnabled = true;
            txtModel.Focus();
        }

        //차종
        private void chkModel_Unchecked(object sender, RoutedEventArgs e)
        {
            txtModel.IsEnabled = false;
            btnPfModel.IsEnabled = false;
        }

        //차종
        private void txtModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
            }
        }

        //차종
        private void btnPfModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
        }

        //OrderNo
        private void lblOrder_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrder.IsChecked == true) { chkOrder.IsChecked = false; }
            else { chkOrder.IsChecked = true; }
        }

        //OrderNo
        private void chkOrder_Checked(object sender, RoutedEventArgs e)
        {
            txtOrderNo.IsEnabled = true;
            btnPfOrderNo.IsEnabled = true;
            txtOrderNo.Focus();
        }

        //OrderNo
        private void chkOrder_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrderNo.IsEnabled = false;
            btnPfOrderNo.IsEnabled = false;
        }

        //OrderNo
        private void txtOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (e.Key == Key.Enter)
                {
                    MainWindow.pf.ReturnCode(txtOrderNo, (int)Defind_CodeFind.DCF_ORDER, "");
                }
            }
        }

        //OrderNo
        private void btnPfOrderNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtOrderNo, (int)Defind_CodeFind.DCF_ORDER, "");
        }

        private void rbnOrderNo_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        private void rbnOrderID_Click(object sender, RoutedEventArgs e)
        {
            Check_bdrOrder();
        }

        private void Check_bdrOrder()
        {
            if (rbnOrderID.IsChecked == true)
            {
                tbkOrder.Text = " 관리번호";
                dgdtxtOrderID.Visibility = Visibility.Visible;
                dgdtxtOrderNo.Visibility = Visibility.Hidden;
            }
            else if (rbnOrderNo.IsChecked == true)
            {
                tbkOrder.Text = "Order No.";
                dgdtxtOrderID.Visibility = Visibility.Hidden;
                dgdtxtOrderNo.Visibility = Visibility.Visible;
            }
        }

        //가공구분
        private void lblWork_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkWork.IsChecked == true) { chkWork.IsChecked = false; }
            else { chkWork.IsChecked = true; }
        }

        //가공구분
        private void chkWork_Checked(object sender, RoutedEventArgs e)
        {
            cboWork.IsEnabled = true;
        }

        //가공구분
        private void chkWork_Unchecked(object sender, RoutedEventArgs e)
        {
            cboWork.IsEnabled = false;
        }

        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain != null)
            {
                dgdMain.Items.Clear();
            }
            if (dgdSum != null)
            {
                dgdSum.Items.Clear();
            }
            FillGrid();
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //인쇄
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //인쇄 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        //바로 인쇄
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }

        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;

            string[] lst = new string[2];
            lst[0] = "수주조회";
            lst[1] = dgdMain.Name;
            Lib lib = new Lib();

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    //MessageBox.Show("대분류");
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
            }
            lib = null;
        }

        //실조회 및 하단 합계
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

                sqlParameter.Add("ChkDate", chkOrderDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOrderDay.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkOrderDay.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustom", chkCustom.IsChecked == true ? (txtCustom.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? (txtCustom.Tag != null ? txtCustom.Tag.ToString() : txtCustom.Text) : "");
                sqlParameter.Add("ChkArticle", 0);// chkArticle.IsChecked == true ? (txtArticle.Tag != null ? 1 : 2) : 0);
                sqlParameter.Add("ArticleID", "");//chkArticle.IsChecked == true ? (txtArticle.Tag != null ? txtArticle.Tag.ToString() : txtArticle.Text) : "");
                sqlParameter.Add("ChkOrderID", chkOrder.IsChecked == true ? (rbnOrderID.IsChecked == true ? 1 : 2) : 0);
                sqlParameter.Add("OrderID", txtOrderNo.Text);
                sqlParameter.Add("ChkClose", int.Parse(cboOrderStatus.SelectedValue != null ? cboOrderStatus.SelectedValue.ToString() : ""));
                sqlParameter.Add("nChkOrderFlag", 0);
                sqlParameter.Add("ChkFiber", 0);
                sqlParameter.Add("ChkBuyerModel", chkModel.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sBuyerModelID", chkModel.IsChecked == true ?
                    (txtModel.Tag != null ? txtModel.Tag.ToString() : "") : "");
                sqlParameter.Add("ChkWorkID", chkWork.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sWorkID", chkWork.IsChecked == true ? (cboWork.SelectedValue != null ? cboWork.SelectedValue.ToString() : "") : "");
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sOrderTotal", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    //dataGrid.Items.Clear();
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int OrderSum = 0;
                        int InsertSum = 0;
                        double InspectSum = 0;
                        double PassSum = 0;
                        double DefectSum = 0;
                        double OutSum = 0;
                        double OasSum = 0;

                        foreach (DataRow item in drc)
                        {
                            var Window_OrderClose_DTO = new Win_ord_OrderClose_U_CodeView()
                            {
                                IsCheck = false,
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = item["OrderNO"] as string,
                                CustomID = item["CustomID"] as string,
                                KCustom = item["KCustom"] as string,

                                DvlyDate = item["DvlyDate"] as string,
                                CloseClss = item["CloseClss"] as string,
                                //ChunkRate = item["ChunkRate"].ToString(),
                                //LossRate = item["LossRate"] as string,
                                Article = item["Article"] as string,

                                WorkName = item["WorkName"].ToString(),
                                //WorkWidth = item["WorkWidth"] as string,
                                OrderQty = item["OrderQty"].ToString(),
                                ColorQty = item["ColorQty"].ToString(),
                                 
                                UnitClss = item["UnitClss"] as string,  //주문기준 value
                                InspectQty = item["InspectQty"].ToString(),

                                PassQty = item["PassQty"].ToString(),
                                DefectQty = item["DefectQty"].ToString(),
                                OutQty = item["OutQty"].ToString(),
                                BuyerModel = item["BuyerModel"] as string,
                                BuyerModelID = item["BuyerModelID"] as string,

                                BuyerArticleNo = item["BuyerArticleNo"] as string,
                                UnitClssName = item["UnitClssName"] as string,
                                p1StartWorkDate = item["p1StartWorkDate"] as string,
                                p1StartWorkDTime = item["p1StartWorkDTime"] as string,
                                p1WorkQty = item["p1WorkQty"].ToString(),

                                p1ProcessID = item["p1ProcessID"] as string,
                                p1ProcessName = item["p1ProcessName"] as string,
                                //ArticleID = item["ArticleID"] as string,
                                //AcptDate = item["AcptDate"] as string,
                                Num = i + 1,

                            };

                            Window_OrderClose_DTO.OrderID_CV = Window_OrderClose_DTO.OrderID.Substring(0, 4) + "-" +
                                Window_OrderClose_DTO.OrderID.Substring(4, 2) + "-" + Window_OrderClose_DTO.OrderID.Substring(6, 4);

                            Window_OrderClose_DTO.OverAndShort = double.Parse(Window_OrderClose_DTO.OrderQty) - double.Parse(Window_OrderClose_DTO.PassQty);

                            i++;

                            if (Window_OrderClose_DTO.OrderQty == null || Window_OrderClose_DTO.OrderQty.Equals("") || Window_OrderClose_DTO.OrderQty.Substring(0, 1).Equals("0"))
                            {
                                OrderSum += 0;
                            }
                            else
                            {
                                OrderSum += int.Parse(Window_OrderClose_DTO.OrderQty);
                            }

                            if (Window_OrderClose_DTO.p1WorkQty == null || Window_OrderClose_DTO.p1WorkQty.Equals("") || Window_OrderClose_DTO.p1WorkQty.Substring(0, 1).Equals("0"))
                            {
                                InsertSum += 0;
                            }
                            else
                            {
                                InsertSum += (int)(double.Parse(Window_OrderClose_DTO.p1WorkQty));
                            }

                            if (Window_OrderClose_DTO.p1WorkQty != null && Lib.Instance.IsNumOrAnother(Window_OrderClose_DTO.p1WorkQty))
                            {
                                if (Window_OrderClose_DTO.p1WorkQty.Contains("."))
                                {
                                    Window_OrderClose_DTO.p1WorkQty = Window_OrderClose_DTO.p1WorkQty.Substring(0, Window_OrderClose_DTO.p1WorkQty.IndexOf("."));
                                }
                            }

                            if (Window_OrderClose_DTO.InspectQty == null || Window_OrderClose_DTO.InspectQty.Equals("") || Window_OrderClose_DTO.InspectQty.Substring(0, 1).Equals("0"))
                            {
                                InspectSum += 0;
                            }
                            else
                            {
                                InspectSum += double.Parse(Window_OrderClose_DTO.InspectQty);
                            }

                            if (Window_OrderClose_DTO.InspectQty != null && Lib.Instance.IsNumOrAnother(Window_OrderClose_DTO.InspectQty))
                            {
                                if (Window_OrderClose_DTO.InspectQty.Contains("."))
                                {
                                    Window_OrderClose_DTO.InspectQty = Window_OrderClose_DTO.InspectQty.Substring(0, Window_OrderClose_DTO.InspectQty.IndexOf("."));
                                }
                            }

                            if (Window_OrderClose_DTO.PassQty == null || Window_OrderClose_DTO.PassQty.Equals("") || Window_OrderClose_DTO.PassQty.Substring(0, 1).Equals("0"))
                            {
                                PassSum += 0;
                            }
                            else
                            {
                                PassSum += double.Parse(Window_OrderClose_DTO.PassQty);
                            }

                            if (Window_OrderClose_DTO.PassQty != null && Lib.Instance.IsNumOrAnother(Window_OrderClose_DTO.PassQty))
                            {
                                if (Window_OrderClose_DTO.PassQty.Contains("."))
                                {
                                    Window_OrderClose_DTO.PassQty = Window_OrderClose_DTO.PassQty.Substring(0, Window_OrderClose_DTO.PassQty.IndexOf("."));
                                }
                            }

                            if (Window_OrderClose_DTO.DefectQty == null || Window_OrderClose_DTO.DefectQty.Equals("") || Window_OrderClose_DTO.DefectQty.Substring(0, 1).Equals("0"))
                            {
                                DefectSum += 0;
                            }
                            else
                            {
                                DefectSum += double.Parse(Window_OrderClose_DTO.DefectQty);
                            }

                            if (Window_OrderClose_DTO.DefectQty != null && Lib.Instance.IsNumOrAnother(Window_OrderClose_DTO.DefectQty))
                            {
                                if (Window_OrderClose_DTO.DefectQty.Contains("."))
                                {
                                    Window_OrderClose_DTO.DefectQty = Window_OrderClose_DTO.DefectQty.Substring(0, Window_OrderClose_DTO.DefectQty.IndexOf("."));
                                }
                            }

                            if (Window_OrderClose_DTO.OutQty == null || Window_OrderClose_DTO.OutQty.Equals("") || Window_OrderClose_DTO.OutQty.Substring(0, 1).Equals("0"))
                            {
                                OutSum += 0;
                            }
                            else
                            {
                                OutSum += double.Parse(Window_OrderClose_DTO.OutQty);
                            }

                            if (Window_OrderClose_DTO.OutQty != null && Lib.Instance.IsNumOrAnother(Window_OrderClose_DTO.OutQty))
                            {
                                if (Window_OrderClose_DTO.OutQty.Contains("."))
                                {
                                    Window_OrderClose_DTO.OutQty = Window_OrderClose_DTO.OutQty.Substring(0, Window_OrderClose_DTO.OutQty.IndexOf("."));
                                }
                            }

                            OasSum += Window_OrderClose_DTO.OverAndShort;

                            //중간 납기일에 들어가는 '-' 를 위해 체크한 후 잘라주거나 그냥 넣어준다.
                            if (Window_OrderClose_DTO.DvlyDate != null && Window_OrderClose_DTO.DvlyDate.ToString().Trim() != "")
                            {
                                Window_OrderClose_DTO.DvlyDateEdit = item["DvlyDate"].ToString().Substring(0, 4) + "-" + item["DvlyDate"].ToString().Substring(4, 2) + "-" + item["DvlyDate"].ToString().Substring(6, 2);
                            }
                            else
                            {
                                Window_OrderClose_DTO.DvlyDateEdit = " ";
                            }

                            //중간에 투입일시의 정규식을 넣기가 힘들어 노가다...
                            if (Window_OrderClose_DTO.p1StartWorkDate != null && !Window_OrderClose_DTO.p1StartWorkDate.Equals("") && Window_OrderClose_DTO.p1StartWorkDTime != null && !Window_OrderClose_DTO.p1StartWorkDTime.Equals(""))
                            {
                                Window_OrderClose_DTO.DayAndTime = item["p1StartWorkDate"].ToString().Substring(4, 2) + "-" + item["p1StartWorkDate"].ToString().Substring(6) + " "
                                + item["p1StartWorkDTime"].ToString().Substring(0, 2) + ":" + item["p1StartWorkDTime"].ToString().Substring(2);
                            }

                            Window_OrderClose_DTO.DefectQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.DefectQty);
                            Window_OrderClose_DTO.OrderQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.OrderQty);
                            Window_OrderClose_DTO.InspectQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.InspectQty);
                            Window_OrderClose_DTO.OutQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.OutQty);
                            Window_OrderClose_DTO.p1WorkQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.p1WorkQty);
                            Window_OrderClose_DTO.PassQty = Lib.Instance.returnNumStringZero(Window_OrderClose_DTO.PassQty);
                            dgdMain.Items.Add(Window_OrderClose_DTO);
                            rowHeaderNum = i.ToString();
                        }

                        var ThisOrderSum = new dgOrderSum
                        {
                            Count = i,
                            OrderSum = OrderSum,
                            InsertSum = InsertSum,
                            InspectSum = InspectSum,
                            PassSum = PassSum,
                            DefectSum = DefectSum,
                            OutSum = OutSum,
                            OasSum = OasSum,
                            TextData = "합계"
                        };
                        dgdSum.Items.Add(ThisOrderSum);
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

        //전체선택
        private void btnAllCheck_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_ord_OrderClose_U_CodeView woccv in dgdMain.Items)
            {
                woccv.IsCheck = true;
            }
        }

        //선택해제
        private void btnAllNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (Win_ord_OrderClose_U_CodeView woccv in dgdMain.Items)
            {
                woccv.IsCheck = false;
            }
        }

        //인쇄 실질 동작
        private void PrintWork(bool preview_click)
        {
            Lib lib2 = new Lib();

            try
            {
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\수주진행현황(영업관리).xls";
                //MyBookPath = MyBookPath.Substring(0, MyBookPath.LastIndexOf("\\")) + "\\order_standard.xls";
                //string MyBookPath = "C:/Users/Administrator/Desktop/order_standard.xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];

                //상단의 일자 
                if (chkOrderDay.IsChecked == true)
                {
                    workrange = worksheet.get_Range("E2", "Q2");//셀 범위 지정
                    workrange.Value2 = dtpSDate.Text + "~" + dtpEDate.Text;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }
                else
                {
                    workrange = worksheet.get_Range("E2", "K2");//셀 범위 지정
                    workrange.Value2 = "전체"; //"" + "~" + "";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }


                //오더번호 혹은 관리번호 
                if (rbnOrderNo.IsChecked == true)
                {
                    workrange = worksheet.get_Range("C5", "F5");//셀 범위 지정
                    workrange.Value2 = "오더번호";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }
                else
                {
                    workrange = worksheet.get_Range("C5", "F5");//셀 범위 지정
                    workrange.Value2 = "관리번호";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //workrange.Font.Size = 10;
                }

                //하단의 회사명
                workrange = worksheet.get_Range("AN35", "AU35");//셀 범위 지정
                workrange.Value2 = "주식회사 지엘에스";
                workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workrange.Font.Size = 11;


                /////////////////////////
                int Page = 0;
                int DataCount = 0;
                int copyLine = 0;

                copysheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                DT = lib2.DataGirdToDataTable(dgdMain);

                string str_Num = string.Empty;
                string str_OrderID = string.Empty;
                string str_OrderID_CV = string.Empty;
                string str_KCustom = string.Empty;
                string str_Article = string.Empty;
                string str_Model = string.Empty;
                string str_ArticleNo = string.Empty;
                string str_DvlyDate = string.Empty;
                string str_Work = string.Empty;
                string str_OrderQty = string.Empty;
                string str_UnitClssName = string.Empty;
                string str_DayAndTime = string.Empty;
                string str_p1WorkQty = string.Empty;
                string str_InspectQty = string.Empty;
                string str_PassQty = string.Empty;
                string str_DefectQty = string.Empty;
                string str_OutQty = string.Empty;

                int TotalCnt = dgdMain.Items.Count;
                int canInsert = 27; //데이터가 입력되는 행 수 27개

                int PageCount = (int)Math.Ceiling(1.0 * TotalCnt / canInsert);

                var Sum = new dgOrderSum();

                //while (dgdMain.Items.Count > DataCount + 1)
                for (int k = 0; k < PageCount; k++)
                {
                    Page++;
                    if (Page != 1) { DataCount++; }  //+1
                    copyLine = (Page - 1) * 38;
                    copysheet.Select();
                    copysheet.UsedRange.Copy();
                    pastesheet.Select();
                    workrange = pastesheet.Cells[copyLine + 1, 1];
                    workrange.Select();
                    pastesheet.Paste();

                    int j = 0;
                    for (int i = DataCount; i < dgdMain.Items.Count; i++)
                    {
                        if (j == 27) { break; }
                        int insertline = copyLine + 7 + j;

                        str_Num = (j + 1).ToString();
                        str_OrderID = DT.Rows[i][1].ToString();
                        str_OrderID_CV = DT.Rows[i][2].ToString();
                        str_KCustom = DT.Rows[i][3].ToString();
                        str_Article = DT.Rows[i][4].ToString();
                        str_Model = DT.Rows[i][5].ToString();
                        str_ArticleNo = DT.Rows[i][6].ToString();
                        str_DvlyDate = DT.Rows[i][7].ToString();
                        str_Work = DT.Rows[i][8].ToString();
                        str_OrderQty = DT.Rows[i][9].ToString();
                        str_UnitClssName = DT.Rows[i][10].ToString();
                        str_DayAndTime = DT.Rows[i][11].ToString();
                        str_p1WorkQty = DT.Rows[i][12].ToString();
                        str_InspectQty = DT.Rows[i][13].ToString();
                        str_PassQty = DT.Rows[i][14].ToString();
                        str_DefectQty = DT.Rows[i][15].ToString();
                        str_OutQty = DT.Rows[i][16].ToString();

                        workrange = pastesheet.get_Range("A" + insertline, "B" + insertline);    //순번
                        workrange.Value2 = str_Num;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.3;

                        if (dgdtxtOrderID.ToString().Equals("오더번호"))
                        {
                            workrange = pastesheet.get_Range("C" + insertline, "F" + insertline);    //오더번호
                            workrange.Value2 = str_OrderID;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;
                            workrange.ColumnWidth = 1.8;
                        }
                        else
                        {
                            workrange = pastesheet.get_Range("C" + insertline, "F" + insertline);    //관리번호
                            workrange.Value2 = str_OrderID_CV;
                            workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            workrange.Font.Size = 9;
                            workrange.ColumnWidth = 1.8;
                        }

                        workrange = pastesheet.get_Range("G" + insertline, "J" + insertline);     //거래처
                        workrange.Value2 = str_KCustom;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 9;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("K" + insertline, "N" + insertline);    //품명
                        workrange.Value2 = str_Article;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("O" + insertline, "R" + insertline);    //차종
                        workrange.Value2 = str_Model;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 0.9;

                        workrange = pastesheet.get_Range("S" + insertline, "V" + insertline);    //품번
                        workrange.Value2 = str_ArticleNo;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 2.7;

                        workrange = pastesheet.get_Range("W" + insertline, "Y" + insertline);    //가공구분
                        workrange.Value2 = str_Work;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.8;

                        workrange = pastesheet.get_Range("Z" + insertline, "AA" + insertline);    //납기일
                        workrange.Value2 = str_DvlyDate;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 3.8;

                        workrange = pastesheet.get_Range("AB" + insertline, "AC" + insertline);    //투입일

                        if (str_DayAndTime.Length > 5)
                        {
                            workrange.Value2 = str_DayAndTime.Substring(0, 5);
                        }
                        else
                        {
                            workrange.Value2 = str_DayAndTime;
                        }

                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 3.8;

                        workrange = pastesheet.get_Range("AD" + insertline, "AF" + insertline);    //수주량
                        workrange.Value2 = str_OrderQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.7;

                        workrange = pastesheet.get_Range("AG" + insertline, "AI" + insertline);    //투입량
                        workrange.Value2 = str_p1WorkQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AJ" + insertline, "AL" + insertline);    //검사량
                        workrange.Value2 = str_InspectQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AM" + insertline, "AO" + insertline);    //합격량
                        workrange.Value2 = str_PassQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AP" + insertline, "AR" + insertline);    //불합격량
                        workrange.Value2 = str_DefectQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        workrange = pastesheet.get_Range("AS" + insertline, "AU" + insertline);    //출고량
                        workrange.Value2 = str_OutQty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 10;
                        workrange.ColumnWidth = 1.2;

                        DataCount = i;
                        j++;

                        // 합계 누적
                        Sum.OrderSum += ConvertInt(str_OrderQty);
                        Sum.InsertSum += ConvertInt(str_p1WorkQty);

                        Sum.InspectSum += ConvertDouble(str_InspectQty);
                        Sum.PassSum += ConvertDouble(str_PassQty);
                        Sum.DefectSum += ConvertDouble(str_DefectQty);
                        Sum.OutSum += ConvertDouble(str_OutQty);


                    }

                    // 합계 출력
                    int totalLine = 34 + ((Page - 1) * 38);

                    Sum.Count = DataCount + 1;


                    workrange = pastesheet.get_Range("AB" + totalLine, "AC" + totalLine);    // 건수
                    workrange.Value2 = Sum.Count + " 건";
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AD" + totalLine, "AF" + totalLine);    // 총 수주량
                    workrange.Value2 = Sum.OrderSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AG" + totalLine, "AI" + totalLine);    // 총 투입량
                    workrange.Value2 = Sum.InsertSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AJ" + totalLine, "AL" + totalLine);    // 총 검수량
                    workrange.Value2 = Sum.InspectSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AM" + totalLine, "AO" + totalLine);    // 총 통과량
                    workrange.Value2 = Sum.PassSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AP" + totalLine, "AR" + totalLine);    // 총 불합격량
                    workrange.Value2 = Sum.DefectSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                    workrange = pastesheet.get_Range("AS" + totalLine, "AU" + totalLine);    // 총 출고량
                    workrange.Value2 = Sum.OutSum;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    workrange.Font.Size = 10;

                }

                pastesheet.PageSetup.TopMargin = 0;
                pastesheet.PageSetup.BottomMargin = 0;
                //pastesheet.PageSetup.Zoom = 43;

                msg.Hide();

                if (preview_click == true)
                {
                    excelapp.Visible = true;
                    pastesheet.PrintPreview();
                }
                else
                {
                    excelapp.Visible = true;
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                lib2.ReleaseExcelObject(workbook);
                lib2.ReleaseExcelObject(worksheet);
                lib2.ReleaseExcelObject(pastesheet);
                lib2.ReleaseExcelObject(excelapp);
                lib2 = null;
            }
        }

        private int ConvertInt(string str)
        {
            int result = 0;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    result = Int32.Parse(str);
                }
            }

            return result;
        }

        private Double ConvertDouble(string str)
        {
            Double result = 0;
            Double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
                {
                    result = Double.Parse(str);
                }
            }

            return result;
        }

        private void re_Search()
        {
            FillGrid();

            //if (dgdMain.Items.Count > 0)
            //{
            //    dgdMain.SelectedIndex = rowNum;
            //}
            //else
            //{
            //    MessageBox.Show("조회된 데이터가 없습니다.");
            //    return;
            //}
        }

        //데이터 그리드 더블 클릭하면 월 납품계획 등록 화면 호출
        private void DgdMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 넘겨줄 데이터를 넣어주시죠
            var Order = dgdMain.SelectedItem as Win_ord_OrderClose_U_CodeView;

            if (Order != null)
            {
                string OrderID = Order.OrderID;
                string sDate = dtpSDate.SelectedDate.Value.ToString("yyyyMMdd");
                string eDate = dtpEDate.SelectedDate.Value.ToString("yyyyMMdd");
                string chkYN = chkOrderDay.IsChecked == true ? "Y" : "N";

                MainWindow.tempContent.Clear();
                MainWindow.tempContent.Add(OrderID);
                MainWindow.tempContent.Add(sDate);
                MainWindow.tempContent.Add(eDate);
                MainWindow.tempContent.Add(chkYN);

                int i = 0;
                foreach (MenuViewModel mvm in MainWindow.mMenulist)
                {
                    if (mvm.Menu.Equals("월간 생산계획등록"))
                    //if (mvm.Menu.Equals("월 납품계획 등록"))
                    {
                        break;
                    }
                    i++;
                }
                try
                {
                    if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[i].subProgramID as MdiChild))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Focus();
                    }
                    else
                    {
                        Type type = Type.GetType("WizMes_YongCheon." + MainWindow.mMenulist[i].ProgramID.Trim(), true);
                        object uie = Activator.CreateInstance(type);

                        MainWindow.mMenulist[i].subProgramID = new MdiChild()
                        {
                            Title = "SamJoo [" + MainWindow.mMenulist[i].MenuID.Trim() + "] " + MainWindow.mMenulist[i].Menu.Trim() +
                                    " (→" + MainWindow.mMenulist[i].ProgramID + ")",
                            Height = SystemParameters.PrimaryScreenHeight * 0.8,
                            MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                            Width = SystemParameters.WorkArea.Width * 0.85,
                            MaxWidth = SystemParameters.WorkArea.Width,
                            Content = uie as UIElement,
                            Tag = MainWindow.mMenulist[i]
                        };

                        Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[i].MenuID, MainWindow.mMenulist[i].Menu, MainWindow.mMenulist[i].subProgramID);
                        MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[i].subProgramID as MdiChild);


                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("해당 화면이 존재하지 않습니다.");
                }
            }
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
    }

    class Win_ord_OrderClose_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public bool IsCheck { get; set; }
        public string OrderNo { get; set; }
        public string OrderID { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string DvlyDate { get; set; }
        public string CloseClss { get; set; }
        public string ChunkRate { get; set; }
        public string LossRate { get; set; }
        public string Article { get; set; }

        public string WorkName { get; set; }

        //public string ArticleID { get; set; }
        public string WorkWidth { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string InspectQty { get; set; }
        public string PassQty { get; set; }
        public string DefectQty { get; set; }
        public string OutQty { get; set; }
        public string ColorQty { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string UnitClssName { get; set; }
        public string p1StartWorkDate { get; set; }
        public string p1StartWorkDTime { get; set; }
        public string p1WorkQty { get; set; }
        public string p1ProcessID { get; set; }
        public string p1ProcessName { get; set; }
        public string DayAndTime { get; set; }
        public string DvlyDateEdit { get; set; }
        //public string AcptDate { get; set; }
        public double OverAndShort { get; set; }
        

        public string OrderID_CV { get; set; }
        public int Num { get; set; }
    }

    public class dgOrderSum
    {
        public int Count { get; set; }
        public int OrderSum { get; set; }
        public int InsertSum { get; set; }
        public double InspectSum { get; set; }
        public double PassSum { get; set; }
        public double DefectSum { get; set; }
        public double OutSum { get; set; }
        public double OasSum { get; set; }

        public string TextData { get; set; }

        //public int Count { get; set; }
        //public int OrderSum { get; set; }
        //public int InsertSum { get; set; }
        //public double InspectSum { get; set; }
        //public double PassSum { get; set; }
        //public double DefectSum { get; set; }
        //public double OutSum { get; set; }
        //public double OasSum { get; set; }
    }
}

