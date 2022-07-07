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
    /// Win_ord_OutWare_Product_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OutWare_Product_Q : UserControl
    {
        Win_ord_OutWare_Product_QView wopqv = new Win_ord_OutWare_Product_QView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        public Win_ord_OutWare_Product_Q()
        {
            InitializeComponent();
            this.DataContext = wopqv;
        }

        private void Window_OutwareProduct_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }

        #region 시작 첫 스텝 // 날짜용 버튼 // ComboSetting // 조회용 체크박스 이벤트
        // 시작 첫 단추.
        private void First_Step()
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            // 시작 지정 및 사용불가 설정.
            chkOutwareDay.IsChecked = true;

            cboArticleGroup.IsEnabled = false;
            txtArticle.IsEnabled = false;
            txtCustomer.IsEnabled = false;
            txtOrderID.IsEnabled = false;
            //cboOutClss.IsEnabled = false;  //체크된 채로 로드되기 때문.
            btnArticle.IsEnabled = false;
            btnCustomer.IsEnabled = false;

            rbnManageNumber.IsChecked = true;

            // 폼 하단 안쓰는 버튼들 가리기.
            chkBuyCustom.Visibility = Visibility.Hidden;
            tbkInsertSheetNO.Visibility = Visibility.Hidden;
            txtBuyCustom.Visibility = Visibility.Hidden;
            txtInsertSheetNO.Visibility = Visibility.Hidden;
            btnBuyCustom.Visibility = Visibility.Hidden;
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



        // 콤보박스 두개 목록 불러오기.  (제품그룹, 출고구분)
        private void ComboBoxSetting()
        {
            cboArticleGroup.Items.Clear();
            cboOutClss.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            //ObservableCollection<CodeView> cbOutClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");

            this.cboArticleGroup.ItemsSource = cbArticleGroup;
            this.cboArticleGroup.DisplayMemberPath = "code_name";
            this.cboArticleGroup.SelectedValuePath = "code_id";
            this.cboArticleGroup.SelectedIndex = 3;  //제품이보이게



            //this.cboOutClss.ItemsSource = cbOutClss;
            //this.cboOutClss.DisplayMemberPath = "code_id_plus_code_name";
            //this.cboOutClss.SelectedValuePath = "code_id";
            //this.cboOutClss.SelectedIndex = 0;

            List<string> cbOutClss = new List<string>();
            cbOutClss.Add("01.제품정상출고");
            cbOutClss.Add("11.제품출고반품");
            cbOutClss.Add("08.예외출고");
            cbOutClss.Add("18.예외출고반품");

            ObservableCollection<CodeView> cboOutClass = ComboBoxUtil.Instance.Direct_SetComboBox(cbOutClss);
            this.cboOutClss.ItemsSource = cboOutClass;
            this.cboOutClss.DisplayMemberPath = "code_name";
            this.cboOutClss.SelectedValuePath = "code_id";
            this.cboOutClss.SelectedIndex = 0;
                       
        }

        //출고일자(날짜) 체크
        private void chkOutwareDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
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
        private void chkOutwareDay_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
            {
                chkOutwareDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkOutwareDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }
        //제품그룹
        private void chkArticleGroup_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticleGroup.IsChecked == true)
            {
                cboArticleGroup.IsEnabled = true;
                cboArticleGroup.Focus();
            }
            else { cboArticleGroup.IsEnabled = false; }
        }
        //제품그룹
        private void chkArticleGroup_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkArticleGroup.IsChecked == true)
            {
                chkArticleGroup.IsChecked = false;
                cboArticleGroup.IsEnabled = false;
            }
            else
            {
                chkArticleGroup.IsChecked = true;
                cboArticleGroup.IsEnabled = true;
                cboArticleGroup.Focus();
            }
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
        private void chkCustomer_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
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
        //관리번호
        private void chkOrderID_Click(object sender, RoutedEventArgs e)
        {
            if (chkOrderID.IsChecked == true)
            {
                txtOrderID.IsEnabled = true;
                txtOrderID.Focus();
            }
            else { txtOrderID.IsEnabled = false; }
        }
        //관리번호
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
        //출고구분
        private void chkOutClss_Click(object sender, RoutedEventArgs e)
        {
            if (chkOutClss.IsChecked == true)
            {
                cboOutClss.IsEnabled = true;
                cboOutClss.Focus();
            }
            else { cboOutClss.IsEnabled = false; }
        }
        //출고구분
        private void chkOutClss_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkOutClss.IsChecked == true)
            {
                chkOutClss.IsChecked = false;
                cboOutClss.IsEnabled = false;
            }
            else
            {
                chkOutClss.IsChecked = true;
                cboOutClss.IsEnabled = true;
                cboOutClss.Focus();
            }
        }

        // 주요관심품목
        private void chkMainInterestItems_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkMainInterestItems.IsChecked == true)
            {
                chkMainInterestItems.IsChecked = false;
            }
            else
            {
                chkMainInterestItems.IsChecked = true;
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

        #endregion


        #region 플러스 파인더
        //플러스 파인더

        //거래처
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");  // 매출거래처만 표기되도록 변경(0 -> 68).
        }

        // 품명
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 81, txtArticle.Text);
        }

        #endregion


        #region 조회 // 조회용 프로시저

        // 검색버튼 클릭. (조회)
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

            string outclssGBN = string.Empty;
            dgdTotal.Items.Clear();

            if (cboOutClss.SelectedIndex == 0) { outclssGBN = "01"; }     //제품정상출고
            else if(cboOutClss.SelectedIndex == 1) { outclssGBN = "11"; } //제품출고반품
            else if(cboOutClss.SelectedIndex == 2) { outclssGBN = "08"; } //예외출고
            else if(cboOutClss.SelectedIndex == 3) { outclssGBN = "18"; } //예외출고반품
                       
            int ChkDate = 0;
            if (chkOutwareDay.IsChecked == true) { ChkDate = 1; }
            string SearchFromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string SearchToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");         //출고일자

            int ChkCustomID = 0;
            if (chkCustomer.IsChecked == true) { ChkCustomID = 1; }                             
            else { txtCustomer.Tag = ""; }                                                      //거래처
            int ChkArticleID = 0;
            if (chkArticle.IsChecked == true) { ChkArticleID = 1; }                            
            else { txtArticle.Tag = ""; }                                                       //품명
            int ChkOrder = 0;
            if (chkOrderID.IsChecked == true)
            {
                if (rbnManageNumber.IsChecked == true) { ChkOrder = 1; }
                else if (rbnOrderNO.IsChecked == true) { ChkOrder = 2; }
            }                                                                               //관리번호
            int chkArticleGrpID = 0;
            if (chkArticleGroup.IsChecked == true) { chkArticleGrpID = 1; }                 //제품그룹
            int int_chkOutClss = 0;
            if (chkOutClss.IsChecked == true) { int_chkOutClss = 1; }                       //출고구분
            int interestitems = 0;
            if (chkMainInterestItems.IsChecked == true) { interestitems = 1; }              //주요관심품

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", ChkDate);
                sqlParameter.Add("SDate", SearchFromDate);
                sqlParameter.Add("EDate", SearchToDate);
                sqlParameter.Add("ChkCustomID", ChkCustomID);
                sqlParameter.Add("CustomID", txtCustomer.Tag.ToString());
                sqlParameter.Add("ChkArticleID", 0); // ChkArticleID);
                sqlParameter.Add("ArticleID", ""); // txtArticle.Tag.ToString());
                sqlParameter.Add("ChkOrder", ChkOrder);
                sqlParameter.Add("Order", txtOrderID.Text);
                sqlParameter.Add("OrderFlag", 0);       //무쓸모..
                sqlParameter.Add("chkArticleGrpID", chkArticleGrpID);
                sqlParameter.Add("sArticleGrpID", cboArticleGroup.SelectedValue.ToString());
                sqlParameter.Add("sProductYN", "Y"); // 제품여부 Y인데 빈값넣으니까 됐어 왜지???
                sqlParameter.Add("chkOutClss", int_chkOutClss);
                sqlParameter.Add("OutClss", outclssGBN); //cboOutClss.SelectedValue.ToString()
                sqlParameter.Add("nMainItem", interestitems);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.ToString().Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sOutwareProduct", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    int i = 1;
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        dgdOutware.Items.Clear();
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            if (item["Depth"].ToString() == "0")
                            {
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = item["OutwareID"].ToString(),
                                    OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = item["KCustom"].ToString(),

                                    OrderNo = item["OrderNo"].ToString(),
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    Article = item["Article"].ToString(),
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = item["OrderQty"].ToString(),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = item["OutClssname"].ToString(),
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    LotID = item["LotID"].ToString(),
                                    ColorGreen = "false",
                                    ColorRed = "false"

                                };
                                dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }
                            else if (item["Depth"].ToString() == "2")
                            {
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = "",
                                    OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = item["KCustom"].ToString(),

                                    OrderNo = "오더계",
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    Article = item["Article"].ToString(),
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = "",
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    ColorGreen = "false",
                                    ColorRed = "false"
                                };
                                dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }
                            else if (item["Depth"].ToString() == "3")
                            {
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = "",
                                    OutDate = item["OutDate"].ToString().Substring(4, 2) + "/" + item["OutDate"].ToString().Substring(6, 2),
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = "거래처 계",

                                    OrderNo = "",
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    Article = "",
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = "",
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    ColorGreen = "true",
                                    ColorRed = "false"
                                };
                                dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }
                            else if (item["Depth"].ToString() == "4")
                            {
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = "",
                                    OutDate = "일계",
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = "",

                                    OrderNo = "",
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    Article = "",
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = "",
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    ColorGreen = "true",
                                    ColorRed = "false"
                                };
                                dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }
                            else if (item["Depth"].ToString() == "5")
                            {
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = "",
                                    OutDate = item["OutDate"].ToString().Substring(4, 2) + "월계",
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = "",

                                    OrderNo = "",
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    Article = "",
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = "",
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    ColorGreen = "false",
                                    ColorRed = "true"
                                };
                                dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }
                            else if (item["Depth"].ToString() == "6")
                            {
                                
                                var window_OutwareProductViewInsert = new Win_ord_OutWare_Product_QView()
                                {
                                    NUM = i,
                                    Depth = item["Depth"].ToString(),
                                    OutwareID = "",
                                    OutDate = "총 합계",
                                    CustomID = item["CustomID"].ToString(),
                                    KCustom = "",

                                    OrderNo = "",
                                    OrderID = item["OrderID"].ToString(),
                                    OutCustom = item["OutCustom"].ToString(),

                                    Article = "",
                                    WorkName = item["WorkName"].ToString(),

                                    OrderQty = stringFormatN0(item["OrderQty"]),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    LabelID = item["LabelID"].ToString(),
                                    LabelGubun = item["LabelGubun"].ToString(),

                                    FromLocName = item["FromLocName"].ToString(),
                                    TOLocname = item["TOLocname"].ToString(),
                                    OutClssname = "",
                                    OutRoll = item["OutRoll"].ToString(),
                                    OutQty = stringFormatN0(item["OutQty"]),

                                    UnitPrice = item["UnitPrice"].ToString(),
                                    Amount = item["Amount"].ToString(),
                                    VatAmount = item["VatAmount"].ToString(),
                                    TotAmount = item["TotAmount"].ToString(),
                                    Remark = item["Remark"].ToString(),
                                    ColorGreen = "false",
                                    ColorRed = "true"
                                };
                                dgdTotal.Items.Add(window_OutwareProductViewInsert);
                                //dgdOutware.Items.Add(window_OutwareProductViewInsert);
                                i++;
                            }

                        }

                    }

                    if (i == 2) //총계라서
                    {
                        MessageBox.Show("조회결과가 없습니다.");
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

        #endregion


        #region 엑셀
        // 엑셀 버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib2 = new Lib();
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdOutware.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdOutware.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdOutware);
                    else
                        dt = lib2.DataGirdToDataTable(dgdOutware);

                    Name = dgdOutware.Name;
                    
                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
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
            lib2 = null;

        }

        #endregion

        //닫기 버튼 클릭./
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
                        break;
                    }
                }
                i++;
            }
        }


        //정렬.
        private void btnMultiSort_Click(object sender, RoutedEventArgs e)
        {
            PopUp.MultiLevelSort MLS = new PopUp.MultiLevelSort(dgdOutware);
            MLS.ShowDialog();

            if (MLS.DialogResult.HasValue)
            {
                string targetSortProperty = string.Empty;
                int targetColIndex;
                dgdOutware.Items.SortDescriptions.Clear();

                for (int x = 0; x < MLS.ColName.Count; x++)
                {
                    targetSortProperty = MLS.SortingProperty[x];
                    targetColIndex = MLS.ColIndex[x];
                    var targetCol = dgdOutware.Columns[targetColIndex];

                    if (targetSortProperty == "UP")
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Ascending));
                        targetCol.SortDirection = ListSortDirection.Ascending;
                    }
                    else
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Descending));
                        targetCol.SortDirection = ListSortDirection.Descending;
                    }
                }
                dgdOutware.Refresh();
            }
        }



        // 사용자 편의. 엔터키로 플러스파인더 호출.
        private void txtCustomer_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnCustomer_Click(null, null);
            }
        }
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnArticle_Click(null, null);
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

        #region 기타 메서드 모음

        // 천 단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천 단위 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return result;
        }

        // Int로 변환
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

        // 소수로 변환 가능한지 체크 이벤트
        private bool CheckConvertDouble(string str)
        {
            bool flag = false;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                if (Double.TryParse(str, out chkDouble) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 숫자로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 소수로 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

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






        #endregion

    }







    /// <summary>
    /// //////////////////////////////////////////////////////////////////////
    /// </summary>

    class Win_ord_OutWare_Product_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 값.    
        public string Depth { get; set; }
        public string OutwareID { get; set; }
        public string OutDate { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string OrderNo { get; set; }
        public string OrderID { get; set; }
        public string OutCustom { get; set; }

        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string Sabun { get; set; }

        public string WorkName { get; set; }

        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string LabelID { get; set; }
        public string LabelGubun { get; set; }

        public string FromLocName { get; set; }
        public string TOLocname { get; set; }
        public string OutClssname { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }

        public string UnitPrice { get; set; }
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string TotAmount { get; set; }
        public string Remark { get; set; }


        //순번
        public int NUM { get; set; }

        //컬러 칠하기
        public string ColorGreen { get; set; }
        public string ColorRed { get; set; }

        public string LotID { get; set; }
    }
}
