using WizMes_WooJung.PopUP;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;
using System.Windows.Shapes;
using System.Windows.Media;
using System.Threading;
using System.Threading.Tasks;

namespace WizMes_WooJung
{
    /// <summary>
    /// Win_ord_InOutSum_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_InOutSum_Q : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 그리드 셀렉트 도전(2018_08_09)
        int Clicked_row = 0;
        int Clicked_col = 0;
        List<Rectangle> PreRect = new List<Rectangle>();

        //전역변수는 이럴때 쓰는거 아니겠어??!!?
        private DataTable PeriodDataTable = null;
        private DataTable DaysDataTable = null;
        private DataTable MonthDataTable = null;
        private DataTable SpreadMonthDataTable = null;


        public Win_ord_InOutSum_Q()
        {
            InitializeComponent();
        }

        // 화면 첫 시작.
        private void Window_InOutTotalGrid_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }

        #region  첫 스텝 // 일자버튼 // 초기설정 // 조회용 체크박스 컨트롤 
        private void First_Step()
        {
            // 월별 가로집계 최근 3개월 지정하기.
            List<MonthChange> MC = new List<MonthChange>();
            MC.Add(new MonthChange()
            {
                H_MON1 = DateTime.Now.ToString("yyyy-MM"),
                H_MON2 = DateTime.Now.AddMonths(-1).ToString("yyyy-MM"),
                H_MON3 = DateTime.Now.AddMonths(-2).ToString("yyyy-MM"),
            });

            this.DataContext = MC;
            //////////////////////////////////////

            lblThisMonth.Content = DateTime.Now.ToString("yyyy-MM");
            lblMinusOneMonth.Content = DateTime.Now.AddMonths(-1).ToString("yyyy-MM");
            lblMinusTwoMonth.Content = DateTime.Now.AddMonths(-2).ToString("yyyy-MM");

            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtblMessage.Visibility = Visibility.Hidden;

            // 전부 노 체크 노 enable로 시작.
            txtCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            btnArticle.IsEnabled = false;
            cboInOutGubun.IsEnabled = false;
            cboInInspectGubun.IsEnabled = false;
        }

        // 어제.(전일)
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringYesterdayDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                dtpFromDate.SelectedDate = dtpFromDate.SelectedDate.Value.AddDays(-1);
                dtpToDate.SelectedDate = dtpFromDate.SelectedDate;
            }
            else
            {
                dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
                dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
            }
        }
        // 오늘(금일)
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        // 지난 달(전월)
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringLastMonthDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                dtpFromDate.SelectedDate = LastMonth1;
                dtpToDate.SelectedDate = LastMonth31;
            }

        }
        // 이번 달(금월)
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthDatetime();

            dtpFromDate.Text = receiver[0];
            dtpToDate.Text = receiver[1];
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
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
        }
        // 거래처
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
        // 거래처
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
        //입출고구분
        private void chkInOutGubun_Click(object sender, RoutedEventArgs e)
        {
            if (chkInOutGubun.IsChecked == true)
            {
                cboInOutGubun.IsEnabled = true;
                cboInOutGubun.Focus();
            }
            else { cboInOutGubun.IsEnabled = false; }
        }
        //입출고구분
        private void chkInOutGubun_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInOutGubun.IsChecked == true)
            {
                chkInOutGubun.IsChecked = false;
                cboInOutGubun.IsEnabled = false;
            }
            else
            {
                chkInOutGubun.IsChecked = true;
                cboInOutGubun.IsEnabled = true;
                cboInOutGubun.Focus();
            }
        }
        //입고검수구분
        private void chkInInspectGubun_Click(object sender, RoutedEventArgs e)
        {
            if (chkInInsepectGubun.IsChecked == true)
            {
                cboInInspectGubun.IsEnabled = true;
                cboInInspectGubun.Focus();
            }
            else { cboInInspectGubun.IsEnabled = false; }
        }
        //입고검수구분
        private void chkInInspectGubun_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInInsepectGubun.IsChecked == true)
            {
                chkInInsepectGubun.IsChecked = false;
                cboInInspectGubun.IsEnabled = false;
            }
            else
            {
                chkInInsepectGubun.IsChecked = true;
                cboInInspectGubun.IsEnabled = true;
                cboInInspectGubun.Focus();
            }
        }
        #endregion

        #region 콤보박스 세팅
        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboInOutGubun.Items.Clear();
            cboInInspectGubun.Items.Clear();

            string[] DirectCombo = new string[2];
            DirectCombo[0] = "Y";
            DirectCombo[1] = "합격";
            string[] DirectCombo1 = new string[2];
            DirectCombo1[0] = "N";
            DirectCombo1[1] = "불합격";

            List<string[]> DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInInspectGubun = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            DirectCombo = new string[2];
            DirectCombo[0] = "1";
            DirectCombo[1] = "입고";
            DirectCombo1 = new string[2];
            DirectCombo1[0] = "2";
            DirectCombo1[1] = "출고";

            DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInOutGubun = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            this.cboInOutGubun.ItemsSource = cbInOutGubun;
            this.cboInOutGubun.DisplayMemberPath = "code_name";
            this.cboInOutGubun.SelectedValuePath = "code_id";
            this.cboInOutGubun.SelectedIndex = 0;

            this.cboInInspectGubun.ItemsSource = cbInInspectGubun;
            this.cboInInspectGubun.DisplayMemberPath = "code_name";
            this.cboInInspectGubun.SelectedValuePath = "code_id";
            this.cboInInspectGubun.SelectedIndex = 0;

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
            pf.ReturnCode(txtArticle, 84, txtArticle.Text);
        }

        #endregion


        // 검색(조회) 버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            TabItem NowTI = tabconGrid.SelectedItem as TabItem;

            if (NowTI.Header.ToString() == "기간집계") { FillGrid_Period(); }
            else if (NowTI.Header.ToString() == "일일집계") { FillGrid_Day(); }
            else if (NowTI.Header.ToString() == "월별집계(세로)") { FillGrid_Month_V(); }
            else if (NowTI.Header.ToString() == "월별집계(가로)") { FillGrid_Month_H(); }

            for (int i = 0; i < 3; i++)
            {
                lib.Delay(10);
                Header_Content_Width_Adjust(NowTI.Header.ToString());
                lib.Delay(10);
                Re_Header_Content_Width_Adjust(NowTI.Header.ToString());
            }
        }

        #region 기간집계 조회
        //기간집계 조회
        private void FillGrid_Period()
        {
            //클리어 후 데이터를 뿌려줘야 하니까
            grdMerge_Period.Children.Clear();
            grdPeriod.Items.Clear();

            string SearchFromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string SearchToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");         //기준일자
            int ChkCustomID = 0;
            if (chkCustomer.IsChecked == true) { ChkCustomID = 1; }
            else { txtCustomer.Tag = ""; txtCustomer.Text = ""; }                                                      //거래처
            int ChkArticleID = 0;
            if (chkArticle.IsChecked == true) { ChkArticleID = 1; }
            else { txtArticle.Tag = ""; txtArticle.Text = ""; }                                                       //품명
            int nGubun = 0;
            if (chkInOutGubun.IsChecked == true)
            {
                if (cboInOutGubun.SelectedValue.ToString() == "1") { nGubun = 1; }
                if (cboInOutGubun.SelectedValue.ToString() == "2") { nGubun = 2; }
            }                                                                                   //입출고구분
            int nMainItem = 0;
            if (chkMainInterestItem.IsChecked == true) { nMainItem = 1; }                       //주요관심품목
            int nCustomItem = 0;
            if (chkCustomsEnrollItem.IsChecked == true) { nCustomItem = 1; }                    //거래처등록품목
            int chkInspect = 0;
            string sInspect = string.Empty;
            if (chkInInsepectGubun.IsChecked == true)
            {
                chkInspect = 1;
                if (cboInInspectGubun.SelectedValue.ToString() == "Y") { sInspect = "Y"; }
                if (cboInInspectGubun.SelectedValue.ToString() == "N") { sInspect = "N"; }
            }                                                                                   //입고검수구분

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", SearchFromDate);
                sqlParameter.Add("EDate", SearchToDate);
                sqlParameter.Add("ChkCustomID", ChkCustomID);
                sqlParameter.Add("CustomID", ChkCustomID == 1 ? txtCustomer.Tag.ToString() : "");
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0); //0);// ChkArticleID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : string.Empty);//""); //txtArticle.Tag.ToString());
                sqlParameter.Add("nGubun", nGubun);
                sqlParameter.Add("nMainItem", nMainItem);
                sqlParameter.Add("nCustomItem", nCustomItem);
                sqlParameter.Add("chkInspect", chkInspect);
                sqlParameter.Add("sInspect", sInspect);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Period", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    PeriodDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Period.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        Style cellStyleLeft = new Style(typeof(TextBox));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleCenter = new Style(typeof(TextBox));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleRight = new Style(typeof(TextBox));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleRight.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleRight.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Right));
                        cellStyleRight.Setters.Add(new Setter(TextBox.FocusableProperty, true));


                        string strOTemp = string.Empty;
                        string str_O_Temp = string.Empty;

                        string strPTemp = string.Empty;

                        int k;
                        int m = 1; //var textbox02용
                        int o = 1; //var textbox03용
                        int p = 1; //var textbox04용

                        var textbox03 = new TextBox();
                        var textbox04 = new TextBox();

                        for (int i = 0; drc.Count > i; i++)
                        {
                            k = 1 + i;
                            var Win_ord_InOutSum_QView_Insert = new Win_ord_InOutSum_QView()
                            {
                                P_NUM = k,
                                P_cls = drc[i]["cls"].ToString(),
                                P_Gbn = drc[i]["Gbn"].ToString(),
                                P_IODate = drc[i]["IODate"].ToString(),
                                P_CustomID = drc[i]["CustomID"].ToString(),
                                P_CustomName = drc[i]["CustomName"].ToString(),
                                P_BuyerArticleNo = drc[i]["BuyerArticleNo"].ToString(),
                                P_ArticleID = drc[i]["ArticleID"].ToString(),
                                P_Article = drc[i]["Article"].ToString(),
                                P_Roll = drc[i]["Roll"].ToString(),
                                P_Qty = drc[i]["Qty"].ToString().Split('.')[0].Trim(),
                                P_UnitClss = drc[i]["UnitClss"].ToString(),
                                P_UnitClssName = drc[i]["UnitClssName"].ToString(),
                                P_UnitPrice = drc[i]["UnitPrice"].ToString(),
                                P_PriceClss = drc[i]["PriceClss"].ToString(),
                                P_PriceClssName = drc[i]["PriceClssName"].ToString(),
                                P_Amount = drc[i]["Amount"].ToString(),
                                P_VatAmount = drc[i]["VatAmount"].ToString(),
                                P_TotAmount = drc[i]["TotAmount"].ToString(),
                                P_CustomRate = lib.returnNumStringTwo(drc[i]["CustomRate"].ToString()),
                                P_CustomRateOrder = drc[i]["CustomRateOrder"].ToString(),
                            };

                            Win_ord_InOutSum_QView_Insert.P_Roll = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_Roll);
                            Win_ord_InOutSum_QView_Insert.P_Qty = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_Qty);
                            Win_ord_InOutSum_QView_Insert.P_UnitPrice = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_UnitPrice);
                            Win_ord_InOutSum_QView_Insert.P_Amount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_Amount);
                            Win_ord_InOutSum_QView_Insert.P_VatAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_VatAmount);
                            Win_ord_InOutSum_QView_Insert.P_TotAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_TotAmount);
                            Win_ord_InOutSum_QView_Insert.P_CustomRate = lib.returnNumString(Win_ord_InOutSum_QView_Insert.P_CustomRate);

                            if (Win_ord_InOutSum_QView_Insert.P_Article == "zzzzzzzzz")
                            {
                                Win_ord_InOutSum_QView_Insert.P_Article = "";
                            }

                            grdPeriod.Items.Add(Win_ord_InOutSum_QView_Insert);

                            RowDefinition row = new RowDefinition();
                            row.Height = GridLength.Auto;
                            grdMerge_Period.RowDefinitions.Add(row);


                            //순번
                            var textbox01 = new TextBox();
                            textbox01.Text = k.ToString();
                            textbox01.Style = cellStyleCenter;
                            textbox01.IsReadOnly = true;
                            textbox01.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox01.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox01);
                            Grid.SetRow(textbox01, i);
                            Grid.SetColumn(textbox01, 0);


                            //기간
                            var textbox02 = new TextBox();
                            textbox02.Text = SearchFromDate + "~" + SearchToDate;
                            textbox02.Style = cellStyleCenter;
                            textbox02.IsReadOnly = true;
                            textbox02.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox02.PreviewMouseUp += Textbox_Period_previewmouseup;

                            if (k != drc.Count) { m++; }

                            else if (k == drc.Count)
                            {
                                grdMerge_Period.Children.Add(textbox02);
                                Grid.SetRow(textbox02, 0);
                                Grid.SetRowSpan(textbox02, m);
                                Grid.SetColumn(textbox02, 1);

                                m = 1;
                            }

                            // textbox 03 시작지점.
                            if (Win_ord_InOutSum_QView_Insert.P_Gbn == "1")
                            {
                                if (Win_ord_InOutSum_QView_Insert.P_cls == "0")
                                {
                                    str_O_Temp = "입고 총계";
                                    Win_ord_InOutSum_QView_Insert.P_Gbn = str_O_Temp;
                                }
                                else
                                {
                                    str_O_Temp = "입고";
                                    Win_ord_InOutSum_QView_Insert.P_Gbn = str_O_Temp;
                                }
                            }
                            else if (Win_ord_InOutSum_QView_Insert.P_Gbn == "2")
                            {
                                if (Win_ord_InOutSum_QView_Insert.P_cls == "0")
                                {
                                    str_O_Temp = "출고 총계";
                                    Win_ord_InOutSum_QView_Insert.P_Gbn = str_O_Temp;
                                }
                                else
                                {
                                    //if (Win_ord_InOutSum_QView_Insert)
                                    //{
                                    //    str_O_Temp = "예외출고";
                                    //    Win_ord_InOutSum_QView_Insert.P_Gbn = str_O_Temp;
                                    //}
                                    //else {
                                    str_O_Temp = "출고";
                                    Win_ord_InOutSum_QView_Insert.P_Gbn = str_O_Temp;
                                    //}
                                }
                            }

                            if (i == 0)
                            {
                                //구분
                                textbox03.Text = str_O_Temp;
                                textbox03.Style = cellStyleCenter;
                                textbox03.IsReadOnly = true;
                                textbox03.PreviewMouseDown += Textbox_Period_previewmousedown;
                                textbox03.PreviewMouseUp += Textbox_Period_previewmouseup;

                                grdMerge_Period.Children.Add(textbox03);
                                Grid.SetRow(textbox03, i);
                                Grid.SetColumn(textbox03, 2);
                                strOTemp = str_O_Temp;
                            }

                            else if (k == drc.Count)
                            {
                                o++;
                                Grid.SetRowSpan(textbox03, o);
                            }

                            else
                            {
                                if (strOTemp.Equals(str_O_Temp))
                                {
                                    o++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox03, o);
                                    o = 1;

                                    //구분
                                    textbox03 = new TextBox();
                                    textbox03.Text = str_O_Temp;
                                    textbox03.Style = cellStyleCenter;
                                    textbox03.IsReadOnly = true;
                                    textbox03.PreviewMouseDown += Textbox_Period_previewmousedown;
                                    textbox03.PreviewMouseUp += Textbox_Period_previewmouseup;

                                    grdMerge_Period.Children.Add(textbox03);
                                    Grid.SetRow(textbox03, i);
                                    Grid.SetColumn(textbox03, 2);
                                    strOTemp = str_O_Temp;
                                }
                            }


                            // textbox 04 시작지점.
                            if (i == 0)
                            {
                                //거래처
                                textbox04.Text = Win_ord_InOutSum_QView_Insert.P_CustomName;
                                textbox04.Style = cellStyleLeft;
                                textbox04.IsReadOnly = true;
                                textbox04.PreviewMouseDown += Textbox_Period_previewmousedown;
                                textbox04.PreviewMouseUp += Textbox_Period_previewmouseup;

                                grdMerge_Period.Children.Add(textbox04);
                                Grid.SetRow(textbox04, i);
                                Grid.SetColumn(textbox04, 3);
                                strPTemp = Win_ord_InOutSum_QView_Insert.P_CustomName;
                            }
                            else if (k == drc.Count)
                            {
                                p++;
                                Grid.SetRowSpan(textbox04, p);
                            }
                            else
                            {
                                if (strPTemp.Equals(Win_ord_InOutSum_QView_Insert.P_CustomName))
                                {
                                    p++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox04, p);
                                    p = 1;

                                    //거래처
                                    textbox04 = new TextBox();
                                    textbox04.Text = Win_ord_InOutSum_QView_Insert.P_CustomName;
                                    textbox04.Style = cellStyleLeft;
                                    textbox04.IsReadOnly = true;
                                    textbox04.PreviewMouseDown += Textbox_Period_previewmousedown;
                                    textbox04.PreviewMouseUp += Textbox_Period_previewmouseup;

                                    grdMerge_Period.Children.Add(textbox04);
                                    Grid.SetRow(textbox04, i);
                                    Grid.SetColumn(textbox04, 3);
                                    strPTemp = Win_ord_InOutSum_QView_Insert.P_CustomName;
                                }
                            }

                            //품번
                            var textbox05 = new TextBox();
                            textbox05.Text = Win_ord_InOutSum_QView_Insert.P_BuyerArticleNo;
                            textbox05.Style = cellStyleLeft;
                            textbox05.IsReadOnly = true;
                            textbox05.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox05.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox05);
                            Grid.SetRow(textbox05, i);
                            Grid.SetColumn(textbox05, 4);




                            //품명
                            var textbox06 = new TextBox();
                            textbox06.Text = Win_ord_InOutSum_QView_Insert.P_Article;
                            textbox06.Style = cellStyleLeft;
                            textbox06.IsReadOnly = true;
                            textbox06.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox06.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox06);
                            Grid.SetRow(textbox06, i);
                            Grid.SetColumn(textbox06, 5);

                            //품명코드
                            var textbox07 = new TextBox();
                            textbox07.Text = Win_ord_InOutSum_QView_Insert.P_ArticleID;
                            textbox07.Style = cellStyleCenter;
                            textbox07.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox07.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox07);
                            Grid.SetRow(textbox07, i);
                            Grid.SetColumn(textbox07, 6);

                            //건수
                            var textbox08 = new TextBox();
                            textbox08.Text = Win_ord_InOutSum_QView_Insert.P_Roll;
                            textbox08.Style = cellStyleRight;
                            textbox08.IsReadOnly = true;
                            textbox08.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox08.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox08);
                            Grid.SetRow(textbox08, i);
                            Grid.SetColumn(textbox08, 7);

                            //수량
                            var textbox09 = new TextBox();
                            textbox09.Text = Win_ord_InOutSum_QView_Insert.P_Qty;
                            textbox09.Style = cellStyleRight;
                            textbox09.IsReadOnly = true;
                            textbox09.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox09.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox09);
                            Grid.SetRow(textbox09, i);
                            Grid.SetColumn(textbox09, 8);

                            //단위
                            var textbox10 = new TextBox();
                            textbox10.Text = Win_ord_InOutSum_QView_Insert.P_UnitClssName;
                            textbox10.Style = cellStyleCenter;
                            textbox10.IsReadOnly = true;
                            textbox10.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox10.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox10);
                            Grid.SetRow(textbox10, i);
                            Grid.SetColumn(textbox10, 9);


                            //점유율
                            var textbox11 = new TextBox();
                            double tb12 = Win_ord_InOutSum_QView_Insert.P_CustomRate.Equals("") ? 0 : Convert.ToDouble(Win_ord_InOutSum_QView_Insert.P_CustomRate);
                            if ((Win_ord_InOutSum_QView_Insert.P_ArticleID == "") || (Win_ord_InOutSum_QView_Insert.P_ArticleID == "거래처계"))
                            {
                                textbox11.Text = String.Format("{0:0.##;0:#;#}", tb12);
                            }
                            else { textbox11.Text = ""; }
                            textbox11.Style = cellStyleRight;
                            textbox11.IsReadOnly = true;
                            textbox11.PreviewMouseDown += Textbox_Period_previewmousedown;
                            textbox11.PreviewMouseUp += Textbox_Period_previewmouseup;

                            grdMerge_Period.Children.Add(textbox11);
                            Grid.SetRow(textbox11, i);
                            Grid.SetColumn(textbox11, 10);

                        }
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

        #region 일일집계 조회
        //일일집계 조회
        private void FillGrid_Day()
        {
            grdMerge_Days.Children.Clear();
            grdMergeDays.Items.Clear();

            string SearchFromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string SearchToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");         //기준일자
            int ChkCustomID = 0;
            if (chkCustomer.IsChecked == true) { ChkCustomID = 1; }
            else { txtCustomer.Tag = ""; txtCustomer.Text = ""; }                                                      //거래처
            int ChkArticleID = 0;
            if (chkArticle.IsChecked == true) { ChkArticleID = 1; }
            else { txtArticle.Tag = ""; txtArticle.Text = ""; }                                                       //품명
            int nGubun = 0;
            if (chkInOutGubun.IsChecked == true)
            {
                if (cboInOutGubun.SelectedValue.ToString() == "1") { nGubun = 1; }
                if (cboInOutGubun.SelectedValue.ToString() == "2") { nGubun = 2; }
            }                                                                                   //입출고구분
            int nMainItem = 0;
            if (chkMainInterestItem.IsChecked == true) { nMainItem = 1; }                       //주요관심품목
            int nCustomItem = 0;
            if (chkCustomsEnrollItem.IsChecked == true) { nCustomItem = 1; }                    //거래처등록품목
            int chkInspect = 0;
            string sInspect = string.Empty;
            if (chkInInsepectGubun.IsChecked == true)
            {
                chkInspect = 1;
                if (cboInInspectGubun.SelectedValue.ToString() == "Y") { sInspect = "Y"; }
                if (cboInInspectGubun.SelectedValue.ToString() == "N") { sInspect = "N"; }
            }                                                                                   //입고검수구분

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", SearchFromDate);
                sqlParameter.Add("EDate", SearchToDate);
                sqlParameter.Add("ChkCustomID", ChkCustomID);
                sqlParameter.Add("CustomID", txtCustomer.Tag.ToString());
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0); // ChkArticleID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : string.Empty); // txtArticle.Tag.ToString());
                sqlParameter.Add("nGubun", nGubun);
                sqlParameter.Add("nMainItem", nMainItem);
                sqlParameter.Add("nCustomItem", nCustomItem);
                sqlParameter.Add("chkInspect", chkInspect);
                sqlParameter.Add("sInspect", sInspect);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Day", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    DaysDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {

                        //grdMerge_Days.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        Style cellStyleLeft = new Style(typeof(TextBox));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleCenter = new Style(typeof(TextBox));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleRight = new Style(typeof(TextBox));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleRight.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleRight.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Right));
                        cellStyleRight.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        int k;

                        var textbox02 = new TextBox();
                        int m = 1;                      //var textbox02용
                        string strMTemp = string.Empty; //var textbox02용

                        var textbox03 = new TextBox();
                        int o = 1;                      //var textbox03용
                        string strOTemp = string.Empty; //var textbox03용

                        var textbox04 = new TextBox();
                        int p = 1;                      //var textbox04용
                        string strPTemp = string.Empty; //var textbox04용

                        for (int i = 0; drc.Count > i; i++)
                        {
                            k = 1 + i;
                            var Win_ord_InOutSum_QView_Insert = new Win_ord_InOutSum_QView()
                            {
                                D_NUM = k,
                                D_cls = drc[i]["cls"].ToString(),
                                D_Gbn = drc[i]["Gbn"].ToString(),
                                D_IODate = drc[i]["IODate"].ToString(),
                                D_CustomID = drc[i]["CustomID"].ToString(),
                                D_CustomName = drc[i]["CustomName"].ToString(),
                                D_ArticleID = drc[i]["ArticleID"].ToString(),
                                D_BuyerArticleNo = drc[i]["BuyerArticleNo"].ToString(),
                                D_Article = drc[i]["Article"].ToString(),
                                D_Roll = drc[i]["Roll"].ToString(),
                                D_Qty = drc[i]["Qty"].ToString().Split('.')[0].Trim(),
                                D_UnitClss = drc[i]["UnitClss"].ToString(),
                                D_UnitClssName = drc[i]["UnitClssName"].ToString(),
                                D_UnitPrice = drc[i]["UnitPrice"].ToString(),
                                D_PriceClss = drc[i]["PriceClss"].ToString(),
                                D_PriceClssName = drc[i]["PriceClssName"].ToString(),
                                D_Amount = drc[i]["Amount"].ToString(),
                                D_VatAmount = drc[i]["VatAmount"].ToString().Split('.')[0].Trim(),
                                D_TotAmount = drc[i]["TotAmount"].ToString().Split('.')[0].Trim(),
                                D_CustomRate = lib.returnNumStringTwo(drc[i]["CustomRate"].ToString()),
                            };

                            Win_ord_InOutSum_QView_Insert.D_Roll = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_Roll);
                            Win_ord_InOutSum_QView_Insert.D_Qty = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_Qty);
                            Win_ord_InOutSum_QView_Insert.D_UnitPrice = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_UnitPrice);
                            Win_ord_InOutSum_QView_Insert.D_Amount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_Amount);
                            Win_ord_InOutSum_QView_Insert.D_VatAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_VatAmount);
                            Win_ord_InOutSum_QView_Insert.D_TotAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_TotAmount);
                            Win_ord_InOutSum_QView_Insert.D_CustomRate = lib.returnNumString(Win_ord_InOutSum_QView_Insert.D_CustomRate);

                            if (Win_ord_InOutSum_QView_Insert.D_IODate == "00000000")
                            {
                                if (Win_ord_InOutSum_QView_Insert.D_Gbn == "1")
                                {
                                    Win_ord_InOutSum_QView_Insert.D_IODate = "입고총계";
                                }
                                else if (Win_ord_InOutSum_QView_Insert.D_Gbn == "2")
                                {
                                    Win_ord_InOutSum_QView_Insert.D_IODate = "출고총계";
                                }
                            }

                            if (Win_ord_InOutSum_QView_Insert.D_Gbn == "1")
                            {
                                Win_ord_InOutSum_QView_Insert.D_Gbn = "입고";
                            }
                            else if (Win_ord_InOutSum_QView_Insert.D_Gbn == "2")
                            {
                                Win_ord_InOutSum_QView_Insert.D_Gbn = "출고";
                            }

                            if (Win_ord_InOutSum_QView_Insert.D_CustomName == "zzzzzzzzzzz")
                            {
                                Win_ord_InOutSum_QView_Insert.D_CustomName = "";
                            }
                            else if (Win_ord_InOutSum_QView_Insert.D_CustomName == "zzzzzzzzzzzzz")  //살짜기 더 김.
                            {
                                Win_ord_InOutSum_QView_Insert.D_CustomName = "일계";
                            }

                            if (Win_ord_InOutSum_QView_Insert.D_ArticleID == "99999")
                            {
                                Win_ord_InOutSum_QView_Insert.D_ArticleID = "";
                            }

                            if (Win_ord_InOutSum_QView_Insert.D_Article == "zzzzzzzzz")
                            {
                                Win_ord_InOutSum_QView_Insert.D_Article = "";
                            }

                            if (Win_ord_InOutSum_QView_Insert.D_Article == "zzzzzzzzzz")
                            {
                                Win_ord_InOutSum_QView_Insert.D_Article = "";
                            }

                            grdMergeDays.Items.Add(Win_ord_InOutSum_QView_Insert);

                            RowDefinition row = new RowDefinition();
                            row.Height = GridLength.Auto;
                            grdMerge_Days.RowDefinitions.Add(row);


                            //순번
                            var textbox01 = new TextBox();
                            textbox01.Text = k.ToString();
                            textbox01.Style = cellStyleCenter;
                            textbox01.IsReadOnly = true;
                            textbox01.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox01.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox01);
                            Grid.SetRow(textbox01, i);
                            Grid.SetColumn(textbox01, 0);


                            // textbox 02 시작지점.
                            if (i == 0)
                            {
                                //일자
                                textbox02.Text = Win_ord_InOutSum_QView_Insert.D_IODate;
                                textbox02.Style = cellStyleCenter;
                                textbox02.IsReadOnly = true;
                                textbox02.PreviewMouseDown += Textbox_Days_previewmousedown;
                                textbox02.PreviewMouseUp += Textbox_Days_previewmouseup;

                                grdMerge_Days.Children.Add(textbox02);
                                Grid.SetRow(textbox02, i);
                                Grid.SetColumn(textbox02, 1);
                                strMTemp = Win_ord_InOutSum_QView_Insert.D_IODate;
                            }
                            else if (k == drc.Count)
                            {
                                m++;
                                Grid.SetRowSpan(textbox02, m);
                            }
                            else
                            {
                                if (strMTemp.Equals(Win_ord_InOutSum_QView_Insert.D_IODate))
                                {
                                    m++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox02, m);
                                    m = 1;

                                    //일자
                                    textbox02 = new TextBox();
                                    textbox02.Text = Win_ord_InOutSum_QView_Insert.D_IODate;
                                    textbox02.Style = cellStyleCenter;
                                    textbox02.IsReadOnly = true;
                                    textbox02.PreviewMouseDown += Textbox_Days_previewmousedown;
                                    textbox02.PreviewMouseUp += Textbox_Days_previewmouseup;

                                    grdMerge_Days.Children.Add(textbox02);
                                    Grid.SetRow(textbox02, i);
                                    Grid.SetColumn(textbox02, 1);
                                    strMTemp = Win_ord_InOutSum_QView_Insert.D_IODate;
                                }
                            }


                            // textbox 03 시작지점.
                            if (i == 0)
                            {
                                //구분
                                textbox03.Text = Win_ord_InOutSum_QView_Insert.D_Gbn;
                                textbox03.Style = cellStyleCenter;
                                textbox03.IsReadOnly = true;
                                textbox03.PreviewMouseDown += Textbox_Days_previewmousedown;
                                textbox03.PreviewMouseUp += Textbox_Days_previewmouseup;

                                grdMerge_Days.Children.Add(textbox03);
                                Grid.SetRow(textbox03, i);
                                Grid.SetColumn(textbox03, 2);
                                strOTemp = Win_ord_InOutSum_QView_Insert.D_Gbn;
                            }
                            else if (k == drc.Count)
                            {
                                o++;
                                Grid.SetRowSpan(textbox03, o);
                            }
                            else
                            {
                                if (strOTemp.Equals(Win_ord_InOutSum_QView_Insert.D_Gbn))
                                {
                                    o++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox03, o);
                                    o = 1;

                                    //구분
                                    textbox03 = new TextBox();
                                    textbox03.Text = Win_ord_InOutSum_QView_Insert.D_Gbn;
                                    textbox03.Style = cellStyleCenter;
                                    textbox03.IsReadOnly = true;
                                    textbox03.PreviewMouseDown += Textbox_Days_previewmousedown;
                                    textbox03.PreviewMouseUp += Textbox_Days_previewmouseup;

                                    grdMerge_Days.Children.Add(textbox03);
                                    Grid.SetRow(textbox03, i);
                                    Grid.SetColumn(textbox03, 2);
                                    strOTemp = Win_ord_InOutSum_QView_Insert.D_Gbn;
                                }
                            }


                            // textbox 04 시작지점.
                            if (i == 0)
                            {
                                //거래처
                                textbox04.Text = Win_ord_InOutSum_QView_Insert.D_CustomName;
                                textbox04.Style = cellStyleLeft;
                                textbox04.IsReadOnly = true;
                                textbox04.PreviewMouseDown += Textbox_Days_previewmousedown;
                                textbox04.PreviewMouseUp += Textbox_Days_previewmouseup;

                                grdMerge_Days.Children.Add(textbox04);
                                Grid.SetRow(textbox04, i);
                                Grid.SetColumn(textbox04, 3);
                                strPTemp = Win_ord_InOutSum_QView_Insert.D_CustomName;
                            }
                            else if (k == drc.Count)
                            {
                                p++;
                                Grid.SetRowSpan(textbox04, p);
                            }
                            else
                            {
                                if (strPTemp.Equals(Win_ord_InOutSum_QView_Insert.D_CustomName))
                                {
                                    p++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox04, p);
                                    p = 1;

                                    //거래처
                                    textbox04 = new TextBox();
                                    textbox04.Text = Win_ord_InOutSum_QView_Insert.D_CustomName;
                                    textbox04.Style = cellStyleLeft;
                                    textbox04.IsReadOnly = true;
                                    textbox04.PreviewMouseDown += Textbox_Days_previewmousedown;
                                    textbox04.PreviewMouseUp += Textbox_Days_previewmouseup;

                                    grdMerge_Days.Children.Add(textbox04);
                                    Grid.SetRow(textbox04, i);
                                    Grid.SetColumn(textbox04, 3);
                                    strPTemp = Win_ord_InOutSum_QView_Insert.D_CustomName;
                                }
                            }
                            //품번
                            var textbox05 = new TextBox();
                            textbox05.Text = Win_ord_InOutSum_QView_Insert.D_BuyerArticleNo;
                            textbox05.Style = cellStyleLeft;
                            textbox05.IsReadOnly = true;
                            textbox05.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox05.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox05);
                            Grid.SetRow(textbox05, i);
                            Grid.SetColumn(textbox05, 4);


                            //품명
                            var textbox06 = new TextBox();
                            textbox06.Text = Win_ord_InOutSum_QView_Insert.D_Article;
                            textbox06.Style = cellStyleLeft;
                            textbox06.IsReadOnly = true;
                            textbox06.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox06.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox06);
                            Grid.SetRow(textbox06, i);
                            Grid.SetColumn(textbox06, 5);


                            //건수
                            var textbox07 = new TextBox();
                            textbox07.Text = Win_ord_InOutSum_QView_Insert.D_Roll;
                            textbox07.Style = cellStyleRight;
                            textbox07.IsReadOnly = true;
                            textbox07.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox07.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox07);
                            Grid.SetRow(textbox07, i);
                            Grid.SetColumn(textbox07, 6);


                            //수량
                            var textbox08 = new TextBox();
                            textbox08.Text = Win_ord_InOutSum_QView_Insert.D_Qty;
                            textbox08.Style = cellStyleRight;
                            textbox08.IsReadOnly = true;
                            textbox08.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox08.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox08);
                            Grid.SetRow(textbox08, i);
                            Grid.SetColumn(textbox08, 7);

                            //단위
                            var textbox09 = new TextBox();
                            textbox09.Text = Win_ord_InOutSum_QView_Insert.D_UnitClssName;
                            textbox09.Style = cellStyleCenter;
                            textbox09.IsReadOnly = true;
                            textbox09.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox09.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox09);
                            Grid.SetRow(textbox09, i);
                            Grid.SetColumn(textbox09, 8);


                            //금액
                            var textbox10 = new TextBox();
                            textbox10.Text = Win_ord_InOutSum_QView_Insert.D_Amount;
                            textbox10.Style = cellStyleRight;
                            textbox10.IsReadOnly = true;
                            textbox10.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox10.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox10);
                            Grid.SetRow(textbox10, i);
                            Grid.SetColumn(textbox10, 9);


                            //부가세
                            var textbox11 = new TextBox();
                            textbox11.Text = Win_ord_InOutSum_QView_Insert.D_VatAmount;
                            textbox11.Style = cellStyleRight;
                            textbox11.IsReadOnly = true;
                            textbox11.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox11.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox11);
                            Grid.SetRow(textbox11, i);
                            Grid.SetColumn(textbox11, 10);


                            //합계금액
                            var textbox12 = new TextBox();
                            textbox12.Text = Win_ord_InOutSum_QView_Insert.D_TotAmount;
                            textbox12.Style = cellStyleRight;
                            textbox12.IsReadOnly = true;
                            textbox12.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox12.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox12);
                            Grid.SetRow(textbox12, i);
                            Grid.SetColumn(textbox12, 11);

                            //점유율
                            var textbox13 = new TextBox();
                            double tb13 = Win_ord_InOutSum_QView_Insert.D_CustomRate.Equals("") ? 0 : Convert.ToDouble(Win_ord_InOutSum_QView_Insert.D_CustomRate);
                            /*Convert.ToDouble(Win_ord_InOutSum_QView_Insert.D_CustomRate);*/
                            textbox13.Text = String.Format("{0:0.##;0:#;#}", tb13);
                            textbox13.Style = cellStyleRight;
                            textbox13.IsReadOnly = true;
                            textbox13.PreviewMouseDown += Textbox_Days_previewmousedown;
                            textbox13.PreviewMouseUp += Textbox_Days_previewmouseup;

                            grdMerge_Days.Children.Add(textbox13);
                            Grid.SetRow(textbox13, i);
                            Grid.SetColumn(textbox13, 12);
                        }
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

        #region 월별집계 (세로) 조회
        //월별집계 (세로) 조회
        private void FillGrid_Month_V()
        {
            grdMerge_Month_V.Children.Clear();
            grdMergeMonth_V.Items.Clear();

            string SearchFromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
            string SearchToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");         //기준일자
            int ChkCustomID = 0;
            if (chkCustomer.IsChecked == true) { ChkCustomID = 1; }
            else { txtCustomer.Tag = ""; txtCustomer.Text = ""; }                                                      //거래처
            int ChkArticleID = 0;
            if (chkArticle.IsChecked == true) { ChkArticleID = 1; }
            else { txtArticle.Tag = ""; txtArticle.Text = ""; }                                                       //품명
            int nGubun = 0;
            if (chkInOutGubun.IsChecked == true)
            {
                if (cboInOutGubun.SelectedValue.ToString() == "1") { nGubun = 1; }
                if (cboInOutGubun.SelectedValue.ToString() == "2") { nGubun = 2; }
            }                                                                                   //입출고구분
            int nMainItem = 0;
            if (chkMainInterestItem.IsChecked == true) { nMainItem = 1; }                       //주요관심품목
            int nCustomItem = 0;
            if (chkCustomsEnrollItem.IsChecked == true) { nCustomItem = 1; }                    //거래처등록품목
            int chkInspect = 0;
            string sInspect = string.Empty;
            if (chkInInsepectGubun.IsChecked == true)
            {
                chkInspect = 1;
                if (cboInInspectGubun.SelectedValue.ToString() == "Y") { sInspect = "Y"; }
                if (cboInInspectGubun.SelectedValue.ToString() == "N") { sInspect = "N"; }
            }                                                                                   //입고검수구분

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", SearchFromDate);
                sqlParameter.Add("EDate", SearchToDate);
                sqlParameter.Add("ChkCustomID", ChkCustomID);
                sqlParameter.Add("CustomID", txtCustomer.Tag.ToString());
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);// ChkArticleID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : string.Empty);//txtArticle.Tag.ToString());
                sqlParameter.Add("nGubun", nGubun);
                sqlParameter.Add("nMainItem", nMainItem);
                sqlParameter.Add("nCustomItem", nCustomItem);
                sqlParameter.Add("chkInspect", chkInspect);
                sqlParameter.Add("sInspect", sInspect);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Month", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    MonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_V.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        Style cellStyleLeft = new Style(typeof(TextBox));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleCenter = new Style(typeof(TextBox));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleRight = new Style(typeof(TextBox));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleRight.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleRight.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Right));
                        cellStyleRight.Setters.Add(new Setter(TextBox.FocusableProperty, true));


                        int k;

                        var textbox02 = new TextBox();
                        int m = 1;                      //var textbox02용
                        string strMTemp = string.Empty; //var textbox02용

                        var textbox03 = new TextBox();
                        int o = 1;                      //var textbox03용
                        string strOTemp = string.Empty; //var textbox03용

                        var textbox04 = new TextBox();
                        int p = 1;                      //var textbox04용
                        string strPTemp = string.Empty; //var textbox04용


                        for (int i = 0; drc.Count > i; i++)
                        {
                            k = 1 + i;
                            var Win_ord_InOutSum_QView_Insert = new Win_ord_InOutSum_QView()
                            {
                                V_NUM = k,
                                V_cls = drc[i]["cls"].ToString(),
                                V_Gbn = drc[i]["Gbn"].ToString(),
                                V_IODate = drc[i]["IODate"].ToString(),
                                V_CustomID = drc[i]["CustomID"].ToString(),
                                V_CustomName = drc[i]["CustomName"].ToString(),
                                V_BuyerArticleNo = drc[i]["BuyerArticleNo"].ToString(),
                                V_ArticleID = drc[i]["ArticleID"].ToString(),
                                V_Article = drc[i]["Article"].ToString(),
                                V_Roll = drc[i]["Roll"].ToString(),
                                V_Qty = drc[i]["Qty"].ToString().Split('.')[0].Trim(),
                                V_UnitClss = drc[i]["UnitClss"].ToString(),
                                V_UnitClssName = drc[i]["UnitClssName"].ToString(),
                                V_UnitPrice = drc[i]["UnitPrice"].ToString(),
                                V_PriceClss = drc[i]["PriceClss"].ToString(),
                                V_PriceClssName = drc[i]["PriceClssName"].ToString(),
                                V_Amount = drc[i]["Amount"].ToString(),
                                V_VatAmount = drc[i]["VatAmount"].ToString().Split('.')[0].Trim(),
                                V_TotAmount = drc[i]["TotAmount"].ToString().Split('.')[0].Trim(),
                                V_CustomRate = lib.returnNumStringTwo(drc[i]["CustomRate"].ToString()),
                                V_CustomRateOrder = drc[i]["CustomRateOrder"].ToString(),
                                V_RN = drc[i]["RN"].ToString()
                            };

                            Win_ord_InOutSum_QView_Insert.V_Roll = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_Roll);
                            Win_ord_InOutSum_QView_Insert.V_Qty = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_Qty);
                            Win_ord_InOutSum_QView_Insert.V_UnitPrice = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_UnitPrice);
                            Win_ord_InOutSum_QView_Insert.V_Amount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_Amount);
                            Win_ord_InOutSum_QView_Insert.V_VatAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_VatAmount);
                            Win_ord_InOutSum_QView_Insert.V_TotAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_TotAmount);
                            Win_ord_InOutSum_QView_Insert.V_CustomRate = lib.returnNumString(Win_ord_InOutSum_QView_Insert.V_CustomRate);

                            if (Win_ord_InOutSum_QView_Insert.V_IODate == "00000000")
                            {
                                if (Win_ord_InOutSum_QView_Insert.V_Gbn == "2")
                                {
                                    Win_ord_InOutSum_QView_Insert.V_IODate = "출고 총계";
                                }
                                else if (Win_ord_InOutSum_QView_Insert.V_Gbn == "1")
                                {
                                    Win_ord_InOutSum_QView_Insert.V_IODate = "입고 총계";
                                }
                            }

                            if (Win_ord_InOutSum_QView_Insert.V_Gbn == "1")
                            {
                                Win_ord_InOutSum_QView_Insert.V_Gbn = "입고";
                            }
                            else if (Win_ord_InOutSum_QView_Insert.V_Gbn == "2")
                            {
                                Win_ord_InOutSum_QView_Insert.V_Gbn = "출고";
                            }

                            if (Win_ord_InOutSum_QView_Insert.V_Article == "zzzzzzzzz")
                            {
                                if ((Win_ord_InOutSum_QView_Insert.V_IODate == "출고 총계") ||
                                        (Win_ord_InOutSum_QView_Insert.V_IODate == "입고 총계"))
                                {
                                    Win_ord_InOutSum_QView_Insert.V_Article = "";
                                }
                                else
                                {
                                    Win_ord_InOutSum_QView_Insert.V_Article = "거래처 계";
                                }
                            }

                            if (Win_ord_InOutSum_QView_Insert.V_cls == "1")
                            {
                                Win_ord_InOutSum_QView_Insert.V_CustomRate = "0";
                            }

                            grdMergeMonth_V.Items.Add(Win_ord_InOutSum_QView_Insert);

                            RowDefinition row = new RowDefinition();
                            row.Height = GridLength.Auto;
                            grdMerge_Month_V.RowDefinitions.Add(row);

                            //순번
                            var textbox01 = new TextBox();
                            textbox01.Text = k.ToString();
                            textbox01.Style = cellStyleCenter;
                            textbox01.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox01);
                            Grid.SetRow(textbox01, i);
                            Grid.SetColumn(textbox01, 0);


                            // textbox 02 시작지점.
                            if (i == 0)
                            {
                                //월 
                                textbox02.Text = Win_ord_InOutSum_QView_Insert.V_IODate;
                                textbox02.Style = cellStyleCenter;
                                textbox02.IsReadOnly = true;

                                grdMerge_Month_V.Children.Add(textbox02);
                                Grid.SetRow(textbox02, i);
                                Grid.SetColumn(textbox02, 1);
                                strMTemp = Win_ord_InOutSum_QView_Insert.V_IODate;
                            }
                            else if (k == drc.Count)
                            {
                                m++;
                                Grid.SetRowSpan(textbox02, m);
                            }
                            else
                            {
                                if (strMTemp.Equals(Win_ord_InOutSum_QView_Insert.V_IODate))
                                {
                                    m++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox02, m);
                                    m = 1;

                                    //월
                                    textbox02 = new TextBox();
                                    textbox02.Text = Win_ord_InOutSum_QView_Insert.V_IODate;
                                    textbox02.Style = cellStyleCenter;
                                    textbox02.IsReadOnly = true;

                                    grdMerge_Month_V.Children.Add(textbox02);
                                    Grid.SetRow(textbox02, i);
                                    Grid.SetColumn(textbox02, 1);
                                    strMTemp = Win_ord_InOutSum_QView_Insert.V_IODate;
                                }
                            }


                            // textbox 03 시작지점.
                            if (i == 0)
                            {
                                //구분
                                textbox03.Text = Win_ord_InOutSum_QView_Insert.V_Gbn;
                                textbox03.Style = cellStyleCenter;
                                textbox03.IsReadOnly = true;

                                grdMerge_Month_V.Children.Add(textbox03);
                                Grid.SetRow(textbox03, i);
                                Grid.SetColumn(textbox03, 2);
                                strOTemp = Win_ord_InOutSum_QView_Insert.V_Gbn;
                            }
                            else if (k == drc.Count)
                            {
                                o++;
                                Grid.SetRowSpan(textbox03, o);
                            }
                            else
                            {
                                if (strOTemp.Equals(Win_ord_InOutSum_QView_Insert.V_Gbn))
                                {
                                    o++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox03, o);
                                    o = 1;

                                    //구분
                                    textbox03 = new TextBox();
                                    textbox03.Text = Win_ord_InOutSum_QView_Insert.V_Gbn;
                                    textbox03.Style = cellStyleCenter;
                                    textbox03.IsReadOnly = true;

                                    grdMerge_Month_V.Children.Add(textbox03);
                                    Grid.SetRow(textbox03, i);
                                    Grid.SetColumn(textbox03, 2);
                                    strOTemp = Win_ord_InOutSum_QView_Insert.V_Gbn;
                                }
                            }


                            // textbox 04 시작지점.
                            if (i == 0)
                            {
                                //거래처
                                textbox04.Text = Win_ord_InOutSum_QView_Insert.V_CustomName;
                                textbox04.Style = cellStyleLeft;
                                textbox04.IsReadOnly = true;

                                grdMerge_Month_V.Children.Add(textbox04);
                                Grid.SetRow(textbox04, i);
                                Grid.SetColumn(textbox04, 3);
                                strPTemp = Win_ord_InOutSum_QView_Insert.V_CustomName;
                            }
                            else if (k == drc.Count)
                            {
                                p++;
                                Grid.SetRowSpan(textbox04, p);
                            }
                            else
                            {
                                if (strPTemp.Equals(Win_ord_InOutSum_QView_Insert.V_CustomName))
                                {
                                    p++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox04, p);
                                    p = 1;

                                    //거래처
                                    textbox04 = new TextBox();
                                    textbox04.Text = Win_ord_InOutSum_QView_Insert.V_CustomName;
                                    textbox04.Style = cellStyleLeft;
                                    textbox04.IsReadOnly = true;

                                    grdMerge_Month_V.Children.Add(textbox04);
                                    Grid.SetRow(textbox04, i);
                                    Grid.SetColumn(textbox04, 3);
                                    strPTemp = Win_ord_InOutSum_QView_Insert.V_CustomName;
                                }
                            }

                            //품번
                            var textbox05 = new TextBox();
                            textbox05.Text = Win_ord_InOutSum_QView_Insert.V_BuyerArticleNo;
                            textbox05.Style = cellStyleLeft;
                            textbox05.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox05);
                            Grid.SetRow(textbox05, i);
                            Grid.SetColumn(textbox05, 4);


                            //품명
                            var textbox06 = new TextBox();
                            textbox06.Text = Win_ord_InOutSum_QView_Insert.V_Article;
                            textbox06.Style = cellStyleLeft;
                            textbox06.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox06);
                            Grid.SetRow(textbox06, i);
                            Grid.SetColumn(textbox06, 5);


                            //품명코드
                            var textbox07 = new TextBox();
                            textbox07.Text = Win_ord_InOutSum_QView_Insert.V_ArticleID;
                            textbox07.Style = cellStyleCenter;
                            textbox07.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox07);
                            Grid.SetRow(textbox07, i);
                            Grid.SetColumn(textbox07, 6);

                            //건수
                            var textbox08 = new TextBox();
                            textbox08.Text = Win_ord_InOutSum_QView_Insert.V_Roll;
                            textbox08.Style = cellStyleRight;
                            textbox08.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox08);
                            Grid.SetRow(textbox08, i);
                            Grid.SetColumn(textbox08, 7);


                            //수량
                            var textbox09 = new TextBox();
                            textbox09.Text = Win_ord_InOutSum_QView_Insert.V_Qty;
                            textbox09.Style = cellStyleRight;
                            textbox09.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox09);
                            Grid.SetRow(textbox09, i);
                            Grid.SetColumn(textbox09, 8);

                            //단위
                            var textbox10 = new TextBox();
                            textbox10.Text = Win_ord_InOutSum_QView_Insert.V_UnitClssName;
                            textbox10.Style = cellStyleCenter;
                            textbox10.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox10);
                            Grid.SetRow(textbox10, i);
                            Grid.SetColumn(textbox10, 9);


                            //점유율
                            var textbox11 = new TextBox();
                            double tb11 = Win_ord_InOutSum_QView_Insert.V_CustomRate.Equals("") ? 0 : Convert.ToDouble(Win_ord_InOutSum_QView_Insert.V_CustomRate);
                            /*Convert.ToDouble(Win_ord_InOutSum_QView_Insert.V_CustomRate);*/
                            textbox11.Text = String.Format("{0:0.##;0:#;#}", tb11);
                            textbox11.Style = cellStyleRight;
                            textbox11.IsReadOnly = true;

                            grdMerge_Month_V.Children.Add(textbox11);
                            Grid.SetRow(textbox11, i);
                            Grid.SetColumn(textbox11, 10);

                        }
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

        #region 월별집계 최근 3개월 가로집계
        // 월별집계 (가로) (최근 3개월)
        private void FillGrid_Month_H()
        {
            grdMerge_Month_H.Children.Clear();
            grdMergeMonth_H.Items.Clear();

            string SearchToDate = DateTime.Now.ToString("yyyy-MM-dd").Substring(0, 10).Replace("-", "");         //오늘날짜 자동세팅
            int ChkCustomID = 0;
            if (chkCustomer.IsChecked == true) { ChkCustomID = 1; }
            else { txtCustomer.Tag = ""; txtCustomer.Text = ""; }                                                      //거래처
            int ChkArticleID = 0;
            if (chkArticle.IsChecked == true) { ChkArticleID = 1; }
            else { txtArticle.Tag = ""; txtArticle.Text = ""; }                                                       //품명
            int nGubun = 0;
            if (chkInOutGubun.IsChecked == true)
            {
                if (cboInOutGubun.SelectedValue.ToString() == "1") { nGubun = 1; }
                if (cboInOutGubun.SelectedValue.ToString() == "2") { nGubun = 2; }
            }                                                                                   //입출고구분
            int nMainItem = 0;
            if (chkMainInterestItem.IsChecked == true) { nMainItem = 1; }                       //주요관심품목
            int nCustomItem = 0;
            if (chkCustomsEnrollItem.IsChecked == true) { nCustomItem = 1; }                    //거래처등록품목
            int chkInspect = 0;
            string sInspect = string.Empty;
            if (chkInInsepectGubun.IsChecked == true)
            {
                chkInspect = 1;
                if (cboInInspectGubun.SelectedValue.ToString() == "Y") { sInspect = "Y"; }
                if (cboInInspectGubun.SelectedValue.ToString() == "N") { sInspect = "N"; }
            }                                                                                   //입고검수구분

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("EDate", SearchToDate);
                sqlParameter.Add("ChkCustomID", ChkCustomID);
                sqlParameter.Add("CustomID", txtCustomer.Tag.ToString());
                sqlParameter.Add("ChkArticleID", chkArticle.IsChecked == true ? 1 : 0);// ChkArticleID);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : string.Empty);// txtArticle.Tag.ToString());
                sqlParameter.Add("nGubun", nGubun);
                sqlParameter.Add("nMainItem", nMainItem);
                sqlParameter.Add("nCustomItem", nCustomItem);
                sqlParameter.Add("chkInspect", chkInspect);
                sqlParameter.Add("sInspect", sInspect);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text.Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_MonthSpread3", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    SpreadMonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_H.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        Style cellStyleLeft = new Style(typeof(TextBox));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
                        cellStyleLeft.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleCenter = new Style(typeof(TextBox));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
                        cellStyleCenter.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        Style cellStyleRight = new Style(typeof(TextBox));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BorderBrushProperty, System.Windows.Media.Brushes.Black));
                        cellStyleRight.Setters.Add(new Setter(TextBox.BackgroundProperty, System.Windows.Media.Brushes.White));
                        cellStyleRight.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                        cellStyleRight.Setters.Add(new Setter(TextBox.HorizontalContentAlignmentProperty, HorizontalAlignment.Right));
                        cellStyleRight.Setters.Add(new Setter(TextBox.FocusableProperty, true));

                        int k;

                        var textbox02 = new TextBox();
                        int m = 1;                      //var textbox02용
                        string strMTemp = string.Empty; //var textbox02용

                        var textbox03 = new TextBox();
                        int o = 1;                      //var textbox03용
                        string strOTemp = string.Empty; //var textbox03용



                        for (int i = 0; drc.Count > i; i++)
                        {
                            k = 2 + i;
                            var Win_ord_InOutSum_QView_Insert = new Win_ord_InOutSum_QView()
                            {
                                H_NUM = (k - 1),
                                H_cls = drc[i]["cls"].ToString(),
                                H_Gbn = drc[i]["Gbn"].ToString(),
                                H_CustomID = drc[i]["CustomID"].ToString(),
                                H_CustomName = drc[i]["CustomName"].ToString(),
                                H_BuyerArticleNo = drc[i]["BuyerArticleNo"].ToString(),
                                H_ArticleID = drc[i]["ArticleID"].ToString(),
                                H_Article = drc[i]["Article"].ToString(),
                                H_UnitClss = drc[i]["UnitClss"].ToString(),
                                H_UnitClssName = drc[i]["UnitClssName"].ToString(),
                                H_UnitPrice = drc[i]["UnitPrice"].ToString(),
                                H_PriceClss = drc[i]["PriceClss"].ToString(),
                                H_PriceClssName = drc[i]["PriceClssName"].ToString(),
                                H_YYYYMM1 = drc[i]["YYYYMM1"].ToString(),
                                H_YYYYMM2 = drc[i]["YYYYMM2"].ToString(),
                                H_YYYYMM3 = drc[i]["YYYYMM3"].ToString(),
                                H_YYYYMM4 = drc[i]["YYYYMM4"].ToString(),
                                H_YYYYMM5 = drc[i]["YYYYMM5"].ToString(),
                                H_YYYYMM6 = drc[i]["YYYYMM6"].ToString(),
                                H_YYYYMM7 = drc[i]["YYYYMM7"].ToString(),
                                H_YYYYMM8 = drc[i]["YYYYMM8"].ToString(),
                                H_YYYYMM9 = drc[i]["YYYYMM9"].ToString(),

                                H_YYYYMM10 = drc[i]["YYYYMM10"].ToString(),
                                H_roll10 = drc[i]["roll10"].ToString(),
                                H_Qty10 = drc[i]["Qty10"].ToString().Split('.')[0].Trim(),
                                H_Amount10 = drc[i]["Amount10"].ToString(),
                                H_VatAmount10 = drc[i]["VatAmount10"].ToString(),

                                H_YYYYMM11 = drc[i]["YYYYMM11"].ToString(),
                                H_roll11 = drc[i]["roll11"].ToString(),
                                H_Qty11 = drc[i]["Qty11"].ToString().Split('.')[0].Trim(),
                                H_Amount11 = drc[i]["Amount11"].ToString(),
                                H_VatAmount11 = drc[i]["VatAmount11"].ToString(),

                                H_YYYYMM12 = drc[i]["YYYYMM12"].ToString(),
                                H_roll12 = drc[i]["roll12"].ToString(),
                                H_Qty12 = drc[i]["Qty12"].ToString().Split('.')[0].Trim(),
                                H_Amount12 = drc[i]["Amount12"].ToString(),
                                H_VatAmount12 = drc[i]["VatAmount12"].ToString(),

                                H_roll13 = drc[i]["roll13"].ToString(),
                                H_Qty13 = drc[i]["Qty13"].ToString().Split('.')[0].Trim(),
                                H_Amount13 = drc[i]["Amount13"].ToString(),
                                H_VatAmount13 = drc[i]["VatAmount13"].ToString(),

                                H_RN = drc[i]["RN"].ToString(),
                                H_CustomRate = lib.returnNumStringTwo(drc[i]["CustomRate"].ToString()),
                                H_CustomAmount = drc[i]["CustomAmount"].ToString(),
                                H_AllTotalAmount = drc[i]["AllTotalAmount"].ToString(),
                            };

                            Win_ord_InOutSum_QView_Insert.H_UnitPrice = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_UnitPrice);

                            Win_ord_InOutSum_QView_Insert.H_roll10 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_roll10);
                            Win_ord_InOutSum_QView_Insert.H_Qty10 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Qty10);
                            Win_ord_InOutSum_QView_Insert.H_Amount10 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Amount10);
                            Win_ord_InOutSum_QView_Insert.H_VatAmount10 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_VatAmount10);
                            Win_ord_InOutSum_QView_Insert.H_roll11 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_roll11);
                            Win_ord_InOutSum_QView_Insert.H_Qty11 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Qty11);
                            Win_ord_InOutSum_QView_Insert.H_Amount11 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Amount11);
                            Win_ord_InOutSum_QView_Insert.H_VatAmount11 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_VatAmount11);

                            Win_ord_InOutSum_QView_Insert.H_roll12 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_roll12);
                            Win_ord_InOutSum_QView_Insert.H_Qty12 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Qty12);
                            Win_ord_InOutSum_QView_Insert.H_Amount12 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Amount12);
                            Win_ord_InOutSum_QView_Insert.H_VatAmount12 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_VatAmount12);
                            Win_ord_InOutSum_QView_Insert.H_roll13 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_roll13);
                            Win_ord_InOutSum_QView_Insert.H_Qty13 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Qty13);
                            Win_ord_InOutSum_QView_Insert.H_Amount13 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_Amount13);
                            Win_ord_InOutSum_QView_Insert.H_VatAmount13 = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_VatAmount13);

                            Win_ord_InOutSum_QView_Insert.H_RN = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_RN);
                            Win_ord_InOutSum_QView_Insert.H_CustomRate = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_CustomRate);
                            Win_ord_InOutSum_QView_Insert.H_CustomAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_CustomAmount);
                            Win_ord_InOutSum_QView_Insert.H_AllTotalAmount = lib.returnNumString(Win_ord_InOutSum_QView_Insert.H_AllTotalAmount);

                            if (Win_ord_InOutSum_QView_Insert.H_Gbn == "1")
                            {
                                if (Win_ord_InOutSum_QView_Insert.H_cls == "0")
                                {
                                    Win_ord_InOutSum_QView_Insert.H_Gbn = "입고총계";
                                }
                                else
                                {
                                    Win_ord_InOutSum_QView_Insert.H_Gbn = "입고";
                                }
                            }
                            else if (Win_ord_InOutSum_QView_Insert.H_Gbn == "2")
                            {
                                if (Win_ord_InOutSum_QView_Insert.H_cls == "0")
                                {
                                    Win_ord_InOutSum_QView_Insert.H_Gbn = "출고총계";
                                }
                                else
                                {
                                    Win_ord_InOutSum_QView_Insert.H_Gbn = "출고";
                                }
                            }

                            grdMergeMonth_H.Items.Add(Win_ord_InOutSum_QView_Insert);

                            RowDefinition row = new RowDefinition();
                            row.Height = GridLength.Auto;
                            grdMerge_Month_H.RowDefinitions.Add(row);

                            //순번
                            var textbox01 = new TextBox();
                            textbox01.Text = (k - 1).ToString();
                            textbox01.Style = cellStyleCenter;
                            textbox01.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox01);
                            Grid.SetRow(textbox01, i);
                            Grid.SetColumn(textbox01, 0);


                            // textbox 02 시작지점.
                            if (i == 0)
                            {
                                //구분
                                textbox02.Text = Win_ord_InOutSum_QView_Insert.H_Gbn;
                                textbox02.Style = cellStyleCenter;
                                textbox02.IsReadOnly = true;

                                grdMerge_Month_H.Children.Add(textbox02);
                                Grid.SetRow(textbox02, i);
                                Grid.SetColumn(textbox02, 1);
                                strMTemp = Win_ord_InOutSum_QView_Insert.H_Gbn;
                            }
                            else if ((k - 1) == drc.Count)
                            {
                                m++;
                                Grid.SetRowSpan(textbox02, m);
                            }
                            else
                            {
                                if (strMTemp.Equals(Win_ord_InOutSum_QView_Insert.H_Gbn))
                                {
                                    m++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox02, m);
                                    m = 1;

                                    //구분
                                    textbox02 = new TextBox();
                                    textbox02.Text = Win_ord_InOutSum_QView_Insert.H_Gbn;
                                    textbox02.Style = cellStyleCenter;
                                    textbox02.IsReadOnly = true;

                                    grdMerge_Month_H.Children.Add(textbox02);
                                    Grid.SetRow(textbox02, i);
                                    Grid.SetColumn(textbox02, 1);
                                    strMTemp = Win_ord_InOutSum_QView_Insert.H_Gbn;
                                }
                            }


                            // textbox 03 시작지점.
                            if (i == 0)
                            {
                                //거래처
                                textbox03.Text = Win_ord_InOutSum_QView_Insert.H_CustomName;
                                textbox03.Style = cellStyleLeft;
                                textbox03.IsReadOnly = true;

                                grdMerge_Month_H.Children.Add(textbox03);
                                Grid.SetRow(textbox03, i);
                                Grid.SetColumn(textbox03, 2);
                                strOTemp = Win_ord_InOutSum_QView_Insert.H_CustomName;
                            }
                            else if ((k - 1) == drc.Count)
                            {
                                o++;
                                Grid.SetRowSpan(textbox03, o);
                            }
                            else
                            {
                                if (strOTemp.Equals(Win_ord_InOutSum_QView_Insert.H_CustomName))
                                {
                                    o++;
                                }
                                else
                                {
                                    Grid.SetRowSpan(textbox03, o);
                                    o = 1;

                                    //거래처
                                    textbox03 = new TextBox();
                                    textbox03.Text = Win_ord_InOutSum_QView_Insert.H_CustomName;
                                    textbox03.Style = cellStyleLeft;
                                    textbox03.IsReadOnly = true;

                                    grdMerge_Month_H.Children.Add(textbox03);
                                    Grid.SetRow(textbox03, i);
                                    Grid.SetColumn(textbox03, 2);
                                    strOTemp = Win_ord_InOutSum_QView_Insert.H_CustomName;
                                }
                            }


                            //품번
                            var textbox04 = new TextBox();
                            textbox04.Text = Win_ord_InOutSum_QView_Insert.H_BuyerArticleNo;
                            textbox04.Style = cellStyleLeft;
                            textbox04.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox04);
                            Grid.SetRow(textbox04, i);
                            Grid.SetColumn(textbox04, 3);


                            //품명
                            var textbox05 = new TextBox();
                            textbox05.Text = Win_ord_InOutSum_QView_Insert.H_Article;
                            textbox05.Style = cellStyleLeft;
                            textbox05.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox05);
                            Grid.SetRow(textbox05, i);
                            Grid.SetColumn(textbox05, 4);


                            //단위
                            var textbox06 = new TextBox();
                            textbox06.Text = Win_ord_InOutSum_QView_Insert.H_UnitClssName;
                            textbox06.Style = cellStyleCenter;
                            textbox06.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox06);
                            Grid.SetRow(textbox06, i);
                            Grid.SetColumn(textbox06, 5);

                            //var textbox06 = new TextBox();
                            //textbox06.Text = Win_ord_InOutSum_QView_Insert.H_PriceClssName;
                            //textbox06.Style = cellStyle;

                            //grdMerge_Month_H.Children.Add(textbox06);
                            //Grid.SetRow(textbox06, i);
                            //Grid.SetColumn(textbox06, 5);

                            //점유율
                            var textbox07 = new TextBox();
                            double tb07 = Win_ord_InOutSum_QView_Insert.H_CustomRate.Equals("") ? 0 : Convert.ToDouble(Win_ord_InOutSum_QView_Insert.H_CustomRate);
                            /* Convert.ToDouble(Win_ord_InOutSum_QView_Insert.H_CustomRate);*/
                            textbox07.Text = String.Format("{0:0.##;0:#;#}", tb07);
                            textbox07.Style = cellStyleRight;
                            textbox07.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox07);
                            Grid.SetRow(textbox07, i);
                            Grid.SetColumn(textbox07, 6);

                            //건수
                            var textbox08 = new TextBox();
                            textbox08.Text = Win_ord_InOutSum_QView_Insert.H_roll13;
                            textbox08.Style = cellStyleRight;
                            textbox08.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox08);
                            Grid.SetRow(textbox08, i);
                            Grid.SetColumn(textbox08, 7);

                            //수량
                            var textbox09 = new TextBox();
                            textbox09.Text = Win_ord_InOutSum_QView_Insert.H_Qty13;
                            textbox09.Style = cellStyleRight;
                            textbox09.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox09);
                            Grid.SetRow(textbox09, i);
                            Grid.SetColumn(textbox09, 8);

                            //건수
                            var textbox10 = new TextBox();
                            textbox10.Text = Win_ord_InOutSum_QView_Insert.H_roll10;
                            textbox10.Style = cellStyleRight;
                            textbox10.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox10);
                            Grid.SetRow(textbox10, i);
                            Grid.SetColumn(textbox10, 9);

                            //수량
                            var textbox11 = new TextBox();
                            textbox11.Text = Win_ord_InOutSum_QView_Insert.H_Qty10;
                            textbox11.Style = cellStyleRight;
                            textbox11.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox11);
                            Grid.SetRow(textbox11, i);
                            Grid.SetColumn(textbox11, 10);

                            //건수
                            var textbox12 = new TextBox();
                            textbox12.Text = Win_ord_InOutSum_QView_Insert.H_roll11;
                            textbox12.Style = cellStyleRight;
                            textbox12.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox12);
                            Grid.SetRow(textbox12, i);
                            Grid.SetColumn(textbox12, 11);

                            //수량
                            var textbox13 = new TextBox();
                            textbox13.Text = Win_ord_InOutSum_QView_Insert.H_Qty11;
                            textbox13.Style = cellStyleRight;
                            textbox13.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox13);
                            Grid.SetRow(textbox13, i);
                            Grid.SetColumn(textbox13, 12);

                            //건수
                            var textbox14 = new TextBox();
                            textbox14.Text = Win_ord_InOutSum_QView_Insert.H_roll12;
                            textbox14.Style = cellStyleRight;
                            textbox14.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox14);
                            Grid.SetRow(textbox14, i);
                            Grid.SetColumn(textbox14, 13);

                            //수량
                            var textbox15 = new TextBox();
                            textbox15.Text = Win_ord_InOutSum_QView_Insert.H_Qty12;
                            textbox15.Style = cellStyleRight;
                            textbox15.IsReadOnly = true;

                            grdMerge_Month_H.Children.Add(textbox15);
                            Grid.SetRow(textbox15, i);
                            Grid.SetColumn(textbox15, 14);

                        }

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


        #region 월별 가로집계 셀렉션 체인지 이벤트
        // 탬 컨트롤 셀렉션 체인지 이벤트.
        private void tabconGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string sNowTI = ((sender as TabControl).SelectedItem as TabItem).Header as string;

            switch (sNowTI)
            {
                case "기간집계":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "일일집계":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "월별집계(세로)":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "월별집계(가로)":
                    txtblMessage.Visibility = Visibility.Visible;
                    dtpFromDate.IsEnabled = false;
                    dtpToDate.IsEnabled = false;
                    break;
                default: return;
            }
        }


        #endregion


        //닫기 버튼 클릭.
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


        #region 엑셀

        // 엑셀 버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            string sNowTI = (tabconGrid.SelectedItem as TabItem).Header as string;
            string Listname1 = string.Empty;
            string Listname2 = string.Empty;
            DataTable choicedt = null;
            Lib lib2 = new Lib();

            if (PeriodDataTable != null)
            {
                switch (sNowTI)
                {
                    case "기간집계":
                        if (PeriodDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 기간집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "기간집계";
                        Listname2 = "PeriodData";
                        choicedt = PeriodDataTable;
                        break;
                    case "일일집계":
                        if (DaysDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 일일집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "일일집계";
                        Listname2 = "DayData";
                        choicedt = DaysDataTable;
                        break;
                    case "월별집계(세로)":
                        if (MonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(세로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(세로)집계";
                        Listname2 = "MonthData";
                        choicedt = MonthDataTable;
                        break;
                    case "월별집계(가로)":
                        if (SpreadMonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(가로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(가로)집계";
                        Listname2 = "SpreadMonthData";
                        choicedt = SpreadMonthDataTable;
                        break;
                    default: return;
                }

                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = Listname1;
                lst[2] = Listname2;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                // 어쨋든 머든 여기서 dt로 만들어서 주면 된다는 거네.
                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(Listname2))
                    {
                        Name = Listname2;
                        if (lib2.GenerateExcel(choicedt, Name))
                        {
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else
                    {
                        if (choicedt != null)
                        {
                            choicedt.Clear();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("엑설로 변환할 자료가 없습니다.");
            }

            lib2 = null;
        }





        #endregion


        #region 그리드 SELECT 시 BACKGROUND 그리기.
        // 일반 그리드는 select의 개념이 없습니다.
        // 그래서 클릭했을때, DataGrid처럼 파란 background를 그려주지 않습니다.
        // 그렇기에, 일반 그리드는 일일히 그려주어야 합니다.
        /// <summary>
        /// //////////////////////////////////////////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Textbox_Period_previewmousedown(object sender, MouseButtonEventArgs e)
        {
            if (PreRect.Count != 0)
            {
                for (int k = 0; k < grdMerge_Period.ColumnDefinitions.Count; k++)
                {
                    grdMerge_Period.Children.Remove(PreRect[k]);
                }
            }

            var point = Mouse.GetPosition(grdMerge_Period);

            Clicked_row = 0;
            Clicked_col = 0;
            double accumulatedHeight = 0.0;
            double accumulatedWidth = 0.0;

            foreach (var rowDefinition in grdMerge_Period.RowDefinitions)
            {
                accumulatedHeight += rowDefinition.ActualHeight;
                if (accumulatedHeight >= point.Y)
                    break;
                Clicked_row++;
            }
            foreach (var columnDefinition in grdMerge_Period.ColumnDefinitions)
            {
                accumulatedWidth += columnDefinition.ActualWidth;
                if (accumulatedWidth >= point.X)
                    break;
                Clicked_col++;
            }
        }


        private void Textbox_Period_previewmouseup(object sender, MouseButtonEventArgs e)
        {
            // selected row
            for (int k = 0; k < grdMerge_Period.ColumnDefinitions.Count; k++)
            {
                Rectangle rect = new Rectangle
                {
                    Height = grdMerge_Period.RowDefinitions[Clicked_row].ActualHeight,
                    Width = grdMerge_Period.ColumnDefinitions[k].ActualWidth,
                    Fill = new SolidColorBrush(Colors.Blue),
                    Opacity = 0.4
                };

                PreRect.Insert(k, rect);

                Grid.SetColumn(rect, k);
                Grid.SetRow(rect, Clicked_row);
                grdMerge_Period.Children.Add(rect);
            }

            // cell 클릭 용
            //Grid.SetColumn(rect, Clicked_col);
            //Grid.SetRow(rect, Clicked_row);
            //grdMerge_Period.Children.Add(rect);           
        }

        private void Textbox_Days_previewmousedown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (PreRect.Count != 0)
                {
                    for (int k = 0; k < grdMerge_Days.ColumnDefinitions.Count; k++)
                    {
                        grdMerge_Days.Children.Remove(PreRect[k]);
                    }
                }

                var point = Mouse.GetPosition(grdMerge_Days);

                Clicked_row = 0;
                Clicked_col = 0;
                double accumulatedHeight = 0.0;
                double accumulatedWidth = 0.0;

                foreach (var rowDefinition in grdMerge_Days.RowDefinitions)
                {
                    accumulatedHeight += rowDefinition.ActualHeight;
                    if (accumulatedHeight >= point.Y)
                        break;
                    Clicked_row++;
                }
                foreach (var columnDefinition in grdMerge_Period.ColumnDefinitions)
                {
                    accumulatedWidth += columnDefinition.ActualWidth;
                    if (accumulatedWidth >= point.X)
                        break;
                    Clicked_col++;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        private void Textbox_Days_previewmouseup(object sender, MouseButtonEventArgs e)
        {
            // selected row
            for (int k = 0; k < grdMerge_Days.ColumnDefinitions.Count; k++)
            {
                Rectangle rect = new Rectangle
                {
                    Height = grdMerge_Days.RowDefinitions[Clicked_row].ActualHeight,
                    Width = grdMerge_Days.ColumnDefinitions[k].ActualWidth,
                    Fill = new SolidColorBrush(Colors.Blue),
                    Opacity = 0.4
                };

                PreRect.Insert(k, rect);

                Grid.SetColumn(rect, k);
                Grid.SetRow(rect, Clicked_row);
                grdMerge_Days.Children.Add(rect);
            }

            // cell 클릭 용
            //Grid.SetColumn(rect, Clicked_col);
            //Grid.SetRow(rect, Clicked_row);
            //grdMerge_Period.Children.Add(rect);           
        }



        private void Textbox_Month_V_previewmousedown(object sender, MouseButtonEventArgs e)
        {
            if (PreRect.Count != 0)
            {
                for (int k = 0; k < grdMerge_Month_V.ColumnDefinitions.Count; k++)
                {
                    grdMerge_Month_V.Children.Remove(PreRect[k]);
                }
            }

            var point = Mouse.GetPosition(grdMerge_Month_V);

            Clicked_row = 0;
            Clicked_col = 0;
            double accumulatedHeight = 0.0;
            double accumulatedWidth = 0.0;

            foreach (var rowDefinition in grdMerge_Month_V.RowDefinitions)
            {
                accumulatedHeight += rowDefinition.ActualHeight;
                if (accumulatedHeight >= point.Y)
                    break;
                Clicked_row++;
            }
            foreach (var columnDefinition in grdMerge_Month_V.ColumnDefinitions)
            {
                accumulatedWidth += columnDefinition.ActualWidth;
                if (accumulatedWidth >= point.X)
                    break;
                Clicked_col++;
            }
        }


        private void Textbox_Month_V_previewmouseup(object sender, MouseButtonEventArgs e)
        {
            // selected row
            for (int k = 0; k < grdMerge_Month_V.ColumnDefinitions.Count; k++)
            {
                Rectangle rect = new Rectangle
                {
                    Height = grdMerge_Month_V.RowDefinitions[Clicked_row].ActualHeight,
                    Width = grdMerge_Month_V.ColumnDefinitions[k].ActualWidth,
                    Fill = new SolidColorBrush(Colors.Blue),
                    Opacity = 0.4
                };

                PreRect.Insert(k, rect);

                Grid.SetColumn(rect, k);
                Grid.SetRow(rect, Clicked_row);
                grdMerge_Month_V.Children.Add(rect);
            }
        }


        private void Textbox_Month_H_previewmousedown(object sender, MouseButtonEventArgs e)
        {
            if (PreRect.Count != 0)
            {
                for (int k = 0; k < grdMerge_Month_H.ColumnDefinitions.Count; k++)
                {
                    grdMerge_Month_H.Children.Remove(PreRect[k]);
                }
            }

            var point = Mouse.GetPosition(grdMerge_Month_H);

            Clicked_row = 0;
            Clicked_col = 0;
            double accumulatedHeight = 0.0;
            double accumulatedWidth = 0.0;

            foreach (var rowDefinition in grdMerge_Month_H.RowDefinitions)
            {
                accumulatedHeight += rowDefinition.ActualHeight;
                if (accumulatedHeight >= point.Y)
                    break;
                Clicked_row++;
            }
            foreach (var columnDefinition in grdMerge_Month_H.ColumnDefinitions)
            {
                accumulatedWidth += columnDefinition.ActualWidth;
                if (accumulatedWidth >= point.X)
                    break;
                Clicked_col++;
            }
        }


        private void Textbox_Month_H_previewmouseup(object sender, MouseButtonEventArgs e)
        {
            // selected row
            for (int k = 0; k < grdMerge_Month_H.ColumnDefinitions.Count; k++)
            {
                Rectangle rect = new Rectangle
                {
                    Height = grdMerge_Month_H.RowDefinitions[Clicked_row].ActualHeight,
                    Width = grdMerge_Month_H.ColumnDefinitions[k].ActualWidth,
                    Fill = new SolidColorBrush(Colors.Blue),
                    Opacity = 0.4
                };

                PreRect.Insert(k, rect);

                Grid.SetColumn(rect, k);
                Grid.SetRow(rect, Clicked_row);
                grdMerge_Month_H.Children.Add(rect);
            }
        }


        #endregion


        private void Header_Content_Width_Adjust(string nowti)
        {
            if (nowti == "기간집계")
            {
                for (int i = 0; i < grdMerge_Period.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Period.ColumnDefinitions[i].ActualWidth;
                    grdHeader_Period.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
                }
            }
            else if (nowti == "일일집계")
            {
                for (int i = 0; i < grdMerge_Days.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Days.ColumnDefinitions[i].ActualWidth;
                    grdHeader_Days.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
                }
            }
            else if (nowti == "월별집계(세로)")
            {
                for (int i = 0; i < grdMerge_Month_V.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Month_V.ColumnDefinitions[i].ActualWidth;
                    grdHeader_Month_V.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
                }
            }
            else if (nowti == "월별집계(가로)")
            {
                for (int i = 0; i < grdMerge_Month_H.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Month_H.ColumnDefinitions[i].ActualWidth;
                    grdHeader_Month_H.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
                }
            }
        }


        private void Re_Header_Content_Width_Adjust(string nowti)
        {
            //if (nowti == "기간집계")
            //{
            //    for (int i = 0; i < grdMerge_Period.ColumnDefinitions.Count; i++)
            //    {
            //        double ContentWidth = grdMerge_Period.ColumnDefinitions[i].ActualWidth;
            //        grdHeader_Period.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
            //    }
            //}
            if (nowti == "일일집계")
            {
                for (int i = 0; i < grdMerge_Days.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Days.ColumnDefinitions[i].ActualWidth;
                    double HeaderWidth = grdHeader_Days.ColumnDefinitions[i].ActualWidth;
                    if ((ContentWidth + 1) < HeaderWidth)
                    {
                        grdMerge_Days.ColumnDefinitions[i].Width = new GridLength(HeaderWidth, GridUnitType.Pixel);
                    }
                }
            }
            //else if (nowti == "월별집계(세로)")
            //{
            //    for (int i = 0; i < grdMerge_Month_V.ColumnDefinitions.Count; i++)
            //    {
            //        double ContentWidth = grdMerge_Month_V.ColumnDefinitions[i].ActualWidth;
            //        grdHeader_Month_V.ColumnDefinitions[i].Width = new GridLength(ContentWidth, GridUnitType.Star);
            //    }
            //}
            else if (nowti == "월별집계(가로)")
            {
                for (int i = 0; i < grdMerge_Month_H.ColumnDefinitions.Count; i++)
                {
                    double ContentWidth = grdMerge_Month_H.ColumnDefinitions[i].ActualWidth;
                    double HeaderWidth = grdHeader_Month_H.ColumnDefinitions[i].ActualWidth;
                    if ((ContentWidth + 1) < HeaderWidth)
                    {
                        grdMerge_Month_H.ColumnDefinitions[i].Width = new GridLength(HeaderWidth, GridUnitType.Pixel);
                    }
                }
            }
        }



        private void txtCustomer_KeyDown(object sender, KeyEventArgs e)
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

        private void grdPeriod_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void grdMergeDays_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void grdMergeMonth_V_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void grdPeriod_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {

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





    /// <summary>
    /// /////////////////////////////////////////////////////////////////////
    /// </summary>


    public class MonthChange
    {
        //SpreadMonth 월 기간 확인용
        public string H_MON1 { get; set; }
        public string H_MON2 { get; set; }
        public string H_MON3 { get; set; }

    }



    class Win_ord_InOutSum_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 - 기간집계용 ( P_ (Period))
        public int P_NUM { get; set; }
        public string P_cls { get; set; }
        public string P_Gbn { get; set; }
        public string P_IODate { get; set; }
        public string P_CustomID { get; set; }
        public string P_CustomName { get; set; }

        public string P_Sabun { get; set; }

        public string P_BuyerArticleNo { get; set; }
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public string P_Roll { get; set; }
        public string P_Qty { get; set; }
        public string P_UnitClss { get; set; }

        public string P_UnitClssName { get; set; }
        public string P_UnitPrice { get; set; }
        public string P_PriceClss { get; set; }
        public string P_PriceClssName { get; set; }
        public string P_Amount { get; set; }

        public string P_VatAmount { get; set; }
        public string P_TotAmount { get; set; }
        public string P_CustomRate { get; set; }
        public string P_CustomRateOrder { get; set; }



        // 조회 - 일별집계용 ( D_ (Day))
        public int D_NUM { get; set; }
        public string D_cls { get; set; }
        public string D_Gbn { get; set; }
        public string D_IODate { get; set; }
        public string D_CustomID { get; set; }
        public string D_CustomName { get; set; }

        public string D_BuyerArticleNo { get; set; }
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Roll { get; set; }
        public string D_Qty { get; set; }
        public string D_UnitClss { get; set; }

        public string D_Sabun { get; set; }

        public string D_UnitClssName { get; set; }
        public string D_UnitPrice { get; set; }
        public string D_PriceClss { get; set; }
        public string D_PriceClssName { get; set; }
        public string D_Amount { get; set; }

        public string D_VatAmount { get; set; }
        public string D_TotAmount { get; set; }
        public string D_CustomRate { get; set; }
        public string D_CustomRateOrder { get; set; }


        // 조회 - 월별집계용 _V ( V_ (V_Month))
        public int V_NUM { get; set; }
        public string V_cls { get; set; }
        public string V_Gbn { get; set; }
        public string V_IODate { get; set; }
        public string V_CustomID { get; set; }
        public string V_CustomName { get; set; }

        public string V_BuyerArticleNo { get; set; }
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public string V_Roll { get; set; }
        public string V_Qty { get; set; }
        public string V_UnitClss { get; set; }

        public string V_Sabun { get; set; }

        public string V_UnitClssName { get; set; }
        public string V_UnitPrice { get; set; }
        public string V_PriceClss { get; set; }
        public string V_PriceClssName { get; set; }
        public string V_Amount { get; set; }

        public string V_VatAmount { get; set; }
        public string V_TotAmount { get; set; }
        public string V_CustomRate { get; set; }
        public string V_CustomRateOrder { get; set; }
        public string V_RN { get; set; }



        // 조회 - 월별집계용 _H ( H_ (H_Month))
        public int H_NUM { get; set; }
        public string H_cls { get; set; }
        public string H_Gbn { get; set; }
        public string H_CustomID { get; set; }
        public string H_CustomName { get; set; }

        public string H_BuyerArticleNo { get; set; }
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public string H_UnitClss { get; set; }
        public string H_UnitClssName { get; set; }
        public string H_UnitPrice { get; set; }

        public string H_Sabun { get; set; }

        public string H_PriceClss { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_YYYYMM1 { get; set; }
        public string H_YYYYMM2 { get; set; }
        public string H_YYYYMM3 { get; set; }

        public string H_YYYYMM4 { get; set; }
        public string H_YYYYMM5 { get; set; }
        public string H_YYYYMM6 { get; set; }
        public string H_YYYYMM7 { get; set; }
        public string H_YYYYMM8 { get; set; }

        public string H_YYYYMM9 { get; set; }
        public string H_YYYYMM10 { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_Amount10 { get; set; }

        public string H_VatAmount10 { get; set; }
        public string H_YYYYMM11 { get; set; }
        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_Amount11 { get; set; }

        public string H_VatAmount11 { get; set; }
        public string H_YYYYMM12 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_Amount12 { get; set; }

        public string H_VatAmount12 { get; set; }
        public string H_YYYYMM13 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }
        public string H_Amount13 { get; set; }

        public string H_VatAmount13 { get; set; }
        public string H_RN { get; set; }
        public string H_CustomRate { get; set; }
        public string H_CustomAmount { get; set; }
        public string H_AllTotalAmount { get; set; }


        public List<P_listmodel> P_listmodel { get; set; }
        public List<D_gbnmodel> D_gbnmodel { get; set; }
        public List<V_gbnmodel> V_gbnmodel { get; set; }
        public List<H_custommodel> H_custommodel { get; set; }


    }

    public class D_gbnmodel
    {
        public string D_Gbn { get; set; }
        public string D_YesColor { get; set; }

        public List<D_custommodel> D_custommodel { get; set; }
    }

    public class V_gbnmodel
    {
        public string V_Gbn { get; set; }
        public List<V_custommodel> V_custommodel { get; set; }
    }



    public class D_custommodel
    {
        public string D_CustomName { get; set; }
        public List<D_listmodel> D_listmodel { get; set; }
    }

    public class V_custommodel
    {
        public string V_CustomName { get; set; }
        public string V_YesColor { get; set; }
        public List<V_listmodel> V_listmodel { get; set; }
    }

    public class H_custommodel
    {
        public string H_CustomName { get; set; }
        public List<H_listmodel> H_listmodel { get; set; }
    }



    public class D_listmodel
    {
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Roll { get; set; }
        public string D_Qty { get; set; }
        public string D_UnitClssName { get; set; }
        public string D_PriceClssName { get; set; }

        public string D_VatAmount { get; set; }
        public string D_TotAmount { get; set; }
        public string D_CustomRate { get; set; }

    }

    public class P_listmodel
    {
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public string P_Roll { get; set; }
        public string P_Qty { get; set; }
        public string P_UnitClssName { get; set; }
        public string P_CustomRate { get; set; }

        public string P_YesColor { get; set; }

    }

    public class V_listmodel
    {
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public string V_Roll { get; set; }
        public string V_Qty { get; set; }
        public string V_UnitClssName { get; set; }
        public string V_CustomRate { get; set; }
    }



    public class H_listmodel
    {
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public string H_UnitClssName { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_CustomRate { get; set; }

        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }

    }


}

