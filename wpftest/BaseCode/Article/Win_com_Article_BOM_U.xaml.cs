using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WizMes_WooJung.PopUp;
using WizMes_WooJung.PopUP;
//using System.Threading.Tasks;
using System.Threading;

namespace WizMes_WooJung
{
    /// <summary>
    /// Win_com_ArticleBOM_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Article_BOM_U : UserControl
    {
        string strFlag = string.Empty;
        string strDirection = string.Empty;
        int rowNum = 0;
        int ArticleBomCnt = 0;
        Lib lib = new Lib();

        Win_com_ArticleBOM_ItemList WinArticleBomList = new Win_com_ArticleBOM_ItemList();
        Win_com_ArticleBOM_Code_U WinArticleBomCode = new Win_com_ArticleBOM_Code_U();
        DataTable dataTableArticle = null;
        DataTable dataTableBOM = null;

        public Win_com_Article_BOM_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            SetComboBox();
            strDirection = "B";
            dataTableArticle = Procedure.Instance.GetArticle("", 1, "", "");
        }

        private void SetComboBox()
        {
            //품명그룹(조회, 입력)
            ObservableCollection<CodeView> ovcArticleGrp = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            cboArticleGrpSrh.ItemsSource = ovcArticleGrp;
            cboArticleGrpSrh.DisplayMemberPath = "code_name";
            cboArticleGrpSrh.SelectedValuePath = "code_id";

            cbosArticleGrpP.ItemsSource = ovcArticleGrp;
            cbosArticleGrpP.DisplayMemberPath = "code_name";
            cbosArticleGrpP.SelectedValuePath = "code_id";

            cbosArticleGrpS.ItemsSource = ovcArticleGrp;
            cbosArticleGrpS.DisplayMemberPath = "code_name";
            cbosArticleGrpS.SelectedValuePath = "code_id";

            //단위
            ObservableCollection<CodeView> ovcUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "MTRUNIT", "Y", "");
            this.cboUnitClss.ItemsSource = ovcUnitClss;
            this.cboUnitClss.DisplayMemberPath = "code_name";
            this.cboUnitClss.SelectedValuePath = "code_id";
        }

        //품명그룹 검색 조건 사용체크
        private void lblArticleGrpSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleGrpSrh.IsChecked == true) { chkArticleGrpSrh.IsChecked = false; }
            else { chkArticleGrpSrh.IsChecked = true; }
        }

        //품명그룹 검색 조건 사용체크
        private void chkArticleGrpSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsEnabled = true;
            cboArticleGrpSrh.IsEnabled = true;
        }

        //품명그룹 검색 조건 사용체크
        private void chkArticleGrpSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleGrpSrh.IsEnabled = true;
            cboArticleGrpSrh.IsEnabled = false;
        }

        //품명 검색 조건 사용체크
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명 검색 조건 사용체크
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        //품명 검색 조건 사용체크
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
        }

        //
        private void TxtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
            }
        }

        //
        private void BtnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");
        }

        //검색 조건 라디오 버튼 체크(상향식)
        private void rbnTwoSrh_Click(object sender, RoutedEventArgs e)
        {
            rbnTwoSrh.IsChecked = true;        
            strDirection = "B";
            tblArticleSrh.Text = "상위품";

        }
        private void tblRbnTwoSrh(object sender, MouseButtonEventArgs e)
        {
            rbnTwoSrh.IsChecked = true;
            strDirection = "B";
            tblArticleSrh.Text = "상위품";

        }

        //검색 조건 라디오 버튼 체크(하향식)
        private void rbnThreeSrh_Click(object sender, RoutedEventArgs e)
        {
            rbnThreeSrh.IsChecked = true;
            strDirection = "U";
            tblArticleSrh.Text = "하위품";
            
        }
        private void tblRbnThreeSrh(object sender, MouseButtonEventArgs e)
        {
            rbnThreeSrh.IsChecked = true;
            strDirection = "U";
            tblArticleSrh.Text = "하위품";
        }

        // 사용안함 포함 체크박스 이벤트
        private void lblNotUseSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNoUse.IsChecked == true)
            {
                chkNoUse.IsChecked = false;
            }
            else
            {
                chkNoUse.IsChecked = true;
            }
        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            // 수정 후, 수정 취소시 상위품목, 품명은 입력 가능하도록 설정
            txtParentArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtParentArticle.IsReadOnly = false;
            btnPfParentArticle.IsEnabled = true;

            txtArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");
            txtArticle.IsReadOnly = false;
            btnPfArticle.IsEnabled = true;

            Lib.Instance.UiButtonEnableChange_IUControl(this);
            gbxInput.IsHitTestVisible = false;
            gbxDataGrid.IsEnabled = false; // 상위품목, 하위품목 리스트는 추가 상태일때만 활성화
            tlvItemList.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            gbxInput.IsHitTestVisible = true;
            gbxDataGrid.IsEnabled = true;
            tlvItemList.IsHitTestVisible = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;
            cboUnitClss.SelectedIndex = 1;
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today.AddYears(500);

            txtParentArticle.Tag = null;
            txtArticle.Tag = null;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList != null)
                {
                    tbkMsg.Text = "자료 수정 중";
                    CantBtnControl();
                    strFlag = "U";

                    // 상위품목, 품명은 수정 불가능 설정
                    txtParentArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtParentArticle.IsReadOnly = true;
                    btnPfParentArticle.IsEnabled = false;

                    txtArticle.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");
                    txtArticle.IsReadOnly = true;
                    btnPfArticle.IsEnabled = false;

                    // 상위품목, 하위품목 리스트는 수정일때는 비활성화
                    gbxDataGrid.IsEnabled = false;

                    this.DataContext = WinArticleBomList;
                }
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                }
                else if (WinBomList.Items.Count == 1)
                {
                    MessageBox.Show("해당 품목의 하위품 먼저 삭제해주세요.");
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        //삭제 전 체크
                        if (!DeleteDataCheck(WinArticleBomList.ArticleID, WinArticleBomList.PARENTArticleID))
                            return;

                        if (tlvItemList.Items.Count > 0 && tlvItemList.SelectedItem != null)
                        {
                            rowNum = 0;
                        }

                        if (Procedure.Instance.DeleteData(WinArticleBomList.ArticleID, WinArticleBomList.PARENTArticleID,
                            "sArticleID", "sParentArticleID", "xp_Article_dArticleBOM"))
                        {
                            this.DataContext = null;
                            //rowNum -= 1;
                            re_Search(rowNum);
                        }
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;
            
            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);
                
                ArticleBomCnt = 0;
                rowNum = 0;

                re_Search(rowNum);

                if (tlvItemList.Items.Count > 0)
                {
                    tblArticleBomCnt.Text = "▶검색결과 : " + stringFormatN0(ArticleBomCnt) + "건";
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {

                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);

            
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CanBtnControl();

                lblMsg.Visibility = Visibility.Hidden;
                rowNum = 0;
                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();

            txtParentArticle.Text = "";
            txtParentArticle.Tag = null;
            txtArticle.Text = "";
            txtArticle.Tag = null;
            txtQty.Text = "";
            txtLossQty.Text = "";


            cbosArticleGrpP.SelectedValue = false;
            cbosArticleGrpS.SelectedValue = false;

            strFlag = string.Empty;

            re_Search(rowNum);


            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }


            if (dgdArticleC.Items.Count > 0)
            {
                dgdArticleC.Items.Clear();
            }
            //re_Search_C(rowNum);
            //re_Search_P(rowNum);


            //dgdArticleP = null;
            //dgdArticleC = null;

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[6];
            lst[0] = "품명BOM 리스트";
            lst[1] = "품명 BOM_상위품목";
            lst[2] = "품명 BOM_하위품목";
            lst[3] = tlvItemList.Name;
            lst[4] = dgdArticleP.Name;
            lst[5] = dgdArticleC.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(tlvItemList.Name))
                {
                    //if (ExpExc.Check.Equals("Y"))
                    //    dt = Lib.Instance.DataGridToDTinHidden(dataTableBOM);
                    //else
                    //    dt = Lib.Instance.DataGirdToDataTable(tlvItemList);

                    Name = tlvItemList.Name;
                    Lib.Instance.GenerateExcel(dataTableBOM, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdArticleP.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdArticleP);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdArticleP);

                    Name = dgdArticleP.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
                }
                else if (ExpExc.choice.Equals(dgdArticleC.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdArticleC);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdArticleC);

                    Name = dgdArticleC.Name;
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

        //재검색
        private void re_Search(int selectedIndex)
        {
            FillGrid2();
        }

        private void re_Search_P(int selectedIndex)
        {
            FillGridP();

            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }
        }
        private void re_Search_C(int selectedIndex)
        {
            FillGridC();
        }


        #region 2020-02-11 이전 조회
        //private void FillGrid()
        //{
        //    TreeViewItem mTreeViewItem = null;
        //    TreeViewItem mTreeViewItem0 = null;
        //    TreeViewItem mTreeViewItem1 = null;
        //    TreeViewItem mTreeViewItem3 = null;
        //    TreeViewItem mTreeViewItem4 = null;
        //    string strOneParentArticleID = string.Empty;
        //    string strTwoParentArticleID = string.Empty;
        //    string strThreeParentArticleID = string.Empty;
        //    string strFoursParentArticleID = string.Empty;

        //    if (tlvItemList.Items.Count > 0)
        //    {
        //        tlvItemList.Items.Clear();
        //    }

        //    try
        //    {
        //        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
        //        sqlParameter.Add("sArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSrh.SelectedValue != null? cboArticleGrpSrh.SelectedValue.ToString() : "");
        //        sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "" );
        //        sqlParameter.Add("sDirection", strDirection);
        //        sqlParameter.Add("sIncNotuse", chkNoUse.IsChecked == true ? 1 : 0);

        //        DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleBOM", sqlParameter, false);

        //        if (ds != null)
        //        {
        //            DataTable dt = ds.Tables[0];

        //            if (dt.Rows.Count > 0)
        //            {
        //                ArticleBomCnt = dt.Rows.Count;

        //                int i = 0;
        //                DataRowCollection drc = dt.Rows;

        //                foreach (DataRow dr in drc)
        //                {
        //                    i++;
        //                    var ItemList = new Win_com_ArticleBOM_ItemList
        //                    {
        //                        Num = i,
        //                        Article = dr["Article"].ToString(),
        //                        ArticleID = dr["ArticleID"].ToString(),
        //                        LVL = dr["LVL"].ToString(),
        //                        ArticleGrpID = dr["ArticleGrpID"].ToString(),
        //                        ArticleP = dr["ArticleP"].ToString(),
        //                        ChildBuyerArticleNO = dr["ChildBuyerArticleNO"].ToString(),
        //                        ChildCnt = dr["ChildCnt"].ToString(),
        //                        FromDate = dr["FromDate"].ToString(),
        //                        LossPcntClss = dr["LossPcntClss"].ToString(),
        //                        LossQty = dr["LossQty"].ToString(),
        //                        LvlPad = dr["LvlPad"].ToString(),
        //                        ord = dr["ord"].ToString(),
        //                        PARENTArticleID = dr["PARENTArticleID"].ToString(),
        //                        ParentArticleIDS = dr["ParentArticleIDS"].ToString(),
        //                        ParentBuyerArticleNO = dr["ParentBuyerArticleNO"].ToString(),
        //                        PcntClss = dr["PcntClss"].ToString(),
        //                        Qty = dr["Qty"].ToString(),
        //                        ScraptRate = dr["ScraptRate"].ToString(),
        //                        ToDate = dr["ToDate"].ToString(),
        //                        UnitClss = dr["UnitClss"].ToString(),
        //                        UnitClssName = dr["UnitClssName"].ToString(),
        //                        Weight = dr["Weight"].ToString(),
        //                        UseYN = dr["UseYN"].ToString()
        //                    };

        //                    //if (int.Parse(ItemList.FromDate) <= int.Parse(DateTime.Today.ToString("yyyyMMdd"))
        //                    //        && int.Parse(ItemList.ToDate) >= int.Parse(DateTime.Today.ToString("yyyyMMdd")))
        //                    //{
        //                    //    ItemList.UseYN = "Y";
        //                    //}
        //                    //else
        //                    //{
        //                    //    ItemList.UseYN = "N";
        //                    //}

        //                    ItemList.Qty = Lib.Instance.returnNumString(ItemList.Qty);
        //                    ItemList.Weight = Lib.Instance.returnNumString(ItemList.Weight);
        //                    ItemList.LossQty = Lib.Instance.returnNumString(ItemList.LossQty);
        //                    ItemList.ScraptRate = Lib.Instance.returnNumString(ItemList.ScraptRate);

        //                    ItemList.FromDate_CV = Lib.Instance.StrDateTimeBar(ItemList.FromDate);
        //                    ItemList.ToDate_CV = Lib.Instance.StrDateTimeBar(ItemList.ToDate);
        //                    ItemList.FirstColumnCV = ItemList.LvlPad + "(" + ItemList.ArticleID + ")" + ItemList.BuyerArticleNo;

        //                    if (ItemList.LVL.Equals("1"))
        //                    {
        //                        mTreeViewItem = new TreeViewItem() { Header = ItemList, IsExpanded = true };
        //                        if (mTreeViewItem != null) { tlvItemList.Items.Add(mTreeViewItem); }
        //                        strOneParentArticleID = ItemList.ArticleID;
        //                    }
        //                    else
        //                    {
        //                        if (ItemList.PARENTArticleID == strOneParentArticleID)
        //                        {
        //                            mTreeViewItem0 = new TreeViewItem() { Header = ItemList, IsExpanded = true };
        //                            if (mTreeViewItem0 != null) { mTreeViewItem.Items.Add(mTreeViewItem0); }
        //                            strTwoParentArticleID = ItemList.ArticleID;
        //                        }

        //                        if (ItemList.PARENTArticleID == strTwoParentArticleID)
        //                        {
        //                            mTreeViewItem1 = new TreeViewItem() { Header = ItemList, IsExpanded = true };
        //                            if (mTreeViewItem1 != null) { mTreeViewItem0.Items.Add(mTreeViewItem1); }
        //                            strThreeParentArticleID = ItemList.ArticleID;
        //                        }

        //                        if (ItemList.PARENTArticleID == strThreeParentArticleID)
        //                        {
        //                            mTreeViewItem3 = new TreeViewItem() { Header = ItemList };
        //                            if (mTreeViewItem3 != null) { mTreeViewItem1.Items.Add(mTreeViewItem3); }
        //                            strFoursParentArticleID = ItemList.ArticleID;
        //                        }

        //                        if (ItemList.PARENTArticleID == strFoursParentArticleID)
        //                        { 
        //                            mTreeViewItem4 = new TreeViewItem() { Header = ItemList };
        //                            if (mTreeViewItem4 != null) { mTreeViewItem3.Items.Add(mTreeViewItem4); }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
        //    }
        //}
        #endregion

        #region 2020-02-11 신규 조회 FillGrid2()

        private void FillGrid2()
        {
            TreeViewItem mTreeViewItem = null;

            List<TreeViewItem> lstTree = new List<TreeViewItem>();

            List<string> lstParentArticleID = new List<string>();


            if (tlvItemList.Items.Count > 0)
            {
                tlvItemList.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", chkArticleGrpSrh.IsChecked == true && cboArticleGrpSrh.SelectedValue != null ? cboArticleGrpSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true && txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : "");
                sqlParameter.Add("sDirection", strDirection);
                sqlParameter.Add("sIncNotuse", chkNoUse.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleBOM", sqlParameter, false);

                if (ds != null)
                {
                    DataTable dt = ds.Tables[0];
                    //GLS에서 요청 2021-10-21
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    if (dt.Rows.Count > 0)
                    {
                        ArticleBomCnt = dt.Rows.Count;

                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        int BeforeLVL = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var ItemList = new Win_com_ArticleBOM_ItemList
                            {
                                Num = i,
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                LVL = ConvertInt(dr["LVL"].ToString()),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleP = dr["ArticleP"].ToString(),
                                ChildBuyerArticleNO = dr["ChildBuyerArticleNO"].ToString(),
                                ChildCnt = dr["ChildCnt"].ToString(),
                                FromDate = dr["FromDate"].ToString(),
                                LossPcntClss = dr["LossPcntClss"].ToString(),
                                LossQty = dr["LossQty"].ToString(),
                                LvlPad = dr["LvlPad"].ToString(),
                                ord = dr["ord"].ToString(),
                                PARENTArticleID = dr["PARENTArticleID"].ToString(),
                                ParentArticleIDS = dr["ParentArticleIDS"].ToString(),
                                ParentBuyerArticleNO = dr["ParentBuyerArticleNO"].ToString(),
                                PcntClss = dr["PcntClss"].ToString(),
                                Qty = dr["Qty"].ToString(),
                                ScraptRate = dr["ScraptRate"].ToString(),
                                ToDate = dr["ToDate"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                UseYN = dr["UseYN"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString()
                            };

                            ItemList.Qty = Lib.Instance.returnNumString(ItemList.Qty);
                            ItemList.Weight = Lib.Instance.returnNumString(ItemList.Weight);
                            ItemList.LossQty = Lib.Instance.returnNumString(ItemList.LossQty);
                            ItemList.ScraptRate = Lib.Instance.returnNumString(ItemList.ScraptRate);

                            ItemList.FromDate_CV = Lib.Instance.StrDateTimeBar(ItemList.FromDate);
                            ItemList.ToDate_CV = Lib.Instance.StrDateTimeBar(ItemList.ToDate);
                            ItemList.FirstColumnCV = ItemList.LvlPad + "(" + ItemList.ArticleID + ")" + ItemList.BuyerArticleNo;

                            if (ItemList.LvlPad != "")
                            {
                                ItemList.FirstColumnCV2 = ItemList.LvlPad + " " + ItemList.Article;
                            }
                            else
                            {
                                ItemList.FirstColumnCV2 = ItemList.Article;
                            }

                            //if (ItemList.LVL != 1
                            //    && BeforeLVL + 1 != ItemList.LVL)
                            //{
                            //    continue;
                            //}


                            if (ItemList.LVL == 1)
                            {
                                mTreeViewItem = new TreeViewItem() { Header = ItemList, IsExpanded = true };
                                if (mTreeViewItem != null)
                                {
                                    tlvItemList.Items.Add(mTreeViewItem);
                                }

                                if (lstParentArticleID.Count == 0)
                                {
                                    lstParentArticleID.Add(ItemList.ArticleID);

                                    lstTree.Add(mTreeViewItem);
                                }
                                else
                                {
                                    lstParentArticleID = new List<string>();
                                    lstParentArticleID.Add(ItemList.ArticleID);

                                    lstTree = new List<TreeViewItem>();
                                    lstTree.Add(mTreeViewItem);
                                }
                            }
                            else
                            {
                                string CParentArticleID = ItemList.PARENTArticleID;
                                //string PParentArticleID = ;

                                var mTreeViewItemC = new TreeViewItem() { Header = ItemList, IsExpanded = true };

                                if (lstParentArticleID.Count < ItemList.LVL)
                                {
                                    // 1 다음이 3 ~ 9 라면..?
                                    #region 1안 → 봉인
                                    // → LVL 값을 조정
                                    //while (true)
                                    //{
                                    //    if (lstParentArticleID.Count != ItemList.LVL - 1)
                                    //    {
                                    //        ItemList.LVL -= 1;
                                    //        if (ItemList.LvlPad.Length > 1)
                                    //        {
                                    //            ItemList.LvlPad = ItemList.LvlPad.Substring(0, ItemList.LvlPad.Length - 1);
                                    //        }
                                    //    }
                                    //    else
                                    //    {
                                    //        break;
                                    //    }
                                    //}

                                    //ItemList.FirstColumnCV = ItemList.LvlPad + "(" + ItemList.ArticleID + ")" + ItemList.BuyerArticleNo;
                                    #endregion
                                }

                                // 첫 다음 레벨
                                if (lstParentArticleID.Count == ItemList.LVL - 1)
                                {
                                    if (lstParentArticleID[ItemList.LVL - 2].Equals(ItemList.PARENTArticleID))
                                    {
                                        lstTree[ItemList.LVL - 2].Items.Add(mTreeViewItemC);

                                        lstParentArticleID.Add(ItemList.ArticleID);
                                        lstTree.Add(mTreeViewItemC);
                                    }
                                }
                                else // 중복?
                                {
                                    //레벨에 벗어나고 상위품번이 없는경우
                                    if (lstParentArticleID.Count == 0)
                                    {
                                        mTreeViewItem = new TreeViewItem() { Header = ItemList, IsExpanded = true };
                                        if (mTreeViewItem != null)
                                        {
                                            tlvItemList.Items.Add(mTreeViewItem);
                                        }
                                        lstParentArticleID.Add(ItemList.ArticleID);
                                        lstTree.Add(mTreeViewItem);
                                    }

                                    else if(lstParentArticleID.Count <= ItemList.LVL - 2)
                                    {
                                        continue;
                                    }

                                    else if (lstParentArticleID[ItemList.LVL - 2].Equals(ItemList.PARENTArticleID))
                                    {
                                        lstTree[ItemList.LVL - 2].Items.Add(mTreeViewItemC);

                                        lstParentArticleID[ItemList.LVL - 1] = ItemList.ArticleID;
                                        lstTree[ItemList.LVL - 1] = mTreeViewItemC;
                                    }
                                }
                            }

                            BeforeLVL = ItemList.LVL;

                        } //foreach 문 끝 
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13 현달씨 DBClose
            }
        }

        #endregion

        private void TlvItemList_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var WinBomList = tlvItemList.SelectedItem as TreeViewItem;

            if (WinBomList != null)
            {
                WinArticleBomList = WinBomList.Header as Win_com_ArticleBOM_ItemList;

                if (WinArticleBomList != null)
                {
                    this.DataContext = WinArticleBomList;
                    txtArticle.Tag = WinArticleBomList.ArticleID;

                    if (WinArticleBomList.UseYN.Equals("Y"))
                    {
                        chkUseClss.IsChecked = false;
                    }
                    else
                    {
                        chkUseClss.IsChecked = true;
                    }
                }
            }
        }

        // 사용안함 ViewBox 클릭 이벤트
        private void vbUseClss_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkUseClss.IsChecked == true)
            {
                chkUseClss.IsChecked = false;
            }
            else
            {
                chkUseClss.IsChecked = true;
            }
        }
        // 퍼센트
        private void vbPercent_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkPercent.IsChecked == true)
            {
                chkPercent.IsChecked = false;
            }
            else
            {
                chkPercent.IsChecked = true;
            }
        }
        private void vbLossPercent_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkLossPercent.IsChecked == true)
            {
                chkLossPercent.IsChecked = false;
            }
            else
            {
                chkLossPercent.IsChecked = true;
            }
        }

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strYYYY"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    // 만약에 상위품 하위품의 품명이 같다면!!!!!!!!!! → 그대로 들어가면 문제 생김
                    // 상위품을 지우고 하위품명으로만 들어가도록
                    if (txtParentArticle.Tag != null
                        && !txtParentArticle.Text.Trim().Equals("")
                        && txtArticle.Tag.ToString().Trim().Equals(txtParentArticle.Tag.ToString().Trim()))
                    {
                        txtParentArticle.Tag = "";
                    }

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("sArticleID", txtArticle.Tag.ToString());
                    sqlParameter.Add("sParentArticleID", txtParentArticle.Tag != null && !txtParentArticle.Text.Trim().Equals("") ? txtParentArticle.Tag.ToString() : "");
                    sqlParameter.Add("Qty", ConvertDouble(txtQty.Text));
                    sqlParameter.Add("PcntClss", chkPercent.IsChecked==true ? "*":"");
                    sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());
                    sqlParameter.Add("LossQty", ConvertDouble(txtLossQty.Text));
                    sqlParameter.Add("LossPcntClss", chkLossPercent.IsChecked==true ? "*" :"");
                    sqlParameter.Add("FromDate", dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("ToDate", dtpToDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("UseYN", chkUseClss.IsChecked == true ? "N":"Y");
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);
                    sqlParameter.Add("OutMsg", "");

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_iArticleBOM";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sArticleID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Article_uArticleBOM";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sArticleID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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

                    #endregion
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

            return flag;
        }

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            //if (txtParentArticle.Text.Length <= 0 || txtParentArticle.Text.Equals(""))
            //{
            //    MessageBox.Show("상위품명이 입력되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            if (txtArticle.Tag == null || txtArticle.Text.Trim().Equals(""))
            {
                MessageBox.Show("품명이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtQty.Text.Length <= 0 || txtQty.Text.Equals(""))
            {
                MessageBox.Show("소요량이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (txtLossQty.Text.Length <= 0 || txtLossQty.Text.Equals(""))
            //{
            //    MessageBox.Show("로스량이 입력되지 않았습니다.");
            //    flag = false;
            //    return flag;
            //}

            // 만약에 상위품 하위품의 품명이 같다면!!!!!!!!!! → 그대로 들어가면 문제 생김
            // 상위품 하위품 같은 품명으로 못들어가도록
            if (txtParentArticle.Tag != null
                && !txtParentArticle.Text.Trim().Equals("")
                && txtArticle.Tag.ToString().Trim().Equals(txtParentArticle.Tag.ToString().Trim()))
            {
                MessageBox.Show("상위품과 하위품이 같은 품명으로 저장이 불가능합니다.\r상위품을 지우거나 하위품을 변경해주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }
        //삭제체크
        private bool DeleteDataCheck(string strArticleID, string strPArticleID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sPArticleID", strPArticleID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Article_dArticleBOM_Check", sqlParameter, false);
                string[] resultSplit;

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    Flag = true;
                }
                else
                {
                    resultSplit = result[1].Split('/');

                    if (resultSplit.Length == 2)
                    {
                        if (Convert.ToInt32(resultSplit[0]) <= 10)
                        {
                            if (MessageBox.Show(resultSplit[1] + " 무시하고 계속 진행하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                Flag = true;
                            }
                        }
                        else
                        {
                            MessageBox.Show(resultSplit[1]);
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

            return Flag;
        }


        //상위품명
        private void txtParentArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                if (txtParentArticle.Tag != null)
                {
                    getArticleInfo(txtParentArticle.Tag.ToString());
                }

            }
        }

        //상위품명
        private void btnPfParentArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtParentArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

            if (txtParentArticle.Tag != null)
            {
                getArticleInfo(txtParentArticle.Tag.ToString());
            }
        }

        //하위품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());
                }
            }
        }

        //하위품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_BuyerArticleNo, "");

                if (txtArticle.Tag != null)
                {
                    getArticleInfo(txtArticle.Tag.ToString());
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }

        }

        #region 품명 선택시 단위 자동으로 선택

        // 위의 getArticleIdByReq 로 가져온
        // ArticleID 로 Article 정보 가져오기
        private void getArticleInfo(string setArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", setArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        var getArticleInfo = new BOMArticleInfo
                        {
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            PartGBNID = dr["PartGBNID"].ToString(),
                            ProductGrpID = dr["ProductGrpID"].ToString()
                        };

                        cboUnitClss.SelectedValue = getArticleInfo.UnitClss;
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

        #endregion // 품명 선택시 단위 자동으로 선택

        //트리뷰 펼쳐보기
        private void btnExpanding_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < tlvItemList.Items.Count; i++)
            {
                var OneTree = tlvItemList.Items[i] as TreeViewItem;
                OneTree.IsExpanded = true;
            }
        }

        //트리뷰 접어보기
        private void btnFolding_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < tlvItemList.Items.Count; i++)
            {
                var OneTree = tlvItemList.Items[i] as TreeViewItem;
                OneTree.IsExpanded = false;
            }
        }

        // 상위품목-제품 반제품 리스트 조회
        private void FillGridP()
        {
            string sql = string.Empty;

            if (dgdArticleP.Items.Count > 0)
            {
                dgdArticleP.Items.Clear();
            }

            try
            {
                DataTable dt = dataTableArticle;
                sql += " ArticleGrpID = '" + cbosArticleGrpP.SelectedValue.ToString() + "' ";
                sql += "and Article like '%" + txtSrhArticleP.Text + "%' ";
                
                //GLS에서 요청 2021-10-21
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                }

                if (dt.Rows.Count > 0)
                {
                    int i = 0;

                    foreach (DataRow dr in dt.Select(sql))
                    {
                        i++;
                        var ItemList = new Win_com_ArticleBOM_Code_U
                        {
                            Num = i.ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            //Qty = dr["Qty"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            UnitClssName = dr["UnitClssName"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            //ArticleGrp = dr["ArticleGrp"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            BuySaleMainYN = dr["BuySaleMainYN"].ToString(),
                            CoatingSpec = dr["CoatingSpec"].ToString(),
                            //DyeingID = dr["DyeingID"].ToString(),
                            FTAMgrYN = dr["FTAMgrYN"].ToString(),
                            HSCODE = dr["HSCODE"].ToString(),
                            LabelPrintYN = dr["LabelPrintYN"].ToString(),
                            NeedStockQty = dr["NeedStockQty"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString(),
                            PART_ATTR = dr["PART_ATTR"].ToString(),
                            PatternID = dr["PatternID"].ToString(),
                            //Process = dr["Process"].ToString(),
                            //ProcessID = dr["ProcessID"].ToString(),
                            //QtyPerBox = dr["QtyPerBox"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            //StuffWidth = dr["StuffWidth"].ToString(),
                            SupplyType = dr["SupplyType"].ToString(),
                            SupplyTypeName = dr["SupplyTypeName"].ToString(),
                            //Thread = dr["Thread"].ToString(),
                            //ThreadID = dr["ThreadID"].ToString(),
                            Unitprice = dr["Unitprice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UseClss = dr["UseClss"].ToString(),
                            UseingType = dr["UseingType"].ToString(),
                            UseingTypeName = dr["UseingTypeName"].ToString()
                        };

                        dgdArticleP.Items.Add(ItemList);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13
            }
        }

        // 하위품 리스트 조회
        private void FillGridC()
        {
            string sql = string.Empty;

            if (dgdArticleC.Items.Count > 0)
            {
                dgdArticleC.Items.Clear();
            }

            try
            {
                DataTable dt = dataTableArticle;
                sql += " ArticleGrpID = '" + cbosArticleGrpS.SelectedValue.ToString() + "' ";
                sql += "and Article like '%" + txtSrhArticleC.Text + "%' ";
                
                //GLS에서 요청 2021-10-21
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                }
                
                if (dt.Rows.Count > 0)
                {
                    int i = 0;

                    foreach (DataRow dr in dt.Select(sql))
                    {
                        i++;
                        var ItemList = new Win_com_ArticleBOM_Code_U
                        {
                            Num = i.ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            //Qty = dr["Qty"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            UnitClssName = dr["UnitClssName"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            //ArticleGrp = dr["ArticleGrp"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            BuySaleMainYN = dr["BuySaleMainYN"].ToString(),
                            CoatingSpec = dr["CoatingSpec"].ToString(),
                            Comments = dr["Comments"].ToString(),
                            //DyeingID = dr["DyeingID"].ToString(),
                            FTAMgrYN = dr["FTAMgrYN"].ToString(),
                            HSCODE = dr["HSCODE"].ToString(),
                            //ImageName = dr["ImageName"].ToString(),
                            LabelPrintYN = dr["LabelPrintYN"].ToString(),
                            NeedStockQty = dr["NeedStockQty"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString(),
                            PART_ATTR = dr["PART_ATTR"].ToString(),
                            PatternID = dr["PatternID"].ToString(),
                            //Process = dr["Process"].ToString(),
                            //ProcessID = dr["ProcessID"].ToString(),
                            //QtyPerBox = dr["QtyPerBox"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            //StuffWidth = dr["StuffWidth"].ToString(),
                            SupplyType = dr["SupplyType"].ToString(),
                            SupplyTypeName = dr["SupplyTypeName"].ToString(),
                            //Thread = dr["Thread"].ToString(),
                            //ThreadID = dr["ThreadID"].ToString(),
                            Unitprice = dr["Unitprice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UseClss = dr["UseClss"].ToString(),
                            UseingType = dr["UseingType"].ToString(),
                            UseingTypeName = dr["UseingTypeName"].ToString()
                        };

                        dgdArticleC.Items.Add(ItemList);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13
            }
        }

        // 상위품목-제품 반제품 리스트 조회
        private void btnSearchP_Click(object sender, RoutedEventArgs e)
        {
            if (cbosArticleGrpP.SelectedValue == null)
            {
                MessageBox.Show("품명그룹을 선택하세요");
                return;
            }

            FillGridP();
        }

        // 하위품 리스트 조회
        private void btnSearchC_Click(object sender, RoutedEventArgs e)
        {
            if (cbosArticleGrpS.SelectedValue == null)
            {
                MessageBox.Show("품명그룹을 선택하세요");
                return;
            }

            FillGridC();
        }

        private void DgdArticleP_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ArticleP = dgdArticleP.SelectedItem as Win_com_Article_U_CodeView;

            if (ArticleP != null)
            {
                txtParentArticle.Text = ArticleP.Article;
                txtParentArticle.Tag = ArticleP.ArticleID;
            }
        }

        private void DgdArticleC_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ArticleC = dgdArticleC.SelectedItem as Win_com_Article_U_CodeView;

            if (ArticleC != null)
            {
                txtArticle.Text = ArticleC.Article;
                txtArticle.Tag = ArticleC.ArticleID;
            }
        }


        // 상위품목-제품 반제품 리스트 선택버튼 클릭 이벤트
        private void btnArticleP_Click(object sender, MouseButtonEventArgs e)
        {
            WinArticleBomCode = dgdArticleP.SelectedItem as Win_com_ArticleBOM_Code_U;
            txtParentArticle.Tag = WinArticleBomCode.ArticleID.ToString();
            txtParentArticle.Text = WinArticleBomCode.Article.ToString();
        }

        // 하위품목 리스트 선택버튼 클릭 이벤트
        private void btnArticleC_Click(object sender, MouseButtonEventArgs e)
        {
            WinArticleBomCode = dgdArticleC.SelectedItem as Win_com_ArticleBOM_Code_U;
            txtArticle.Tag = WinArticleBomCode.ArticleID.ToString();
            txtArticle.Text = WinArticleBomCode.Article.ToString();
        }

        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 두자리
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

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            FillGrid2();
        }

        //소요량 숫자만 입력
        private void txtQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        //스크랩양 숫자만 입력
        private void txtLossQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
        }

        
    }

    class Win_com_ArticleBOM_ItemList : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }

        public int LVL { get; set; }
        public string ArticleID { get; set; }
        public string PARENTArticleID { get; set; }
        public string ord { get; set; }
        public string Article { get; set; }

        public string ArticleP { get; set; }
        public string ArticleGrpID { get; set; }
        public string ChildCnt { get; set; }
        public string Qty { get; set; }
        public string FromDate { get; set; }

        public string ToDate { get; set; }
        public string UnitClss { get; set; }
        public string PcntClss { get; set; }
        public string ParentArticleIDS { get; set; }
        public string LvlPad { get; set; }

        public string LossQty { get; set; }
        public string LossPcntClss { get; set; }
        public string ScraptRate { get; set; }
        public string UnitClssName { get; set; }
        public string Weight { get; set; }

        public string ChildBuyerArticleNO { get; set; }
        public string ParentBuyerArticleNO { get; set; }

        public string FromDate_CV { get; set; }
        public string ToDate_CV { get; set; }
        public string FirstColumnCV { get; set; }
        public string UseYN { get; set; }

        public string BuyerArticleNo { get; set; }

        public string FirstColumnCV2 { get; set; }

    }

    class Win_com_ArticleBOM_Code_U
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string ArticleGrpID { get; set; }
        public string Qty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string Weight { get; set; }
        public string ArticleGrp { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuySaleMainYN { get; set; }
        public string CoatingSpec { get; set; }
        public string Comments { get; set; }
        public string DyeingID { get; set; }
        public string FTAMgrYN { get; set; }
        public string HSCODE { get; set; }
        public string ImageName { get; set; }
        public string LabelPrintYN { get; set; }
        public string NeedStockQty { get; set; }
        public string OutUnitPrice { get; set; }
        public string PART_ATTR { get; set; }
        public string PatternID { get; set; }
        public string Process { get; set; }
        public string ProcessID { get; set; }
        public string QtyPerBox { get; set; }
        public string Spec { get; set; }
        public string StuffWidth { get; set; }
        public string SupplyType { get; set; }
        public string SupplyTypeName { get; set; }
        public string Thread { get; set; }
        public string ThreadID { get; set; }
        public string Unitprice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UseClss { get; set; }
        public string UseingType { get; set; }
        public string UseingTypeName { get; set; }

    }

    class BOMArticleInfo : BaseView
    {
        public string ArticleGrpID { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UnitClss { get; set; }
        public string PartGBNID { get; set; }
        public string ProductGrpID { get; set; }
    }
}
