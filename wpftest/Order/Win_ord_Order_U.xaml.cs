using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WooJung.PopUp;
using WizMes_WooJung.PopUP;

namespace WizMes_WooJung
{
    /// <summary>
    /// Win_ord_Order_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_Order_U : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string strFlag = string.Empty;
        int rowNum = 0;

        Win_ord_Order_U_CodeView OrderView = new Win_ord_Order_U_CodeView();
        Win_ord_Order_U_CodeView WinOrder = new Win_ord_Order_U_CodeView();

        Win_ord_Order_U_Sub_CodeView ComboboxSub = new Win_ord_Order_U_Sub_CodeView();
        List<Win_ord_Order_U_Sub_CodeView> winordorderusubcodeview = new List<Win_ord_Order_U_Sub_CodeView>();

        OrderArticle OrderArticle = new OrderArticle();
        ArticleData articleData = new ArticleData();
        CustomData customData = new CustomData();
        string PrimaryKey = string.Empty;

        int rowSubNum = 0;   //서브데이터그리드 rowNum

        //FTP 활용모음
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;
        string strDelFileName = string.Empty;

        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트
        List<string[]> lstExistFtpFile = new List<string[]>();

        // 촤! FTP Server 에 있는 폴더 + 파일 경로를 저장해놓고 그걸로 다운 및 업로드하자 마!
        // 이미지 이름 : 폴더이름
        Dictionary<string, string> lstFtpFilePath = new Dictionary<string, string>();

        private FTP_EX _ftp = null;
        string SketchPath = null;


        List<string[]> listFtpFile = new List<string[]>();
        private List<UploadFileInfo> _listFileInfo = new List<UploadFileInfo>();

        internal struct UploadFileInfo          //FTP.
        {
            public string Filename { get; set; }
            public FtpFileType Type { get; set; }
            public DateTime LastModifiedTime { get; set; }
            public long Size { get; set; }
            public string Filepath { get; set; }
        }

        internal enum FtpFileType
        {
            None,
            DIR,
            File
        }

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/InspectAuto";

        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Order";
        string ForderName = "Order";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        ////string FTP_ADDRESS = "ftp://222.104.222.145:25000/ImageData/McRegularInspect";
        //string FTP_ADDRESS = "ftp://192.168.0.95/ImageData/McRegularInspect";


        public Win_ord_Order_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {

            Lib.Instance.UiLoading(sender);
            btnToday_Click(null, null);
            SetComboBox();

            txtOrderCount.Text = "0 건";
            txtOrderYds.Text = "0";

            if (MainWindow.tempContent != null
                && MainWindow.tempContent.Count > 0)
            {
                string OrderId = MainWindow.tempContent[0];
                string sDate = MainWindow.tempContent[1];
                string eDate = MainWindow.tempContent[2];
                string chkYN = MainWindow.tempContent[3];


                if (chkYN.Equals("Y"))
                {
                    ChkDateSrh.IsChecked = true;
                }
                else
                {
                    ChkDateSrh.IsChecked = false;
                }

                dtpSDate.SelectedDate = DateTime.Parse(sDate.Substring(0, 4) + "-" + sDate.Substring(4, 2) + "-" + sDate.Substring(6, 2));
                dtpEDate.SelectedDate = DateTime.Parse(eDate.Substring(0, 4) + "-" + eDate.Substring(4, 2) + "-" + eDate.Substring(6, 2));

                chkOrderIDSrh.IsChecked = true;
                txtOrderIDSrh.Text = OrderId;

                //;

                rowNum = 0;
                re_Search(rowNum);

                MainWindow.tempContent.Clear();
            }
        }

        //콤보박스 만들기
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcWork = ComboBoxUtil.Instance.GetCode_SetComboBox("Work", null);
            cboWork.ItemsSource = ovcWork;
            cboWork.DisplayMemberPath = "code_name";
            cboWork.SelectedValuePath = "code_id";

            cboWorkSrh.ItemsSource = ovcWork;
            cboWorkSrh.DisplayMemberPath = "code_name";
            cboWorkSrh.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> oveOrderForm = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDFRM", "Y", "", "");
            cboOrderForm.ItemsSource = oveOrderForm;
            cboOrderForm.DisplayMemberPath = "code_name";
            cboOrderForm.SelectedValuePath = "code_id";


            ObservableCollection<CodeView> oveOrderNo = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDFLG", "Y", "", "");
            cboOrderNO.ItemsSource = oveOrderNo;
            cboOrderNO.DisplayMemberPath = "code_name";
            cboOrderNO.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcOrderClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDGBN", "Y", "", "");
            cboOrderClss.ItemsSource = ovcOrderClss;
            cboOrderClss.DisplayMemberPath = "code_name";
            cboOrderClss.SelectedValuePath = "code_id";

            cboOrderClassSrh.ItemsSource = ovcOrderClss;
            cboOrderClassSrh.DisplayMemberPath = "code_name";
            cboOrderClassSrh.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcWorkUnitClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMUNIT", "Y", "", "");
            //cboWorkUnitClss.ItemsSource = ovcWorkUnitClss;
            //cboWorkUnitClss.DisplayMemberPath = "code_name";
            //cboWorkUnitClss.SelectedValuePath = "code_id";

            cboUnitClss.ItemsSource = ovcWorkUnitClss;
            cboUnitClss.DisplayMemberPath = "code_name";
            cboUnitClss.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> ovcArticleGrpID = ComboBoxUtil.Instance.GetArticleCode_SetComboBox("", 0);
            cboArticleGroup.ItemsSource = ovcArticleGrpID;
            cboArticleGroup.DisplayMemberPath = "code_name";
            cboArticleGroup.SelectedValuePath = "code_id";

            //List<string> strValue = new List<string>();
            //strValue.Add("진행");
            //strValue.Add("완료");

            List<string[]> strNewYN = new List<string[]>();
            string[] strYes = { "0", "Y" };
            string[] strNo = { "1", "N" };
            strNewYN.Add(strYes);
            strNewYN.Add(strNo);

            ObservableCollection<CodeView> cboNewYN = ComboBoxUtil.Instance.Direct_SetComboBox(strNewYN);

            this.cboCloseClssSrh.ItemsSource = cboNewYN;
            this.cboCloseClssSrh.DisplayMemberPath = "code_name";
            this.cboCloseClssSrh.SelectedValuePath = "code_id";


            List<string[]> strArray = new List<string[]>();
            string[] strOne = { "", "진행" };
            string[] strTwo = { "1", "완료" };
            strArray.Add(strOne);
            strArray.Add(strTwo);

            ObservableCollection<CodeView> cboCloseClssSrh = ComboBoxUtil.Instance.Direct_SetComboBox(strArray);
            this.cboCloseClssSrh.ItemsSource = cboCloseClssSrh;
            this.cboCloseClssSrh.DisplayMemberPath = "code_name";
            this.cboCloseClssSrh.SelectedValuePath = "code_id";

            List<string> strVAT_Value = new List<string>();
            strVAT_Value.Add("Y");
            strVAT_Value.Add("N");
            strVAT_Value.Add("0");

            ObservableCollection<CodeView> cboVAT_YN = ComboBoxUtil.Instance.Direct_SetComboBox(strVAT_Value);
            this.cboVAT_YN.ItemsSource = cboVAT_YN;
            this.cboVAT_YN.DisplayMemberPath = "code_name";
            this.cboVAT_YN.SelectedValuePath = "code_name";

            ObservableCollection<CodeView> oveOrderFlag = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "ORDFLG", "Y", "", "");
            //영업, 생산오더만 보여주기 위해.            
            oveOrderFlag.RemoveAt(2);
            //카운트 4에서 하나 지우고 나면 카운트 3돼서 또 2번 지움
            oveOrderFlag.RemoveAt(2);

            //검색조건 수주구분 콤보박스
            cboOrderFlag.ItemsSource = oveOrderFlag;
            cboOrderFlag.DisplayMemberPath = "code_name";
            cboOrderFlag.SelectedValuePath = "code_id";

            cboOrderFlag.SelectedIndex = 1;   //영업오더가 먼저 보이는 게 맞겠지???

            //추가, 수정시 수구주분 콤보박스
            cboOrderNO.ItemsSource = oveOrderFlag;
            cboOrderNO.DisplayMemberPath = "code_name";
            cboOrderNO.SelectedValuePath = "code_id";






        }

        #region 체크박스 연동동작(상단)

        //수주일자
        private void lblDateSrh_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ChkDateSrh.IsChecked == true) { ChkDateSrh.IsChecked = false; }
            else { ChkDateSrh.IsChecked = true; }
        }

        //수주일자
        private void ChkDateSrh_Checked(object sender, RoutedEventArgs e)
        {
            if (dtpSDate != null && dtpEDate != null)
            {
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        //수주일자
        private void ChkDateSrh_Unchecked(object sender, RoutedEventArgs e)
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

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        //거래처
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true) { chkCustomSrh.IsChecked = false; }
            else { chkCustomSrh.IsChecked = true; }
        }

        //거래처
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
        }

        //거래처
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }

        //거래처
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            //if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            //else { chkArticleSrh.IsChecked = true; }
        }

        //품명
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            //txtArticleSrh.IsEnabled = true;
            //btnPfArticleSrh.IsEnabled = true;
        }

        //품명
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            //txtArticleSrh.IsEnabled = false;
            //btnPfArticleSrh.IsEnabled = false;
        }

        //품명
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_Article, "");
            //}
        }

        //품명
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            //MainWindow.pf.ReturnCode(txtArticleSrh, (int)Defind_CodeFind.DCF_Article, "");
        }

        //차종
        private void lblModelIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerModelIDSrh.IsChecked == true) { chkBuyerModelIDSrh.IsChecked = false; }
            else { chkBuyerModelIDSrh.IsChecked = true; }
        }

        //차종
        private void chkBuyerModelIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtBuyerModelIDSrh.IsEnabled = true;
            btnPfBuyerModelIDSrh.IsEnabled = true;
        }

        //차종
        private void chkBuyerModelIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBuyerModelIDSrh.IsEnabled = false;
            btnPfBuyerModelIDSrh.IsEnabled = false;
        }

        //차종
        private void txtBuyerModelIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerModelIDSrh, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
            }
        }

        //차종
        private void btnPfBuyerModelIDSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerModelIDSrh, (int)Defind_CodeFind.DCF_BUYERMODEL, "");
        }

        //품명 라벨 클릭
        private void LabelArticleSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (CheckBoxArticleSearch.IsChecked == true)
                {
                    CheckBoxArticleSearch.IsChecked = false;
                }
                else
                {
                    CheckBoxArticleSearch.IsChecked = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //품명 체크박스 체크
        private void CheckBoxArticleSearch_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextBoxArticleSearch.IsEnabled = true;
                ButtonArticleSearch.IsEnabled = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //품명 체크박스 체크해제
        private void CheckBoxArticleSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextBoxArticleSearch.IsEnabled = false;
                ButtonArticleSearch.IsEnabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //품명 텍스트박스 키다운 이벤트
        private void TextBoxArticleSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(TextBoxArticleSearch, 77, "");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //품명 플러스파인더 버튼
        private void ButtonArticleSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(TextBoxArticleSearch, 77, "");
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 라벨 클릭
        private void LabelBuyerArticleNoSearch_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (CheckBoxBuyerArticleNoSearch.IsChecked == true)
                {
                    CheckBoxBuyerArticleNoSearch.IsChecked = false;
                }
                else
                {
                    CheckBoxBuyerArticleNoSearch.IsChecked = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 체크박스 체크
        private void CheckBoxBuyerArticleNoSearch_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextBoxBuyerArticleNoSearch.IsEnabled = true;
                ButtonBuyerArticleNoSearch.IsEnabled = true;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 체크박스 체크해제
        private void CheckBoxBuyerArticleNoSearch_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextBoxBuyerArticleNoSearch.IsEnabled = false;
                ButtonBuyerArticleNoSearch.IsEnabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 텍스트박스 키다운 이벤트
        private void TextBoxBuyerArticleNoSearch_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //검색조건 - 품번 플러스파인더 버튼
        private void ButtonBuyerArticleNoSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pf.ReturnCode(TextBoxBuyerArticleNoSearch, 76, TextBoxBuyerArticleNoSearch.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - " + ee.ToString());
            }
        }

        //OrderNo 텍스트박스 키다운 이벤트
        private void TextBoxOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    txtCustom.Focus();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - TextBoxOrderNo_KeyDown : " + ee.ToString());
            }
        }

        //수주번호
        private void lblOrderIDSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderIDSrh.IsChecked == true) { chkOrderIDSrh.IsChecked = false; }
            else { chkOrderIDSrh.IsChecked = true; }
        }

        //수주번호
        private void chkOrderIDSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtOrderIDSrh.IsEnabled = true;
        }

        //수주번호
        private void chkOrderIDSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtOrderIDSrh.IsEnabled = false;
        }

        //완료구분
        private void lblCloseClssSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCloseClssSrh.IsChecked == true) { chkCloseClssSrh.IsChecked = false; }
            else { chkCloseClssSrh.IsChecked = true; }
        }

        //완료구분
        private void chkCloseClssSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboCloseClssSrh.IsEnabled = true;
        }

        //완료구분
        private void chkCloseClssSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboCloseClssSrh.IsEnabled = false;
        }

        //가공구분
        private void lblWorkSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkWorkSrh.IsChecked == true) { chkWorkSrh.IsChecked = false; }
            else { chkWorkSrh.IsChecked = true; }
        }

        //가공구분
        private void chkWorkSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboWorkSrh.IsEnabled = true;
        }

        //가공구분
        private void chkWorkSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboWorkSrh.IsEnabled = false;
        }

        //주문구분
        private void lblOrderClassSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderClassSrh.IsChecked == true) { chkOrderClassSrh.IsChecked = false; }
            else { chkOrderClassSrh.IsChecked = true; }
        }

        //주문구분
        private void chkOrderClassSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboOrderClassSrh.IsEnabled = true;
        }

        //주문구분
        private void chkOrderClassSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboOrderClassSrh.IsEnabled = false;
        }

        #endregion

        #region 수주일괄등록복사

        //수주일괄등록복사
        private void btnMassEnrollment_Click(object sender, RoutedEventArgs e)
        {
            popPreviousOrder.IsOpen = true;
        }

        private void popPreviousOrder_Opened(object sender, EventArgs e)
        {
            dtpPreviousMonth.SelectedDate = DateTime.Today.AddMonths(-1);
            dtpThisMonth.SelectedDate = DateTime.Today;
        }

        private void btnPreOrderOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "1");
                sqlParameter.Add("EDate", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "31");
                sqlParameter.Add("ChkCustom", 0);
                sqlParameter.Add("CustomID", "");

                sqlParameter.Add("ChkArticleID", 0);
                sqlParameter.Add("ArticleID", "");
                sqlParameter.Add("ChkBuyerModelID", "");
                sqlParameter.Add("BuyerModelID", "");
                sqlParameter.Add("ChkOrderID", 0);

                sqlParameter.Add("OrderID", "");
                sqlParameter.Add("ChkCloseClss", "");
                sqlParameter.Add("CloseClss", "");
                sqlParameter.Add("ChkWorkID", 0);
                sqlParameter.Add("WorkID", "");

                sqlParameter.Add("ChkOrderClss", 0);
                sqlParameter.Add("OrderClss", "");
                sqlParameter.Add("ChkOrderFlag", 0);
                sqlParameter.Add("OrderFlag", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ord_sOrder", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        if (MessageBox.Show(dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dt.Rows.Count.ToString() + " 건이 있습니다. " +
                            "무시하고 진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                        {
                            OrderCopy();
                        }
                    }
                    else
                    {
                        if (MessageBox.Show(dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "월의 수주로 복사됩니다." +
                            "진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                        {
                            OrderCopy();
                        }
                    }
                }
                else
                {
                    if (MessageBox.Show(dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM") + " 월의 수주가 " + dtpThisMonth.SelectedDate.Value.ToString("yyyyMM") + "월의 수주로 복사됩니다." +
                        "진행하시겠습니까?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                    {
                        OrderCopy();
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

            popPreviousOrder.IsOpen = false;
        }

        private void btnPreOrderCC_Click(object sender, RoutedEventArgs e)
        {
            popPreviousOrder.IsOpen = false;
        }

        private void OrderCopy()
        {
            bool Inresult = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("FromYYYYMM", dtpPreviousMonth.SelectedDate.Value.ToString("yyyyMM"));
                sqlParameter.Add("ToYYYYMM", dtpThisMonth.SelectedDate.Value.ToString("yyyyMM"));  //후에 Tag.Text 로 바꿔야 한다
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Order_iOrderCopy";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "OrderID";
                pro1.OutputLength = "10";

                Prolist.Add(pro1);
                ListParameter.Add(sqlParameter);

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                }
                else
                {
                    MessageBox.Show("수주 복사가 완료 되었습니다.");
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

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            btnExcel.Visibility = Visibility.Visible;
            grdInput.IsHitTestVisible = false;
            lblMsg.Visibility = Visibility.Hidden;
            dgdMain.IsHitTestVisible = true;

            //서브그리드 false
            SubGridAdd.IsEnabled = false;
            DataGridSub.IsHitTestVisible = false;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            btnExcel.Visibility = Visibility.Hidden;
            grdInput.IsHitTestVisible = true;
            lblMsg.Visibility = Visibility.Visible;
            dgdMain.IsHitTestVisible = false;

            //서브그리드 true
            SubGridAdd.IsEnabled = true;
            DataGridSub.IsHitTestVisible = true;

            if (strFlag.Equals("I"))
            {
                cboOrderForm.SelectedIndex = 1;
                cboOrderClss.SelectedIndex = 0;
                cboUnitClss.SelectedIndex = 0;
                //cboWorkUnitClss.SelectedIndex = 0;
            }
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "I";
            this.DataContext = null;

            //cboVAT_YN 부가세별도 값은 Y로 기본 값 셋팅
            cboVAT_YN.SelectedIndex = 0;

            //자동검사여부는 기본값 N으로
            cboAutoInspect.SelectedItem = cboN;

            //혹시 모르니까 납기일자의 체크박스가 체크되어 있을 수도 있으니까 해제
            chkDvlyDate.IsChecked = false;

            CantBtnControl();

            if (dgdNeedStuff.Items.Count > 0)
            {
                dgdNeedStuff.Items.Clear();
            }

            cboOrderForm.SelectedIndex = 0;
            cboOrderNO.SelectedIndex = 0;
            cboArticleGroup.SelectedIndex = 0;
            //cboAutoInspect.SelectedIndex = 0;
            cboOrderClss.SelectedIndex = 0;
            cboUnitClss.SelectedIndex = 0;
            cboWork.SelectedIndex = 0;

            dtpAcptDate.SelectedDate = DateTime.Today;
            dtpDvlyDate.SelectedDate = DateTime.Today;
            btnNeedStuff.IsEnabled = true;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdMain.SelectedIndex;

            //추가버튼을 누르면 cboOrderNO 콤보박스에 커서가 가도록
            //내수에 콤보박스 열리는 거 불편하다 해서 주석처리 하고, 포커스만 줘야겠다.
            //아몰라.... 그냥 거래처에 포커스 줄래. 
            //cboOrderNO.IsDropDownOpen = true;
            //cboOrderNO.Focus();
            txtCustom.Focus();


            

        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            OrderView = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

            if (OrderView != null)
            {
                //rowNum = dgdMain.SelectedIndex;
                dgdMain.IsHitTestVisible = false;
                btnNeedStuff.IsEnabled = true;
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
                CantBtnControl();
                PrimaryKey = OrderView.OrderID;
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            OrderView = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

            if (OrderView == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                string sql = "select OrderID from pl_Input where OrderID = " + OrderView.OrderID;

                DataSet ds = DataStore.Instance.QueryToDataSet(sql);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        sql = "select OrderID from OutWare where OrderID = " + OrderView.OrderID;

                        ds = DataStore.Instance.QueryToDataSet(sql);

                        if (ds != null && ds.Tables.Count > 0)
                        {
                            dt = ds.Tables[0];
                            if (dt.Rows.Count > 0)
                            {
                                MessageBox.Show("해당 수주 건은 생산 진행중이오니, 삭제하시려면 생산부터 작업지시까지 먼저 삭제해주세요.");
                            }
                            else
                            {
                                MessageBox.Show("해당 수주 건은 작업지시 진행중이오니, 삭제하시려면 작업지시 먼저 삭제해주세요.");
                            }
                        }
                    }
                    else
                    {
                        if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                            {
                                rowNum = dgdMain.SelectedIndex;
                            }

                            if (DeleteData(OrderView.OrderID))
                            {
                                rowNum -= 1;
                                re_Search(rowNum);
                            }
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

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CanBtnControl();
                lblMsg.Visibility = Visibility.Hidden;
                dgdMain.IsHitTestVisible = true;
                btnNeedStuff.IsEnabled = false;
                //re_Search(rowNum); //2021-04-28 저장 후 재조회 안되게 막음
                re_Search(rowNum);
                PrimaryKey = string.Empty;
                rowNum = 0;
                MessageBox.Show("저장이 완료되었습니다."); //2021-04-28 저장되면 저장이 완료되었다는 메세지 띄우기
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();

            //혹시 모르니까 납기일자의 체크박스가 체크되어 있을 수도 있으니까 해제
            chkDvlyDate.IsChecked = false;

            dgdMain.IsHitTestVisible = true;
            btnNeedStuff.IsEnabled = false;

            if (strFlag.Equals("U"))
            {
                re_Search(rowNum);
            }
            else
            {
                rowNum = 0;
                re_Search(rowNum);
            }

            strFlag = string.Empty;

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[2];
            lst[0] = "수주 조회 목록";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
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
            }
            lib = null;
        }

        //
        private int SelectItem(string strPrimary, DataGrid dataGrid)
        {
            int index = 0;

            try
            {
                for (int i = 0; i < dataGrid.Items.Count; i++)
                {
                    var Item = dataGrid.Items[i] as Win_ord_Order_U_CodeView;

                    if (strPrimary.Equals(Item.OrderID))
                    {
                        index = i;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return index;
        }

        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                if (PrimaryKey.Equals(string.Empty))
                {
                    dgdMain.SelectedIndex = selectedIndex;
                }
                else
                {
                    dgdMain.SelectedIndex = SelectItem(PrimaryKey, dgdMain);
                }
            }
            else
            {
                this.DataContext = null;
            }

            //CalculGridSum();
        }

        //실조회
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            if (dgdNeedStuff.Items.Count > 0)
            {
                dgdNeedStuff.Items.Clear();
            }

            try
            {

                string BuyerModelSrh = string.Empty;

                if (chkBuyerModelIDSrh.IsChecked == true && txtBuyerModelIDSrh.Tag != null
                    && txtBuyerModelIDSrh.Text.Length > 0)
                    BuyerModelSrh = txtBuyerModelIDSrh.Tag.ToString();

                //string nCloseClss = string.Empty;
                //if(chkCloseClssSrh.IsChecked == true && cboCloseClssSrh.SelectedValue.ToString() == "1")
                //{
                //    nCloseClss = "1";
                //}
                //else if(chkCloseClssSrh.IsChecked == true && cboCloseClssSrh.SelectedValue.ToString() == "0")
                //{
                //    nCloseClss = "";
                //}
                //else
                //{
                //    nCloseClss = "2";
                //}

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", ChkDateSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", ChkDateSrh.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", ChkDateSrh.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomSrh.IsChecked == true ? (txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : "") : "");

                sqlParameter.Add("ChkArticleID", CheckBoxArticleSearch.IsChecked == true ? 1 : 0); 
                sqlParameter.Add("ArticleID", CheckBoxArticleSearch.IsChecked == true ? (TextBoxArticleSearch.Tag == null ? "" : TextBoxArticleSearch.Tag.ToString()) : "");
                sqlParameter.Add("ChkBuyerModelID", chkBuyerModelIDSrh.IsChecked == true ? 1 : 0 );
                sqlParameter.Add("BuyerModelID", chkBuyerModelIDSrh.IsChecked == true ? (txtBuyerModelIDSrh.Tag == null ? "" : txtBuyerModelIDSrh.Tag.ToString()) : "");
                sqlParameter.Add("ChkOrderID", chkOrderIDSrh.IsChecked == true ? 1 : 0);
                
                sqlParameter.Add("OrderID", chkOrderIDSrh.IsChecked == true ? (txtOrderIDSrh.Text == string.Empty ? "" : txtOrderIDSrh.Text) : "");
                sqlParameter.Add("ChkCloseClss", chkCloseClssSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CloseClss", chkCloseClssSrh.IsChecked == true ? (cboCloseClssSrh.SelectedValue == null ? "" : cboCloseClssSrh.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkWorkID", chkWorkSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("WorkID", chkWorkSrh.IsChecked == true ? (cboWorkSrh.SelectedValue == null ? "" : cboWorkSrh.SelectedValue.ToString()) : "");
                
                sqlParameter.Add("ChkOrderClss", chkOrderClassSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("OrderClss", chkOrderClassSrh.IsChecked == true ? (cboOrderClassSrh.SelectedValue == null ? "" : cboOrderClassSrh.SelectedValue.ToString()) : "");
                sqlParameter.Add("ChkOrderFlag", chkOrderFlag.IsChecked == true ? 1 : 0);
                sqlParameter.Add("OrderFlag", chkOrderFlag.IsChecked == true ? (cboOrderFlag.SelectedValue == null ? "" : cboOrderFlag.SelectedValue.ToString()) : "");



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ord_sOrder", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var OrderCodeView = new Win_ord_Order_U_CodeView
                            {
                                Num = i,

                                OrderID = dr["OrderID"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                CompanyID = dr["CompanyID"].ToString(),
                                BuyerID = dr["BuyerID"].ToString(),

                                InCustomID = dr["InCustomID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                PoNo = dr["PoNo"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),

                                OrderForm = dr["OrderForm"].ToString(),
                                OrderFormName = dr["OrderFormName"].ToString(),
                                OrderClss = dr["OrderClss"].ToString(),
                                OrderClssName = dr["OrderClssName"].ToString(),
                                BrandClss = dr["BrandClss"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                DvlyDate = dr["DvlyDate"].ToString(),

                                DvlyPlace = dr["DvlyPlace"].ToString(),
                                WorkID = dr["WorkID"].ToString(),
                                WorkName = dr["WorkName"].ToString(),
                                ExchRate = dr["ExchRate"].ToString(),
                                OrderQty = dr["OrderQty"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),

                                Vat_IND_YN = dr["Vat_IND_YN"].ToString(),
                                ModifyClss = dr["ModifyClss"].ToString(),
                                ModifyRemark = dr["ModifyRemark"].ToString(),
                                CancelRemark = dr["CancelRemark"].ToString(),

                                Remark = dr["Remark"].ToString(),
                                CloseClss = dr["CloseClss"].ToString(),
                                //CloseDate = dr["CloseDate"].ToString(),
                                OrderFlag = dr["OrderFlag"].ToString(),
                                //OrderEnd = dr["OrderEnd"].ToString(),

                                OrderSpec = dr["OrderSpec"].ToString(),
                                KInCustom = dr["KInCustom"].ToString(),
                                
                                ProductAutoInspectYN = dr["ProductAutoInspectYN"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                UnitPrice = stringFormatN0(dr["UnitPrice"]),

                                //NewArticleQty = dr["NewArticleQty"].ToString(),
                                //RePolishingQty = dr["RePolishingQty"].ToString(),


                            };

                            if (Lib.Instance.IsNumOrAnother(OrderCodeView.Amount))
                            {
                                OrderCodeView.Amount_CV = Lib.Instance.returnNumStringZero(OrderCodeView.Amount);
                            }

                            if (OrderCodeView.DvlyDate != null && !OrderCodeView.DvlyDate.Equals(string.Empty))
                            {
                                OrderCodeView.DvlyDate_CV = Lib.Instance.StrDateTimeBar(OrderCodeView.DvlyDate);
                            }

                            if (OrderCodeView.AcptDate != null && !OrderCodeView.AcptDate.Equals(string.Empty))
                            {
                                OrderCodeView.AcptDate_CV = Lib.Instance.StrDateTimeBar(OrderCodeView.AcptDate);
                            }

                            OrderCodeView.OrderQty = Lib.Instance.returnNumStringZero(OrderCodeView.OrderQty);
                            //OrderCodeView.NewArticleQty = Lib.Instance.returnNumStringZero(OrderCodeView.NewArticleQty);
                            //OrderCodeView.RePolishingQty = Lib.Instance.returnNumStringZero(OrderCodeView.RePolishingQty);
                            //OrderCodeView.UnitPrice = Lib.Instance.returnNumStringZero(OrderCodeView.UnitPrice);

                            dgdMain.Items.Add(OrderCodeView);
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

        private DataRow FillOneOrderData(string strOrderID)
        {
            DataRow dr = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("OrderID", strOrderID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sOrderOne", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        dr = drc[0];
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

            return dr;
        }

        //그리드 하단 합계 표시
        private void CalculGridSum()
        {
            Int64 numYDS = 0;
            double numTotal = 0;

            txtOrderCount.Text = string.Format("{0:N0}", dgdMain.Items.Count) + " 건";
            if (dgdMain.Items.Count > 0)
            {
                Win_ord_Order_U_CodeView WinOrder = new Win_ord_Order_U_CodeView();

                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    WinOrder = dgdMain.Items[i] as Win_ord_Order_U_CodeView;

                    if (WinOrder.UnitClss.Equals("0"))
                    {
                        numYDS += Int64.Parse(WinOrder.OrderQty.Replace(",", ""));
                        numTotal += double.Parse(WinOrder.Amount.Replace(",", ""));
                    }
                    else
                    {
                        numYDS += Int64.Parse(WinOrder.OrderQty.Replace(",", ""));
                        numTotal += double.Parse(WinOrder.Amount.Replace(",", ""));
                    }
                }
            }

            txtOrderYds.Text = string.Format("{0:N0}", numYDS) + " EA";
            //txtOrderAmount.Text = string.Format("{0:0,0.0}", numTotal) + " 원";
            txtOrderAmount.Text = string.Format("{0:N0}", numTotal) + " 원";
        }

        #region 안씀
        //
        //private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    OrderView = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

        //    if (OrderView != null)
        //    {
        //        //셀렉트한 값으로 태그값 넣어주기. 혹시 모르니까.
        //        txtCustom.Tag = OrderView.CustomID;
        //        txtArticle.Tag = OrderView.ArticleID;

        //        this.DataContext = OrderView;
        //        if (OrderView.ProductAutoInspectYN.Equals("Y"))
        //        {
        //            cboAutoInspect.SelectedItem = cboY;
        //        }
        //        else
        //        {
        //            cboAutoInspect.SelectedItem = cboN;
        //        }
        //        CallArticleData(OrderView.ArticleID);
        //        CallCustomData(OrderView.CustomID);
        //        FillNeedStockQty(OrderView.ArticleID, txtAmount.Text.Replace(",", ""));


        //        //납기일자에 값이 있으면 체크, 그게 아니면 해제
        //        if (!OrderView.DvlyDate.Trim().Equals(""))
        //        {
        //            chkDvlyDate.IsChecked = true;
        //        }
        //        else
        //        {
        //            chkDvlyDate.IsChecked = false;
        //        }


        //    }
        //}
        #endregion
        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", strID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Order_dOrder", sqlParameter, false);

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

        /// <summary>
        /// 실저장
        /// </summary>
        /// <param name="strID"></param>
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
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("OrderID", txtOrderID.Text == string.Empty ? "" : txtOrderID.Text);
                    sqlParameter.Add("CustomID", txtCustom.Tag.ToString());
                    sqlParameter.Add("OrderNO", TextBoxOrderNo.Text == string.Empty ? "" : TextBoxOrderNo.Text);
                    sqlParameter.Add("PoNo", txtPONO.Text);
                    sqlParameter.Add("OrderForm", cboOrderForm.SelectedValue.ToString());

                    sqlParameter.Add("OrderClss", cboOrderClss.SelectedValue.ToString());
                    sqlParameter.Add("AcptDate", dtpAcptDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("DvlyDate", chkDvlyDate.IsChecked == true ? dtpDvlyDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    //sqlParameter.Add("ArticleID", txtArticle.Tag.ToString());
                    sqlParameter.Add("ArticleGrpID", cboArticleGroup.SelectedValue.ToString());

                    sqlParameter.Add("DvlyPlace", txtDylvLoc.Text);
                    sqlParameter.Add("WorkID", cboWork.SelectedValue.ToString());
                    // sqlParameter.Add("PriceClss", 0);
                    sqlParameter.Add("ExchRate", 0.00);
                    sqlParameter.Add("Vat_IND_YN", cboVAT_YN.SelectedValue.ToString());

                    sqlParameter.Add("OrderQty", int.Parse(txtAmount.Text.Replace(",", "")));

                    //sqlParameter.Add("RePolishingQty", txtRePolishing.Text != null && !txtRePolishing.Text.Trim().Equals("") ? ConvertDouble(txtRePolishing.Text) : 0);
                    //sqlParameter.Add("NewArticleQty", txtNewArticle.Text != null && !txtNewArticle.Text.Trim().Equals("") ? ConvertDouble(txtNewArticle.Text) : 0);

                    sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());

                    sqlParameter.Add("Remark", txtComments.Text);


                    sqlParameter.Add("OrderFlag", 0);

                    sqlParameter.Add("InCustomID", txtCustom.Tag.ToString());
                    //sqlParameter.Add("UnitPriceClss", articleData.UnitPriceClss);
                    sqlParameter.Add("UnitPriceClss", "0");
                    //sqlParameter.Add("WorkUnitClss", cboWorkUnitClss.SelectedValue.ToString());

                    sqlParameter.Add("OrderSpec", "");
                    sqlParameter.Add("BuyerModelID", txtModel.Tag != null ? txtModel.Tag.ToString() : "");
                    sqlParameter.Add("ProductAutoInspectYN", cboAutoInspect.SelectedItem == cboY ? "Y" : "N");
                    sqlParameter.Add("UnitPrice", ConvertDouble(txtUnitPrice.Text));

                    #region 추가

                    if (strFlag.Equals("I"))
                    {

                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Order_iOrder";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "OrderID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "OrderID")
                                {
                                    sGetID = kv.value;

                                    PrimaryKey = sGetID;

                                    flag = true;

       
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[1].value.ToString());
                            //flag = false;
                            return false;
                        }

                        Prolist.Clear();
                        ListParameter.Clear();

                        //Sub 그리드 추가
                        
                        

                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();

                        sqlParameter.Add("OrderID", sGetID);
                        sqlParameter.Add("OrderSeq", 0);
                        sqlParameter.Add("ArticleID", txtBuyerArticleNO.Tag.ToString());
                        sqlParameter.Add("NewProductYN", "");
                        sqlParameter.Add("Remark", "");
                        sqlParameter.Add("UnitPriceClss", "0");
                        sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());

                        //sqlParameter.Add("ColorQty", OrderSub.ColorQty == string.Empty ? "0" : OrderSub.ColorQty.ToString().Replace(",", ""));
                        sqlParameter.Add("ColorQty", int.Parse(txtAmount.Text.Replace(",", "")));

                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro2 = new Procedure();
                        pro2.Name = "xp_ord_iOrderSub";
                        pro2.OutputUseYN = "N";
                        pro2.OutputName = "OrderID";
                        pro2.OutputLength = "10";

                        Prolist.Add(pro2);
                        ListParameter.Add(sqlParameter);

                        
                    }
                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_Order_uOrder";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "xp_Order_uOrder";
                        pro3.OutputLength = "10";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);

                        // 모든것을 삭제한 후에, 새롭게 추가
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("OrderID", txtOrderID.Text);
                        //sqlParameter.Add("Seq", "");

                        Procedure pro4 = new Procedure();
                        pro4.Name = "xp_Order_dOrderColorAll";
                        pro4.OutputUseYN = "N";
                        pro4.OutputName = "OrderID";
                        pro4.OutputLength = "10";

                        Prolist.Add(pro4);
                        ListParameter.Add(sqlParameter);


                        //Sub 그리드 추가
                        for (int i = 0; i < DataGridSub.Items.Count; i++)
                        {
                            var OrderSub = DataGridSub.Items[i] as Win_ord_Order_U_Sub_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("OrderID", txtOrderID.Text);
                            sqlParameter.Add("OrderSeq", i + 1);
                            //sqlParameter.Add("ArticleID", OrderSub.ArticleID);
                            sqlParameter.Add("ArticleID", txtBuyerArticleNO.Tag.ToString());
                            sqlParameter.Add("NewProductYN", OrderSub.NewProductYN);
                            sqlParameter.Add("Remark", OrderSub.Remark);
                            sqlParameter.Add("UnitPriceClss", "0");
                            sqlParameter.Add("UnitClss", cboUnitClss.SelectedValue.ToString());

                            sqlParameter.Add("ColorQty", OrderSub.ColorQty == string.Empty ? "0" : OrderSub.ColorQty.ToString().Replace(",", ""));

                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_ord_iOrderSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "OrderID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }
                    }
                    #endregion

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

        private bool CheckData()
        {
            bool flag = true;

            if (txtCustom.Text.Length <= 0 || txtCustom.Tag == null)
            {
                MessageBox.Show("거래처가 입력되지 않았습니다. 먼저 거래처를 입력해주세요");
                flag = false;
                return flag;
            }



            if (txtAmount.Text.Length <= 0)
            {
                MessageBox.Show("총 주문량이 입력되지 않았습니다. 먼저 총 주문량을 입력해주세요");
                flag = false;
                return flag;
            }

            if (cboOrderForm.SelectedValue == null)
            {
                MessageBox.Show("주문형태가 선택되지 않았습니다. 먼저 주문형태를 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboOrderClss.SelectedValue == null)
            {
                MessageBox.Show("주문구분이 선택되지 않았습니다. 먼저 주문구분을 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboUnitClss.SelectedValue == null)
            {
                MessageBox.Show("주문기준이 선택되지 않았습니다. 먼저 주문기준을 선택해주세요");
                flag = false;
                return flag;
            }

            //if (cboWorkUnitClss.SelectedValue == null)
            //{
            //    MessageBox.Show("수불기준이 선택되지 않았습니다. 먼저 수불기준을 선택해주세요");
            //    flag = false;
            //    return flag;
            //}

            //if (cboArticleGroup.SelectedValue == null)
            //{
            //    MessageBox.Show("품명종류가 선택되지 않았습니다. 먼저 품명종류를 선택해주세요");
            //    flag = false;
            //    return flag;
            //}

            if (cboWork.SelectedValue == null)
            {
                MessageBox.Show("가공구분이 선택되지 않았습니다. 먼저 가공구분을 선택해주세요");
                flag = false;
                return flag;
            }

            if (cboVAT_YN.SelectedValue == null)
            {
                MessageBox.Show("부가세별도여부가 선택되지 않았습니다. 먼저 부가세별도여부를 선택해주세요");
                flag = false;
                return flag;
            }


            ////수정시, 작업지시가 내려간 수주라면 품명을 변경하지 못하도록. 2020.03.25, 장가빈
            //if (strFlag.Equals("U"))
            //{

            //    List<Procedure> Prolist = new List<Procedure>();
            //    List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            //    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //    sqlParameter.Clear();
            //    sqlParameter.Add("OrderID", txtOrderID.Text);
            //    sqlParameter.Add("NewArticleID", txtArticle.Tag.ToString().Trim());
            //    sqlParameter.Add("sMessage", "");

            //    Procedure pro1 = new Procedure();
            //    pro1.Name = "xp_Order_chkuOrder";
            //    pro1.OutputUseYN = "Y";
            //    pro1.OutputName = "sMessage";
            //    pro1.OutputLength = "1000";

            //    Prolist.Add(pro1);
            //    ListParameter.Add(sqlParameter);

            //    //동운씨가 만든 아웃풋 값 찾는 방법
            //    List<KeyValue> list_Result = new List<KeyValue>();
            //    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);

            //    //Prolist.RemoveAt(0);
            //    //ListParameter.RemoveAt(0);

            //    string sGetID = string.Empty;

            //    if (list_Result[0].key.ToLower() == "success")
            //    {
            //        //list_Result.RemoveAt(0);
            //        for (int i = 0; i < list_Result.Count; i++)
            //        {
            //            KeyValue kv = list_Result[i];
            //            if (kv.key == "sMessage")
            //            {
            //                sGetID = kv.value;

            //                if (sGetID.Equals(""))
            //                {
            //                    continue;
            //                }

            //                MessageBox.Show("알림 : " + sGetID.ToString());
            //                flag = false;


            //                //저장된 원래의 tag값 다시 넣어주기 2020.04.03, 장가빈
            //                //품명 변경 시도한 tag가 실패 후 계속 남아있는 것을 해결하기 위함.
            //                txtArticle.Tag = OrderView.ArticleID;


            //                //strFlag = string.Empty;
            //            }
            //        }
            //    }
            //    Prolist.Clear();
            //    ListParameter.Clear();
            //    return flag;
            //}
            return flag;
        }

        #region 입력시 Event

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

                if (txtCustom.Tag != null)
                {
                    CallCustomData(txtCustom.Tag.ToString());
                    txtDylvLoc.Text = txtCustom.Text;
                    txtInCustom.Text = txtCustom.Text;
                    txtInCustom.Tag = txtCustom.Tag;

                }

                //납품거래처 -> 납품 장소 커서이동
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    txtDylvLoc.Focus();
                }
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");

            if (txtCustom.Tag != null)
            {
                CallCustomData(txtCustom.Tag.ToString());
                txtDylvLoc.Text = txtCustom.Text;
            }

            //플러스 파인더 선택 후 납품 장소로 커서 이동
            txtDylvLoc.Focus();

        }

        //품명
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {

                    if (txtCustom != null && txtCustom.Text != "")
                    {   //선택된 납품거래처에 따른 품명만 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, 57, txtCustom.Tag.ToString().Trim());

                        //품번을 품명처럼 쓴다고 해서 품번을 조회하도록 2020.03.17, 장가빈
                        MainWindow.pf.ReturnCodeGLS(txtArticle, 7070, txtCustom.Tag.ToString().Trim());
                        
                    }
                    else
                    {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_Article, "");

                        //품번을 품명처럼 쓴다고 해서 품번을 조회하도록 2020.03.17, 장가빈
                        MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");
                    }


                    if (txtArticle.Tag != null)
                    {
                        CallArticleData(txtArticle.Tag.ToString());
                        //품명종류 대입(ex.제품 등)
                        cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                        //품번 대입
                        //txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                        //품명 대입
                        txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                        //단가 대입
                        txtUnitPrice.Text = articleData.OutUnitPrice;

                    }

                    //플러스 파인더 작동 후 규격으로 커서 이동
                    txtSpec.Focus();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (txtCustom != null && txtCustom.Text != "")
                {   //선택된 납품거래처에 따른 품명만 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtArticle, 7070, txtCustom.Tag.ToString().Trim());
                }
                else
                {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtArticle, 7071, "");
                }

                if (txtArticle.Tag != null)
                {
                    CallArticleData(txtArticle.Tag.ToString());
                    //품명종류 대입(ex.제품 등)
                    cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                    //품번 대입
                    //txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                    //품명 대입
                    //txtBuyerArticleNO.Text = articleData.Article;
                    //단가 대입
                    txtUnitPrice.Text = articleData.OutUnitPrice;
                }

                //플러스 파인더 작동 후 규격으로 커서 이동
                txtSpec.Focus();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //차종 키다운 
        private void txtModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");

                //주문 형태 콤보박스 열기
                //cboOrderForm.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
                cboOrderForm.Focus();

            }
        }

        //차종
        private void btnPfModel_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtModel, (int)Defind_CodeFind.DCF_BUYERMODEL, "");

            //주문 형태 콤보박스 열기
            //cboOrderForm.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
        }

        private void CallArticleData(string strArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        articleData = new ArticleData
                        {
                            ArticleID = dr["ArticleID"].ToString(),
                            Article = dr["Article"].ToString(),
                            ThreadID = dr["ThreadID"].ToString(),
                            thread = dr["thread"].ToString(),
                            StuffWidth = dr["StuffWidth"].ToString(),
                            DyeingID = dr["DyeingID"].ToString(),
                            Weight = dr["Weight"].ToString(),
                            Spec = dr["Spec"].ToString(),
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            Code_Name = dr["Code_Name"].ToString(),
                            //ProcessName = dr["ProcessName"].ToString(),
                            //HSCode = dr["HSCode"].ToString(),
                            OutUnitPrice = dr["OutUnitPrice"].ToString()
                        };
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

        private void CallCustomData(string strCustomID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("CustomID", strCustomID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sCustomData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        customData = new CustomData
                        {
                            CalClss = dr["CalClss"].ToString(),
                            LossClss = dr["LossClss"].ToString(),
                            PointClss = dr["PointClss"].ToString(),
                            SpendingClss = dr["SpendingClss"].ToString(),
                            WorkingClss = dr["WorkingClss"].ToString()
                        };
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

        private void chkDvlyDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpDvlyDate.IsEnabled = true;
        }

        private void chkDvlyDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpDvlyDate.IsEnabled = false;
        }

        ////자재필요량조회 
        //private void btnNeedStuff_Click(object sender, RoutedEventArgs e)
        //{
        //    if (txtArticle.Tag == null)
        //    {
        //        MessageBox.Show("먼저 품명을 선택해주세요");
        //        return;
        //    }

        //    if (txtAmount.Text.Replace(" ", "").Equals(""))
        //    {
        //        MessageBox.Show("먼저 총 주문량을 입력해주세요");
        //        return;
        //    }

        //    //자재필요량조회에 필요한 파라미터 값을 넘겨주자, 품명이랑 주문량
        //    FillNeedStockQty(txtArticle.Tag.ToString(), txtAmount.Text.Replace(",", ""));
        //}

        //자재필요량조회
        private void FillNeedStockQty(string strArticleID, string strQty)
        {
            if (dgdNeedStuff.Items.Count > 0)
            {
                dgdNeedStuff.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleNeedStockQty", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var NeedStockQty = new ArticleNeedStockQty()
                            {
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                NeedQty = dr["NeedQty"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString()
                            };

                            if (Lib.Instance.IsNumOrAnother(NeedStockQty.NeedQty))
                            {
                                if (Lib.Instance.IsNumOrAnother(strQty))
                                {
                                    double doubleTemp = double.Parse(NeedStockQty.NeedQty) * double.Parse(strQty);
                                    NeedStockQty.NeedQty = string.Format("{0:N0}", doubleTemp);
                                }
                            }

                            dgdNeedStuff.Items.Add(NeedStockQty);
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

        private void dgdMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (btnUpdate.IsEnabled == true)
            {
                if (e.ClickCount == 2)
                {
                    btnUpdate_Click(null, null);
                }
            }
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

        #region keyDown 이벤트(커서이동)

        //Order No. -> 납품거래처
        //private void TxtOrderNO_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        e.Handled = true;
        //        txtCustom.Focus();
        //    }
        //}

        ////납품 장소 -> 품명
        //private void TxtDylvLoc_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        e.Handled = true;
        //        txtArticle.Focus();
        //    }
        //}

        //규격 -> 접수 일자 
        //private void TxtSpec_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        e.Handled = true;

        //        //접수일자의 캘린더를 펼치고 
        //        //dtpAcptDate.IsDropDownOpen = true; 

        //        //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함

        //        //접수일자에 그냥 커서 이동
        //        dtpAcptDate.Focus();

        //    }
        //}

        private void DtpAcptDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                //날짜는 직접 입력하면 되니까, 그냥 가공 구분으로 커서 이동 하자.
                //dtpAcptDate.IsDropDownOpen = true;

                //가공구분 콤보박스에 커서만 이동되도록
                cboWork.Focus();
            }
        }

        //접수일자 캘린더가 닫히면 납기일자 체크박스로 포커스 이동
        private void DtpAcptDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            //chkDvlyDate.Focus();

            //가공구분 콤보박스에 커서만 이동되도록
            cboWork.Focus();
        }

        //납기일자 체크박스 -> 가공 구분
        private void ChkDvlyDate_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            //체크박스에서 스페이스바를 누르면
            if (e.Key == Key.Space)
            {
                e.Handled = true;

                //체크박스 체크가 되고
                chkDvlyDate.IsChecked = true;

                //납기일자 데이트피커로 포커스 이동
                dtpDvlyDate.IsDropDownOpen = true;

            }
            //그게 아니고 체크박스에서 엔터를 누르면
            else if (e.Key == Key.Enter)
            {
                e.Handled = true;

                //체크박스 체크가 해제인 채
                chkDvlyDate.IsChecked = false;

                //가공구분 콤보박스 열기
                //cboWork.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
            }
        }

        //납기일자 데이트피커가 닫혔을 때 가공구분 콤보박스를 열기
        private void DtpDvlyDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            //가공구분 콤보박스 열기
            //cboWork.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
        }

        //가공구분 콤보박스 닫히면 P/O NO.로 포커스 이동
        private void CboWork_DropDownClosed(object sender, EventArgs e)
        {
            txtPONO.Focus();
        }

        //P/O NO. -> 차종
        private void TxtPONO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtModel.Focus();
            }
        }

        //품번 -> 주문형태
        //private void TxtBuyerArticleNO_PreviewKeyDown(object sender, KeyEventArgs e)
        //{
        //    //주문 형태 콤보박스 열기
        //    cboOrderForm.IsDropDownOpen = true;
        //}

        //주문형태 콤보박스가 닫히면 주문 구분 콤보박스로 열기
        private void CboOrderForm_DropDownClosed(object sender, EventArgs e)
        {
            //cboOrderClss.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
        }

        //주문 구분 콤보박스가 닫히면 총 주문량 텍스트 박스 커서 이동
        private void CboOrderClss_DropDownClosed(object sender, EventArgs e)
        {
            txtAmount.Focus();
        }

        //총 주문량 -> 주문기준 콤보박스 열기
        private void TxtAmount_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                //cboUnitClss.IsDropDownOpen = true;
                cboUnitClss.Focus();
            }

            //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함

        }

        //주문 기준 콤보박스가 닫히면 품명 종류 콤보박스 열기
        private void CboUnitClss_DropDownClosed(object sender, EventArgs e)
        {
            //cboArticleGroup.IsDropDownOpen = true; //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함
        }

        //품명 종류 콤보박스가 닫히면 단가 텍스트 박스 열기
        private void CboArticleGroup_DropDownClosed(object sender, EventArgs e)
        {
            txtUnitPrice.Focus();
        }

        //단가 텍스트박스에서 부가세 별도 콤보박스 열기
        private void TxtUnitPrice_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            //if (e.Key == Key.Enter)
            //{
            //    e.Handled = true;

            //    cboVAT_YN.Focus();
            //    cboVAT_YN.IsDropDownOpen = true;
            //}

            //2020.02.14 장가빈, 수정시 콤보박스 자동 열리는 것 불편하대서 주석처리 함


            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                cboVAT_YN.Focus();
            }
        }

        //부가세별도 콤보박스 닫히면 비고사항으로 포커스 이동
        private void CboVAT_YN_DropDownClosed(object sender, EventArgs e)
        {
            txtComments.Focus();
        }

        //콤보박스 닫히면 납품거래처에 커서가 가도록
        private void CboOrderNO_DropDownClosed(object sender, EventArgs e)
        {
            txtCustom.Focus();
        }


        //주문구분(내수, 수출, 시가공 어쩌구..)
        private void CboOrderNO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCustom.Focus();
            }
        }

        //가공구분 키다운
        private void CboWork_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtPONO.Focus();
            }
        }

        //주문형태 키다운
        private void CboOrderForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboOrderClss.Focus();
            }
        }

        //주문구분 키다운
        private void CboOrderClss_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtAmount.Focus();
            }
        }

        //주문 기준 키다운
        private void CboUnitClss_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                cboArticleGroup.Focus();
            }
        }
        //품명 종류 키다운
        private void CboArticleGroup_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtUnitPrice.Focus();
            }
        }

        //부가세별도 키다운
        private void CboVAT_YN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtComments.Focus();
            }
        }

        //숫자 외에 다른 문자열 못들어오도록
        public bool IsNumeric(string source)
        {

            Regex regex = new Regex("[^0-9.-]+");
            return !regex.IsMatch(source);
        }

        //총주문량 숫자 외에 못들어가게 
        private void TxtAmount_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumeric(e.Text);
        }

        //단가 숫자 외에 못들어가게
        private void TxtUnitPrice_TextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsNumeric(e.Text);
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

        //수주구분 라벨
        private void LblOrderFlag_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkOrderFlag.IsChecked == true)
            {
                chkOrderFlag.IsChecked = false;
                cboOrderFlag.IsEnabled = false;
            }
            else
            {
                chkOrderFlag.IsChecked = true;
                cboOrderFlag.IsEnabled = true;
            }
        }

        //수주구분 체크박스 체크
        private void ChkOrderFlag_Checked(object sender, RoutedEventArgs e)
        {
            //cboOrderFlag.IsEnabled = true;
        }

        //수주구분 체크박스 체크 해제
        private void ChkOrderFlag_Unchecked(object sender, RoutedEventArgs e)
        {
            //cboOrderFlag.IsEnabled = false;
        }

        //매출거래처 
        private void txtInCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtInCustom, 72, "");
            }
        }

        //매출거래처
        private void btnPfInCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtInCustom, 72, "");
        }
        #endregion keydown 이벤트


        #region 서브 그리드 키다운 이벤트
        //서브 데이터 그리드 품명 텍스트박스 키다운 이벤트
        private void DataGridSubTextBoxArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (strFlag == "I" || strFlag == "U")
                {
                    var SubItem = DataGridSub.CurrentItem as Win_ord_Order_U_Sub_CodeView;

                    if (e.Key == Key.Enter)
                    {
                        TextBox tb1 = sender as TextBox;
                        pf.ReturnCode(tb1, 5000, SubItem.Article);

                        if (tb1.Tag != null)
                        {
                            SubItem.ArticleID = tb1.Tag.ToString();
                            SubItem.Article = tb1.Text;
                        }

                        string FindText = tb1.Text;
                        string FindTag = tb1.Tag.ToString();
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("ArticleID", FindTag);
                        DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ord_sArticle", sqlParameter, false);

                        DataTable dt = null;
                        dt = ds.Tables[0];
                        if (dt.Rows.Count != 0)
                        {
                            //SubItem.Color = dt.Rows[0]["Color"].ToString();
                            //SubItem.UnitPrice = dt.Rows[0]["UnitPrice"].ToString();
                            SubItem.Spec = dt.Rows[0]["Spec"].ToString();
                            //SubItem.Weight = dt.Rows[0]["Weight"].ToString();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSubTextBoxArticle_KeyDown : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 수량 숫자만 입력
        private void DataGridTextBoxColorQty_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                Lib.Instance.CheckIsNumeric((TextBox)sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_PreviewTextInput : " + ee.ToString());
            }
        }

        //서브 데이터 그리드 수량 변경 이벤트
        private void DataGridTextBoxColorQty_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                SumColorQty();
                SumNewArticleQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridTextBoxColorQty_TextChanged : " + ee.ToString());
            }
        }
        #endregion

        #region 서브 그리드 행 추가, 삭제
        //데이터 그리드 서브 행 추가 버튼
        private void ButtonDataGridSubRowAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowAdd();

                //int colCount = DataGridSub.Columns.IndexOf(d)
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowAdd_Click : " + ee.ToString());
            }
        }

        //데이터 그리드 서브 행 삭제 버튼
        private void ButtonDataGridSubRowDel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SubRowDel();
                SumColorQty();
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - ButtonDataGridSubRowDel_Click : " + ee.ToString());
            }
        }

        //서브 그리드 추가
        private void SubRowAdd()
        {
            try
            {
                int index = DataGridSub.Items.Count;

                var WOOUSC = new Win_ord_Order_U_Sub_CodeView()
                {
                    Num = index + 1,
                    OrderID = "",
                    OrderSeq = "",
                    ArticleID = "",
                    Article = "",

                    NewProductYN = "",
                    ColorQty = "",
                    Remark = "",


                    //Color = "",
                    //UnitPrice = "",
                    //UnitPriceClss = "",
                    //OrderEnd = "",

                    //ProdQty = "",
                    //ProdDate = "",
                    //ReProdQty = "",
                    //ReProdDate = "",
                    //RGB = "",
                };
                DataGridSub.Items.Add(WOOUSC);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SubRowAdd : " + ee.ToString());
            }
        }

        //서브 그리드 삭제
        private void SubRowDel()
        {
            try
            {
                if (DataGridSub.Items.Count > 0)
                {
                    if (DataGridSub.SelectedItem != null)
                    {
                        if (DataGridSub.CurrentItem != null)
                        {
                            DataGridSub.Items.Remove(DataGridSub.CurrentItem as Win_ord_Order_U_Sub_CodeView);
                        }
                        else
                        {
                            winordorderusubcodeview.Add(DataGridSub.SelectedItem as Win_ord_Order_U_Sub_CodeView);
                            DataGridSub.Items.Remove((DataGridSub.Items[DataGridSub.SelectedIndex]) as Win_ord_Order_U_Sub_CodeView);

                            //DataGridSub.Items.Remove((DataGridSub.Items[DataGridSub.Items.Count - 1]) as Win_ord_Order_U_Sub_CodeView);  //마지막 행만 삭제
                        }

                        DataGridSub.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("삭제할 데이터를 먼저 선택하세요.");
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SubRowDel : " + ee.ToString());
            }
        }

        #endregion

        #region 서브 그리드 수량 합계
        private void SumColorQty()
        {
            try
            {
                double ColorQty = 0;

                for (int i = 0; i < DataGridSub.Items.Count; i++)
                {
                    var label = DataGridSub.Items[i] as Win_ord_Order_U_Sub_CodeView;
                    if (label.ColorQty != null)
                    {
                        ColorQty += lib.returnDouble(label.ColorQty.ToString());
                    }
                }

                txtAmount.Text = lib.returnNumStringZero(ColorQty.ToString());

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumQty : " + ee.ToString());
            }
        }

        #endregion


        #region 신품수량합계
        private void SumNewArticleQty()
        {
            try
            {
                double NewArticleQty = 0; //신품
                double RePolishing = 0; //재연마

                for (int i = 0; i < DataGridSub.Items.Count; i++)
                {
                    var label = DataGridSub.Items[i] as Win_ord_Order_U_Sub_CodeView;

                    //신품 카운트 (신품Y고 수량이 Null 아닐 때)
                    if (label.ColorQty != null && label.NewProductYN.Equals("Y"))
                    {
                        NewArticleQty += lib.returnDouble(label.ColorQty.ToString());
                    }

                    //재연마 카운트 (신품 N 이고 수량이 Null 아닐 떄 )
                    if (label.ColorQty != null && label.NewProductYN.Equals("N"))
                    {
                        RePolishing += lib.returnDouble(label.ColorQty.ToString());
                    }

                }

                //txtNewArticle.Text = lib.returnNumStringZero(NewArticleQty.ToString());

                //txtRePolishing.Text = lib.returnNumStringZero(RePolishing.ToString());

            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - SumQty : " + ee.ToString());
            }
        }

        #endregion

        #region 서브 데이터그리드 방향키 이동 및 셀 포커스
        private void DataGridSub_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGridSub_KeyDown(sender, e);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_PreviewKeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                var SubItem = DataGridSub.CurrentItem as Win_ord_Order_U_Sub_CodeView;
                int rowCount = DataGridSub.Items.IndexOf(DataGridSub.CurrentItem);
                int colCount = DataGridSub.Columns.IndexOf(DataGridSub.CurrentCell.Column);
                int StartColumnCount = 1; //DataGridSub.Columns.IndexOf(dgdtpeMCoperationRateScore);
                int EndColumnCount = 5; //DataGridSub.Columns.IndexOf(dgdtpeComments);

                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount == colCount && DataGridSub.Items.Count - 1 > rowCount)
                    {
                        DataGridSub.SelectedIndex = rowCount + 1;
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount + 1], DataGridSub.Columns[StartColumnCount]);
                    }
                    else if (EndColumnCount > colCount && DataGridSub.Items.Count - 1 > rowCount)
                    {
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount], DataGridSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount && DataGridSub.Items.Count - 1 == rowCount)
                    {
                        btnSave.Focus();
                    }
                    else if (EndColumnCount > colCount && DataGridSub.Items.Count - 1 == rowCount)
                    {
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount], DataGridSub.Columns[colCount + 1]);
                    }
                    else
                    {
                        MessageBox.Show("있으면 찾아보자...");
                    }
                }
                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (DataGridSub.Items.Count - 1 > rowCount)
                    {
                        DataGridSub.SelectedIndex = rowCount + 1;
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount + 1], DataGridSub.Columns[colCount]);
                    }
                    else if (DataGridSub.Items.Count - 1 == rowCount)
                    {
                        if (EndColumnCount > colCount)
                        {
                            DataGridSub.SelectedIndex = 0;
                            DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[0], DataGridSub.Columns[colCount + 1]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (rowCount > 0)
                    {
                        DataGridSub.SelectedIndex = rowCount - 1;
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount - 1], DataGridSub.Columns[colCount]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (colCount > 0)
                    {
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount], DataGridSub.Columns[colCount - 1]);
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (EndColumnCount > colCount)
                    {
                        DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount], DataGridSub.Columns[colCount + 1]);
                    }
                    else if (EndColumnCount == colCount)
                    {
                        if (DataGridSub.Items.Count - 1 > rowCount)
                        {
                            DataGridSub.SelectedIndex = rowCount + 1;
                            DataGridSub.CurrentCell = new DataGridCellInfo(DataGridSub.Items[rowCount + 1], DataGridSub.Columns[StartColumnCount]);
                        }
                        else
                        {
                            btnSave.Focus();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_KeyDown " + ee.ToString());
            }
        }

        private void DataGridSub_TextFocus(object sender, KeyEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINControlFocus(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_TextFocus " + ee.ToString());
            }
        }

        private void DataGridSub_GotFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    DataGridCell cell = sender as DataGridCell;
                    cell.IsEditing = true;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_GotFocus " + ee.ToString());
            }
        }

        private void DataGridSub_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Lib.Instance.DataGridINBothByMouseUP(sender, e);
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridSub_MouseUp " + ee.ToString());
            }
        }
        #endregion

        private void ChekNewYN_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            Win_ord_Order_U_Sub_CodeView senderOCReqSub = chkSender.DataContext as Win_ord_Order_U_Sub_CodeView;
            senderOCReqSub.NewProductYN = "Y";

            senderOCReqSub.NewYNChecked = true;
        }

        private void ChekNewYN_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            Win_ord_Order_U_Sub_CodeView senderOCReqSub = chkSender.DataContext as Win_ord_Order_U_Sub_CodeView;
            senderOCReqSub.NewProductYN = "N";

            senderOCReqSub.NewYNChecked = false;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboboxSub = DataGridSub.CurrentItem as Win_ord_Order_U_Sub_CodeView;
            ComboBox cboNewYN = (ComboBox)sender;

            if (ComboboxSub == null)
            {
                ComboboxSub = DataGridSub.Items[rowSubNum] as Win_ord_Order_U_Sub_CodeView;
            }

            if (cboNewYN.SelectedValue != null && !cboNewYN.SelectedValue.ToString().Equals(""))
            {

                var theView = cboNewYN.SelectedItem as CodeView;
                if (theView != null)
                {
                    ComboboxSub.NewRecord = theView.code_name;
                    ComboboxSub.NewProductYN = theView.code_id;
                    //cboPriceClss.SelectedValue = 0;

                }
                sender = cboNewYN;
            }

        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            ComboBox cboNewYN = (ComboBox)sender;

            //매입,매출 화폐단위(입력)
            List<string[]> listNewYN = new List<string[]>();
            string[] New01 = new string[] { "Y", "Y" };
            string[] New02 = new string[] { "N", "N" };
            listNewYN.Add(New01);
            listNewYN.Add(New02);

            ObservableCollection<CodeView> ovcNewYN = ComboBoxUtil.Instance.Direct_SetComboBox(listNewYN);
            cboNewYN.ItemsSource = ovcNewYN;
            cboNewYN.DisplayMemberPath = "code_name";
            cboNewYN.SelectedValuePath = "code_id";
            cboNewYN.SelectedIndex = 0;

        }

        private void btnNeedStuff_Click(object sender, RoutedEventArgs e)
        {
            if (txtArticle.Tag == null)
            {
                MessageBox.Show("먼저 품명을 선택해주세요");
                return;
            }

            if (txtAmount.Text.Replace(" ", "").Equals(""))
            {
                MessageBox.Show("먼저 총 주문량을 입력해주세요");
                return;
            }

            //자재필요량조회에 필요한 파라미터 값을 넘겨주자, 품명이랑 주문량
            FillNeedStockQty(txtArticle.Tag.ToString(), txtAmount.Text.Replace(",", ""));
        }

        //메인 데이터그리드 선택 이벤트
        private void DataGridMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var OrderInfo = dgdMain.SelectedItem as Win_ord_Order_U_CodeView;

                if (OrderInfo != null)
                {
                    this.DataContext = OrderInfo;

                    String OrderID = OrderInfo.OrderID;
                    FillGridSub(OrderID);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - DataGridMain_SelectionChanged : " + ee.ToString());
            }
        }


        #region 조회Sub
        private void FillGridSub(string strOrderID)
        {
            if (DataGridSub.Items.Count > 0)
            {
                DataGridSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", strOrderID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ord_sOrderSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("Sub 데이터가 없습니다.");
                    }
                    else
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WOOUSC = new Win_ord_Order_U_Sub_CodeView
                            {
                                Num = i,

                                OrderID = dr["OrderID"].ToString(),
                                OrderSeq = dr["OrderSeq"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                ColorQty = dr["ColorQty"].ToString(),
                                UnitPriceClss = dr["UnitPriceClss"].ToString(),
                                NewProductYN = dr["NewProductYN"].ToString(),
                                Remark = dr["Remark"].ToString(),
                                Spec = dr["Spec"].ToString(),
                            };
                            
                            DataGridSub.Items.Add(WOOUSC);
                            WOOUSC.ColorQty = Lib.Instance.returnNumStringZero(WOOUSC.ColorQty);
                        }

                        
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - FillGridSub : " + ee.ToString());
            }
        }


        #endregion

        private void txtBuyerArticle_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {

                    if (txtCustom != null && txtCustom.Text != "")
                    {   //선택된 납품거래처에 따른 품명만 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, 57, txtCustom.Tag.ToString().Trim());

                        //품번을 품명처럼 쓴다고 해서 품번을 조회하도록 2020.03.17, 장가빈
                        MainWindow.pf.ReturnCodeGLS(txtBuyerArticleNO, 7070, txtCustom.Tag.ToString().Trim());

                    }
                    else
                    {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                        //MainWindow.pf.ReturnCode(txtArticle, (int)Defind_CodeFind.DCF_Article, "");

                        //품번을 품명처럼 쓴다고 해서 품번을 조회하도록 2020.03.17, 장가빈
                        MainWindow.pf.ReturnCodeGLS(txtBuyerArticleNO, 7071, "");
                    }


                    if (txtBuyerArticleNO.Tag != null)
                    {
                        CallArticleData(txtBuyerArticleNO.Tag.ToString());
                        //품명종류 대입(ex.제품 등)
                        cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                        //품번 대입
                        //txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                        //품명 대입
                        txtArticle.Text = articleData.Article;
                        //단가 대입
                        txtUnitPrice.Text = articleData.OutUnitPrice;

                    }

                    //플러스 파인더 작동 후 규격으로 커서 이동
                    txtSpec.Focus();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

        }

        private void btnPfBuyerArticle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (txtCustom != null && txtCustom.Text != "")
                {   //선택된 납품거래처에 따른 품명만 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtBuyerArticleNO, 7070, txtCustom.Tag.ToString().Trim());
                }
                else
                {   //선택된 납품거래처가 없다면 전체 품명 다 보여주게
                    MainWindow.pf.ReturnCodeGLS(txtBuyerArticleNO, 7071, "");
                }

                if (txtBuyerArticleNO.Tag != null)
                {
                    CallArticleData(txtBuyerArticleNO.Tag.ToString());
                    //품명종류 대입(ex.제품 등)
                    cboArticleGroup.SelectedValue = articleData.ArticleGrpID;
                    //품번 대입
                    //txtBuyerArticleNO.Text = articleData.BuyerArticleNo;
                    //품명 대입
                    txtArticle.Text = articleData.Article;
                    //단가 대입
                    txtUnitPrice.Text = articleData.OutUnitPrice;
                }

                //플러스 파인더 작동 후 규격으로 커서 이동
                txtSpec.Focus();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
    }







    public class Win_ord_Order_U_CodeView : BaseView
    {
        public string OrderID { get; set; }
        public string OrderNO { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string CloseClss { get; set; }

        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string Article { get; set; }
        public string ChunkRate { get; set; }
        public string PatternID { get; set; }

        public string Amount { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerArticleNo { get; set; }
        public string PONO { get; set; }

        public string OrderForm { get; set; }
        public string OrderClss { get; set; }
        public string InCustomID { get; set; }
        public string AcptDate { get; set; }
        public string DvlyDate { get; set; }

        public string ArticleID { get; set; }
        public string DvlyPlace { get; set; }
        public string WorkID { get; set; }
        public string PriceClss { get; set; }
        public string ExchRate { get; set; }

        public string Vat_IND_YN { get; set; }
        public string ColorCnt { get; set; }
        public string StuffWidth { get; set; }
        public string StuffWeight { get; set; }
        public string CutQty { get; set; }

        public string WorkWidth { get; set; }
        public string WorkWeight { get; set; }
        public string WorkDensity { get; set; }
        public string LossRate { get; set; }
        public string ReduceRate { get; set; }

        public string TagClss { get; set; }
        public string LabelID { get; set; }
        public string BandID { get; set; }
        public string EndClss { get; set; }
        public string MadeClss { get; set; }

        public string SurfaceClss { get; set; }
        public string ShipClss { get; set; }
        public string AdvnClss { get; set; }
        public string LotClss { get; set; }
        public string EndMark { get; set; }

        public string TagArticle { get; set; }
        public string TagArticle2 { get; set; }
        public string TagOrderNo { get; set; }
        public string TagRemark { get; set; }
        public string Tag { get; set; }

        public string BasisID { get; set; }
        public string BasisUnit { get; set; }
        public string SpendingClss { get; set; }
        public string DyeingID { get; set; }
        public string WorkingClss { get; set; }

        public string BTID { get; set; }
        public string BTIDSeq { get; set; }
        public string ChemClss { get; set; }
        public string AccountClss { get; set; }
        public string ModifyClss { get; set; }

        public string ModifyRemark { get; set; }
        public string CancelRemark { get; set; }
        public string Remark { get; set; }
        public string ActiveClss { get; set; }
        public string ModifyDate { get; set; }

        public string OrderFlag { get; set; }
        public string TagRemark2 { get; set; }
        public string TagRemark3 { get; set; }
        public string TagRemark4 { get; set; }
        public string UnitPriceClss { get; set; }

        public string WeightPerYard { get; set; }
        public string WorkUnitClss { get; set; }
        public string ArticleGrpID { get; set; }
        public string OrderSpec { get; set; }
        public string UnitPrice { get; set; }

        public string CompleteArticleFile { get; set; }
        public string CompleteArticlePath { get; set; }
        public string FirstArticleFile { get; set; }
        public string FirstArticlePath { get; set; }
        public string MediumArticleFIle { get; set; }

        public string MediumArticlePath { get; set; }
        public string sketch1Path { get; set; }
        public string sketch1file { get; set; }
        public string sketch2Path { get; set; }
        public string sketch2file { get; set; }

        public string sketch3Path { get; set; }
        public string sketch3file { get; set; }
        public string sketch4Path { get; set; }
        public string sketch4file { get; set; }
        public string sketch5Path { get; set; }

        public string sketch5file { get; set; }
        public string sketch6Path { get; set; }
        public string sketch6file { get; set; }
        public string ProductAutoInspectYN { get; set; }
        public string kBuyer { get; set; }

        public string BuyerID { get; set; }
        public int Num { get; set; }
        public string AcptDate_CV { get; set; }
        public string DvlyDate_CV { get; set; }
        public string Amount_CV { get; set; }

        public string KInCustom { get; set; }
        public string SketchFile { get; set; }
        public string SketchPath { get; set; }
        public string ImageName { get; set; }

        public string CompanyID { get; set; }
        public string OrderNo { get; set; }
        public string PoNo { get; set; }
        public string OrderFormName { get; set; }
        public string BrandClss { get; set; }
        public string WorkName { get; set; }
        public string OrderClssName { get; set; }

        public string NewArticleQty { get; set; }
        public string RePolishingQty { get; set; }


    }

    public class OrderArticle : BaseView
    {
        public string OrderID { get; set; }
        public string CustomID { get; set; }
        public string OrderNo { get; set; }
        public string KCustom { get; set; }
        public string PONO { get; set; }

        public string OrderForm { get; set; }
        public string OrderClss { get; set; }
        public string InCustomID { get; set; }
        public string AcptDate { get; set; }
        public string DvlyDate { get; set; }

        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string DvlyPlace { get; set; }
        public string WorkID { get; set; }
        public string PriceClss { get; set; }

        public string ExchRate { get; set; }
        public string Vat_IND_YN { get; set; }
        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string ColorCnt { get; set; }

        public string StuffWidth { get; set; }
        public string StuffWeight { get; set; }
        public string CutQty { get; set; }
        public string WorkWidth { get; set; }
        public string WorkWeight { get; set; }

        public string WorkDensity { get; set; }
        public string ChunkRate { get; set; }
        public string LossRate { get; set; }
        public string ReduceRate { get; set; }
        public string TagClss { get; set; }

        public string LabelID { get; set; }
        public string BandID { get; set; }
        public string EndClss { get; set; }
        public string MadeClss { get; set; }
        public string SurfaceClss { get; set; }

        public string ShipClss { get; set; }
        public string AdvnClss { get; set; }
        public string LotClss { get; set; }
        public string EndMark { get; set; }
        public string TagArticle { get; set; }

        public string TagArticle2 { get; set; }
        public string TagOrderNo { get; set; }
        public string TagRemark { get; set; }
        public string Tag { get; set; }
        public string BasisID { get; set; }

        public string BasisUnit { get; set; }
        public string SpendingClss { get; set; }
        public string DyeingID { get; set; }
        public string WorkingClss { get; set; }
        public string PatternID { get; set; }

        public string BTID { get; set; }
        public string BTIDSeq { get; set; }
        public string ChemClss { get; set; }
        public string AccountClss { get; set; }
        public string ModifyClss { get; set; }

        public string ModifyRemark { get; set; }
        public string CancelRemark { get; set; }
        public string Remark { get; set; }
        public string ActiveClss { get; set; }
        public string CloseClss { get; set; }

        public string ModifyDate { get; set; }
        public string OrderFlag { get; set; }
        public string TagRemark2 { get; set; }
        public string TagRemark3 { get; set; }
        public string TagRemark4 { get; set; }

        public string UnitPriceClss { get; set; }
        public string WeightPerYard { get; set; }
        public string WorkUnitClss { get; set; }
        public string ArticleGrpID { get; set; }
        public string OrderSpec { get; set; }

        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string BuyerArticleNo { get; set; }
        public string UnitPrice { get; set; }
        public string CompleteArticleFile { get; set; }

        public string CompleteArticlePath { get; set; }
        public string FirstArticleFile { get; set; }
        public string FirstArticlePath { get; set; }
        public string MediumArticleFIle { get; set; }
        public string MediumArticlePath { get; set; }

        public string sketch1Path { get; set; }
        public string sketch1file { get; set; }
        public string sketch2Path { get; set; }
        public string sketch2file { get; set; }
        public string sketch3Path { get; set; }

        public string sketch3file { get; set; }
        public string sketch4Path { get; set; }
        public string sketch4file { get; set; }
        public string sketch5Path { get; set; }
        public string sketch5file { get; set; }

        public string sketch6Path { get; set; }
        public string sketch6file { get; set; }
        public string ProductAutoInspectYN { get; set; }

        public string InCustom { get; set; }
        public string BuyerID { get; set; }
        public string kBuyer { get; set; }
    }

    public class ArticleData : BaseView
    {
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string ThreadID { get; set; }
        public string thread { get; set; }
        public string StuffWidth { get; set; }
        public string DyeingID { get; set; }
        public string Weight { get; set; }
        public string Spec { get; set; }
        public string ArticleGrpID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string UnitClss { get; set; }
        public string ProcessName { get; set; }
        public string HSCode { get; set; }
        public string OutUnitPrice { get; set; }
        public string Code_Name { get; set; }
    }

    public class CustomData : BaseView
    {
        public string LossClss { get; set; }
        public string SpendingClss { get; set; }
        public string WorkingClss { get; set; }
        public string CalClss { get; set; }
        public string PointClss { get; set; }
    }

    public class ArticleNeedStockQty : BaseView
    {
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string NeedQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
    }

    class Win_ord_Order_U_Sub_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string OrderID { get; set; }
        public string OrderSeq { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string Color { get; set; }
        public string DesignNo { get; set; }
        public string ColorQty { get; set; }
        public string UnitPrice { get; set; }
        public string UnitPriceClss { get; set; }
        public string OrderEnd { get; set; }
        public string PatternID { get; set; }
        public string Remark { get; set; }
        public string ProdQty { get; set; }
        public string ProdDate { get; set; }
        public string ReProdQty { get; set; }
        public string ReProdDate { get; set; }
        public string RGB { get; set; }
        public string Spec { get; set; }
        public string Weight { get; set; }

        public string NewProductYN { get; set; }
        public string NewRecord { get; set; }
        public bool NewYNChecked { get; set; }
    }
}

