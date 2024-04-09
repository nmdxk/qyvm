#define THREE_PHASE
//#define DBQ4
//#define VIRTUAL
using System;
using System.Collections.Generic;
using FirstFloor.ModernUI.Presentation;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Automation.BDaq;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot;
using System.Collections;
using System.Numerics;
using System.Diagnostics;
using System.Timers;
using HandyControl.Collections;
using System.Windows;
using PropertyChanged;
using System.Collections.ObjectModel;
using System.Windows.Documents;
using System.Windows.Threading;
using DatabaseManager;
using System.IO;
using System.Linq.Expressions;
using System.Data;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data.SQLite;
using System.Security.Cryptography;
using FirstFloor.ModernUI.App.qunying;
using System.Collections.Concurrent;
using Microsoft.SqlServer.Server;
using System.Runtime.InteropServices;
using HandyControl.Controls;
using MessageBox = System.Windows.MessageBox;
using static FirstFloor.ModernUI.App.qunying.Page7;
using FirstFloor.ModernUI.App.Content.ModernFrame;

namespace FirstFloor.ModernUI.App.ViewModel.qunying
{
    public class SpinLockExample
    {
        private int _lock = 0;

        public void Enter()
        {
            while (Interlocked.Exchange(ref _lock, 1) == 1)
            {
            }
        }

        public void Exit()
        {
            _lock = 0;
        }
    }
    class MultimediaTimer
    {
        // Windows Multimedia Timer API
        [DllImport("winmm.dll")]
        public static extern uint timeSetEvent(uint uDelay, uint uResolution, TimerCallback lpTimeProc, UIntPtr dwUser, uint fuEvent);

        [DllImport("winmm.dll")]
        public static extern uint timeKillEvent(uint uTimerID);

        // TimerCallback delegate
        public delegate void TimerCallback(uint uTimerID, uint uMsg, UIntPtr dwUser, UIntPtr dw1, UIntPtr dw2);

        // Timer ID
        private uint timerID;

        // Start the multimedia timer
        public void Start(uint delay, uint resolution, TimerCallback callback)
        {
            // Set the timer event
            timerID = timeSetEvent(delay, resolution, callback, UIntPtr.Zero, 1 /* TIME_PERIODIC */);

            if (timerID == 0)
            {
                throw new Exception("Failed to start multimedia timer.");
            }
        }

        // Stop the multimedia timer
        public void Stop()
        {
            if (timerID != 0)
            {
                // Kill the timer event
                timeKillEvent(timerID);
                timerID = 0;
            }
        }
    }
    public enum StatusModel
    {
        dynamicc = 1,
        staticc,
        converc,
        close
    }
    [AddINotifyPropertyChangedInterface]
    public class Struce//工位自检
    {
        public string fault_index { get; set; } = string.Empty;
        public string energize_open_ { get; set; }//动合断
        public string release_stickiness_ { get; set; }//静合粘
        public string energize_stickiness_ { get; set; }//动合粘
        public string release_open_ { get; set; }//静合断
        public string open_ { get; set; }//静合断
        public string stickiness_ { get; set; }//静合断
        public string phase_ { get; set; }//静合断
    }
    [AddINotifyPropertyChangedInterface]
    public class CoilTestStruc//工位自检
    {
        public int channel_order_ { get; set; }
        public string sample_code_ { get; set; } = string.Empty;
        public double self_retain_value_ { get; set; }
        public double come_back_value_ { get; set; }
    }
    [AddINotifyPropertyChangedInterface]
    public class WpaTes
    {
        public string channel_ { get; set; } = string.Empty;
        public double energize_open_value_ { get; set; }//动合断
        public double release_stickiness_value_ { get; set; }//静合粘
        public double energize_stickiness_value_ { get; set; }//动合粘
        public double release_open_value_ { get; set; }//静合断
    }
    [AddINotifyPropertyChangedInterface]
    public static class BasicSetting
    {
        public static int delay_time_ { get; set; } = 3000;
        public static int conver_time_ { get; set; } = 2000;
    }
    [AddINotifyPropertyChangedInterface]
    public class DataBaseTemple//数据存储中转站 每次解析完赋值 然后存储
    {
        public string typedfswitch { get; set; } = "";
        public string time { get; set; } = "";
        public string falutname { get; set; } = "";
        public bool status { get; set; } = false;
        public double buckvalue { get; set; } = -1;
        public double voltagevalue { get; set; } = -1;
        public double currentvalue { get; set; } = -1;
    }
    [AddINotifyPropertyChangedInterface]
    public class QyViewModel
    {
        //    ulong times = 0;
        //    ulong time = 0;
        //public static AutoResetEvent ResetEventClearPlotData = new AutoResetEvent(true);
        //  static SpinLockExample spinLock = new SpinLockExample();
        public static bool clearn_plotdata_ = true;
        #region 解析数据
        public static string[] name_array = new string[33] { "Ua", "Ub", "Uc", "IaInstant", "IbInstant", "IcInstant", "IaRate", "IbRate", "IcRate", "DC1H", "DC2H", "DC3H", "DC4H", "DC5H", "DC6H", "DC7H", "DC8H", "DC1D", "DC2D", "DC3D", "DC4D", "DC5D", "DC6D", "DC7D", "DC8D", "Temp1", "Temp3", "Temp5", "Temp7", "Temp2", "Temp4", "Temp6", "Temp8" };

        public static string[] name_array_aux = new string[] { "DC1H", "DC2H", "DC3H", "DC4H", "DC5H", "DC6H", "DC7H", "DC8H", "DC1D", "DC2D", "DC3D", "DC4D", "DC5D", "DC6D", "DC7D", "DC8D" };
        static string[] name_array_input = new string[] { "Ua", "Ub", "Uc", "IaInstant", "IbInstant", "IcInstant", "IaRate", "IbRate", "IcRate" };
        public static string[] name_array_all = new string[] { "Ua", "Ub", "Uc", "DC1H", "DC2H", "DC3H", "DC4H", "DC5H", "DC6H", "DC7H", "DC8H", "DC1D", "DC2D", "DC3D", "DC4D", "DC5D", "DC6D", "DC7D", "DC8D" };
        static Dictionary<string, int> auxmap_statu_name_ = new Dictionary<string, int>() { { "DC1H", 1 }, { "DC2H", 2 }, { "DC3H", 3 }, { "DC4H", 4 }, { "DC5H", 5 }, { "DC6H", 6 }, { "DC7H", 7 }, { "DC8H", 8 }, { "DC1D", 1 }, { "DC2D", 2 }, { "DC3D", 3 }, { "DC4D", 4 }, { "DC5D", 5 }, { "DC6D", 6 }, { "DC7D", 7 }, { "DC8D", 8 } };



        static Dictionary<string, ObservableCollection<Struce>> keyValuePairs = new Dictionary<string, ObservableCollection<Struce>>();
        static Dictionary<string, Queue<double>> fault_save_data = new Dictionary<string, Queue<double>>() {
                {"DC1H",new Queue<double>()},{"DC4H",new Queue<double>() },{"DC8H",new Queue<double>()},
                {"DC1D",new Queue<double>()},{"DC4D",new Queue<double>() },{"DC8D",new Queue<double>()},
                {"DC2H",new Queue<double>()},{"DC5H",new Queue<double>() },{"DC7H",new Queue<double>()},
                {"DC2D",new Queue<double>()},{"DC5D",new Queue<double>() },{"DC7D",new Queue<double>()},
                {"DC3H",new Queue<double>()},{"DC6D",new Queue<double>() },{"DC6H",new Queue<double>()},
                {"DC3D",new Queue<double>()},{"Ua",new Queue<double>() },{"Ub",new Queue<double>()},
                {"Uc",new Queue<double>()},
            };
        static Dictionary<string, int> fault_count = new Dictionary<string, int>() {
                {"DC1HH断故障",0},{"DC4HH断故障",0},{"DC8HH断故障",0},
                {"DC1DD断故障",0},{"DC4DD断故障",0},{"DC8DD断故障",0},
                {"DC2HH断故障",0},{"DC5HH断故障",0},{"DC7HH断故障",0},
                {"DC2DD断故障",0},{"DC5DD断故障",0},{"DC7DD断故障",0},
                {"DC3HH断故障",0},{"DC6DD断故障",0},{"DC6HH断故障",0},
                {"DC3DD断故障",0},
                {"DC1HH粘故障",0},{"DC4HH粘故障",0},{"DC8HH粘故障",0},
                {"DC1DD粘故障",0},{"DC4DD粘故障",0},{"DC8DD粘故障",0},
                {"DC2HH粘故障",0},{"DC5HH粘故障",0},{"DC7HH粘故障",0},
                {"DC2DD粘故障",0},{"DC5DD粘故障",0},{"DC7DD粘故障",0},
                {"DC3HH粘故障",0},{"DC6DD粘故障",0},{"DC6HH粘故障",0},
                {"DC3DD粘故障",0},
                {"Ua断故障",0}, {"Ua粘故障",0},
                {"Ub断故障",0}, {"Ub粘故障",0},
                {"Uc断故障",0}, {"Uc粘故障",0},
            };
        //
        static Dictionary<string, int> fault_count_save = new Dictionary<string, int>() {
                {"DC1HH断故障",0},{"DC4HH断故障",0},{"DC8HH断故障",0},
                {"DC1DD断故障",0},{"DC4DD断故障",0},{"DC8DD断故障",0},
                {"DC2HH断故障",0},{"DC5HH断故障",0},{"DC7HH断故障",0},
                {"DC2DD断故障",0},{"DC5DD断故障",0},{"DC7DD断故障",0},
                {"DC3HH断故障",0},{"DC6DD断故障",0},{"DC6HH断故障",0},
                {"DC3DD断故障",0},
                {"DC1HH粘故障",0},{"DC4HH粘故障",0},{"DC8HH粘故障",0},
                {"DC1DD粘故障",0},{"DC4DD粘故障",0},{"DC8DD粘故障",0},
                {"DC2HH粘故障",0},{"DC5HH粘故障",0},{"DC7HH粘故障",0},
                {"DC2DD粘故障",0},{"DC5DD粘故障",0},{"DC7DD粘故障",0},
                {"DC3HH粘故障",0},{"DC6DD粘故障",0},{"DC6HH粘故障",0},
                {"DC3DD粘故障",0},
                {"Ua断故障",0}, {"Ua粘故障",0},
                {"Ub断故障",0}, {"Ub粘故障",0},
                {"Uc断故障",0}, {"Uc粘故障",0},
            };

        static Dictionary<string, Deque<double>> fault_data = new Dictionary<string, Deque<double>>() {
                {"DC1HH断故障",new Deque<double>()},{"DC4HH断故障",new Deque<double>()},{"DC8HH断故障",new Deque<double>()},
                {"DC1DD断故障",new Deque<double>()},{"DC4DD断故障",new Deque<double>()},{"DC8DD断故障",new Deque<double>()},
                {"DC2HH断故障",new Deque<double>()},{"DC5HH断故障",new Deque<double>()},{"DC7HH断故障",new Deque<double>()},
                {"DC2DD断故障",new Deque<double>()},{"DC5DD断故障",new Deque<double>()},{"DC7DD断故障",new Deque<double>()},
                {"DC3HH断故障",new Deque<double>()},{"DC6DD断故障",new Deque<double>()},{"DC6HH断故障",new Deque<double>()},
                {"DC3DD粘故障",new Deque<double>()},
                {"DC3DD断故障",new Deque<double>()},
                {"DC1HH粘故障",new Deque<double>()},{"DC4HH粘故障",new Deque<double>()},{"DC8HH粘故障",new Deque<double>()},
                {"DC1DD粘故障",new Deque<double>()},{"DC4DD粘故障",new Deque<double>()},{"DC8DD粘故障",new Deque<double>()},
                {"DC2HH粘故障",new Deque<double>()},{"DC5HH粘故障",new Deque<double>()},{"DC7HH粘故障",new Deque<double>()},
                {"DC2DD粘故障",new Deque<double>()},{"DC5DD粘故障",new Deque<double>()},{"DC7DD粘故障",new Deque<double>()},
                {"DC3HH粘故障",new Deque<double>()},{"DC6DD粘故障",new Deque<double>()},{"DC6HH粘故障",new Deque<double>()},

                {"Ua断故障",new Deque<double>()}, {"Ua粘故障",new Deque<double>()},{"Ub断故障",new Deque<double>()}, {"Ub粘故障",new Deque<double>()},{"Uc断故障",new Deque<double>()}, {"Uc粘故障",new Deque<double>()},
            };
        #endregion
        public int xoy_refresh = 200;
        public string xoyplot_refresh
        {
            get { return xoy_refresh.ToString(); }
            set
            {
                xoy_refresh = Convert.ToInt16(value);
                //XoyPlotMd.readDataTimer.Stop();
                //XoyPlotMd.readDataTimer.Start();
            }
        }


        public PlotModel Model { get; set; }


        #region COMBOBOX数据
        private static IEnumerable<string> data_test_condition_ { get; set; } = new List<string>(){
            "高温","低温"};
        private static IEnumerable<string> data_load_property_ { get; set; } = new List<string>(){
            "阻性负载","过负载","极限通断能力","模拟电动机负载","阶梯负载","限时电流继电特性"};
        private static IEnumerable<string> load_wave_info_ { get; set; } = new List<string>()
        {
        };
        private static IEnumerable<string> data_test_project_ { get; set; } = new List<string>(){
            "项目1","项目2"};
        private static IEnumerable<string> data_relay_type_ { get; set; } = new List<string>(){
            "普通型","极化型","磁保持型","交流型"};
        private static IEnumerable<string> data_On_Off_ratio_ { get; set; } = new List<string>(){
            "10%","20%","30%","40%",
            "50%","60%","70%","80%","90%", };

        private static IEnumerable<string> data_thresholds_stop_ { get; set; } = new List<string>(){
            "停机","不停机",};
        private static IEnumerable<string> data_fault_stop_ { get; set; } = new List<string>(){
            "停机","不停机"};
        private static IEnumerable<string> data_load_voltage_monitoring_ { get; set; } = new List<string>(){
            "是","否"};
        private static IEnumerable<string> data_excitation_mode_ { get; set; } = new List<string>(){
            "激励方式1","激励方式2"};
        #endregion
        #region 基础设置界面参数
        static public string load_vol_ { get; set; } = "100";
        public string load_cur_ { get; set; } = "100";
        public string load_sour_vol_ { get; set; } = "24";
        public string load_sour_cur_ { get; set; } = "1";
        public int aux_count_ { get; set; } = 8;
        #region 文件导入参数
        public string designator_ { get; set; } = string.Empty;
        public string batch_number_ { get; set; } = string.Empty;
        public string run_times_ { get; set; } = string.Empty;
        #endregion
        #region DO输出时序控制变量
        static BitArray bitArray_do { get; set; } = new BitArray(16);
        static BitArray bitArray_do_2 { get; set; } = new BitArray(16);
        volatile static bool connect_statu_ = false;

        volatile static private bool data_is_efficent = false;
        volatile static private bool data_is_efficent_his = false;
        static Byte[] ByteArray = new byte[2];
        static Byte[] ByteArray_2 = new byte[2];
        static byte index_temp = 0; //索引 温度切换
        public bool start_capture { get; set; } = false; //切换
        #endregion
        #region 采集卡配置参数

#if VIRTUAL
        public string deviceDescription_input = "DemoDevice,BID#0";//采集卡标识
        public string deviceDescription_aux = "DemoDevice,BID#1";//采集卡标识
        public const int s_channel_count_ = 8;//采集数据的通道数
#else
        public const int s_channel_count_ = 16;//采集数据的通道数
        public string deviceDescription_input = "";// "PCI-1716L,BID#0";//采集卡标识DemoDevice,BID#0
        public string deviceDescription_aux = "";//"PCI-1716L,BID#8";//采集卡标识
#endif
        public static int s_start_channel_ = 0;//采集起始通道
        public static int s_section_count_ = 0;//采集次数 0为无限,否则采集buff总数达到次数*长度 则停止采集
        public static int s_section_length_ { get; set; }//数据长度
        public static int s_section_length_pr { get; set; }//数据长度
        public static int count_input = 0;//数据长度
        public static double[] s_pci_buffer_input;//接收数据的缓存 来自采集卡
        public static double[] s_pci_buffer_aux;//接收数据的缓存 来自采集卡
        public static double s_convert_clk_rate_ { get; set; } = 4000;//1000.0;//始终转换频率 Hz
        public static int bufSize = 0;
        ErrorCode errorCode = ErrorCode.Success;
        ErrorCode errorCode2 = ErrorCode.Success;
        #endregion
        #region 常规参数
        public int time_save_temperat = 30;
        public string test_condition { get; set; } = string.Empty;
        public string test_project { get; set; } = string.Empty;
        public string test_project_ { get; set; } = string.Empty;
        public string test_batch_ { get; set; } = string.Empty;
        public string test_rly_type_ { get; set; } = string.Empty;
        public string test_dates_ { get; set; } = string.Empty;
        public string test_staff_ { get; set; } = string.Empty;
        public string load_property { get; set; } = string.Empty;
        public string file_path_ { get; set; } = string.Empty;
        public string specification_ { get; set; } = string.Empty;
        public string audits_ { get; set; } = string.Empty;
        public ObservableCollection<string> item_test_project_ { get; set; } = new ObservableCollection<string>(data_test_project_);
        public ObservableCollection<string> test_condition_ { get; set; } = new ObservableCollection<string>(data_test_condition_);
        public static ObservableCollection<string> load_property_ { get; set; } = new ObservableCollection<string>(data_load_property_);
        #endregion
        #region 运行参数
        static private int data_sensitivity_percent_stick_ { get; set; } = 80;
        static private int data_sensitivity_percent_sever_ { get; set; } = 20;

        private double data_monit_time_ = 50;
        public string fault_stop { get; set; } = string.Empty;
        public double data_monit_time
        {
            get { return data_monit_time_; }
            set
            {
                if (value > 50) data_monit_time_ = 50;
                else if (value < 10) data_monit_time_ = 10;
                else data_monit_time_ = value;

            }
        }
        public int data_sensitivity_percent_stick
        {
            get { return data_sensitivity_percent_stick_; }
            set
            {
                if (value < 80) data_sensitivity_percent_stick_ = 80;
                else if (value > 95) data_sensitivity_percent_stick_ = 95;
                else
                {
                    data_sensitivity_percent_stick_ = value;
                }
            }
        }
        public int data_sensitivity_percent_sever
        {
            get { return data_sensitivity_percent_sever_; }
            set
            {
                if (value < 5) data_sensitivity_percent_sever_ = 5;
                else if (value > 20) data_sensitivity_percent_sever_ = 20;
                else
                {
                    data_sensitivity_percent_sever_ = value;
                }
            }
        }
        public string thresholds_stop { get; set; } = string.Empty;
        public string sensitivity_percent { get; set; } = string.Empty;
        public string monit_time { get; set; } = string.Empty;
        //  public string On_Off_ratio { get; set; } = "50%";
        public string relay_type { get; set; } = string.Empty;
        public string load_voltage_monitoring { get; set; } = string.Empty;
        public int formulation_times_ { get; set; } = 1000000;
        // public int current_run_times { get; set; } = 0;.
        static private Int32 on_off_cycle = 0;
        public double working_frequency_ { get; set; } = 60;
        public double sensitivity_ { get; set; } = 0.3;
        static public int fault_upper_limit_ { get; set; } = 99;
        public string cf_connect_ { get; set; } = "100";//通断比左参数
        public string cf_fault_upper_limit_ { get; set; } = "100";//通断比右参数
        public ObservableCollection<string> relay_type_ { get; set; } = new ObservableCollection<string>(data_relay_type_);
        public ObservableCollection<string> fault_stop_ { get; set; } = new ObservableCollection<string>(data_fault_stop_);
        public ObservableCollection<string> thresholds_stop_ { get; set; } = new ObservableCollection<string>(data_thresholds_stop_);
        //  public ObservableCollection<string> on_off_ratio_ { get; set; } = new ObservableCollection<string>(data_On_Off_ratio_);
        public ObservableCollection<string> load_voltage_monitoring_ { get; set; } = new ObservableCollection<string>(data_load_voltage_monitoring_);
        private readonly object _lock = new object();
        // public static readonly object _lock_recive_data = new object();
        #endregion
        #endregion
        #region 线圈测试
        public static int coil_excitation_current_ { get; set; } = 0;
        public static int coil_rated_voltage_ { get; set; } = 0;
        public static int coil_released_voltage_ { get; set; } = 0;
        public static int coil_keep_recovery_voltage_ { get; set; } = 0;
        public static string coil_relay_type { get; set; } = string.Empty;
        public static ObservableCollection<string> coil_relay_type_ { get; set; } = new ObservableCollection<string>(data_relay_type_);
        #endregion
        #region 加电保温
        public string thermos_output_voltage_ { get; set; } = string.Empty;
        public string thermos_run_time_ { get; set; } = "保温运行时间：";
        public string thermos_output_current_ { get; set; } = string.Empty;
        public int thermos_rated_voltage_ { get; set; } = 0;
        public int thermos_rated_current_ { get; set; } = 0;
        public int thermos_time_ { get; set; } = 0;
        public ObservableCollection<string> thermos_relay_type_ { get; set; } = new ObservableCollection<string>(data_relay_type_);
        public ObservableCollection<string> thermos_excitation_mode_ { get; set; } = new ObservableCollection<string>(data_excitation_mode_);
        #endregion
        #region 数据处理&逻辑判断成员变量
        //临时存储触点数据
        // public static Dictionary<int, DataBaseTemple> dic_data_base_temple_ = new Dictionary<int, DataBaseTemple>();
        private static System.Timers.Timer aTimer_ChangeTempData_;//定时切换
        private static System.Timers.Timer aTimer_SaveData_;//存储运行时数据
        private static System.Timers.Timer aTimer_SaveData_Temp_;//存储温度

        //private static System.Timers.Timer aTimer_t;//定时转换读取温度
        //试验测试
        List<string> datagrid_index_ { get; set; } = new List<string>();
        //接触器状态记录
        public static Dictionary<int, StatusModel> s_state_contactor_dic_ = new Dictionary<int, StatusModel>() {
            { 1, StatusModel.close }, { 2, StatusModel.close },
            { 8, StatusModel.close }, { 3, StatusModel.close },
            { 7, StatusModel.close }, { 4, StatusModel.close },
            { 6, StatusModel.close }, { 5, StatusModel.close } };

        //用来区分采集数据的周期
        private static double his_data_last = 0; //记录一帧结束后最后一个点位，用来与下一个周期第一个点判断 如果"-"则上一周期完整,如果"+"用                                               his_data_组合；
        public ObservableCollection<WpaTes> temp_Wpa { set; get; } = new ObservableCollection<WpaTes>();
        Dictionary<int, double[]> dic_value_wpa_data_ = new Dictionary<int, double[]>() {
                {0, new double[4]{0.0,0.0,0.0,0.0 } }, { 1, new double[4]{0.0,0.0,0.0,0.0 } },
                {2, new double[4]{0.0,0.0,0.0,0.0 } }, { 3, new double[4]{0.0,0.0,0.0,0.0 } },
                {4, new double[4]{0.0,0.0,0.0,0.0 } }, { 5, new double[4]{0.0,0.0,0.0,0.0 } },
                {6, new double[4]{0.0,0.0,0.0,0.0 } }, { 7, new double[4]{0.0,0.0,0.0,0.0 } }, };
        public static string table_name = "";//数据表后缀
#if true
        ConcurrentQueue<Tuple<string, int, string, List<List<double>>>> data_plots = new ConcurrentQueue<Tuple<string, int, string, List<List<double>>>>();
        ConcurrentQueue<Tuple<double, string, string, bool, double, double, string>> data_plot = new ConcurrentQueue<Tuple<double, string, string, bool, double, double, string>>();

        static public ConcurrentDictionary<string, ConcurrentQueue<Tuple<double, string, bool, Int32>>> s_pci_data_ { get; set; } = new ConcurrentDictionary<string, ConcurrentQueue<Tuple<double, string, bool, Int32>>>();
        static public ConcurrentQueue<Tuple<double[], string, bool, Int32>> pci_aux = new ConcurrentQueue<Tuple<double[], string, bool, int>>();
        static public ConcurrentQueue<Tuple<double[], string, bool, Int32>> pci_input = new ConcurrentQueue<Tuple<double[], string, bool, int>>();
        static public ConcurrentQueue<double[]> pci_aux_plot = new ConcurrentQueue<double[]>();
        static public ConcurrentQueue<double[]> pci_input_plot = new ConcurrentQueue<double[]>();
#endif
        public static ConcurrentDictionary<string, ConcurrentQueue<double>> xoyplot_data_ { get; set; } = new ConcurrentDictionary<string, ConcurrentQueue<double>>();
        static public ConcurrentDictionary<string, Tuple<double, bool, string>> s_pci_data_last_ { get; set; } = new ConcurrentDictionary<string, Tuple<double, bool, string>>();

        static Dictionary<string, int> close_channel = new Dictionary<string, int>() {
            {"DC1H",0 },{"DC2H",1 },{"DC3H",2 },{"DC4H",3 },{"DC5H",4 },{"DC6H",5 },{"DC7H",6 },{"DC8H",7 },
            {"DC1D",8 },{"DC2D",9 },{"DC3D",10 },{"DC4D",11 },{"DC5D",12 },{"DC6D",13 },{"DC7D",14 },{"DC8D",15 },
        };

#if VIRTUAL
#if THREE_PHASE
        static Dictionary<int, string> dic_pci_input = new Dictionary<int, string>() { { 0, "Ua" }, { 1, "Ub" }, { 2, "Uc" }, { 3, "IaRate" }, { 4, "IbRate" }, { 5, "IcRate" }, { 6, "IaInstant" }, { 7, "IbInstant" }, { 8, "IcInstant" }, { 9, "" }, { 10, "" }, { 11, "" }, { 12, "" }, { 13, "" }, { 14, "" }, { 15, "" } };
        static Dictionary<int, string> dic_pci_aux = new Dictionary<int, string>() { { 0, "DC1H" }, { 1, "DC1D" }, { 2, "DC2H" }, { 3, "DC2D" }, { 4, "DC3H" }, { 5, "DC3D" }, { 6, "DC4H" }, { 7, "DC4D" }, { 8, "DC5H" }, { 9, "DC5D" }, { 10, "DC6H" }, { 11, "DC6D" }, { 12, "DC7H" }, { 13, "DC7D" }, { 14, "DC8H" }, { 15, "DC8D" } };
#else
        static Dictionary<int, string> dic_pci_input = new Dictionary<int, string>() { { 0, "Ua" }, { 1, "" }, { 2, "" }, { 3, "IaRate" }, { 4, "IaInstant" }, { 5, "" }, { 6, "" }, { 7, "" }, { 8, "" }, { 9, "" }, { 10, "" }, { 11, "" }, { 12, "" }, { 13, "" }, { 14, "" }, { 15, "" } };
        static Dictionary<int, string> dic_pci_aux = new Dictionary<int, string>() { { 0, "DC1H" }, { 1, "DC1D" }, { 2, "DC2H" }, { 3, "DC2D" }, { 4, "DC3H" }, { 5, "DC3D" }, { 6, "DC4H" }, { 7, "DC4D" }, { 8, "DC5H" }, { 9, "DC5D" }, { 10, "DC6H" }, { 11, "DC6D" }, { 12, "DC7H" }, { 13, "DC7D" }, { 14, "DC8H" }, { 15, "DC8D" } };
#endif
#else
#if THREE_PHASE
        static Dictionary<int, string> dic_pci_input = new Dictionary<int, string>() { { 0, "Ua" }, { 1, "Ub" }, { 2, "Uc" }, { 3, "IaRate" }, { 4, "IbRate" }, { 5, "IcRate" }, { 6, "IaInstant" }, { 7, "IbInstant" }, { 8, "IcInstant" }, { 9, "" }, { 10, "" }, { 11, "" }, { 12, "" }, { 13, "" }, { 14, "" }, { 15, "" } };
        static Dictionary<int, string> dic_pci_aux = new Dictionary<int, string>() { { 0, "DC1H" }, { 1, "DC1D" }, { 2, "DC2H" }, { 3, "DC2D" }, { 4, "DC3H" }, { 5, "DC3D" }, { 6, "DC4H" }, { 7, "DC4D" }, { 8, "DC5H" }, { 9, "DC5D" }, { 10, "DC6H" }, { 11, "DC6D" }, { 12, "DC7H" }, { 13, "DC7D" }, { 14, "DC8H" }, { 15, "DC8D" } };
#else
        static Dictionary<int, string> dic_pci_input = new Dictionary<int, string>() { { 0, "Ua" }, { 1, "" }, { 2, "" }, { 3, "IaRate" }, { 4, "IaInstant" }, { 5, "" }, { 6, "" }, { 7, "" }, { 8, "" }, { 9, "" }, { 10, "" }, { 11, "" }, { 12, "" }, { 13, "" }, { 14, "" }, { 15, "" } };
        static Dictionary<int, string> dic_pci_aux = new Dictionary<int, string>() { { 0, "DC1H" }, { 1, "DC1D" }, { 2, "DC2H" }, { 3, "DC2D" }, { 4, "DC3H" }, { 5, "DC3D" }, { 6, "DC4H" }, { 7, "DC4D" }, { 8, "DC5H" }, { 9, "DC5D" }, { 10, "DC6H" }, { 11, "DC6D" }, { 12, "DC7H" }, { 13, "DC7D" }, { 14, "DC8H" }, { 15, "DC8D" } };
#endif
#endif
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        static private string SearchDicKey_PCI1()
        {
            return string.Format("Temp{0}", index_temp + 1);
        }
        #endregion
        //static private object lck = new object();

#if !DBQ4
        static InstantAiCtrl instantAIContrl_input = new InstantAiCtrl();//DO实例
        static InstantAiCtrl instantAIContrl_aux = new InstantAiCtrl();//DO实例
        static InstantDoCtrl instanDoCtrl_input = new InstantDoCtrl();//DO实例
        static InstantDoCtrl instanDoCtrl_aux = new InstantDoCtrl();//DO实例
        static WaveformAiCtrl waveformAiCtrl_input = new WaveformAiCtrl();//AI stream 实例
        static WaveformAiCtrl waveformAiCtrl_aux = new WaveformAiCtrl();//AI2 stream 实例
#endif
        static double[] get_data_spci_instant = new double[16];
        #region 测试数据
        private double input_falut { get; set; } = 220.0;
        public string Input_falut { get { return input_falut.ToString(); } set { input_falut = Convert.ToDouble(value); } }
        private double aux_falut { get; set; } = 26.0;
        public string Aux_falut { get { return aux_falut.ToString(); } set { aux_falut = Convert.ToDouble(value); } }
        static public int hz_power { get; set; } = 50;// 800;
        public string Hz_power { get { return hz_power.ToString(); } set { hz_power = Convert.ToInt32(value); } }

        private double U_value { get; set; } = 0;
        private double I_value { get; set; } = 0;
        public string U_VALUE { get { return U_value.ToString(); } set { U_value = Convert.ToDouble(value); } }
        public string I_VALUE { get { return I_value.ToString(); } set { I_value = Convert.ToDouble(value); } }
        Dictionary<string, Int32> calc_falut_count_ = new Dictionary<string, Int32>() {
             {"DC1HH断故障",-1},{"DC4HH断故障",-1},{"DC8HH断故障",-1},
                {"DC1DD断故障",-1},{"DC4DD断故障",-1},{"DC8DD断故障",-1},
                {"DC2HH断故障",-1},{"DC5HH断故障",-1},{"DC7HH断故障",-1},
                {"DC2DD断故障",-1},{"DC5DD断故障",-1},{"DC7DD断故障",-1},
                {"DC3HH断故障",-1},{"DC6DD断故障",-1},{"DC6HH断故障",-1},
                {"DC3DD断故障",-1},
                {"DC1HH粘故障",-1},{"DC4HH粘故障",-1},{"DC8HH粘故障",-1},
                {"DC1DD粘故障",-1},{"DC4DD粘故障",-1},{"DC8DD粘故障",-1},
                {"DC2HH粘故障",-1},{"DC5HH粘故障",-1},{"DC7HH粘故障",-1},
                {"DC2DD粘故障",-1},{"DC5DD粘故障",-1},{"DC7DD粘故障",-1},
                {"DC3HH粘故障",-1},{"DC6DD粘故障",-1},{"DC6HH粘故障",-1},
                {"DC3DD粘故障",-1},
                {"Ua断故障",-1}, {"Ua粘故障",-1},
                {"Ub断故障",-1}, {"Ub粘故障",-1},
                {"Uc断故障",-1}, {"Uc粘故障",-1},
        };
        #endregion

        private static int[] index_result_data = new int[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        public QyViewModel()
        {
            using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + "\\config\\PCI-info.txt"))
            {
                string empty = streamReader.ReadLine();
                empty = streamReader.ReadLine();
                deviceDescription_input = streamReader.ReadLine();
                empty = streamReader.ReadLine();
                deviceDescription_aux = streamReader.ReadLine();
                empty = streamReader.ReadLine();
                s_convert_clk_rate_ = Convert.ToDouble(streamReader.ReadLine());
            }
#if !DBQ4
            instantAIContrl_input.SelectedDevice = new DeviceInformation(deviceDescription_input);
            instantAIContrl_aux.SelectedDevice = new DeviceInformation(deviceDescription_aux);
            instanDoCtrl_input.SelectedDevice = new DeviceInformation(deviceDescription_input);
            instanDoCtrl_aux.SelectedDevice = new DeviceInformation(deviceDescription_aux);
#endif
#if THREE_PHASE
            s_section_length_ = Convert.ToInt16(Math.Ceiling(s_convert_clk_rate_ / hz_power));
            s_section_length_ = Convert.ToInt16(s_convert_clk_rate_);// 1000
#else
            s_section_length_ = 6;
#endif
            bufSize = s_channel_count_ * (s_section_length_);
            s_pci_buffer_input = new double[bufSize];
            s_pci_buffer_aux = new double[bufSize];

            DBManager.dbManager.StartDBRoutine();


            test_dates_ = DateTime.Now.ToString("yy-MM-dd");
            foreach (string str in name_array)
            {
                s_pci_data_last_[str] = Tuple.Create(new double(), false, "");
                xoyplot_data_[str] = new ConcurrentQueue<double>();
                s_pci_data_[str] = new ConcurrentQueue<Tuple<double, string, bool, Int32>>();
            }
            keyValuePairs["Ua"] = inputa_falut_data;
            keyValuePairs["Ub"] = inputb_falut_data;
            keyValuePairs["Uc"] = inputc_falut_data;
            keyValuePairs["DC1H"] = aux1_falut_data;
            keyValuePairs["DC1D"] = aux1_falut_data;
            keyValuePairs["DC2H"] = aux2_falut_data;
            keyValuePairs["DC2D"] = aux2_falut_data;
            keyValuePairs["DC3H"] = aux3_falut_data;
            keyValuePairs["DC3D"] = aux3_falut_data;
            keyValuePairs["DC4H"] = aux4_falut_data;
            keyValuePairs["DC4D"] = aux4_falut_data;
            keyValuePairs["DC5H"] = aux5_falut_data;
            keyValuePairs["DC5D"] = aux5_falut_data;
            keyValuePairs["DC6H"] = aux6_falut_data;
            keyValuePairs["DC6D"] = aux6_falut_data;
            keyValuePairs["DC7H"] = aux7_falut_data;
            keyValuePairs["DC7D"] = aux7_falut_data;
            keyValuePairs["DC8H"] = aux8_falut_data;
            keyValuePairs["DC8D"] = aux8_falut_data;


            bitArray_do[7] = true;
            bitArray_do.CopyTo(ByteArray, 0);
            bitArray_do_2.CopyTo(ByteArray_2, 0);
#if !DBQ4
            instanDoCtrl_input.Write(0, 2, ByteArray);
            InitPci();
#endif



            // initexpertest();

        }

        #region 初始化采集卡

        //  static UInt32 count_plot = 500;
        //*/     
        static string dataready_name = "";
        static uint a = 0;
        static ErrorCode ErrorCode = ErrorCode.Success;
        static ErrorCode ErrorCodes = ErrorCode.Success;

        static void Analyze_Data()
        {
            Tuple<double[], string, bool, Int32> data_spci;
            double[] doubles;
            while (true)
            {
                if (pci_aux_plot.Count > 0)
                {
                    while (!pci_aux_plot.TryDequeue(out doubles)) { }
                    for (int j = 0; j < doubles.Length; j++)
                        if (dic_pci_aux[j % s_channel_count_] != "")
                            xoyplot_data_[dic_pci_aux[j % s_channel_count_]].Enqueue(doubles[j]);
                }
                if (pci_input_plot.Count > 0)
                {
                    while (!pci_input_plot.TryDequeue(out doubles)) { }
                    for (int j = 0; j < doubles.Length; j++)
                        if (dic_pci_input[j % s_channel_count_] != "")
                            xoyplot_data_[dic_pci_input[j % s_channel_count_]].Enqueue(doubles[j]);
                }
                if (pci_aux.Count > 0)
                {
                    while (!pci_aux.TryDequeue(out data_spci)) { }
                    for (int j = 0; j < data_spci.Item1.Length; j++)
                        if (dic_pci_aux[j % s_channel_count_] != "")
                            s_pci_data_[dic_pci_aux[j % s_channel_count_]].Enqueue(Tuple.Create(data_spci.Item1[j], data_spci.Item2, data_spci.Item3, data_spci.Item4));
                }
                if (pci_input.Count > 0)
                {
                    while (!pci_input.TryDequeue(out data_spci)) { }
                    for (int j = 0; j < data_spci.Item1.Length; j++)
                        if (dic_pci_input[j % s_channel_count_] != "")
                            s_pci_data_[dic_pci_input[j % s_channel_count_]].Enqueue(Tuple.Create(data_spci.Item1[j], data_spci.Item2, data_spci.Item3, data_spci.Item4));
                }
            }
        }



        #endregion
        #region 工位自检
        public int count_page { get; set; } = 100;

        public ObservableCollection<Struce> inputa_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> inputb_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> inputc_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux1_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux2_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux3_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux4_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux5_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux6_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux7_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> aux8_falut_data { get; set; } = new ObservableCollection<Struce>();
        public ObservableCollection<Struce> datagrid_experti { get; set; } = new ObservableCollection<Struce>();
        private RelayCommand pageUpdatedCmd;
        public RelayCommand PageUpdatedCmd => pageUpdatedCmd ?? (pageUpdatedCmd = new RelayCommand(PageUpdated));
        public void ContentChange(int index_)
        {
            try
            {
            }
            catch (Exception)
            {
            }
        }
        private void PageUpdated(object obj)
        {
            var temp = obj.GetType().GetProperty("Info")?.GetValue(obj);
            if (temp != null) { ContentChange((int)temp); }
        }
        private void InitWpa()
        {
            temp_Wpa.Clear();
            WpaTes wpaTes;
            //Dictionary<int, string> keyValuePairs = new Dictionary<int, string>() { { 0, "A相" }, { 1, "B相" }, { 2, "C相" } };
            //for (int i = 0; i < 3; i++)
            //{
            //    wpaTes = new WpaTes() { channel_ = keyValuePairs[i], energize_open_value_ = dic_value_wpa_data_[i][0], energize_stickiness_value_ = dic_value_wpa_data_[i][1], release_open_value_ = dic_value_wpa_data_[i][2], release_stickiness_value_ = dic_value_wpa_data_[i][3] };
            //    temp_Wpa.Add(wpaTes);
            //}

            for (int i = 0; i < aux_count_; i++)
            {
                if (s_state_contactor_dic_[i + 1] != StatusModel.close)
                {
                    wpaTes = new WpaTes() { channel_ = string.Format("第 {0} 路", (i + 1).ToString()), energize_open_value_ = dic_value_wpa_data_[i][0], energize_stickiness_value_ = dic_value_wpa_data_[i][1], release_open_value_ = dic_value_wpa_data_[i][2], release_stickiness_value_ = dic_value_wpa_data_[i][3] };
                    temp_Wpa.Add(wpaTes);
                }
            }
        }



        #endregion

        #region 线圈测试
        public ObservableCollection<CoilTestStruc> coil_test_ { get; set; } = new ObservableCollection<CoilTestStruc>();
        private void InitCoil()
        {
            for (int i = 0; i < aux_count_; i++)
            {
                CoilTestStruc coilTestStruc = new CoilTestStruc() { channel_order_ = i, sample_code_ = string.Format("编号{0}", i), come_back_value_ = 3.33, self_retain_value_ = 4.44 };
                coil_test_.Add(coilTestStruc);
            }
        }
        #endregion
        #region 数据库
        private bool InitData(string str)
        {
            CreatData creatData = new CreatData();
            if (creatData.TableExists(str))
            {
                return false;
            }
            if (!creatData.CreateTemp(str))
                MessageBox.Show("DB Temp");
            if (!creatData.RuningData(str))
                MessageBox.Show("DB Run");
            if (!creatData.FaultData(str))
                MessageBox.Show("DB Fault");
            if (!creatData.CreatePlots(str))
                MessageBox.Show("DB Plots");
            if (!creatData.CreateTemp(str))
                MessageBox.Show("DB FristPlots");
            if (!creatData.SaveDataPowerOFF(str))
                MessageBox.Show("DB POF");
            return true;
        }
        #endregion
        private Dictionary<int, string> dic_inttotemp_ = new Dictionary<int, string>() {
            { 1,"Temp1"},{ 2,"Temp2"},{ 3,"Temp3"},{ 4,"Temp4"},{ 5,"Temp5"},{ 6,"Temp6"},{ 7,"Temp7"},{ 8,"Temp8"},
        };
        private int index_save_temp_order_ = 1;
        private void SaveTempData(object sender, ElapsedEventArgs e)//存储温度
        {
            foreach (var item in name_array_aux)
            {
                SqlFrame sqlFrame2 = new SqlFrame();
                sqlFrame2.cmd = "Insert";
                sqlFrame2.parameters = new object[2];
                sqlFrame2.parameters[0] = "TempData" + table_name;
                Dictionary<string, object> columms = new Dictionary<string, object>();
                columms["Id"] = "true";
                columms["datetime"] = "DATETIME"; //存储时间
                columms["aux"] = "TEXT";//辅助触点
                columms["temp"] = "REAL";//温度值
                sqlFrame2.parameters[1] = columms;
                DBManager.dbManager.AddSqlQueue(MainWindow.QyDBPath, sqlFrame2);
            }
        }

        public string[] array_str_savedata = new string[] { "DC1H", "DC1D", "DC2H", "DC2D", "DC3H", "DC3D", "DC4H", "DC4D", "DC5H", "DC5D", "DC6H", "DC6D", "DC7H", "DC7D", "DC8H", "DC8D" };

        private void SaveRuningData(object sender, ElapsedEventArgs e)//存储运行数据
        {
            foreach (var item in array_str_savedata)
            {
                SqlFrame sqlFrame2 = new SqlFrame();
                sqlFrame2.cmd = "Insert";
                sqlFrame2.parameters = new object[2];
                sqlFrame2.parameters[0] = "RunTimeData" + table_name;
                Dictionary<string, object> columms = new Dictionary<string, object>();
                columms["datetime"] = "DATETIME"; //存储时间
                columms["typeofswitch"] = "TEXT";//动合 静合
                columms["auxname"] = "TEXT";//动合 静合
                columms["buckvalue"] = "REAL";//压降值
                columms["status"] = "INT"; //通断状态0 1
                columms["voltagevalue_a"] = "REAL";
                columms["voltagevalue_b"] = "REAL";
                columms["voltagevalue_c"] = "REAL";
                columms["currentvalue_a"] = "REAL";
                columms["currentvalue_b"] = "REAL";
                columms["currentvalue_c"] = "REAL";
                columms["currentvalue_ia"] = "REAL";
                columms["currentvalue_ib"] = "REAL";
                columms["currentvalue_ic"] = "REAL";
                sqlFrame2.parameters[1] = columms;
                DBManager.dbManager.AddSqlQueue(MainWindow.QyDBPath, sqlFrame2);
            }
        }
        static double validatatime_cone = 0;
        static double validatatime_discone = 0;

        static double monittime_cone = 0;
        static double monittime_discone = 0;

        private RelayCommand experitest;
        public RelayCommand Experitest => experitest ?? (experitest = new RelayCommand(StartExperiTest));
        private void StartExperiTest(object obj)
        {
            table_name = test_batch_ + "_" + test_rly_type_;
            DetectionChannel();
            
#if THREE_PHASE
            count_input = Convert.ToInt16(Math.Ceiling(s_convert_clk_rate_ / hz_power));
#else
            count_input = 1;
#endif
            Thread thread1 = new Thread(DataAnaluze) { IsBackground = true };
            if (string.IsNullOrEmpty(test_batch_) || string.IsNullOrEmpty(test_rly_type_))
            {
                MessageBox.Show("请输入 常规参数中\'规格型号\'和\'实验批号\'");
                return;
            }
            //if (InitData(table_name))
            //    MessageBox.Show("break off error");
            //else
            //    MessageBox.Show("run ing");
            InitData(table_name);
            start_capture = true;

            int connect_cf = Convert.ToInt32(cf_connect_);//通断比:左
            int disconnect_cf = Convert.ToInt32(cf_fault_upper_limit_);//通断比:右
            double time_switch = (1000.0 * 60.0) / working_frequency_;//通断一次时间
            double connect = time_switch / (disconnect_cf + connect_cf) * connect_cf;//通持续时间
            double dis_connect = time_switch / (disconnect_cf + connect_cf) * disconnect_cf;//断持续时间
            double int_connect = Math.Round(connect);//通持续时间四舍五入
            double int_disconnect = Math.Round(dis_connect);//断持续时间四舍五入
            validatatime_cone = int_connect;
            validatatime_discone = int_disconnect;
            monittime_cone = ((100 - data_monit_time_) / 100) * int_connect;//有效数据时刻:通
            monittime_discone = ((100 - data_monit_time_) / 100) * int_disconnect;//有效数据时刻:断
#if !DBQ4
            switch (relay_type)
            {
                case "普通型":
                    multimediaTimer_rly_general.Start(1, 1, callback_general);
                    break;
                case "极化型":
                    multimediaTimer_rly_polar.Start(1, 1, callback_ploar);
                    break;
                case "磁保持型":
                    break;
                case "交流型":
                    break;
                default:
                    System.Windows.MessageBox.Show("未选择继电器类型");
                    return;
            }
#endif
            thread1.Start();
            Task.Run(() => { InsertFaultData(); });
            Task.Run(() => { InsertFaultDatas(); });
            Settimer();
        }
        static bool pause_ = false;





        private void Settimer()
        {
            if (aTimer_SaveData_ == null)
            {
                aTimer_SaveData_ = new System.Timers.Timer();
                aTimer_SaveData_.Interval = 1000.0 * 60 * 1;//// count_ai;
                aTimer_SaveData_.Elapsed += SaveRuningData;
                aTimer_SaveData_.AutoReset = true;
                aTimer_SaveData_.Enabled = true;
            }
            return;
            if (aTimer_SaveData_Temp_ == null)
            {
                aTimer_SaveData_Temp_ = new System.Timers.Timer();
                aTimer_SaveData_Temp_.Interval = 1000.0 * time_save_temperat * 2;
                aTimer_SaveData_Temp_.Elapsed += SaveTempData;
                aTimer_SaveData_Temp_.AutoReset = true;
                aTimer_SaveData_Temp_.Enabled = true;
            }
            //if (aTimer_ChangeTempData_ == null)
            //{
            //    aTimer_ChangeTempData_ = new System.Timers.Timer();
            //    aTimer_ChangeTempData_.Interval = 1000 * 60 * 60; //BasicSetting.conver_time_;
            //    aTimer_ChangeTempData_.Elapsed += ChangeTempDo;
            //    aTimer_ChangeTempData_.AutoReset = true;
            //    aTimer_ChangeTempData_.Enabled = true;
            //}

        }

        static byte index_temp_calc = 0;
        private void ChangeTempDo(object sender, ElapsedEventArgs e)
        {
            if (index_temp_calc < 6)
                bitArray_do[index_temp_calc % 3] = !bitArray_do[index_temp_calc % 3];
            if (index_temp_calc == 6)
            { bitArray_do[0] = false; bitArray_do[1] = true; bitArray_do[2] = false; }
            if (index_temp_calc == 7)
            { bitArray_do[0] = true; bitArray_do[1] = false; bitArray_do[2] = true; }
            index_temp = 0;
            for (int start = 2; start >= 0; start--)
            {
                index_temp <<= 1;
                if (bitArray_do[start])
                {
                    index_temp |= 1;
                }
            }
            index_temp_calc++;
            if (index_temp_calc == 8)
            {
                index_temp_calc = 0;
                bitArray_do[0] = false;
                bitArray_do[1] = false;
                bitArray_do[2] = false;
            }
        }

#if DEBUG
        public void initexpertest()
        {
            for (int i = 0; i < 99; i++)
            {
                foreach (string name_ in name_array_all)
                {
                    keyValuePairs[name_].Add(new Struce() { fault_index = "1次故障", energize_open_ = string.Format("{0}:energize_open_", i), release_open_ = string.Format("{0}:release_open_", i), energize_stickiness_ = string.Format("{0}:energize_stickiness_", i), release_stickiness_ = string.Format("{0}:release_stickiness_", i) });
                }
            }
        }
#endif

        /// <summary>
        /// 清空绘图缓存
        /// </summary>
        public void ClearnXoyPlotData()
        {
            double out_date;
            foreach (var str in name_array)
            {
                for (int i = 0; i < xoyplot_data_[str].Count; i++)
                {
                    xoyplot_data_[str].TryDequeue(out out_date);
                    xoyplot_data_[str].TryDequeue(out out_date);
                }
            }
        }

        void DetectionChannel()
        {
            for (int i = 1; i < 9; ++i)
            {
                if (s_state_contactor_dic_[i] == StatusModel.close)
                {
                    switch (i)
                    {
                        case 1:
                            dic_pci_aux[0] = "";
                            dic_pci_aux[1] = "";
                            _close--;
                            break;
                        case 2:
                            dic_pci_aux[2] = "";
                            dic_pci_aux[3] = "";
                            _close--;
                            break;
                        case 3:
                            dic_pci_aux[4] = "";
                            dic_pci_aux[5] = "";
                            _close--;
                            break;
                        case 4:
                            dic_pci_aux[6] = "";
                            dic_pci_aux[7] = "";
                            _close--;
                            break;
                        case 5:
                            dic_pci_aux[8] = "";
                            dic_pci_aux[9] = "";
                            _close--;
                            break;
                        case 6:
                            dic_pci_aux[10] = "";
                            dic_pci_aux[11] = "";
                            _close--;
                            break;
                        case 7:
                            dic_pci_aux[12] = "";
                            dic_pci_aux[13] = "";
                            _close--;
                            break;
                        case 8:
                            dic_pci_aux[14] = "";
                            dic_pci_aux[15] = "";
                            _close--;
                            break;
                        default:
                            break;
                    }
                }
                else if (s_state_contactor_dic_[i] == StatusModel.staticc)
                    switch (i)
                    {
                        case 1:
                            dic_pci_aux[0] = "";
                            break;
                        case 2:
                            dic_pci_aux[2] = "";
                            break;
                        case 3:
                            dic_pci_aux[4] = "";
                            break;
                        case 4:
                            dic_pci_aux[6] = "";
                            break;
                        case 5:
                            dic_pci_aux[8] = "";
                            break;
                        case 6:
                            dic_pci_aux[10] = "";
                            break;
                        case 7:
                            dic_pci_aux[12] = "";
                            break;
                        case 8:
                            dic_pci_aux[14] = "";
                            break;
                        default:
                            break;
                    }
                else if (s_state_contactor_dic_[i] == StatusModel.dynamicc)
                    switch (i)
                    {
                        case 1:
                            dic_pci_aux[1] = "";
                            break;
                        case 2:
                            dic_pci_aux[3] = "";
                            break;
                        case 3:
                            dic_pci_aux[5] = "";
                            break;
                        case 4:
                            dic_pci_aux[7] = "";
                            break;
                        case 5:
                            dic_pci_aux[9] = "";
                            break;
                        case 6:
                            dic_pci_aux[11] = "";
                            break;
                        case 7:
                            dic_pci_aux[13] = "";
                            break;
                        case 8:
                            dic_pci_aux[15] = "";
                            break;
                        default:
                            break;
                    }
            }

        }


       
        void ExportDataToWorld()
        {
            InfoT1();
            ExportData exportData = new ExportData(table_name);
            exportData.CreateDocument(DateTime.Now.ToString("yyyy-MM-dd") + test_rly_type_ + test_batch_);
            exportData.CreateParagraph("条件", ExportData.TITLE.headline);
            exportData.InitTableInfo_T1(T1);
            exportData.CreateParagraph("自检及首次试验波形", ExportData.TITLE.headline);
            exportData.CreateParagraph("数据", ExportData.TITLE.subheading);
            exportData.InitTableInfo_T2(T2);
            exportData.CreateParagraph("首次开断试验波形", ExportData.TITLE.subheading);
            for (int i = 0; i < 1; i++)
            {
                exportData.InsertAPicture("F://test.jpg", "附注");
            }

            exportData.CreateParagraph("失效数据统计", ExportData.TITLE.headline);
            exportData.CreateParagraph("失效范围统计", ExportData.TITLE.subheading);
            exportData.InitTableInfo_T3(T3);
            exportData.CreateParagraph("失效数据统计", ExportData.TITLE.subheading);
            exportData.InitTableInfo_T4(T4);

        }
        static string[] T1 = new string[14];
        static List<List<double>> T2 = new List<List<double>>();
        static List<List<double>> T3 = new List<List<double>>();
        static List<List<double>> T4 = new List<List<double>>();
        void InfoT1()
        {
            T1[0] = table_name;
            T1[1] = validatatime_cone.ToString();
            T1[2] = formulation_times_.ToString();
            T1[3] = validatatime_discone.ToString();
            T1[4] = on_off_cycle.ToString();
            T1[5] = "触点升温指标";
            T1[6] = fault_upper_limit_.ToString();
            T1[7] = load_vol_.ToString();
            T1[8] = load_sour_vol_.ToString();
            T1[9] = load_cur_.ToString();
            T1[10] = load_sour_cur_.ToString();
            T1[11] = data_sensitivity_percent_stick_.ToString();
            T1[12] = table_name;
            T1[13] = data_sensitivity_percent_sever_.ToString();
        }


        #region 数据解析
        //解析AI数据判断故障
        int _close = 11;
        ConcurrentDictionary<string, List<Tuple<List<Double>, bool, string>>> insert_to_database = new ConcurrentDictionary<string, List<Tuple<List<Double>, bool, string>>>();
        ConcurrentDictionary<string, List<double>> inser_to_deque = new ConcurrentDictionary<string, List<double>>();
        ConcurrentDictionary<string, Int32> channel_index = new ConcurrentDictionary<string, Int32>();
        ConcurrentDictionary<string, Int32> fault_count_accumul = new ConcurrentDictionary<string, Int32>();
        ConcurrentDictionary<string, bool> save_old_mark_bool = new ConcurrentDictionary<string, bool>(); //周期是否有故障
        ConcurrentDictionary<string, string> save_old_mark_str = new ConcurrentDictionary<string, string>(); //周期故障描述
        ConcurrentDictionary<string, string> save_old_mark_date = new ConcurrentDictionary<string, string>(); //周期故障时间
        ConcurrentDictionary<string, double> save_old_mark_double = new ConcurrentDictionary<string, double>(); //周期故障数值
        ConcurrentDictionary<string, bool> save_old_mark_statu = new ConcurrentDictionary<string, bool>(); //周期通断
        ConcurrentDictionary<string, double> save_old_mark_voltage = new ConcurrentDictionary<string, double>(); //周期故障电压
        ConcurrentDictionary<string, double> save_old_mark_current = new ConcurrentDictionary<string, double>(); //周期故障电流
        ConcurrentDictionary<string, DataBaseTemple> dic_data_base_temple_ = new ConcurrentDictionary<string, DataBaseTemple>();
        //放大倍数
        public static double fac_insi = 1;//瞬时电流
        public static double fac_u = 38.9769;//额定电压
        public static double fac_i = 1;//额定电流
        public static double fac_aux = 7.8;//辅助触点
        /// <summary>
        /// 解析输入数据
        /// </summary>
        private void DataAnaluze()
        {
            double aux_u = 1;
            Tuple<bool, string> resul;
            Tuple<double, string, bool, Int32> tuple = Tuple.Create(0.0, "", false, -1);//采集值,时间,通断,周期
            //false 为故障
            foreach (string name_ in name_array_all)
            {
                fault_count_accumul[name_] = 1;
                channel_index[name_] = 1;
                insert_to_database[name_] = new List<Tuple<List<Double>, bool, string>>();
                inser_to_deque[name_] = new List<double>();
                save_old_mark_bool[name_] = false;
                save_old_mark_str[name_] = "";
                save_old_mark_date[name_] = "";
                save_old_mark_double[name_] = -999.9;
                dic_data_base_temple_[name_] = new DataBaseTemple();
            }

            //a
            double ua = 0, ia = 0, iai = 0;
            double Ua = 0, Ia = 0, Iai = 0;
#if THREE_PHASE
            //b
            double ub = 0, ib = 0, ibi = 0;
            double Ub = 0, Ib = 0, Ibi = 0;
            //c
            double uc = 0, ic = 0, ici = 0;
            double Uc = 0, Ic = 0, Ici = 0;
#endif
            while (start_capture)
            {
                if (s_pci_data_["Ua"].Count >= count_input)
                {
                    #region UA
                    // string data = "";
                    for (int i = 0; i < count_input; i++)
                    {

                        while (!s_pci_data_["Ua"].TryDequeue(out tuple))
                        {
                        }
                        //  data +=string.Format("{0}--{1}-{2}\n", Math.Pow(tuple.Item1, 2).ToString(), tuple.Item4, tuple.Item3.ToString());
                        ua += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IaRate"].TryDequeue(out tuple))
                        {
                        }
                        ia += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IaInstant"].TryDequeue(out tuple))
                        {
                        }
                        iai += Math.Pow(tuple.Item1, 2);
                    }

                    Ia = Math.Sqrt(ia / count_input) * fac_insi;
                    Iai = Math.Sqrt(iai / count_input) * fac_i;
                    Ua = Math.Sqrt(ua / count_input) * fac_u;
                    ua = 0; ia = 0; iai = 0;
                    //using (StreamWriter original = new StreamWriter("D:\\datas.txt", true, Encoding.UTF8))
                    //{
                    //    original.WriteLine("{0}{1}{2}\n", data, Math.Pow(Ua, 2), tuple.Item3.ToString());
                    //}
                    resul = new Tuple<bool, string>(false, "");
                    resul = DataAnalyzeInput(Ua, tuple.Item3);
                    if (channel_index["Ua"] != tuple.Item4)
                    {
                        U_value = Math.Round(Ua, 3);
                        I_value = Math.Round(Ia, 3);
                        if (save_old_mark_bool["Ua"])
                        {
                            double buckvalue = save_old_mark_double["Ua"];
                            string typeofswitch = save_old_mark_str["Ua"];
                            string faultnumber = "Ua";
                            bool status = save_old_mark_statu["Ua"];
                            double voltagevalue = save_old_mark_voltage["Ua"];
                            double currentvalue = save_old_mark_current["Ua"];
                            string datetime = save_old_mark_date["Ua"];
                            data_plot.Enqueue(Tuple.Create(buckvalue, typeofswitch, faultnumber, status, voltagevalue, currentvalue, datetime));
                            // Task.Run(() => { InsertFaultData(); });
                            // Task.Run(() => { InsertFaultData(save_old_mark_double["Ua"], save_old_mark_str["Ua"], "Ua", save_old_mark_statu["Ua"], save_old_mark_voltage["Ua"], save_old_mark_current["Ua"], save_old_mark_date["Ua"]); });
                        }
                        channel_index["Ua"] = tuple.Item4;
                        if (insert_to_database["Ua"].Count < 6)
                        {
                            Tuple<List<Double>, bool, string> tuple1 = new Tuple<List<double>, bool, string>(inser_to_deque["Ua"], save_old_mark_bool["Ua"], save_old_mark_str["Ua"]);
                            insert_to_database["Ua"].Add(tuple1);
                            save_old_mark_bool["Ua"] = false;
                            save_old_mark_str["Ua"] = "";
                        }
                        else if (insert_to_database["Ua"].Count == 6)
                        {
                            List<List<double>> ls = new List<List<double>>();
                            for (int ii = 0; ii < insert_to_database["Ua"].Count; ii++)
                            {
                                ls.Add(new List<double>(insert_to_database["Ua"][ii].Item1));
                            }
                            for (int i = 0; i < 3; i++)
                            {
                                if (insert_to_database["Ua"][i].Item2)
                                {
                                    data_plots.Enqueue(Tuple.Create("Ua", fault_count_accumul["Ua"]++, insert_to_database["Ua"][i].Item3, ls));
                                    //  Task.Run(() => { InsertFaultDatas(); });
                                    fault_count_save["Ua" + insert_to_database["Ua"][i].Item3]++;
                                }
                            }
                            insert_to_database["Ua"].Add(Tuple.Create(inser_to_deque["Ua"], save_old_mark_bool["Ua"], save_old_mark_str["Ua"]));
                            save_old_mark_bool["Ua"] = false;
                            save_old_mark_str["Ua"] = "";
                        }
                        else if (insert_to_database["Ua"].Count > 6)
                        {

                            if (insert_to_database["Ua"][3].Item2)
                            {
                                //存储
                                List<List<double>> ls = new List<List<double>>();
                                for (int i = 0; i < insert_to_database["Ua"].Count; i++)
                                {
                                    ls.Add(new List<double>(insert_to_database["Ua"][i].Item1));
                                }
                                if (++fault_count_save["Ua" + insert_to_database["Ua"][3].Item3] < fault_upper_limit_)
                                {
                                    data_plots.Enqueue(Tuple.Create("Ua", fault_count_accumul["Ua"]++, insert_to_database["Ua"][3].Item3, ls));
                                    // Task.Run(() => { InsertFaultDatas(); });
                                }
                                else
                                {
                                    dic_pci_input[0] = "";
                                    dic_pci_input[3] = "";
                                    dic_pci_input[6] = "";
                                    _close--;
                                    for (int z = 0; z < s_pci_data_["Ua"].Count; z++)
                                    {
                                        Tuple<double, string, bool, Int32> tuples = Tuple.Create(0.0, "", false, -1);
                                        s_pci_data_["Ua"].TryDequeue(out tuples);
                                        s_pci_data_["IaInstant"].TryDequeue(out tuples);
                                        s_pci_data_["IaRate"].TryDequeue(out tuples);
                                    }
                                }
                            }
                            insert_to_database["Ua"].RemoveAt(0);
                            insert_to_database["Ua"].Add(Tuple.Create(inser_to_deque["Ua"], save_old_mark_bool["Ua"], save_old_mark_str["Ua"]));
                            save_old_mark_bool["Ua"] = false;
                        }
                        inser_to_deque["Ua"].Clear();
                        inser_to_deque["Ua"].Add(Ua);
                        //  });
                    }//
                    else
                    {

                        inser_to_deque["Ua"].Add(Ua);
                        if (resul.Item1)
                        {
                            save_old_mark_bool["Ua"] = true;
                            save_old_mark_str["Ua"] = resul.Item2;
                            save_old_mark_double["Ua"] = aux_u;
                            save_old_mark_current["Ua"] = Ia;
                            save_old_mark_voltage["Ua"] = Ua;
                            save_old_mark_statu["Ua"] = tuple.Item3;
                            save_old_mark_date["Ua"] = tuple.Item2;

                        }
                    }
                    if (resul.Item1)
                    {
                        string name_ = "Ua";
                        string fault_name = name_ + resul.Item2;
                        if (calc_falut_count_[fault_name] != tuple.Item4)
                        {
                            calc_falut_count_[fault_name] = tuple.Item4;
                            fault_count[fault_name]++;
                            System.Windows.Application.Current.Dispatcher.Invoke(new Action(() =>
                            {
                                if (keyValuePairs[name_].Count < fault_count[fault_name] && keyValuePairs[name_].Count < fault_upper_limit_)
                                    keyValuePairs[name_].Add(new Struce() { fault_index = string.Format("{0}次故障", fault_count[fault_name]) });
                                int i = on_off_cycle;//该故障次数
                                int count = fault_count[fault_name] - 1;//该故障次数
                                if (count < keyValuePairs[name_].Count)
                                    switch (resul.Item2)
                                    {
                                        case "粘故障":
                                            keyValuePairs[name_][count].stickiness_ = string.Format("{0}/{1:F2}", i, Ua);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        case "断故障":
                                            keyValuePairs[name_][count].open_ = string.Format("{0}/{1:F2}", i, Ua);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        default: break;
                                    }
                            }));
                        }
                    }
                    #endregion
#if THREE_PHASE
                    #region UB
                    for (int i = 0; i < count_input; i++)
                    {
                        while (!s_pci_data_["Ub"].TryDequeue(out tuple))
                        {
                        }
                        ub += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IbRate"].TryDequeue(out tuple))
                        {
                        }
                        ib += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IbInstant"].TryDequeue(out tuple))
                        {
                        }
                        ibi += Math.Pow(tuple.Item1, 2);
                    }
                    Ib = Math.Sqrt(ib / count_input) * fac_i;
                    Ibi = Math.Sqrt(ibi / count_input) * fac_insi;
                    Ub = Math.Sqrt(ub / count_input) * fac_u;

                    resul = new Tuple<bool, string>(false, "");
                    resul = DataAnalyzeInput(Ub, tuple.Item3);

                    if (channel_index["Ub"] != tuple.Item4)
                    {
                        if (save_old_mark_bool["Ub"])
                        {
                            double buckvalue = save_old_mark_double["Ub"];
                            string typeofswitch = save_old_mark_str["Ub"];
                            string faultnumber = "Ub";
                            bool status = save_old_mark_statu["Ub"];
                            double voltagevalue = save_old_mark_voltage["Ub"];
                            double currentvalue = save_old_mark_current["Ub"];
                            string datetime = save_old_mark_date["Ub"];
                            data_plot.Enqueue(Tuple.Create(buckvalue, typeofswitch, faultnumber, status, voltagevalue, currentvalue, datetime));
                            //  Task.Run(() => { InsertFaultData(); });

                            // Task.Run(() => { InsertFaultData(save_old_mark_double["Ub"], save_old_mark_str["Ub"], "Ub", save_old_mark_statu["Ub"], save_old_mark_voltage["Ub"], save_old_mark_current["Ub"], save_old_mark_date["Ub"]); });
                        }
                        channel_index["Ub"] = tuple.Item4;
                        if (insert_to_database["Ub"].Count < 6)
                        {
                            insert_to_database["Ub"].Add(Tuple.Create(inser_to_deque["Ub"], save_old_mark_bool["Ub"], save_old_mark_str["Ub"]));
                            save_old_mark_bool["Ub"] = false;
                            save_old_mark_str["Ub"] = "";
                        }
                        else if (insert_to_database["Ub"].Count == 6)
                        {
                            List<List<double>> ls = new List<List<double>>();
                            for (int ii = 0; ii < insert_to_database["Ua"].Count; ii++)
                            {
                                ls.Add(new List<double>(insert_to_database["Ua"][ii].Item1));
                            }
                            for (int i = 0; i < 3; i++)
                            {
                                if (insert_to_database["Ub"][i].Item2)
                                {
                                    data_plots.Enqueue(Tuple.Create("Ub", fault_count_accumul["Ub"]++, insert_to_database["Ub"][i].Item3, ls));
                                    // Task.Run(() => { InsertFaultDatas(); });
                                    fault_count_save["Ub" + insert_to_database["Ub"][i].Item3]++;
                                }
                            }
                            insert_to_database["Ub"].Add(Tuple.Create(inser_to_deque["Ub"], save_old_mark_bool["Ub"], save_old_mark_str["Ub"]));
                            save_old_mark_bool["Ub"] = false;
                            save_old_mark_str["Ub"] = "";
                        }
                        else if (insert_to_database["Ub"].Count > 6)
                        {

                            if (insert_to_database["Ub"][3].Item2)
                            {
                                //存储
                                List<List<double>> ls = new List<List<double>>();
                                for (int i = 0; i < insert_to_database["Ub"].Count; i++)
                                {
                                    ls.Add(new List<double>(insert_to_database["Ub"][i].Item1));
                                }
                                if (++fault_count_save["Ub" + insert_to_database["Ub"][3].Item3] < fault_upper_limit_)
                                {
                                    data_plots.Enqueue(Tuple.Create("Ub", fault_count_accumul["Ub"]++, insert_to_database["Ub"][3].Item3, ls));
                                    // Task.Run(() => { InsertFaultDatas(); });

                                }
                                else
                                {
                                    _close--;
                                    dic_pci_input[1] = "";
                                    dic_pci_input[4] = "";
                                    dic_pci_input[7] = "";
                                    for (int z = 0; z < s_pci_data_["Ub"].Count; z++)
                                    {
                                        Tuple<double, string, bool, Int32> tuples = Tuple.Create(0.0, "", false, -1);
                                        s_pci_data_["Ub"].TryDequeue(out tuples);
                                        s_pci_data_["IbInstant"].TryDequeue(out tuples);
                                        s_pci_data_["IbRate"].TryDequeue(out tuples);
                                    }
                                }
                            }
                            insert_to_database["Ub"].RemoveAt(0);
                            insert_to_database["Ub"].Add(Tuple.Create(inser_to_deque["Ub"], save_old_mark_bool["Ub"], save_old_mark_str["Ub"]));
                            save_old_mark_bool["Ub"] = false;
                            save_old_mark_str["Ub"] = "";
                        }
                        inser_to_deque["Ub"].Clear();
                        inser_to_deque["Ub"].Add(Ub);
                        //    });

                    }
                    else
                    {
                        inser_to_deque["Ub"].Add(Ub);
                        if (resul.Item1)
                        {
                            save_old_mark_bool["Ub"] = true;
                            save_old_mark_str["Ub"] = resul.Item2;
                            save_old_mark_double["Ub"] = aux_u;
                            save_old_mark_current["Ub"] = Ia;
                            save_old_mark_voltage["Ub"] = Ua;
                            save_old_mark_statu["Ub"] = tuple.Item3;
                            save_old_mark_date["Ub"] = tuple.Item2;
                        }
                    }

                    if (resul.Item1)
                    {
                        string name_ = "Ub";
                        string fault_name = name_ + resul.Item2;
                        if (calc_falut_count_[fault_name] != tuple.Item4)
                        {
                            calc_falut_count_[fault_name] = tuple.Item4;
                            fault_count[fault_name]++;
                            System.Windows.Application.Current.Dispatcher.Invoke(new Action(() =>
                            {
                                if (keyValuePairs[name_].Count < fault_count[fault_name] && keyValuePairs[name_].Count < fault_upper_limit_)
                                    keyValuePairs[name_].Add(new Struce() { fault_index = string.Format("{0}次故障", fault_count[fault_name]) });
                                int i = on_off_cycle;//该故障次数
                                int count = fault_count[fault_name] - 1;//该故障次数
                                if (count < keyValuePairs[name_].Count)
                                    switch (resul.Item2)
                                    {
                                        case "粘故障":
                                            keyValuePairs[name_][count].stickiness_ = string.Format("{0}/{1:F2}", i, Ub);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        case "断故障":
                                            keyValuePairs[name_][count].open_ = string.Format("{0}/{1:F2}", i, Ub);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        default: break;
                                    }
                            }));
                        }

                    }
                    #endregion
                    #region UC
                    for (int i = 0; i < count_input; i++)
                    {
                        while (!s_pci_data_["Uc"].TryDequeue(out tuple))
                        {
                        }
                        uc += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IcRate"].TryDequeue(out tuple))
                        {
                        }
                        ic += Math.Pow(tuple.Item1, 2);
                        while (!s_pci_data_["IcInstant"].TryDequeue(out tuple))
                        {
                        }
                        ici += Math.Pow(tuple.Item1, 2);
                    }
                    Ic = Math.Sqrt(ic / count_input) * fac_i;
                    Ici = Math.Sqrt(ici / count_input) * fac_insi;
                    Uc = Math.Sqrt(uc / count_input) * fac_u;
                    ub = 0; uc = 0; ibi = 0; ici = 0;
                    ic = 0; ib = 0;
                    resul = new Tuple<bool, string>(false, "");
                    resul = DataAnalyzeInput(Uc, tuple.Item3);
                    if (channel_index["Uc"] != tuple.Item4)
                    {
                        if (save_old_mark_bool["Uc"])
                        {
                            double buckvalue = save_old_mark_double["Uc"];
                            string typeofswitch = save_old_mark_str["Uc"];
                            string faultnumber = "Uc";
                            bool status = save_old_mark_statu["Uc"];
                            double voltagevalue = save_old_mark_voltage["Uc"];
                            double currentvalue = save_old_mark_current["Uc"];
                            string datetime = save_old_mark_date["Uc"];
                            data_plot.Enqueue(Tuple.Create(buckvalue, typeofswitch, faultnumber, status, voltagevalue, currentvalue, datetime));
                            // Task.Run(() => { InsertFaultData(); });

                            //   Task.Run(() => { InsertFaultData(save_old_mark_double["Uc"], save_old_mark_str["Uc"], "Uc", save_old_mark_statu["Uc"], save_old_mark_voltage["Uc"], save_old_mark_current["Uc"], save_old_mark_date["Uc"]); });
                        }
                        channel_index["Uc"] = tuple.Item4;
                        if (insert_to_database["Uc"].Count < 6)
                        {
                            insert_to_database["Uc"].Add(Tuple.Create(inser_to_deque["Uc"], save_old_mark_bool["Uc"], save_old_mark_str["Uc"]));
                            save_old_mark_bool["Uc"] = false;
                            save_old_mark_str["Uc"] = "";
                        }
                        else if (insert_to_database["Uc"].Count == 6)
                        {
                            List<List<double>> ls = new List<List<double>>();
                            for (int ii = 0; ii < insert_to_database["Uc"].Count; ii++)
                            {
                                ls.Add(new List<double>(insert_to_database["Uc"][ii].Item1));
                            }
                            for (int i = 0; i < 3; i++)
                            {
                                if (insert_to_database["Uc"][i].Item2)
                                {
                                    data_plots.Enqueue(Tuple.Create("Uc", fault_count_accumul["Uc"]++, insert_to_database["Uc"][i].Item3, ls));
                                    // Task.Run(() => { InsertFaultDatas(); });
                                    fault_count_save["Uc" + insert_to_database["Uc"][i].Item3]++;
                                }
                            }
                            insert_to_database["Uc"].Add(Tuple.Create(inser_to_deque["Uc"], save_old_mark_bool["Uc"], save_old_mark_str["Uc"]));
                            save_old_mark_bool["Uc"] = false;
                            save_old_mark_str["Uc"] = "";
                        }
                        else if (insert_to_database["Uc"].Count > 6)
                        {

                            if (insert_to_database["Uc"][3].Item2)
                            {
                                //存储
                                List<List<double>> ls = new List<List<double>>();
                                for (int i = 0; i < insert_to_database["Uc"].Count; i++)
                                {
                                    ls.Add(new List<double>(insert_to_database["Uc"][i].Item1));
                                }
                                if (++fault_count_save["Uc" + insert_to_database["Uc"][3].Item3] < fault_upper_limit_)
                                {
                                    data_plots.Enqueue(Tuple.Create("Uc", fault_count_accumul["Uc"]++, insert_to_database["Uc"][3].Item3, ls));
                                    // Task.Run(() => { InsertFaultDatas(); });

                                }
                                else
                                {
                                    _close--;
                                    dic_pci_input[2] = "";
                                    dic_pci_input[5] = "";
                                    dic_pci_input[8] = "";
                                    for (int z = 0; z < s_pci_data_["Uc"].Count; z++)
                                    {
                                        Tuple<double, string, bool, Int32> tuples = Tuple.Create(0.0, "", false, -1);
                                        s_pci_data_["Uc"].TryDequeue(out tuples);
                                        s_pci_data_["IcInstant"].TryDequeue(out tuples);
                                        s_pci_data_["IcRate"].TryDequeue(out tuples);
                                    }
                                }
                            }
                            insert_to_database["Uc"].RemoveAt(0);
                            insert_to_database["Uc"].Add(Tuple.Create(inser_to_deque["Uc"], save_old_mark_bool["Uc"], save_old_mark_str["Uc"]));
                            save_old_mark_bool["Uc"] = false;
                            save_old_mark_str["Uc"] = "";
                        }
                        inser_to_deque["Uc"].Clear();
                        inser_to_deque["Uc"].Add(Uc);
                        //   });
                    }
                    else
                    {
                        inser_to_deque["Uc"].Add(Uc);
                        if (resul.Item1)
                        {
                            save_old_mark_bool["Uc"] = true;
                            save_old_mark_str["Uc"] = resul.Item2;
                            save_old_mark_double["Uc"] = aux_u;
                            save_old_mark_current["Uc"] = Ia;
                            save_old_mark_voltage["Uc"] = Ua;
                            save_old_mark_statu["Uc"] = tuple.Item3;
                            save_old_mark_date["Uc"] = tuple.Item2;
                        }
                    }

                    if (resul.Item1)
                    {
                        string name_ = "Uc";
                        string fault_name = name_ + resul.Item2;
                        if (calc_falut_count_[fault_name] != tuple.Item4)
                        {
                            calc_falut_count_[fault_name] = tuple.Item4;
                            fault_count[fault_name]++;
                            System.Windows.Application.Current.Dispatcher.Invoke(new Action(() =>
                            {
                                if (keyValuePairs[name_].Count < fault_count[fault_name] && keyValuePairs[name_].Count < fault_upper_limit_)
                                    keyValuePairs[name_].Add(new Struce() { fault_index = string.Format("{0}次故障", fault_count[fault_name]) });
                                int i = on_off_cycle;//该故障次数
                                int count = fault_count[fault_name] - 1;//该故障次数
                                if (count < keyValuePairs[name_].Count)
                                    switch (resul.Item2)
                                    {
                                        case "粘故障":
                                            keyValuePairs[name_][count].stickiness_ = string.Format("{0}/{1:F2}", i, Uc);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        case "断故障":
                                            keyValuePairs[name_][count].open_ = string.Format("{0}/{1:F2}", i, Uc);
                                            keyValuePairs[name_][count].phase_ = name_;
                                            break;
                                        default: break;
                                    }
                            }));
                        }
                    }
                    #endregion
#endif
                }
                foreach (string name_ in name_array_aux)
                {
                    int index_ = 0; bool continu = true;
                    resul = new Tuple<bool, string>(false, "");
                    if (s_state_contactor_dic_[auxmap_statu_name_[name_]] == StatusModel.close)
                        continue;
                    if (s_pci_data_[name_].Count >= count_input)
                    {
                        for (int ii = 0; ii < count_input; ii++)
                        {
                            index_++;
                            while (!s_pci_data_[name_].TryDequeue(out tuple))
                            {

                            }
                            //if (name_.Contains("DC1"))
                                if (index_ >= count_input / 2 && continu)
                                {
                                    aux_u = tuple.Item1 * fac_aux;
                                    if (s_state_contactor_dic_[auxmap_statu_name_[name_]] == StatusModel.dynamicc && name_.EndsWith("H"))
                                        resul = DataAnalyzeAux_Dynamicc(name_, aux_u, tuple.Item3);
                                    else if (s_state_contactor_dic_[auxmap_statu_name_[name_]] == StatusModel.staticc && name_.EndsWith("D"))
                                        resul = DataAnalyzeAux_Staticc(name_, aux_u, tuple.Item3);
                                    else if (s_state_contactor_dic_[auxmap_statu_name_[name_]] == StatusModel.converc)
                                        resul = DataAnalyzeAux(name_, aux_u, tuple.Item3);

                                    if (channel_index[name_] != tuple.Item4)
                                    {
                                        if (save_old_mark_bool[name_])
                                        {

                                            double buckvalue = save_old_mark_double[name_];
                                            string typeofswitch = save_old_mark_str[name_];
                                            string faultnumber = name_;
                                            bool status = save_old_mark_statu[name_];
                                            double voltagevalue = save_old_mark_voltage[name_];
                                            double currentvalue = save_old_mark_current[name_];
                                            string datetime = save_old_mark_date[name_];
                                            data_plot.Enqueue(Tuple.Create(buckvalue, typeofswitch, faultnumber, status, voltagevalue, currentvalue, datetime));
                                            //Task.Run(() =>
                                            //{

                                            //    InsertFaultData();
                                            //});

                                            // Task.Run(() => { InsertFaultData(save_old_mark_double[name_], save_old_mark_str[name_], name_, save_old_mark_statu[name_], save_old_mark_voltage[name_], save_old_mark_current[name_], save_old_mark_date[name_]); });
                                        }
                                        channel_index[name_] = tuple.Item4;
                                        if (insert_to_database[name_].Count < 6)
                                        {
                                            insert_to_database[name_].Add(Tuple.Create(inser_to_deque[name_], save_old_mark_bool[name_], save_old_mark_str[name_]));
                                            save_old_mark_bool[name_] = false;
                                            save_old_mark_str[name_] = "";
                                        }
                                        else if (insert_to_database[name_].Count == 6)
                                        {
                                            List<List<double>> ls = new List<List<double>>();
                                            for (int x = 0; x < insert_to_database[name_].Count; x++)
                                            {
                                                ls.Add(new List<double>(insert_to_database[name_][x].Item1));
                                            }
                                            for (int i = 0; i < 3; i++)
                                            {

                                                if (insert_to_database[name_][i].Item2)
                                                {
                                                    data_plots.Enqueue(Tuple.Create(name_, fault_count_accumul[name_]++, insert_to_database[name_][i].Item3, ls));
                                                    // Task.Run(() => { InsertFaultDatas(); });
                                                    fault_count_save[name_ + insert_to_database[name_][i].Item3]++;
                                                }
                                            }
                                            insert_to_database[name_].Add(Tuple.Create(inser_to_deque[name_], save_old_mark_bool[name_], save_old_mark_str[name_]));
                                            save_old_mark_bool[name_] = false;
                                            save_old_mark_str[name_] = "";
                                        }
                                        else if (insert_to_database[name_].Count > 6)
                                        {

                                            if (insert_to_database[name_][3].Item2)
                                            {
                                                //存储
                                                List<List<double>> ls = new List<List<double>>();
                                                for (int i = 0; i < insert_to_database[name_].Count; i++)
                                                {
                                                    ls.Add(new List<double>(insert_to_database[name_][i].Item1));
                                                }
                                                if (++fault_count_save[name_ + insert_to_database[name_][3].Item3] < fault_upper_limit_)
                                                {
                                                    data_plots.Enqueue(Tuple.Create(name_, fault_count_accumul[name_]++, insert_to_database[name_][3].Item3, ls));
                                                    // Task.Run(() => { InsertFaultDatas(); });
                                                }
                                                else
                                                {
                                                    _close--;
                                                    dic_pci_aux[close_channel[name_]] = "";
                                                    for (int z = 0; z < s_pci_data_[name_].Count; z++)
                                                    {
                                                        Tuple<double, string, bool, Int32> tuples = Tuple.Create(0.0, "", false, -1);
                                                        s_pci_data_[name_].TryDequeue(out tuples);
                                                    }
                                                }
                                            }
                                            insert_to_database[name_].RemoveAt(0);
                                            insert_to_database[name_].Add(Tuple.Create(inser_to_deque[name_], save_old_mark_bool[name_], save_old_mark_str[name_]));
                                            save_old_mark_bool[name_] = false;
                                            save_old_mark_str[name_] = "";
                                        }
                                        inser_to_deque[name_].Clear();
                                        inser_to_deque[name_].Add(aux_u);
                                        //     });
                                    }
                                    else
                                    {
                                        dic_data_base_temple_[name_].voltagevalue = Ua;
                                        dic_data_base_temple_[name_].currentvalue = Ia;
                                        dic_data_base_temple_[name_].buckvalue = aux_u;
                                        dic_data_base_temple_[name_].typedfswitch = name_;
                                        dic_data_base_temple_[name_].status = tuple.Item3;
                                        dic_data_base_temple_[name_].time = tuple.Item2;

                                        inser_to_deque[name_].Add(aux_u);
                                        if (resul.Item1)
                                        {
                                            save_old_mark_bool[name_] = true;
                                            save_old_mark_str[name_] = resul.Item2;
                                            save_old_mark_double[name_] = aux_u;
                                            save_old_mark_current[name_] = Ia;
                                            save_old_mark_voltage[name_] = Ua;
                                            save_old_mark_statu[name_] = tuple.Item3;
                                            save_old_mark_date[name_] = tuple.Item2;
                                        }
                                    }

                                    if (resul.Item1)
                                    {
                                        string fault_name = name_ + resul.Item2;
                                        if (calc_falut_count_[fault_name] != tuple.Item4)
                                        {
                                            calc_falut_count_[fault_name] = tuple.Item4;
                                            fault_count[fault_name]++;
                                            System.Windows.Application.Current.Dispatcher.Invoke(new Action(() =>
                                            {
                                                if (keyValuePairs[name_].Count < fault_count[fault_name] && keyValuePairs[name_].Count < fault_upper_limit_)
                                                    keyValuePairs[name_].Add(new Struce() { fault_index = string.Format("{0}次故障", fault_count[fault_name]) });
                                                int i = on_off_cycle;//该故障次数
                                                int count = fault_count[fault_name] - 1;//该故障次数
                                                if (count < keyValuePairs[name_].Count)
                                                    switch (resul.Item2)
                                                    {
                                                        case "D断故障":
                                                            keyValuePairs[name_][count].release_open_ = string.Format("{0}/{1:F2}", i, aux_u);
                                                            break;
                                                        case "H断故障":
                                                            keyValuePairs[name_][count].energize_open_ = string.Format("{0}/{1:F2}", i, aux_u);
                                                            break;
                                                        case "D粘故障":
                                                            keyValuePairs[name_][count].release_stickiness_ = string.Format("{0}/{1:F2}", i, aux_u);
                                                            break;
                                                        case "H粘故障":
                                                            keyValuePairs[name_][count].energize_stickiness_ = string.Format("{0}/{1:F2}", i, aux_u);
                                                            break;
                                                        default: break;
                                                    }
                                            }));
                                        }
                                    }
                                    continu = false;
                                }
                        }
                    }
                }
                if (_close < 1)
                {
#if !DBQ4
                    StopExperiTest(null);
#endif
                }
            }
        }

        /// <summary>
        /// 判断输入粘断故障
        /// </summary>
        /// <param name="data_aux">AI值</param>
        /// <param name="connec_disconnec_">通断状态</param>
        private Tuple<bool, string> DataAnalyzeInput(double data_aux, bool connec_disconnec_)
        {
            Tuple<bool, string> result = Tuple.Create(false, "");
            if (connec_disconnec_)//主触点(使能)跟据文档解释是吸合
            {
                if (data_aux > input_falut * data_sensitivity_percent_stick_ * 0.01)
                {
                    result = Tuple.Create(true, "粘故障");
                }
            }
            else
            {
                if (data_aux < input_falut * data_sensitivity_percent_sever_ * 0.01)
                {
                    result = Tuple.Create(true, "断故障");
                }
            }
            return result;
        }

        /// <summary>
        /// 动合类型辅助触点判断
        /// </summary>
        /// <param name="name_">保留</param>
        /// <param name="data_aux">AI数值</param>
        /// <param name="connec_disconnec_">通断状态</param>
        /// <returns>故障标识,故障类型</returns>
        private Tuple<bool, string> DataAnalyzeAux_Dynamicc(string name_, double data_aux, bool connec_disconnec_)
        {
            Tuple<bool, string> result = Tuple.Create(false, "");
            if (connec_disconnec_)//主触点(使能)跟据文档解释是吸合
            {
                if (data_aux > aux_falut * data_sensitivity_percent_stick_ * 0.01)
                {
                    result = Tuple.Create(true, "H粘故障");
                }
            }
            else
            {
                if (data_aux < aux_falut * data_sensitivity_percent_sever_ * 0.01)
                {
                    result = Tuple.Create(true, "H断故障");
                }
            }
            return result;
        }

        /// <summary>
        /// 静合类型触点
        /// </summary>
        /// <param name="name_"></param>
        /// <param name="data_aux">AI值</param>
        /// <param name="connec_disconnec_">通断状态</param>
        /// <returns>故障标识位,故障描述</returns>
        private Tuple<bool, string> DataAnalyzeAux_Staticc(string name_, double data_aux, bool connec_disconnec_)
        {
            Tuple<bool, string> result = Tuple.Create(false, "");
            if (connec_disconnec_)//主触点(使能)跟据文档解释是吸合
            {
                if (data_aux < aux_falut * data_sensitivity_percent_sever_ * 0.01)
                {
                    result = Tuple.Create(true, "D断故障");
                }
            }
            else
            {
                if (data_aux > aux_falut * data_sensitivity_percent_stick_ * 0.01)
                {
                    result = Tuple.Create(true, "D粘故障");
                }
            }
            return result;
        }

        /// <summary>
        /// 转换类型触点判断
        /// </summary>
        /// <param name="name_">判断动合还是静合触点  D：静合  H:动合</param>
        /// <param name="data_aux">AI值</param>
        /// <param name="connec_disconnec_">通断状态</param>
        /// <returns>故障标识,故障类型</returns>
        private Tuple<bool, string> DataAnalyzeAux(string name_, double data_aux, bool connec_disconnec_)
        {
            Tuple<bool, string> result = Tuple.Create(false, "");
            if (name_.EndsWith("D"))
                if (connec_disconnec_)//主触点(使能)跟据文档解释是吸合
                {
                    if (data_aux < aux_falut * data_sensitivity_percent_sever_ * 0.01)
                    {
                        result = Tuple.Create(true, "D断故障");
                    }
                }
                else
                {
                    if (data_aux > aux_falut * data_sensitivity_percent_stick_ * 0.01)
                    {
                        result = Tuple.Create(true, "D粘故障");
                    }
                }
            else if (name_.EndsWith("H"))
                if (connec_disconnec_)//主触点(使能)跟据文档解释是吸合
                {
                    if (data_aux > aux_falut * data_sensitivity_percent_stick_ * 0.01)
                    {
                        result = Tuple.Create(true, "H粘故障");
                    }
                }
                else
                {
                    if (data_aux < aux_falut * data_sensitivity_percent_sever_ * 0.01)
                    {
                        result = Tuple.Create(true, "H断故障");
                    }
                }
            return result;
        }

        /// <summary>
        /// 好像没用
        /// </summary>
        private void TempAnalyze()
        {
            while (start_capture)
            {
                string dic_name = string.Format("Temp{0}", index_temp + 1);
            }
        }
        #endregion
#if !DBQ4
        #region 采集卡
        /// <summary>
        /// 初始化采集卡
        /// </summary>
        /// <exception cref="Exception"></exception>
        void InitPci()
        {
            waveformAiCtrl_input.Overrun += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_Overrun);
            waveformAiCtrl_aux.Overrun += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_Overrun2);
            // waveformAiCtrl_input.DataReady += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_DataReady);
            //   waveformAiCtrl_aux.DataReady += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_DataReady);
            waveformAiCtrl_input.Stopped += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_Stopped);
            waveformAiCtrl_aux.Stopped += new EventHandler<BfdAiEventArgs>(waveformAiCtrl_Stopped);
            waveformAiCtrl_aux.SelectedDevice = new DeviceInformation(deviceDescription_aux);
            waveformAiCtrl_input.SelectedDevice = new DeviceInformation(deviceDescription_input);
            if (BioFailed(errorCode))
            {
                throw new Exception();
            }
            Conversion conversion = waveformAiCtrl_input.Conversion;
            Conversion conversion2 = waveformAiCtrl_aux.Conversion;
            conversion.ChannelStart = s_start_channel_;
            conversion.ChannelCount = s_channel_count_;
            conversion.ClockRate = s_convert_clk_rate_;
            conversion2.ChannelStart = s_start_channel_;
            conversion2.ChannelCount = s_channel_count_;
            conversion2.ClockRate = s_convert_clk_rate_;
            Record record = waveformAiCtrl_input.Record;
            Record record2 = waveformAiCtrl_aux.Record;
            record.SectionCount = s_section_count_;
            record.SectionLength = s_section_length_;
            record2.SectionCount = s_section_count_;
            record2.SectionLength = s_section_length_;
            errorCode = waveformAiCtrl_input.Prepare();
            if (BioFailed(errorCode))
            {
                System.Windows.MessageBox.Show("启动失败");
            }
            errorCode = waveformAiCtrl_aux.Prepare();
            if (BioFailed(errorCode))
            {
                System.Windows.MessageBox.Show("启动失败");
            }
            Thread.Sleep(10);
            Thread Get_WaveAiCtrlData = new Thread(WaveAiCtrl_GetData) { IsBackground = true };
            Get_WaveAiCtrlData.Start();
            errorCode2 = waveformAiCtrl_aux.Start();
            errorCode = waveformAiCtrl_input.Start();
            if (BioFailed(errorCode))
            {
                System.Windows.MessageBox.Show("启动失败");
            }
            if (BioFailed(errorCode2))
            {
                System.Windows.MessageBox.Show("启动失败");
            }
            Thread thread_analyze = new Thread(Analyze_Data) { IsBackground = true };
            thread_analyze.Start();

        }
        static void waveformAiCtrl_Stopped(object sender, BfdAiEventArgs e)
        {
            WaveformAiCtrl waveformAiCtrl_input = (WaveformAiCtrl)sender;
            int s_channel_count_Max = waveformAiCtrl_input.Features.ChannelCountMax;
            int startChan = waveformAiCtrl_input.Conversion.ChannelStart;
            int s_channel_count_ = waveformAiCtrl_input.Conversion.ChannelCount;
            int s_section_length_ = waveformAiCtrl_input.Record.SectionLength;
            int returnedCount = 0, returnedSumCount = 0, getDataCount = 0;
            int remainingCount = e.Count;
            int bufSize = s_section_length_ * s_channel_count_;
            // 缓冲区段长度，当发出 "Stopped "事件信号时，驱动程序更新的数据计数为 e.count。
            if (remainingCount > s_channel_count_)
            {
                s_pci_buffer_input = new double[bufSize];
                do
                {
                    getDataCount = Math.Min(bufSize, remainingCount);
                    waveformAiCtrl_input.GetData(getDataCount, s_pci_buffer_input, 0, out returnedCount);
                    remainingCount -= returnedCount;
                    returnedSumCount += returnedCount;
                } while (returnedCount > 0);
            }
        }
        static void waveformAiCtrl_Stopped2(object sender, BfdAiEventArgs e)
        {
            WaveformAiCtrl waveformAiCtrl_input = (WaveformAiCtrl)sender;
            int s_channel_count_Max = waveformAiCtrl_input.Features.ChannelCountMax;
            int startChan = waveformAiCtrl_input.Conversion.ChannelStart;
            int s_channel_count_ = waveformAiCtrl_input.Conversion.ChannelCount;
            int s_section_length_ = waveformAiCtrl_input.Record.SectionLength;
            int returnedCount = 0, returnedSumCount = 0, getDataCount = 0;
            int remainingCount = e.Count;
            int bufSize = s_section_length_ * s_channel_count_;
            // 缓冲区段长度，当发出 "Stopped "事件信号时，驱动程序更新的数据计数为 e.count。
            if (remainingCount > s_channel_count_)
            {
                s_pci_buffer_input = new double[bufSize];
                do
                {
                    getDataCount = Math.Min(bufSize, remainingCount);
                    waveformAiCtrl_input.GetData(getDataCount, s_pci_buffer_input, 0, out returnedCount);
                    remainingCount -= returnedCount;
                    returnedSumCount += returnedCount;
                } while (returnedCount > 0);
            }
        }
        static bool BioFailed(ErrorCode err)
        {
            return err < ErrorCode.Success && err >= ErrorCode.ErrorHandleNotValid;
        }
        static void waveformAiCtrl_Overrun(object sender, BfdAiEventArgs e)
        {
            using (StreamWriter original = new StreamWriter("D:\\Overrun.txt", true, Encoding.UTF8))
            {
                original.WriteLine("卡1 缓存溢出时间:{0}", DateTime.Now.ToString("d HH:mm:ss.fff"));
            }
        }
        static void waveformAiCtrl_Overrun2(object sender, BfdAiEventArgs e)
        {
            using (StreamWriter original = new StreamWriter("D:\\Overrun2.txt", true, Encoding.UTF8))
            {
                original.WriteLine("卡2 缓存溢出时间:{0}", DateTime.Now.ToString("d HH:mm:ss.fff"));
            }
        }
        #endregion
        #region 工况流程
        /// <summary>
        /// 工位自检
        /// </summary>
        private RelayCommand manualexp;
        public RelayCommand ManualExp => manualexp ?? (manualexp = new RelayCommand(Manual_experiment));
        private void Manual_experiment(object obj)
        {
            bitArray_do[15] = true;
            InitWpa();
            for (int i = 0; i < bitArray_do.Count; i++)
            {
                bitArray_do[i] = false;
            }
            int count_aux = temp_Wpa.Count;
            bitArray_do[6] = !bitArray_do[6];// temp_Wpa[0].energize_open_value_
            bitArray_do.CopyTo(ByteArray, 0);
            instanDoCtrl_input.Write(0, 2, ByteArray);
            Thread.Sleep(Convert.ToInt16(monittime_cone));
            errorCode = instantAIContrl_input.Read(s_start_channel_, s_channel_count_, get_data_spci_instant);
            if (count_aux > 0)
            {
                temp_Wpa[0].energize_stickiness_value_ = Math.Round(get_data_spci_instant[0], 2);
                temp_Wpa[0].release_open_value_ = Math.Round(get_data_spci_instant[1], 2);
            }
            if (count_aux > 1)
            {
                temp_Wpa[1].energize_stickiness_value_ = Math.Round(get_data_spci_instant[2], 2);
                temp_Wpa[1].release_open_value_ = Math.Round(get_data_spci_instant[3], 2);
            }
            if (count_aux > 2)
            {
                temp_Wpa[2].energize_stickiness_value_ = Math.Round(get_data_spci_instant[4], 2);
                temp_Wpa[2].release_open_value_ = Math.Round(get_data_spci_instant[5], 2);
            }
            if (count_aux > 3)
            {
                temp_Wpa[3].energize_stickiness_value_ = Math.Round(get_data_spci_instant[6], 2);
                temp_Wpa[3].release_open_value_ = Math.Round(get_data_spci_instant[7], 2);
            }
            if (count_aux > 4)
            {
                temp_Wpa[4].energize_stickiness_value_ = Math.Round(get_data_spci_instant[8], 2);
                temp_Wpa[4].release_open_value_ = Math.Round(get_data_spci_instant[9], 2);

            }
            if (count_aux > 5)
            {
                temp_Wpa[5].energize_stickiness_value_ = Math.Round(get_data_spci_instant[10], 2);
                temp_Wpa[5].release_open_value_ = Math.Round(get_data_spci_instant[11], 2);

            }
            if (count_aux > 6)
            {
                temp_Wpa[6].energize_stickiness_value_ = Math.Round(get_data_spci_instant[12], 2);
                temp_Wpa[6].release_open_value_ = Math.Round(get_data_spci_instant[13], 2);
            }
            if (count_aux > 7)
            {
                temp_Wpa[7].energize_stickiness_value_ = Math.Round(get_data_spci_instant[14], 2);
                temp_Wpa[7].release_open_value_ = Math.Round(get_data_spci_instant[15], 2);
            }
            bitArray_do[6] = !bitArray_do[6];
            bitArray_do.CopyTo(ByteArray, 0);
            instanDoCtrl_input.Write(0, 2, ByteArray);
            Thread.Sleep(Convert.ToInt16(monittime_discone));
            errorCode = instantAIContrl_input.Read(s_start_channel_, s_channel_count_, get_data_spci_instant);


            if (count_aux > 0)
            {
                temp_Wpa[0].energize_open_value_ = Math.Round(get_data_spci_instant[0], 2);
                temp_Wpa[0].release_stickiness_value_ = Math.Round(get_data_spci_instant[1], 2);
            }
            if (count_aux > 1)
            {
                temp_Wpa[1].energize_open_value_ = Math.Round(get_data_spci_instant[2], 2);
                temp_Wpa[1].release_stickiness_value_ = Math.Round(get_data_spci_instant[3], 2);
            }
            if (count_aux > 2)
            {
                temp_Wpa[2].energize_open_value_ = Math.Round(get_data_spci_instant[4], 2);
                temp_Wpa[2].release_stickiness_value_ = Math.Round(get_data_spci_instant[5], 2);
            }
            if (count_aux > 3)
            {
                temp_Wpa[3].energize_open_value_ = Math.Round(get_data_spci_instant[6], 2);
                temp_Wpa[3].release_stickiness_value_ = Math.Round(get_data_spci_instant[7], 2);
            }
            if (count_aux > 4)
            {
                temp_Wpa[4].release_stickiness_value_ = Math.Round(get_data_spci_instant[8], 2);
                temp_Wpa[4].energize_open_value_ = Math.Round(get_data_spci_instant[9], 2);
            }
            if (count_aux > 5)
            {
                temp_Wpa[5].energize_open_value_ = Math.Round(get_data_spci_instant[10], 2);
                temp_Wpa[5].release_stickiness_value_ = Math.Round(get_data_spci_instant[11], 2);
            }
            if (count_aux > 6)
            {
                temp_Wpa[6].energize_open_value_ = Math.Round(get_data_spci_instant[12], 2);
                temp_Wpa[6].release_stickiness_value_ = Math.Round(get_data_spci_instant[13], 2);
            }
            if (count_aux > 7)
            {
                temp_Wpa[7].energize_open_value_ = Math.Round(get_data_spci_instant[14], 2);
                temp_Wpa[7].release_stickiness_value_ = Math.Round(get_data_spci_instant[15], 2);
            }
        }


        private RelayCommand pause;
        public RelayCommand Pause => pause ?? (pause = new RelayCommand(Startpause));

        /// <summary>
        /// 暂停
        /// </summary>
        private void Startpause(object obj)
        {
            pause_ = !pause_;
            if (pause_)
            {
                waveformAiCtrl_aux.Stop();
                waveformAiCtrl_input.Stop();
            }
            else
            {
                waveformAiCtrl_aux.Start();
                waveformAiCtrl_input.Start();
            }

        }

        /// <summary>
        /// 停止测试
        /// </summary>
        private void Stop()
        {
            errorCode = waveformAiCtrl_input.Stop();
            errorCode = waveformAiCtrl_aux.Stop();
            switch (relay_type)
            {
                case "普通型":
                    multimediaTimer_rly_general.Stop();
                    break;
                case "极化型":
                    multimediaTimer_rly_polar.Stop();
                    break;
                case "磁保持型":
                    multimediaTimer_rly_magentic.Stop();
                    break;
                case "交流型":
                    multimediaTimer_rly_ac.Stop();
                    break;
            }

            if (!(aTimer_ChangeTempData_ == null))
            {
                aTimer_ChangeTempData_.Dispose();
                aTimer_ChangeTempData_ = null;
            }
            if (!(aTimer_SaveData_ == null))//aTimer_SaveData_Temp_
            {
                aTimer_SaveData_.Dispose();
                aTimer_SaveData_ = null;
            }
            if (!(aTimer_SaveData_Temp_ == null))
            {
                aTimer_SaveData_Temp_.Dispose();
                aTimer_SaveData_Temp_ = null;
            }

            instanDoCtrl_input.Write(0, 0x00);
            instanDoCtrl_input.Write(1, 0x00);
            for (int i = 0; i < bitArray_do.Count; i++)
            {
                bitArray_do[i] = false;
            }
            count_millisecond = 0;
        }

        private RelayCommand stopExperitest;
        public RelayCommand StopExperitest => stopExperitest ?? (stopExperitest = new RelayCommand(StopExperiTest));
        private void StopExperiTest(object obj)
        {
            errorCode = waveformAiCtrl_input.Stop();
            errorCode = waveformAiCtrl_aux.Stop();
            switch (relay_type)
            {
                case "普通型":
                    multimediaTimer_rly_general.Stop();
                    break;
                case "极化型":
                    multimediaTimer_rly_polar.Stop();
                    break;
                case "磁保持型":
                    multimediaTimer_rly_magentic.Stop();
                    break;
                case "交流型":
                    multimediaTimer_rly_ac.Stop();
                    break;
            }

            if (!(aTimer_ChangeTempData_ == null))
            {
                aTimer_ChangeTempData_.Dispose();
                aTimer_ChangeTempData_ = null;
            }
            if (!(aTimer_SaveData_ == null))//aTimer_SaveData_Temp_
            {
                aTimer_SaveData_.Dispose();
                aTimer_SaveData_ = null;
            }
            if (!(aTimer_SaveData_Temp_ == null))
            {
                aTimer_SaveData_Temp_.Dispose();
                aTimer_SaveData_Temp_ = null;
            }
            start_capture = false;
            instanDoCtrl_input.Write(0, 0x00);
            instanDoCtrl_input.Write(1, 0x00);
            for (int i = 0; i < bitArray_do.Count; i++)
            {
                bitArray_do[i] = false;
            }
            count_millisecond = 0;
        }
        MultimediaTimer.TimerCallback callback_general = TimerCallbackMethod;
        MultimediaTimer multimediaTimer_rly_general = new MultimediaTimer();
        MultimediaTimer.TimerCallback callback_ploar = TimerCallbackMethodRLY_Polar;
        MultimediaTimer multimediaTimer_rly_polar = new MultimediaTimer();
        MultimediaTimer.TimerCallback callback_magentic = TimerCallbackMethodRLY_Magentic;
        MultimediaTimer multimediaTimer_rly_magentic = new MultimediaTimer();
        MultimediaTimer.TimerCallback callback_ac = TimerCallbackMethodRLY_Ac;
        MultimediaTimer multimediaTimer_rly_ac = new MultimediaTimer();
        static UInt64 count_millisecond = 0;
        static bool run_ = true;

        /// <summary>
        /// 普通型
        /// </summary>
        /// <param name="id"></param>
        /// <param name="msg"></param>
        /// <param name="user"></param>
        /// <param name="dw1"></param>
        /// <param name="dw2"></param>
        static void TimerCallbackMethod(uint id, uint msg, UIntPtr user, UIntPtr dw1, UIntPtr dw2)
        {//定时器内容

            count_millisecond++;
            if (count_millisecond == 1)
            {

                on_off_cycle++;
                data_is_efficent = false;
                bitArray_do[6] = !bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[6];
                run_ = true;

                return;
            }
            if (count_millisecond < monittime_cone)//+ validatatime_cone
            {
                return;
            }
            if (run_)
                data_is_efficent = true;
            if (count_millisecond < validatatime_cone)//+ validatatime_cone
            {
                return;
            }
            if (run_)
            {
                data_is_efficent = false;
                bitArray_do[6] = !bitArray_do[6];
                connect_statu_ = bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                run_ = false;
            }
            if (count_millisecond < validatatime_cone + monittime_discone) { return; }// + validatatime_discone + validatatime_cone
            data_is_efficent = true;
            if (count_millisecond < validatatime_cone + validatatime_discone) { return; }
            data_is_efficent = false;
            count_millisecond = 0;
            if (on_off_cycle > MainWindow.QyViewModel.formulation_times_ || pause_)
                MainWindow.QyViewModel.StopExperiTest(null);
        }

        /// <summary>
        /// 计划
        /// </summary>
        /// <param name="id"></param>
        /// <param name="msg"></param>
        /// <param name="user"></param>
        /// <param name="dw1"></param>
        /// <param name="dw2"></param>
        static void TimerCallbackMethodRLY_Magentic(uint id, uint msg, UIntPtr user, UIntPtr dw1, UIntPtr dw2)
        {//定时器内容

            count_millisecond++;
            if (count_millisecond == 1)
            {
                on_off_cycle++;
                data_is_efficent = false;
                bitArray_do[7] = !bitArray_do[7];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[7];
                run_ = true;
                return;
            }
            if (count_millisecond < monittime_cone)//+ validatatime_cone
            {
                return;
            }
            data_is_efficent = true;
            if (count_millisecond < validatatime_cone)//+ validatatime_cone
            {
                return;
            }
            if (run_)
            {
                data_is_efficent = false;
                bitArray_do[7] = !bitArray_do[7];
                bitArray_do[8] = !bitArray_do[8];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[7];
                run_ = false;
            }
            if (count_millisecond < validatatime_cone + monittime_discone) { return; }// + validatatime_discone + validatatime_cone
            data_is_efficent = true;
            if (count_millisecond < validatatime_cone + validatatime_discone) { return; }
            count_millisecond = 0;
            data_is_efficent = false;
            bitArray_do[8] = false;
            bitArray_do[7] = false;
            bitArray_do.CopyTo(ByteArray, 0);
            instanDoCtrl_input.Write(0, 2, ByteArray);
            if (on_off_cycle > MainWindow.QyViewModel.formulation_times_)
                MainWindow.QyViewModel.StopExperiTest(null);
        }

        /// <summary>
        /// 交流
        /// </summary>
        /// <param name="id"></param>
        /// <param name="msg"></param>
        /// <param name="user"></param>
        /// <param name="dw1"></param>
        /// <param name="dw2"></param>
        static void TimerCallbackMethodRLY_Ac(uint id, uint msg, UIntPtr user, UIntPtr dw1, UIntPtr dw2)
        {//定时器内容

            count_millisecond++;
            if (count_millisecond == 1)
            {
                on_off_cycle++;
                data_is_efficent = false;
                bitArray_do[6] = !bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[6];
                run_ = true;
                return;
            }
            /*
            if (count_millisecond < validatatime_cone)
            {
                return;
            }*/
            data_is_efficent = true;
            if (count_millisecond < monittime_cone)
            {
                return;
            }
            if (run_)
            {
                bitArray_do[6] = !bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                data_is_efficent = false;
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[6];
                run_ = false;
            }


            if (count_millisecond < monittime_discone + monittime_cone) { return; }
            data_is_efficent = true;
            count_millisecond = 0;
            //   if (on_off_cycle > formulation_times_)
            //      StopExperiTest(null);
        }

        /// <summary>
        /// 磁保持
        /// </summary>
        /// <param name="id"></param>
        /// <param name="msg"></param>
        /// <param name="user"></param>
        /// <param name="dw1"></param>
        /// <param name="dw2"></param>
        static void TimerCallbackMethodRLY_Polar(uint id, uint msg, UIntPtr user, UIntPtr dw1, UIntPtr dw2)
        {//定时器内容

            count_millisecond++;
            if (count_millisecond == 1)
            {
                on_off_cycle++;
                data_is_efficent = false;
                bitArray_do[6] = !bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[6];
                run_ = true;
                return;
            }

            data_is_efficent = true;
            if (count_millisecond < monittime_cone)
            {
                return;
            }
            if (run_)
            {
                bitArray_do[6] = !bitArray_do[6];
                bitArray_do.CopyTo(ByteArray, 0);
                data_is_efficent = false;
                instanDoCtrl_input.Write(0, 2, ByteArray);
                connect_statu_ = bitArray_do[6];
                run_ = false;
            }


            if (count_millisecond < monittime_discone + monittime_cone) { return; }
            data_is_efficent = true;

            count_millisecond = 0;
            //   if (on_off_cycle > formulation_times_)
            //      StopExperiTest(null);
        }
        #endregion
#endif
        #region 数据库
        /// <summary>
        /// 用于测试
        /// </summary>
        /// <param name="data"></param>
        /// <param name="time"></param>
        /// <param name="connect"></param>
        /// <param name="name"></param>
        private void InsertDataBase(double data, string time, bool connect, string name)
        {
            SqlFrame sqlFrame2 = new SqlFrame();
            sqlFrame2.cmd = "Insert";
            sqlFrame2.parameters = new object[2];
            sqlFrame2.parameters[0] = "TestData";
            Dictionary<string, object> columms = new Dictionary<string, object>();
            columms["datetime"] = time;
            columms["typeofswitch"] = name;
            columms["buckvalue"] = data;
            columms["status"] = connect.ToString();
            sqlFrame2.parameters[1] = columms;
            DBManager.dbManager.AddSqlQueue(MainWindow.QyDBPath, sqlFrame2);
        }

        /// <summary>
        /// 存储故障数据
        /// </summary>
        private void InsertFaultData()
        {
            Tuple<double, string, string, bool, double, double, string> tuple = null;
            while (start_capture)
            {
                try
                {
                    if (data_plot.TryDequeue(out tuple))
                    {
                        SqlFrame sqlFrame2 = new SqlFrame();
                        sqlFrame2.cmd = "Insert";
                        sqlFrame2.parameters = new object[2];
                        sqlFrame2.parameters[0] = "AbonormalData" + table_name;
                        Dictionary<string, object> columms = new Dictionary<string, object>();
                        columms["buckvalue"] = Math.Round(tuple.Item1, 3).ToString();
                        columms["typeofswitch"] = tuple.Item2;
                        columms["faultnumber"] = tuple.Item3;//故障描述 "一次故障"
                        columms["status"] = tuple.Item4;
                        columms["voltagevalue"] = Math.Round(tuple.Item5, 3).ToString();
                        columms["currentvalue"] = Math.Round(tuple.Item6, 3).ToString();
                        columms["datetime"] = tuple.Item7;
                        sqlFrame2.parameters[1] = columms;
                        DBManager.dbManager.AddSqlQueue(MainWindow.QyDBPath, sqlFrame2);

                        //using (StreamWriter original = new StreamWriter("D:\\InsertFaultData_ov.txt", true, Encoding.UTF8))
                        //{
                        //    original.WriteLine("记录完成{0}{1}{2}{3}", DateTime.Now.ToString("d HH:mm:ss.fff"), tuple.Item2.ToString() + "次故障", tuple.Item3, tuple.Item1);
                        //}
                    }
                    Thread.Sleep(5);
                }
                catch (Exception)
                {
                    using (StreamWriter original = new StreamWriter("D:\\InsertFaultData_er.txt", true, Encoding.UTF8))
                    {
                        original.WriteLine("{0}{1}{2}", tuple.Item2, tuple.Item3, DateTime.Now.ToString("d HH:mm:ss.fff"));
                    }
                }

            }

        }
        /// <summary>
        /// 存储故障绘图
        /// </summary>
        private void InsertFaultDatas()
        {
            string fault_data = null;
            ulong count = 0;
            Tuple<string, int, string, List<List<double>>> tuple = null;
            while (start_capture)
            {
                try
                {
                    fault_data = null;
                    count = 0;
                    if (data_plots.TryDequeue(out tuple))
                    {
                        for (int i = 0; i < tuple.Item4.Count; i++)
                        {
                            for (int y = 0; y < tuple.Item4[i].Count; y++)
                            {
                                fault_data += string.Format("{0:0.000},", tuple.Item4[i][y]);
                                count++;
                            }
                        }
                        SqlFrame sqlFrame2 = new SqlFrame();
                        sqlFrame2.cmd = "Insert";
                        sqlFrame2.parameters = new object[2];
                        sqlFrame2.parameters[0] = "AbonormalDataPlots" + table_name;
                        Dictionary<string, object> columms = new Dictionary<string, object>();
                        columms["dataindex"] = tuple.Item2.ToString() + "次故障";//故障描述
                        columms["typeofswitch"] = tuple.Item3;//故障数据
                        columms["faultnumber"] = tuple.Item1;//
                        columms["dates"] = fault_data;//故障数据
                        sqlFrame2.parameters[1] = columms;
                        DBManager.dbManager.AddSqlQueue(MainWindow.QyDBPath, sqlFrame2);
                        //using (StreamWriter original = new StreamWriter("D:\\InsertFaultDatas_ov.txt", true, Encoding.UTF8))
                        //{
                        //    original.WriteLine("记录完成InsertFaultDatas{0}{1}{2}{3}", DateTime.Now.ToString("d HH:mm:ss.fff"), tuple.Item2.ToString() + "次故障", tuple.Item3, tuple.Item1);
                        //}
                    }
                    Thread.Sleep(10);
                }
                catch (Exception)
                {
                    using (StreamWriter original = new StreamWriter("D:\\InsertFaultDatas_er.txt", true, Encoding.UTF8))
                    {
                        original.WriteLine("{0}{1}{2}{3}{4} ", count, tuple.Item2.ToString() + "次故障", tuple.Item3, tuple.Item1, DateTime.Now.ToString("d HH:mm:ss.fff"));
                    }

                }

            }


        }
        #endregion
        static void WaveAiCtrl_GetData()
        {
            string str_datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.FFF");
            int returnedCount = 0;
            int returnedCount2 = 0;
            while (true)
            {
                str_datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.FFF");
                if (MainWindow.QyViewModel.start_capture)
                {
                    if ((data_is_efficent == data_is_efficent_his) && data_is_efficent)
                    {
                        ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                        if (ErrorCodes == ErrorCode)
                        {
                            if (clearn_plotdata_)
                            {
                                pci_input.Enqueue(Tuple.Create(s_pci_buffer_input, str_datetime, connect_statu_, on_off_cycle));
                            }
                            else
                            {
                                pci_input_plot.Enqueue(s_pci_buffer_input);
                                pci_input.Enqueue(Tuple.Create(s_pci_buffer_input, str_datetime, connect_statu_, on_off_cycle));
                            }
                        }
                        ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (clearn_plotdata_)
                            {
                                pci_aux.Enqueue(Tuple.Create(s_pci_buffer_aux, str_datetime, connect_statu_, on_off_cycle));
                            }
                            else
                            {
                                pci_aux_plot.Enqueue(s_pci_buffer_aux);
                                pci_aux.Enqueue(Tuple.Create(s_pci_buffer_aux, str_datetime, connect_statu_, on_off_cycle));
                            }
                        }
                        data_is_efficent_his = data_is_efficent;
                    }
                    else
                    {
                        ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (!clearn_plotdata_)
                                pci_input_plot.Enqueue(s_pci_buffer_input);
                        }
                        ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (!clearn_plotdata_)
                                pci_aux_plot.Enqueue(s_pci_buffer_aux);
                        }
                        data_is_efficent_his = data_is_efficent;
                    }
                }
                else
                {
                    ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                    if (!clearn_plotdata_ && ErrorCode == ErrorCodes)
                        pci_input_plot.Enqueue(s_pci_buffer_input);
                    ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                    if (!clearn_plotdata_ && ErrorCode == ErrorCodes)
                        pci_aux_plot.Enqueue(s_pci_buffer_aux);
                    data_is_efficent_his = data_is_efficent;
                }
            }
        }

        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, System.Text.StringBuilder retVal, int size, string filePath);
        public string ReadIni(string section, string key, string path)
        {
            // 每次从ini中读取多少字节
            System.Text.StringBuilder temp = new System.Text.StringBuilder(255);
            // section=配置节点名称，key=键名，temp=上面，path=路径
            GetPrivateProfileString(section, key, "", temp, 255, path);
            return temp.ToString();

        }
    }


#if false
    static void waveformAiCtrl_DataReady(object sender, BfdAiEventArgs e)
        {
            int returnedCount = 0;
            int returnedCount2 = 0;
            int remainingCount = e.Count;
            string str_datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.FFF");
            do
            {
                if (MainWindow.QyViewModel.start_capture)
                {
                    if ((data_is_efficent == data_is_efficent_his) && data_is_efficent)
                    {
                        ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                        if (ErrorCodes == ErrorCode)
                        {
                            if (clearn_plotdata_)
                            {
                                pci_input.Enqueue(Tuple.Create(s_pci_buffer_input, str_datetime, connect_statu_, on_off_cycle));
                            }
                            else
                            {
                                pci_input_plot.Enqueue(s_pci_buffer_input);
                                pci_input.Enqueue(Tuple.Create(s_pci_buffer_input, str_datetime, connect_statu_, on_off_cycle));
                            }
                        }
                        ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (clearn_plotdata_)
                            {
                                pci_aux.Enqueue(Tuple.Create(s_pci_buffer_aux, str_datetime, connect_statu_, on_off_cycle));
                            }
                            else
                            {
                                pci_aux_plot.Enqueue(s_pci_buffer_aux);
                                pci_aux.Enqueue(Tuple.Create(s_pci_buffer_aux, str_datetime, connect_statu_, on_off_cycle));
                            }
                        }
                        data_is_efficent_his = data_is_efficent;
                    }
                    else
                    {

                        ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (!clearn_plotdata_)
                                pci_input_plot.Enqueue(s_pci_buffer_input);
                        }
                        ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                        if (ErrorCode == ErrorCodes)
                        {
                            if (!clearn_plotdata_)
                                pci_aux_plot.Enqueue(s_pci_buffer_aux);
                        }

                        data_is_efficent_his = data_is_efficent;
                    }
                }
                else
                {
                    ErrorCodes = waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
                    //if (!clearn_plotdata_ && ErrorCode == ErrorCodes)
                    pci_input_plot.Enqueue(s_pci_buffer_input);
                    ErrorCodes = waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
                    //if (!clearn_plotdata_ && ErrorCode == ErrorCodes)
                    pci_aux_plot.Enqueue(s_pci_buffer_aux);
                    data_is_efficent_his = data_is_efficent;
                }
            } while (returnedCount > 0);
            //do
            //{
            //    waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_input, 0, out returnedCount);
            //    waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
            //} while (returnedCount2 > 0 || returnedCount > 0);
            //using (StreamWriter effective = new StreamWriter("D:\\as.txt", true, Encoding.UTF8))
            //{
            //    effective.Write(i + "\r");
            //}

            //do
            //{
            //    waveformAiCtrl_aux.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount2);
            //    waveformAiCtrl_input.GetData(bufSize, s_pci_buffer_aux, 0, out returnedCount);
            //} while (returnedCount2 > 0 || returnedCount > 0);
            if (remainingCount > s_section_length_ * s_channel_count_)
            {
                using (StreamWriter effective = new StreamWriter("D:\\a.txt", true, Encoding.UTF8))
                {
                    effective.Write("处理超时" + DateTime.Now.ToString() + "-----" + remainingCount + "\r");
                }

            }

        }
#endif

}
