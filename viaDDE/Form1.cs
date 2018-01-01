using NDde.Client;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace viaDDE
{
    public partial class Form1 : Form
    {
        public delegate void TickReceived(Tick tick);

        public event TickReceived OnTick;

        public string[] allSymbols = new string[]{
            "USDCHF",
            "GBPUSD",
            "EURUSD",
            "USDJPY",
            "USDCAD",
            "AUDUSD",
            "EURAUD",
            "EURCHF",
            "EURJPY",
            "GBPCHF",
            "CADJPY",
            "GBPJPY",
            "AUDNZD",
            "AUDCAD",
            "AUDCHF",
            "AUDJPY",
            "CHFJPY",
            "EURNZD",
            "EURCAD",
            "CADCHF",
            "NZDJPY",
            "NZDUSD",
            "GOLD"
            };

        private Dictionary<string, List<double>> dataSource = new Dictionary<string,List<double>>();
        
        
        public Form1()
        {
            InitializeComponent();
            foreach (var symbol in allSymbols)
            {
                dataSource.Add(symbol, new List<double>());
            }
            TheContainer.TheForm = this;
            this.OnTick += Form1_OnTick;
        }

        void Form1_OnTick(Form1.Tick tick)
        {
            if (this.InvokeRequired)
            {
                // Execute the same method, but this time on the GUI thread
                this.BeginInvoke(new TickReceived(Form1_OnTick), tick);

                // we return immedeately
                return;
            }

            if ("USDCHF" == tick.symbol) {
                USDCHFbid.Text = (tick.bid.ToString());
                USDCHFask.Text = (tick.ask.ToString());
                USDCHFgbox.Text = "USDCHF [" + (tick.date.ToString()) + "]";
            }
            if ("GBPUSD" == tick.symbol) {
                GBPUSDbid.Text = (tick.bid.ToString());
                GBPUSDask.Text = (tick.ask.ToString());
                GBPUSDgbox.Text = "GBPUSD [" + (tick.date.ToString()) + "]";
            }
            if ("EURUSD" == tick.symbol) {
                EURUSDbid.Text = (tick.bid.ToString());
                EURUSDask.Text = (tick.ask.ToString());
                EURUSDgbox.Text = "EURUSD [" + (tick.date.ToString()) + "]";
            }
            if ("USDJPY" == tick.symbol) {
                USDJPYbid.Text = (tick.bid.ToString());
                USDJPYask.Text = (tick.ask.ToString());
            }
            if ("USDCAD" == tick.symbol) {
                USDCADbid.Text = (tick.bid.ToString());
                USDCADask.Text = (tick.ask.ToString());
            }
            if ("AUDUSD" == tick.symbol) {
                AUDUSDbid.Text = (tick.bid.ToString());
                AUDUSDask.Text = (tick.ask.ToString());
            }
            if ("EURAUD" == tick.symbol) {
                EURAUDbid.Text = (tick.bid.ToString());
                EURAUDask.Text = (tick.ask.ToString());
            }
            if ("EURCHF" == tick.symbol) {
                EURCHFbid.Text = (tick.bid.ToString());
                EURCHFask.Text = (tick.ask.ToString());
            }
            if ("EURJPY" == tick.symbol) {
                EURJPYbid.Text = (tick.bid.ToString());
                EURJPYask.Text = (tick.ask.ToString());
            }
            if ("GBPCHF" == tick.symbol) {
                GBPCHFbid.Text = (tick.bid.ToString());
                GBPCHFask.Text = (tick.ask.ToString());
            }
            if ("CADJPY" == tick.symbol) {
                CADJPYbid.Text = (tick.bid.ToString());
                CADJPYask.Text = (tick.ask.ToString());
            }
            if ("GBPJPY" == tick.symbol) {
                GBPJPYbid.Text = (tick.bid.ToString());
                GBPJPYask.Text = (tick.ask.ToString());
            }
            if ("AUDNZD" == tick.symbol) {
                AUDNZDbid.Text = (tick.bid.ToString());
                AUDNZDask.Text = (tick.ask.ToString());
            }
            if ("AUDCAD" == tick.symbol) {
                AUDCADbid.Text = (tick.bid.ToString());
                AUDCADask.Text = (tick.ask.ToString());
            }
            if ("AUDCHF" == tick.symbol) {
                AUDCHFbid.Text = (tick.bid.ToString());
                AUDCHFask.Text = (tick.ask.ToString());
            }
            if ("AUDJPY" == tick.symbol) {
                AUDJPYbid.Text = (tick.bid.ToString());
                AUDJPYask.Text = (tick.ask.ToString());
            }
            if ("CHFJPY" == tick.symbol) {
                CHFJPYbid.Text = (tick.bid.ToString());
                CHFJPYask.Text = (tick.ask.ToString());
            }
            if ("EURNZD" == tick.symbol) {
                EURNZDbid.Text = (tick.bid.ToString());
                EURNZDask.Text = (tick.ask.ToString());
            }
            if ("EURCAD" == tick.symbol) {
                EURCADbid.Text = (tick.bid.ToString());
                EURCADask.Text = (tick.ask.ToString());
            }
            if ("CADCHF" == tick.symbol) {
                CADCHFbid.Text = (tick.bid.ToString());
                CADCHFask.Text = (tick.ask.ToString());
            }
            if ("NZDJPY" == tick.symbol) {
                NZDJPYbid.Text = (tick.bid.ToString());
                NZDJPYask.Text = (tick.ask.ToString());
            }
            if ("NZDUSD" == tick.symbol) {
                NZDUSDbid.Text = (tick.bid.ToString());
                NZDUSDask.Text = (tick.ask.ToString());
            }
            if ("GOLD"   == tick.symbol) {
                GOLDbid.Text = (tick.bid.ToString());
                GOLDask.Text = (tick.ask.ToString());
            }
        }

        private void eventLog1_EntryWritten(object sender, System.Diagnostics.EntryWrittenEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitDDEClient();
        }
        
        

        private void OnAdvise(object sender, DdeAdviseEventArgs args)
        {
            string[] fxData = (args.Item + " " + args.Text).Split(new char[] { ' ' });

            Tick tick = null;
            
            try{
                tick = new Tick()
                {
                    //EURUSD 2014/05/06 23:59 1.34455 1.34456
                    symbol = fxData[0],
                    date = DateTime.Parse(fxData[1] + " " + fxData[2]),
                    bid = double.Parse(fxData[3]),
                    ask = double.Parse(fxData[4])
                };   
            }
           
            catch(Exception ex)
            {
                msgLbl.Text = "Invalid tick data - " + ex.Message;
            }
           
            //---------------------------------------------------
            // create data source
            if(tick != null)
            {
                if (dataSource.ContainsKey(tick.symbol))
                {
                    dataSource[tick.symbol].Insert(0, tick.bid);
                 //   Log(tick.bid);
                }
                else
                {
                   dataSource.Add(tick.symbol, new List<double>(){ tick.bid });
                }
                if (OnTick != null)
                {
                    OnTick.Invoke(tick);
                }
            }

        }

        private void OnDisconnected(object sender, DdeDisconnectedEventArgs args)
        {
            msgLbl.Text = (
                "OnDisconnected: " +
                "IsServerInitiated=" + args.IsServerInitiated.ToString() + " " +
                "IsDisposed=" + args.IsDisposed.ToString());
            TheContainer.TheDdeClient = null;
        }

        private void btnStartDDE_Click(object sender, EventArgs e)
        {
            if (!TheContainer.TheDdeClient.IsConnected)
            {
                ConnectToDDE();
                SubscribeAll();
            }
            else
            {
                msgLbl.Text = ("Already connected");

            }
        }

        private void InitDDEClient()
        {
            DdeClient client = new DdeClient("MT4", "QUOTE");
            TheContainer.TheDdeClient = client;

            // Subscribe to the Disconnected event.  This event will notify 
            // the application when a conversation has been terminated.
            client.Disconnected += OnDisconnected;
            client.Advise += OnAdvise;
        }

        private void ConnectToDDE()
        {
            if (!TheContainer.TheDdeClient.IsConnected)
            {
                try
                {
                    TheContainer.TheDdeClient.Connect();
                    msgLbl.Text = ("DDE Client Started");

                }
                catch (Exception)
                {
                    msgLbl.Text = ("An exception was thrown during DDE connection" + 
                        "Ensure Metatrader 4 is running and DDE is enabled" +
                        "To activate the DDE Server go to Tools -> Options" +
                        "On the Server tab, ensure \"Enable DDE server\" is checked");
                }
            }
            else
            {
                msgLbl.Text = ("Already connected");
                
            }
        }

        private void SubscribeAll()
        {
            if (TheContainer.TheDdeClient != null && TheContainer.TheDdeClient.IsConnected)
            {
                var client = TheContainer.TheDdeClient;
                foreach (var symbol in allSymbols)
                {
                    client.StartAdvise(symbol, 1, true, 60000);
                }
              
            }
        }

        public class Tick
        {
            public string symbol;
            public double bid;
            public double ask;
            public DateTime date;
        }
    }

    // Not sure what the purpose of this..Singleton ?
    // these static variables could be on Form1
   
    public class TheContainer
    {
        public static Form1 TheForm;
        public static DdeClient TheDdeClient;
    }

}
