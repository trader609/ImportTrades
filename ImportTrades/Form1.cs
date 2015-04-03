using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportTrades
{
    public enum TransactionType
    {
        BOT,
        SLD
    }
    public enum TradeStatus
    {
        Open,
        Buy,
        Sell,
        Close
    }
    public struct Transaction
    {
        public DateTime TransactionDateTime;
        public string Symbol;
        public TransactionType BuyOrSell;
        public float NumShares;
        public float Price;
        public float Commission;
    }

    public struct ClosedTrade
    {
        public string Symbol;
        public DateTime FirstBuy;
        public DateTime LastSell;
        public float NumShares;
        public float AvgBuyPricePerShare;
        public float AvgSellPricePerShare;
        public float TotalCommissions;
    }

    public struct Trade
    {
        public float NumShares;
        public TradeStatus Status;
    }
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Get_ExcelSheet();
        }

        public void Get_ExcelSheet()
        {
            var fileName = @"C:\autoimport.xlsx";
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            var closedTrades = ReadTrades(connectionString);
        }

        private static List<ClosedTrade> ReadTrades(string connectionString)
        {
            var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
            var ds = new DataSet();

            adapter.Fill(ds, "Transactions");

            var data = ds.Tables["Transactions"].AsEnumerable();
            //
            // Read all the transactions
            //
            // This is how you combine multiple conditions in linq
            // http://stackoverflow.com/questions/15828/reading-excel-files-from-c-sharp
            var transactions = data.Where(x => !String.IsNullOrEmpty(x.Field<string>("Symbol"))).Select(x =>
                new Transaction
                {
                    TransactionDateTime = x.Field<DateTime>("DateTime"),
                    Symbol = x.Field<string>("Symbol"),
                    BuyOrSell = (TransactionType)Enum.Parse(typeof(TransactionType), x.Field<string>("Action"), true),
                    NumShares = (float)x.Field<double>("Number of Shares"),
                    Price = (float)x.Field<double>("Price"),
                    Commission = (float)x.Field<double>("Commissions"),
                }).GroupBy(t => t.Symbol);

                List<Transaction> Buys = new List<Transaction>();
                List<Transaction> Sells = new List<Transaction>();
                List<ClosedTrade> ClosedTrades = new List<ClosedTrade>();
                int numTransProcessed = 0;
                int numShares = 0;

            foreach (var grp in transactions)
            {
                var symbol = grp.Key;
                var allTrans = grp.OrderBy(t => t.TransactionDateTime);

                Buys.Clear();
                Sells.Clear();
                foreach (var trans in allTrans)
                {
                    switch (trans.BuyOrSell)
                    {
                        case TransactionType.BOT:
                            {
                                Buys.Add(trans);
                                numShares += (int)trans.NumShares;
                                break;
                            }
                        case TransactionType.SLD:
                            {
                                Sells.Add(trans);
                                numShares -= (int)trans.NumShares;
                                break;
                            }

                    }
                    if (numShares == 0)
                    {
                        //Trade closed
                        ClosedTrade trade = new ClosedTrade();
                        trade.Symbol = symbol;
                        trade.FirstBuy = Buys.Min(t => t.TransactionDateTime);
                        trade.LastSell = Sells.Max(t => t.TransactionDateTime);
                        trade.NumShares = Buys.Sum(t => t.NumShares);
                        trade.AvgBuyPricePerShare = Buys.Sum(t => t.NumShares * t.Price) / Buys.Sum(t => t.NumShares);
                        trade.AvgSellPricePerShare = Sells.Sum(t => t.NumShares * t.Price) / Sells.Sum(t => t.NumShares);

                        ClosedTrades.Add(trade);
                        Buys.Clear();
                        Sells.Clear();
                    }

                }



            }
            return ClosedTrades;
        }
        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
