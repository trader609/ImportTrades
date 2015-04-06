using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportTrades
{
    public enum TradeType
    {
        BOT,
        SLD
    }
    
    public struct Trade
    {
        public DateTime TradeDateTime;
        public string Symbol;
        public TradeType BuyOrSell;
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
    class Program
    {
        OleDbConnection oleDbConnection;
        string fileName = @"C:\Users\vbaiyya\Google Drive\stocks\autoImport.xlsx";
        string connectionString;

        private void Init(string fileName)
        {
            connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
            oleDbConnection = new OleDbConnection(connectionString);
            this.fileName = fileName;
        }

        private List<ClosedTrade> ReadTrades()
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet3$]", connectionString);
            DataSet ds = new DataSet();

            adapter.Fill(ds, "Table1");

            var data = ds.Tables["Table1"].AsEnumerable();
            //
            // Read all the transactions
            //
            // This is how you combine multiple conditions in linq
            // http://stackoverflow.com/questions/15828/reading-excel-files-from-c-sharp
            var transactions = data.Where(x => !String.IsNullOrEmpty(x.Field<string>("Symbol"))).Select(x =>
                new Trade
                {
                    TradeDateTime = x.Field<DateTime>("DateTime"),
                    Symbol = x.Field<string>("Symbol"),
                    BuyOrSell = (TradeType)Enum.Parse(typeof(TradeType), x.Field<string>("Action"), true),
                    NumShares = (float)x.Field<double>("Number of Shares"),
                    Price = (float)x.Field<double>("Price"),
                    Commission = (float)x.Field<double>("Commissions"),
                }).GroupBy(t => t.Symbol);

            List<Trade> Buys = new List<Trade>();
            List<Trade> Sells = new List<Trade>();
            List<ClosedTrade> ClosedTrades = new List<ClosedTrade>();
            int numShares = 0;

            foreach (var grp in transactions)
            {
                var symbol = grp.Key;
                var allTrans = grp.OrderBy(t => t.TradeDateTime);

                Buys.Clear();
                Sells.Clear();
                foreach (var trans in allTrans)
                {
                    switch (trans.BuyOrSell)
                    {
                        case TradeType.BOT:
                            {
                                Buys.Add(trans);
                                numShares += (int)trans.NumShares;
                                break;
                            }
                        case TradeType.SLD:
                            {
                                Sells.Add(trans);
                                numShares -= (int)trans.NumShares;
                                break;
                            }
                    }
                    //
                    // When sells cancel out the buys, we closed a trade
                    // 
                    if (numShares == 0)
                    {
                        //Trade closed
                        ClosedTrade trade = new ClosedTrade();
                        trade.Symbol = symbol;
                        trade.FirstBuy = Buys.Min(t => t.TradeDateTime);
                        trade.LastSell = Sells.Max(t => t.TradeDateTime);
                        trade.NumShares = Buys.Sum(t => t.NumShares);
                        trade.AvgBuyPricePerShare = Buys.Sum(t => t.NumShares * t.Price) / Buys.Sum(t => t.NumShares);
                        trade.AvgSellPricePerShare = Sells.Sum(t => t.NumShares * t.Price) / Sells.Sum(t => t.NumShares);
                        trade.TotalCommissions = Buys.Sum(t => t.Commission) + Sells.Sum(t => t.Commission);

                        ClosedTrades.Add(trade);
                        Buys.Clear();
                        Sells.Clear();
                    }
                }
            }
            ClosedTrades = ClosedTrades.OrderBy(c => c.FirstBuy).ToList();
            return ClosedTrades;
        }
        static void Main(string[] args)
        {
            Program p = new Program();
            p.Init(args[0]);
            
            var trades = p.ReadTrades();
            p.WriteTrades(trades);
            
        }

        private void WriteTrades(List<ClosedTrade> trades)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [ClosedTradesSheet$]", connectionString);
            DataSet ds = new DataSet();

            adapter.Fill(ds, "ClosedTrades");
            var closedTradesTable = ds.Tables["ClosedTrades"];
            List<DataRow> invalidEntries = new List<DataRow>();
            foreach (DataRow row in closedTradesTable.Rows)
            {
                if (string.IsNullOrEmpty(row["Buy DateTime"].ToString()) || string.IsNullOrEmpty(row["Symbol"].ToString()))
                {
                    invalidEntries.Add(row);
                }
            }

            foreach (DataRow row in invalidEntries)
            {
                ds.Tables["ClosedTrades"].Rows.Remove(row);
            }

            adapter.InsertCommand = new OleDbCommand("Insert into [ClosedTradesSheet$] ([Buy DateTime], [Sell DateTime], [Symbol], [Num Shares], [Buy Price/Share], [Sell Price/Share], Commissions) Values (?,?,?,?,?,?,?)", oleDbConnection);
            adapter.InsertCommand.Parameters.Add("@Buy DateTime", OleDbType.DBTimeStamp, 255, "Buy DateTime");
            adapter.InsertCommand.Parameters.Add("@Sell DateTime", OleDbType.DBTimeStamp, 255, "Sell DateTime");
            adapter.InsertCommand.Parameters.Add("@Symbol", OleDbType.VarChar, 255, "Symbol");
            adapter.InsertCommand.Parameters.Add("@NumShares", OleDbType.Integer, 4,"Num Shares");
            adapter.InsertCommand.Parameters.Add("@Buy Price/Share", OleDbType.Single, 4, "Buy Price/Share");
            adapter.InsertCommand.Parameters.Add("@Sell Price/Share", OleDbType.Single, 4,"Sell Price/Share");
            adapter.InsertCommand.Parameters.Add("@Commissions", OleDbType.Single, 4, "Commissions");

            foreach (ClosedTrade trade in trades)
            {
                DataRow newTradeRow = closedTradesTable.NewRow();
                newTradeRow["Buy DateTime"] = trade.FirstBuy;
                newTradeRow["Sell DateTime"] = trade.LastSell;
                newTradeRow["Symbol"] = trade.Symbol;
                newTradeRow["Num Shares"] = trade.NumShares;
                newTradeRow["Buy Price/Share"] = trade.AvgBuyPricePerShare;
                newTradeRow["Sell Price/Share"] = trade.AvgSellPricePerShare;
                newTradeRow["Commissions"] = trade.TotalCommissions;


                closedTradesTable.Rows.Add(newTradeRow);
            }
            adapter.Update(ds, "ClosedTrades");
        }
    }
}
