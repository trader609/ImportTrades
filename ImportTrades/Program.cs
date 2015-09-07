//
// TODO :
// 1. Make it possible to update the existing trades, by making the buy time and sell time combination as primary key
// 2. Import all the old trades from IB
//
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportTrades
{
    public enum BuyOrSell
    {
        BOT,
        SLD
    }

    public enum TradeType
    {
        LONG,
        SHORT
    }

    public struct Trade
    {
        public DateTime TradeDateTime;
        public string Symbol;
        public BuyOrSell BuyOrSell;
        public float NumShares;
        public float Price;
        public float Commission;
    }

    public struct ClosedTrade
    {
        public string Symbol;
        public TradeType TradeType;
        public DateTime EntryTime;
        public DateTime ExitTime;
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
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Input$]", connectionString);
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
                    BuyOrSell = (BuyOrSell)Enum.Parse(typeof(BuyOrSell), x.Field<string>("Action"), true),
                    NumShares = (float)x.Field<double>("Number of Shares"),
                    Price = (float)x.Field<double>("Price"),
                    Commission = (float)x.Field<double>("Commissions"),
                }).GroupBy(t => t.Symbol);

            List<Trade> Buys = new List<Trade>();
            List<Trade> Sells = new List<Trade>();
            List<ClosedTrade> ClosedTrades = new List<ClosedTrade>();
            int numShares = 0;
            DateTime Jan_1_2005 = Convert.ToDateTime("1-1-2005");
            

            foreach (var grp in transactions)
            {
                var symbol = grp.Key;
                var grpArr = grp.ToArray();
                for (var i = 0; i < grpArr.Count();++i)
                {
                    if (grpArr[i].TradeDateTime < Jan_1_2005)
                    {
                        var tradeTime = grpArr[i].TradeDateTime.ToShortTimeString();
                        grpArr[i].TradeDateTime = DateTime.Parse( DateTime.Now.ToShortDateString() + " " + grpArr[i].TradeDateTime.ToShortTimeString());
                    }
                }
                //foreach (var trans in grp)
                //{
                //    if (trans.TradeDateTime.Date < Convert.ToDateTime("1-1-2005")) ;
                //    {
                //       // trans.TradeDateTime = Convert.ToDateTime(DateTime.Now.Date);
                //    }
                //}
                var allTrans = grpArr.OrderBy(t => t.TradeDateTime);

                Buys.Clear();
                Sells.Clear();
                TradeType tradeType = TradeType.LONG;
                foreach (var trans in allTrans)
                {
                    switch (trans.BuyOrSell)
                    {
                        case BuyOrSell.BOT:
                            {
                                Buys.Add(trans);
                                numShares += (int)trans.NumShares;
                                break;
                            }
                        case BuyOrSell.SLD:
                            {
                                Sells.Add(trans);
                                numShares -= (int)trans.NumShares;
                                break;
                            }
                    }
                    if (numShares < 0)
                    {
                        tradeType = TradeType.SHORT;
                    }
                    else if (numShares > 0)
                    {
                        tradeType = TradeType.LONG;
                    }
                    //
                    // When sells cancel out the buys, we closed a trade
                    // 
                    else if (numShares == 0)
                    {
                        //Trade closed
                        ClosedTrade trade = new ClosedTrade();
                        trade.Symbol = symbol;
                        trade.TradeType = tradeType;
                        if (tradeType == TradeType.LONG)
                        {
                            trade.EntryTime = Buys.Min(t => t.TradeDateTime);
                            trade.ExitTime = Sells.Max(t => t.TradeDateTime);
                        }
                        else if (tradeType == TradeType.SHORT)
                        {
                            trade.EntryTime = Sells.Min(t => t.TradeDateTime);
                            trade.ExitTime = Buys.Max(t => t.TradeDateTime);
                        }
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
            ClosedTrades = ClosedTrades.OrderBy(c => c.EntryTime).ToList();
            return ClosedTrades;
        }
        static void Main(string[] args)
        {
            try
            {
                Program p = new Program();
                p.Init(args[0]);

                var trades = p.ReadTrades();
                p.WriteTrades(trades);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        private void WriteTrades(List<ClosedTrade> trades)
        {
            // Get Existing data into the table
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [ClosedTradesSheet$]", connectionString);
            DataSet ds = new DataSet();
            adapter.Fill(ds, "ClosedTrades");
            
            // prepare the Insert Command
            adapter.InsertCommand = new OleDbCommand("INSERT INTO [ClosedTradesSheet$] ([Symbol], [Long/Short], [Entry DateTime], [Exit DateTime], [Num Shares], [Buy Price/Share], [Sell Price/Share], Commissions) VALUES (?,?,?,?,?,?,?,?)", oleDbConnection);
            adapter.InsertCommand.Parameters.Add("@Symbol", OleDbType.VarChar, 255, "Symbol");
            adapter.InsertCommand.Parameters.Add("@Long/Short", OleDbType.VarChar, 255, "Long/Short");
            adapter.InsertCommand.Parameters.Add("@Entry DateTime", OleDbType.DBTimeStamp, 255, "Entry DateTime");
            adapter.InsertCommand.Parameters.Add("@Exit DateTime", OleDbType.DBTimeStamp, 255, "Exit DateTime");
            adapter.InsertCommand.Parameters.Add("@NumShares", OleDbType.Integer, 4,"Num Shares");
            adapter.InsertCommand.Parameters.Add("@Buy Price/Share", OleDbType.Single, 4, "Buy Price/Share");
            adapter.InsertCommand.Parameters.Add("@Sell Price/Share", OleDbType.Single, 4,"Sell Price/Share");
            adapter.InsertCommand.Parameters.Add("@Commissions", OleDbType.Single, 4, "Commissions");
            //adapter.InsertCommand.Parameters.Add("@Total", OleDbType.Single, 4, "Total");
            //adapter.InsertCommand.Parameters.Add("@Total Commissions", OleDbType.Single, 4, "Total Commissions");
            //adapter.InsertCommand.Parameters.Add("@Total With Commissions", OleDbType.Single, 4, "Total With Commissions");
            //adapter.InsertCommand.Parameters.Add("@Label", OleDbType.VarChar, 255, "Label");

            DateTime previousTrade = trades.First().EntryTime;
            foreach (ClosedTrade trade in trades)
            {
                DataRow newTradeRow = ds.Tables["ClosedTrades"].NewRow();
                newTradeRow["Symbol"] = trade.Symbol;
                newTradeRow["Long/Short"] = (trade.TradeType == TradeType.LONG)?"Long":"Short";
                newTradeRow["Entry DateTime"] = trade.EntryTime;
                newTradeRow["Exit DateTime"] = trade.ExitTime;
                newTradeRow["Num Shares"] = trade.NumShares;
                newTradeRow["Buy Price/Share"] = trade.AvgBuyPricePerShare;
                newTradeRow["Sell Price/Share"] = trade.AvgSellPricePerShare;
                newTradeRow["Commissions"] = trade.TotalCommissions;

                ds.Tables["ClosedTrades"].Rows.Add(newTradeRow);
            }
            adapter.Update(ds, "ClosedTrades");

            //// prepare the Update Command
            //adapter.UpdateCommand = new OleDbCommand("UPDATE [ClosedTradesSheet$] SET Total = @total, [Total Commissions] = @totalCommissions, [Total With Commissions] = @totalWithCommissions, Label = @label WHERE [Entry DateTime] = @buyDateTime AND [Exit DateTime] = @sellDateTime", oleDbConnection);
            //adapter.UpdateCommand.Parameters.AddWithValue("@total", 10.0f);
            //adapter.UpdateCommand.Parameters.AddWithValue("@totalCommissions", 10.0f);
            //adapter.UpdateCommand.Parameters.AddWithValue("@totalWithCommissions", 10.0f);
            //adapter.UpdateCommand.Parameters.AddWithValue("@label", "test");
            //adapter.UpdateCommand.Parameters.AddWithValue("@buyDateTime", trades.First().FirstBuy);
            //adapter.UpdateCommand.Parameters.AddWithValue("@sellDateTime", trades.First().LastSell);
            //oleDbConnection.Open();
            //adapter.UpdateCommand.ExecuteNonQuery();
            //oleDbConnection.Close();

            
        }
    }
}
