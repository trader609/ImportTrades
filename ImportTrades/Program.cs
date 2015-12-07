//
// TODO :
// 1. Make it possible to update the existing trades, by making the buy time and sell time combination as primary key
// 2. Import all the old trades from IB
//
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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

    public struct TradeMarker
    {
        public int MarkerId;
        public DateTime TradeDateTime;
        public string Symbol;
        public BuyOrSell BuyOrSell;
        public float Price;
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

    public struct ClosedPosition
    {
        public string Symbol;
        public TradeType TradeType;
        public DateTime EntryTime;
        public DateTime ExitTime;
        public List<Trade> Entries;
        public List<Trade> Exits;
        public float NumShares;
        public float AvgBuyPricePerShare;
        public float AvgSellPricePerShare;
        public float TotalCommissions;
    }
    class Program
    {
        private enum Action
        {
            SUMMARIZE,
            GENERATETS
        }
        /// <summary>
        /// Reads raw trades from the excel sheet and either writes them in the
        /// summary sheet or generates thinkscript to plot the trades as arrows on charts
        /// </summary>
        /// <param name="args">
        /// arg 0 : path to the source excel file
        /// arg 1 : Write closed trades to sheet or generate thinkscript
        ///     1 : write closed trades
        ///     2 : generate thinkscript
        /// </param>
        static void Main(string[] args)
        {
            string excelFilename = args[0];

            try
            {
                Action action = args.Count() > 1 && int.Parse(args[1]) == 2 ? Action.GENERATETS : Action.SUMMARIZE;
                var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", args[0]);

                var trades = ReadTrades(connectionString);
                trades = FixDates(trades);
                var closedPositions = ComputeClosedPositions(trades);

                switch (action)
                {
                    case Action.SUMMARIZE:
                        WriteTrades(closedPositions, connectionString);
                        break;

                    case Action.GENERATETS:
                        TSGenerator ts = new TSGenerator();
                        var script = ts.GenerateTS(closedPositions);
                        string outFile =args.Count() > 2 ? args[2] : @"d:\scratch\test.ts";
                        File.WriteAllLines(outFile, script);
                        break;
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

      
        /// <summary>
        /// Reads raw trades from the excel sheet and returns a list of them.
        /// </summary>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private static IEnumerable<Trade> ReadTrades(string connectionString)
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
                });

                return transactions;

        }

        /// <summary>
        /// Parses each trade and fixes the date.
        /// Trades made today will not have a year associated with them in the excel sheet which results in
        /// .Net parsing the year as 1900.
        /// This code identifies those trades and fixes the date to today.
        /// </summary>
        /// <param name="transactions"></param>
        /// <returns></returns>
        private static IEnumerable<Trade> FixDates(IEnumerable<Trade> transactions)
        {
            DateTime Jan_1_2005 = Convert.ToDateTime("1-1-2005");

            var transactionsList = transactions.ToArray();
            for (int i = 0; i < transactions.Count(); ++i)
            {
                if (transactionsList[i].TradeDateTime < Jan_1_2005)
                {
                    //var tradeTime = transactionsList[i].TradeDateTime.ToShortTimeString();
                    //transactionsList[i].TradeDateTime = DateTime.Parse(DateTime.Now.ToShortDateString() + " " + transactionsList[i].TradeDateTime.ToLongTimeString());
                    transactionsList[i].TradeDateTime = new DateTime(
                        DateTime.Now.Year,
                        DateTime.Now.Month,
                        DateTime.Now.Day,
                        transactionsList[i].TradeDateTime.Hour,
                        transactionsList[i].TradeDateTime.Minute,
                        transactionsList[i].TradeDateTime.Second);
                }
            }

            return transactionsList;
        }

        /// <summary>
        /// Computes closed trades
        /// </summary>
        /// <param name="transactions"></param>
        /// <returns></returns>
        private static IEnumerable<ClosedPosition> ComputeClosedPositions(IEnumerable<Trade> transactions)
        {
            List<Trade> Buys = new List<Trade>();
            List<Trade> Sells = new List<Trade>();
            List<ClosedPosition> ClosedPositions = new List<ClosedPosition>();
            List<TradeMarker> TradeMarkers = new List<TradeMarker>();
            int numShares = 0;
            int tradeId = 0;
            var transactionsBySymbol = transactions.GroupBy(t => t.Symbol);
            

            foreach (var grp in transactionsBySymbol)
            {
                var symbol = grp.Key;
               
               
                var grpOrderedByTime = grp.OrderBy(t => t.TradeDateTime);

                Buys.Clear();
                Sells.Clear();
                TradeType tradeType = TradeType.LONG;
                foreach (var trans in grpOrderedByTime)
                {
                    TradeMarker tMarker = new TradeMarker();
                    tMarker.Symbol = symbol;
                    tMarker.BuyOrSell = trans.BuyOrSell;
                    tMarker.MarkerId = tradeId;
                    tMarker.Price = trans.Price;
                    tMarker.TradeDateTime = trans.TradeDateTime;
                    TradeMarkers.Add(tMarker);
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
                        ClosedPosition trade = new ClosedPosition();
                        trade.Symbol = symbol;
                        trade.TradeType = tradeType;
                        if (tradeType == TradeType.LONG)
                        {
                            trade.EntryTime = Buys.Min(t => t.TradeDateTime);
                            trade.ExitTime = Sells.Max(t => t.TradeDateTime);
                            trade.Entries = new List<Trade>(Buys);
                            trade.Exits = new List<Trade>(Sells);
                        }
                        else if (tradeType == TradeType.SHORT)
                        {
                            trade.EntryTime = Sells.Min(t => t.TradeDateTime);
                            trade.ExitTime = Buys.Max(t => t.TradeDateTime);
                            trade.Entries = new List<Trade>(Sells);
                            trade.Exits = new List<Trade>(Buys);
                        }
                        trade.NumShares = Buys.Sum(t => t.NumShares);
                        trade.AvgBuyPricePerShare = Buys.Sum(t => t.NumShares * t.Price) / Buys.Sum(t => t.NumShares);
                        trade.AvgSellPricePerShare = Sells.Sum(t => t.NumShares * t.Price) / Sells.Sum(t => t.NumShares);
                        trade.TotalCommissions = Buys.Sum(t => t.Commission) + Sells.Sum(t => t.Commission);

                        ClosedPositions.Add(trade);
                        Buys.Clear();
                        Sells.Clear();
                        tradeId++;
                    }
                }
            }
            return ClosedPositions.OrderBy(c => c.EntryTime);
        }
       
        /// <summary>
        /// Writes the closed trades into the excel sheet.
        /// </summary>
        /// <param name="closedPositions"></param>
        /// <param name="connectionString"></param>
        private static void WriteTrades(IEnumerable<ClosedPosition> closedPositions, string connectionString)
        {
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
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

            DateTime previousTrade = closedPositions.First().EntryTime;
            foreach (ClosedPosition trade in closedPositions)
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
