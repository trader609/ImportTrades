using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportTrades
{
    class TSGenerator
    {
        /// <summary>
        /// AggregationPeriod
        /// </summary>
        private enum AggregationPeriod
        {
            TWO_MIN = 2,
            FIVE_MIN = 5
        };
        private string[] colors = new string[]
        {
            "Color.CYAN",
            "Color.PINK",
            "Color.LIGHT_ORANGE"
        };
        private List<string> _tsScriptLines = new List<string>();

        /// <summary>
        ///  Generates thinkscript for a single closed position
        /// </summary>
        /// <param name="closedPosition"></param>
        /// <returns></returns>
        public List<string> GenerateTS(ClosedPosition closedPosition, int id)
        {
            List<string> tsScriptLines = new List<string>();
            try
            {
                var buys = closedPosition.TradeType == TradeType.LONG ? closedPosition.Entries : closedPosition.Exits;
                var sells = closedPosition.TradeType == TradeType.LONG ? closedPosition.Exits : closedPosition.Entries;
                string correctSymbolStr = string.Format("correctSymbol{0}{1}", closedPosition.Symbol, id);
                tsScriptLines.Add("#################################################################");
                tsScriptLines.Add(string.Format("def {0} = (GetSymbol() == \"{1}\");",correctSymbolStr, closedPosition.Symbol));

                tsScriptLines.AddRange(GeneratePlotTs(id, correctSymbolStr, buys));
                tsScriptLines.AddRange(GeneratePlotTs(id, correctSymbolStr, sells));
                

                return tsScriptLines;
                
            }
            catch (System.FormatException formatEx)
            {
                Console.WriteLine(formatEx.Message);
            }

            return tsScriptLines;
        }

        /// <summary>
        /// Generates plots for the trades
        /// </summary>
        /// <param name="id"></param>
        /// <param name="correctSymbolStr"></param>
        /// <param name="trades"></param>
        /// <returns></returns>
        private List<string> GeneratePlotTs(int id, string correctSymbolStr, List<Trade> trades)
        {
            string plotColor = colors[colors.Length % (id)];
            List<string> tsScript = new List<string>();

            for (var i = 1; i <= trades.Count; ++i)
            {
                var buy = trades[i - 1];
                string plotName = string.Format("{0}Plot{1}{2}_{3}", buy.BuyOrSell == BuyOrSell.BOT ? "buy" : "sell", buy.Symbol, id, i);
                string tradeDateCondition = string.Format("GetYYYYMMDD() == {0}", buy.TradeDateTime.ToString("yyyyMMdd"));
                string tradeTimeCondition5Min = string.Format("SecondsTillTime({0}) == 0", GetRoundedTime(buy.TradeDateTime, AggregationPeriod.FIVE_MIN));
                string tradeTimeCondition2Min = string.Format("SecondsTillTime({0}) == 0", GetRoundedTime(buy.TradeDateTime, AggregationPeriod.TWO_MIN));

                tsScript.Add(string.Format("plot {0};", plotName));
                tsScript.Add(string.Format("{0}.SetPaintingStrategy({1});", plotName, buy.BuyOrSell == BuyOrSell.BOT ? "PaintingStrategy.ARROW_UP" : "PaintingStrategy.ARROW_DOWN"));
                tsScript.Add(string.Format("{0}.SetDefaultColor({1});", plotName, plotColor));
                tsScript.Add(string.Format("{0}.SetLineWeight(3);", plotName));
                tsScript.Add(string.Format("if Is{0}MinChart", 5));
                tsScript.Add("{");
                tsScript.Add(string.Format("\t{0} = if {1} and {2} and {3} then {4} else Double.NaN;", plotName, correctSymbolStr, tradeDateCondition, tradeTimeCondition5Min, buy.Price));
                tsScript.Add("}");
                tsScript.Add(string.Format("else if Is{0}MinChart", 2));
                tsScript.Add("{");
                tsScript.Add(string.Format("\t{0} = if {1} and {2} and {3} then {4} else Double.NaN;", plotName, correctSymbolStr, tradeDateCondition, tradeTimeCondition2Min, buy.Price));
                tsScript.Add("}");
                tsScript.Add("else");
                tsScript.Add("{");
                tsScript.Add(string.Format("\t{0} = Double.NaN;", plotName));
                tsScript.Add("}");
                tsScript.Add("\n");
            }
            return tsScript;
        }

        /// <summary>
        /// Gets the nearest 2 min or 5 min time for the trade
        /// and also adds 3 hours to the time to convert to EST
        /// </summary>
        /// <param name="tradeDateTime"></param>
        /// <param name="chartPeriod"></param>
        /// <returns></returns>
        private static string GetRoundedTime(DateTime tradeDateTime, AggregationPeriod chartPeriod)
        {
            var dateTime = new DateTime(tradeDateTime.Year, tradeDateTime.Month, tradeDateTime.Day, tradeDateTime.Hour, tradeDateTime.Minute - tradeDateTime.Minute % (int)chartPeriod, 0);
            dateTime = dateTime.AddHours(3.0f); /// Convert to EST
            return dateTime.ToString("hhmm");
        }

        /// <summary>
        /// Generates the Ts script for all the closed positions in the list.
        /// It starts by seggregating all the closed positions by date and symbol and then generating ts
        /// for each closed position in the seggregated group.
        /// 
        /// </summary>
        /// <param name="closedPositions"></param>
        /// <returns></returns>
        internal string[] GenerateTS(IEnumerable<ClosedPosition> closedPositions)
        {
            //
            // common header
            //
            AddCommonHeader();
           
            // We want to use different colors for marking different trades on the same symbol on the
            // same day. We do this by seggregating the trades by DayOfYear.
            // Then for each day, we seggregate by the symbol.
            // Then we call GenerateTS for each trade for that symbol
            // 
            var closedPositionsByDate = closedPositions.GroupBy(c => c.EntryTime.DayOfYear);
            foreach(var closedPositionsPerDay in closedPositionsByDate)
            {
                var closedPositionsPerDayPerSym = closedPositionsPerDay.GroupBy(c => c.Symbol);
                foreach(var closedPositionsPerDayPerSymGrp in closedPositionsPerDayPerSym)
                {
                    int i = 1;
                    var closedPositionsPerDayPerSymList = closedPositionsPerDayPerSymGrp.ToList();
                    foreach (var closedPosition in closedPositionsPerDayPerSymList)
                    {
                        var tsScriptLines = GenerateTS(closedPosition, i);
                        _tsScriptLines.AddRange(tsScriptLines);
                        ++i;
                    }
                }
            }
            return _tsScriptLines.ToArray();
        }

        /// <summary>
        /// Common TS header
        /// </summary>
        private void AddCommonHeader()
        {
            _tsScriptLines.Add("#");
            _tsScriptLines.Add("# Common Header");
            _tsScriptLines.Add("#");
            _tsScriptLines.Add(
                     "def Is5MinChart = (GetAggregationPeriod() == AggregationPeriod.FIVE_MIN);");
            _tsScriptLines.Add(
                    "def Is2MinChart = (GetAggregationPeriod() == AggregationPeriod.TWO_MIN);");
            _tsScriptLines.Add(
                    "");
        }
    }
}
