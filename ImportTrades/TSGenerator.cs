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
            List<string> tsScriptLines = null;
            try
            {
                string buyPlotName = string.Format("buyPlot{0}{1}", closedPosition.Symbol, id);
                string sellPlotName = string.Format("sellPlot{0}{1}", closedPosition.Symbol, id);
                string correctSymbolStr = string.Format("correctSymbol{0}{1}", closedPosition.Symbol, id);
                string plotColor = colors[colors.Length % (id)];

                tsScriptLines = new List<string>
                {
                    "##########################################################################",
                    "",
                    "#",
                    string.Format("# {0} trades", closedPosition.Symbol),
                    "#",
                    string.Format("plot {0};", buyPlotName),
                    string.Format("plot {0};", sellPlotName),
                    string.Format("def {0} = (GetSymbol() == \"{1}\");", correctSymbolStr, closedPosition.Symbol),
                    string.Format("{0}.SetPaintingStrategy(PaintingStrategy.ARROW_UP);", buyPlotName),
                    string.Format("{0}.SetDefaultColor({1});", buyPlotName, plotColor),
                    string.Format("{0}.SetLineWeight(2);", buyPlotName),
                    string.Format("{0}.SetPaintingStrategy(PaintingStrategy.ARROW_DOWN);", sellPlotName),
                    string.Format("{0}.SetDefaultColor({1});", sellPlotName, plotColor),
                    string.Format("{0}.SetLineWeight(2);", sellPlotName),
                    "",
                };
                
                // 5 min body
                var tsBody = GenerateTSBody(buyPlotName, sellPlotName, correctSymbolStr, closedPosition, AggregationPeriod.FIVE_MIN);
                tsScriptLines.AddRange(tsBody);

                // 2 min body
                tsBody = GenerateTSBody(buyPlotName, sellPlotName, correctSymbolStr, closedPosition, AggregationPeriod.TWO_MIN);
                tsScriptLines.AddRange(tsBody);

                // final else block
                tsScriptLines.Add("{");
                tsScriptLines.Add(
                    string.Format("\t{0} = Double.NaN;", buyPlotName));// NaN
                tsScriptLines.Add(
                    string.Format("\t{0} = Double.NaN;", sellPlotName));// NaN
                tsScriptLines.Add("}");
            }
            catch (System.FormatException formatEx)
            {
                Console.WriteLine(formatEx.Message);
            }

            return tsScriptLines;
        }

        /// <summary>
        /// Generates TS body for a closed position
        /// </summary>
        /// <param name="buyPlotName"></param>
        /// <param name="sellPlotName"></param>
        /// <param name="correctSymbolStr"></param>
        /// <param name="closedPosition"></param>
        /// <param name="chartPeriod"></param>
        /// <returns></returns>
        private static List<string> GenerateTSBody(
            string buyPlotName,
            string sellPlotName,
            string correctSymbolStr,
            ClosedPosition closedPosition,
            AggregationPeriod chartPeriod)
        {
            List<string> tsScriptBodyLines = new List<string>();

            tsScriptBodyLines.Add(string.Format("# {0} min markers", (int)chartPeriod));
            tsScriptBodyLines.Add(string.Format("if {0} and Is{1}MinChart", correctSymbolStr, (int)chartPeriod));
            tsScriptBodyLines.Add("{");

            List<Trade> buys = null;
            List<Trade> sells = null;
            if (closedPosition.TradeType == TradeType.LONG)
            {
                buys = closedPosition.Entries;
                sells = closedPosition.Exits;
            }
            else if (closedPosition.TradeType == TradeType.SHORT)
            {
                buys = closedPosition.Exits;
                sells = closedPosition.Entries;
            }
            // Script all the buys
            foreach (Trade buy in buys)
            {
                string tradeDateCondition = string.Format("GetYYYYMMDD() == {0}", buy.TradeDateTime.ToString("yyyyMMdd"));
                tsScriptBodyLines.Add(
                    string.Format("\tif {0} and SecondsTillTime({1}) == 0", tradeDateCondition, GetRoundedTime(buy.TradeDateTime, chartPeriod)));// 5 min time
                tsScriptBodyLines.Add("\t{");
                tsScriptBodyLines.Add(
                    string.Format("\t\t{0} = {1};", buyPlotName, buy.Price));// buy price
                tsScriptBodyLines.Add(
                    string.Format("\t\t{0} = Double.NaN;", sellPlotName));// NaN
                tsScriptBodyLines.Add("\t}");
                tsScriptBodyLines.Add("\telse");
            }

            // Script all the sells
            foreach (Trade sell in sells)
            {
                string tradeDateCondition = string.Format("GetYYYYMMDD() == {0}", sell.TradeDateTime.ToString("yyyyMMdd"));
                tsScriptBodyLines.Add(
                    string.Format("\tif {0} and SecondsTillTime({1}) == 0", tradeDateCondition, GetRoundedTime(sell.TradeDateTime, chartPeriod)));// 5 min time
                tsScriptBodyLines.Add("\t{");
                tsScriptBodyLines.Add(
                    string.Format("\t\t{0} = Double.NaN;", buyPlotName));// NaN
                tsScriptBodyLines.Add(
                    string.Format("\t\t{0} = {1};", sellPlotName, sell.Price));// buy price
                tsScriptBodyLines.Add("\t}");
                tsScriptBodyLines.Add("\telse");
            }
            tsScriptBodyLines.Add("\t{");
            tsScriptBodyLines.Add(
                string.Format("\t\t{0} = Double.NaN;", buyPlotName));// NaN
            tsScriptBodyLines.Add(
                string.Format("\t\t{0} = Double.NaN;", sellPlotName));// NaN
            tsScriptBodyLines.Add("\t}");

            tsScriptBodyLines.Add("}");
            tsScriptBodyLines.Add("else");

            return tsScriptBodyLines;
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
