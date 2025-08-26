namespace BatchAndReport.Models
{
    public class StrategyModels
    {
        public string FiscalYearDesc { get; set; }
        public string TopicNo { get; set; }
        public string Topic { get; set; }
        public List<StrategyDetailModels> SubStrategy { get; set; }
    }
  

    public class StrategyDetailModels
    {
      
        public int StrategyNum { get; set; }
        public string StrategyDesc { get; set; }
    }
    public class StrategyResponse
    {
        public List<StrategyModels> result { get; set; }
        public int responseCode { get; set; }
        public string responseMsg { get; set; }
    }

    public class StrategyDataModels
    {
        public string FiscalYearDesc { get; set; }
        public int StrategyId { get; set; }
        public string TopicNo { get; set; }
        public string Topic { get; set; }
        public int StrategyDetailId { get; set; }
        public int StrategyNum { get; set; }
        public string StrategyDesc { get; set; }
    }
}
