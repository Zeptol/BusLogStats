using System;

namespace BusDataStats
{
    public class BusData
    {
        public string LineCode { get; set; }
        public Direction Direction { get; set; }
        public int StationNum { get; set; }
        public string StationCode { get; set; }
        public string VehCode { get; set; }
        public DateTime EstimatedTime { get; set; }
        public DateTime ProcessTime { get; set; }
    }

    public class BusDataStats
    {
        public string LineCode { get; set; }
        public Direction Direction { get; set; }
        public int StationNum { get; set; }
        public string StationCode { get; set; }
        public string VehCode { get; set; }
        public DateTime FirstEstimatedTime { get; set; }
        public DateTime LastEstimatedTime { get; set; }
        public DateTime FirstProcessTime { get; set; }
        public DateTime LastProcessTime { get; set; }

        public double TimeDiffMins
        {
            get
            {
                var res = (LastProcessTime - FirstEstimatedTime).TotalMinutes;
                return -(Math.Abs(res) > 1000 ? res < 0 ? res + 1440 : res > 0 ? res - 1440 : res : res);
                //return - (res < -1000 ? res + 1440 : res);
            }
        }
        public double ActualTotalMins
        {
            get
            {
                var res = (LastProcessTime - FirstProcessTime).TotalMinutes;
                return res < -1000 ? res + 1440 : res;
            }
        }
    }
    public enum Direction
    {
        Up,
        Down
    }
}
