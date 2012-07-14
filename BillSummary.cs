using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EskomParser
{
    public class BillSummary
    {
        #region Properties

        public decimal AdminCharge { get; set; }
        public decimal TransmissionNetworkCharge { get; set; }
        public decimal NetworkAccessCharge { get; set; }
        public decimal NetworkChargeDemand { get; set; }
        public decimal EnergyChargeOffPeak { get; set; }
        public decimal EnergyChargePeak { get; set; }
        public decimal EnergyChargeStandard { get; set; }
        public decimal ReactiveEnergy { get; set; }
        public decimal ElectrificationAndRuralCharge { get; set; }
        public decimal EnvironmentalLevy { get; set; }
        public decimal ExcessNetworkAccessCharge { get; set; }
        public decimal ServiceCharge { get; set; }

        public double EnergyUsageOffPeak { get; set; }
        public double EnergyUsagePeak { get; set; }
        public double EnergyUsageStandard { get; set; }
        public double ReactiveEnergyUsage { get; set; }

        public int Month { get; set; }
        public int Year { get; set; }

        #endregion
    }
}
