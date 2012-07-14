using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EskomParser
{
    /// <summary>
    /// The Premise details.
    /// </summary>
    public class Premise
    {
        #region Properties

        public string PremiseId { get; set; }
        public decimal AdminChargeRate { get; set; }
        public int AdminPeriod { get; set; }
        public decimal TXNetworkAccessChargeRate { get; set; }
        public double TXkVA { get; set; }
        public decimal NetworkAccessChargeRate { get; set; }
        public double NetworkAccesskVA { get; set; }
        public decimal NetworkDemandChargeRate { get; set; }
        public double NetworkDemandkVA { get; set; }
        public decimal SeasonOffPeakEnergyChargeRate { get; set; }
        public double SeasonOffPeakEnergykWh { get; set; }
        public decimal SeasonPeakEnergyChargeRate { get; set; }
        public double SeasonPeakEnergykWh { get; set; }
        public decimal SeasonStandardEnergyChargeRate { get; set; }
        public double SeasonStandardEnergykWh { get; set; }
        public decimal ElectrificationAndRuralSubsidyRate { get; set; }
        public double ElectrificationAndRuralSubsidykWh { get; set; }
        public decimal RetailEnvironmentLevyChargeRate { get; set; }
        public double RetailEnviromentLevykWh { get; set; }
        public decimal SeasonReactiveEnergyChargeRate { get; set; }
        public double SeasonReactiveEnergykvarh { get; set; }
        public bool IsHighSeason { get; set; }

        public double EnergyConsumptionOffPeak { get; set; }
        public double EnergyConsumptionPeak { get; set; }
        public double EnergyConsumptionStandard { get; set; }
        public double EnergyConsumptionTotal { get; set; }

        public double DemandConsumptionOffPeak { get; set; }
        public double DemandConsumptionPeak { get; set; }
        public double DemandConsumptionStandard { get; set; }
        public double DemandConsumptionReading { get; set; }

        public double ReactiveEnergyOffPeak { get; set; }
        public double ReactiveEnergyPeak { get; set; }
        public double ReactiveEnergyStandard { get; set; }
        public double ExcessReactiveEnergy { get; set; }
        public double LoadFactor { get; set; }

        public double NotifiedMaxDemand { get; set; }
        public double UtilisedCapacity { get; set; }
        public decimal TotalCharges { get; set; }

        public double ExcessNACUsage { get; set; }
        public decimal ExcessNACCharge { get; set; }
        public int ExcessNACEvents { get; set; }
        public double ExcessNACExceeded { get; set; }

        public decimal ServiceCharge { get; set; }

        #endregion
    }
}
