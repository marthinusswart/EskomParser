using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using org.pdfbox.pdmodel;
using org.pdfbox.util;
using TextBox = System.Windows.Forms.TextBox;

namespace EskomParser
{
    public class PDF2XLSParser
    {
        #region Delegates
        public delegate void ProgressDelegate(string message);
        public delegate string GetSaveFileDelegate();
        #endregion

        #region Fields
        private Dictionary<string, Premise> _premises = new Dictionary<string, Premise>();
        private BillSummary _billSummary = new BillSummary();
        private ExcelHelper _excel = null;
        private bool _is2009Format = false;
        #endregion

        #region Properties
        public TextBox ProgressOutput { get; set; }
        public string RawText { get; set; }
        public MainForm MainForm { get; set; }
        #endregion

        #region Methods
        public void Parse(string pdfFile, string xlsFile)
        {
            try
            {
                UpdateProgress("Loading the PDF document");
                var doc = PDDocument.load(pdfFile);
                var stripper = new PDFTextStripper();
                UpdateProgress("Stripping the string values");
                string text = stripper.getText(doc);
                RawText = text;
                SaveRawText(text, pdfFile + ".txt");
                UpdateProgress("Loading the summary");
                LoadSummary(RawText);
                UpdateProgress("Loading all premise blocks");
                LoadAllPremises(RawText);
                UpdateProgress("Adding data to Excel");
                AddData(xlsFile);
                UpdateProgress("Done");
            }
            catch (InvalidDataException iex)
            {
                UpdateProgress("An error occured");
                UpdateProgress(iex.Message);
            }
            catch (Exception ex)
            {
                UpdateProgress(string.Format("Unknown error occured.{0}{1}", Environment.NewLine, ex.ToString()));
            }
        }

        private void LoadSummary(string rawText)
        {
            var pos = rawText.IndexOf("ACOUNTMONTHCURENTDUEDATE");

            if (pos == -1)
            {
                pos = rawText.IndexOf("CURRENT DUE DATE");

                if (pos == -1)
                {
                    throw new InvalidDataException("The PDF is in an unknown format");
                }
                else
                {
                    _is2009Format = true;
                }
            }

            GetDate(rawText, pos);
            GetSummaryAdminCharge(rawText);
            GetTransmissionNetworkCharge(rawText);
            GetDistNetworkAccessCharge(rawText);
            GetNetworkChargeDemand(rawText);
            GetNetworkEnergyChargeOffPeak(rawText);
            GetNetworkEnergyChargePeak(rawText);
            GetNetworkEnergyChargeStd(rawText);
            GetElectrificationAndRuralSummary(rawText);
            GetEnvironmentalLevy(rawText);
            GetExcessNetworkCharge(rawText);
            GetServiceCharge(rawText);
            GetReactiveEnergySummary(rawText);
        }

        private void GetServiceCharge(string rawText)
        {
            var searchString = "SERVICE CHARGE";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            substring = substring.Substring(pos + 1);
            pos = substring.IndexOf(".");
            var value = substring.Substring(0, pos + 3).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.ServiceCharge = Decimal.Parse(value, nfi);
        }

        private void GetExcessNetworkCharge(string rawText)
        {
            var searchString = "DX EXCESS NETWORK ACCESS CHARGR";
            var pos = rawText.IndexOf(searchString);

            if (pos > -1)
            {
                var substring = rawText.Substring(pos + searchString.Length);
                pos = substring.IndexOf(".");
                var value = substring.Substring(0, pos + 3).Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                _billSummary.ExcessNetworkAccessCharge = Decimal.Parse(value, nfi);
            }
        }

        private void GetEnvironmentalLevy(string rawText)
        {
            var searchString = "ENVIRONMENTAL LEVY";
            var pos = rawText.IndexOf(searchString);

            if (pos > -1)
            {
                var substring = rawText.Substring(pos + searchString.Length);
                pos = substring.IndexOf("R");
                substring = substring.Substring(pos + 1);
                pos = substring.IndexOf(".");
                var value = substring.Substring(0, pos + 3).Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                _billSummary.EnvironmentalLevy = Decimal.Parse(value, nfi);
            }
        }

        private void GetElectrificationAndRuralSummary(string rawText)
        {
            var searchString = "ELECTRIFICATION AND RURAL SUBS (ALL)";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            substring = substring.Substring(pos + 1);
            pos = substring.IndexOf(".");
            var value = substring.Substring(0, pos + 3).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.ElectrificationAndRuralCharge = Decimal.Parse(value, nfi);
        }

        private void GetReactiveEnergySummary(string rawText)
        {
            var searchString = "REACTIVE ENERGY";
            var pos = rawText.IndexOf(searchString);

            if (pos < 2000)
            {
                var substring = rawText.Substring(pos + searchString.Length);
                pos = substring.IndexOf("R");
                var value = substring.Substring(0, pos).Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                _billSummary.ReactiveEnergyUsage = Double.Parse(value, nfi);

                searchString = "R";
                pos = substring.IndexOf(searchString);
                substring = substring.Substring(pos + searchString.Length);
                pos = substring.IndexOf(".");
                value = substring.Substring(0, pos + 3).Trim();

                nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                _billSummary.ReactiveEnergy = Decimal.Parse(value, nfi);
            }
        }

        private void GetNetworkEnergyChargeStd(string rawText)
        {
            var searchString = "ENERGY CHARGE (STD)";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            var value = substring.Substring(0, pos).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyUsageStandard = Double.Parse(value, nfi);

            searchString = "R";
            pos = substring.IndexOf(searchString);
            substring = substring.Substring(pos + searchString.Length);
            pos = substring.IndexOf(".");
            value = substring.Substring(0, pos + 3).Trim();

            nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyChargeStandard = Decimal.Parse(value, nfi);
        }

        private void GetNetworkEnergyChargePeak(string rawText)
        {
            var searchString = "ENERGY CHARGE (PEAK)";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            var value = substring.Substring(0, pos).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyUsagePeak = Double.Parse(value, nfi);

            searchString = "R";
            pos = substring.IndexOf(searchString);
            substring = substring.Substring(pos + searchString.Length);
            pos = substring.IndexOf(".");
            value = substring.Substring(0, pos + 3).Trim();

            nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyChargePeak = Decimal.Parse(value, nfi);
        }

        private void GetNetworkEnergyChargeOffPeak(string rawText)
        {
            var searchString = "ENERGY CHARGE (OFF)";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            var value = substring.Substring(0, pos).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyUsageOffPeak = Double.Parse(value, nfi);

            searchString = "R";
            pos = substring.IndexOf(searchString);
            substring = substring.Substring(pos + searchString.Length);
            pos = substring.IndexOf(".");
            value = substring.Substring(0, pos + 3).Trim();

            nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.EnergyChargeOffPeak = Decimal.Parse(value, nfi);
        }

        private void GetNetworkChargeDemand(string rawText)
        {
            var searchString = "NETWORK CHARGE DEMAND";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            substring = substring.Substring(pos + 1);
            pos = substring.IndexOf(".");
            var value = substring.Substring(0, pos + 3).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.NetworkChargeDemand = Decimal.Parse(value, nfi);
        }

        private void GetSummaryAdminCharge(string rawText)
        {
            var searchString = "ADMINISTRATION CHARGE";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            substring = substring.Substring(pos + 1);
            pos = substring.IndexOf(".");
            var value = substring.Substring(0, pos + 3).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.AdminCharge = Decimal.Parse(value, nfi);
        }

        private void GetTransmissionNetworkCharge(string rawText)
        {
            var searchString = "TRANSMISSION NETWORK CHARGE";
            var pos = rawText.IndexOf(searchString);

            if (pos > -1)
            {
                var substring = rawText.Substring(pos + searchString.Length);
                pos = substring.IndexOf("R");
                substring = substring.Substring(pos + 1);
                pos = substring.IndexOf(".");
                var value = substring.Substring(0, pos + 3).Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                _billSummary.TransmissionNetworkCharge = Decimal.Parse(value, nfi);
            }
        }

        private void GetDistNetworkAccessCharge(string rawText)
        {
            var searchString = "DIST. NETWORK ACCESS CHARGE";
            var pos = rawText.IndexOf(searchString);
            var substring = rawText.Substring(pos + searchString.Length);
            pos = substring.IndexOf("R");
            substring = substring.Substring(pos + 1);
            pos = substring.IndexOf(".");
            var value = substring.Substring(0, pos + 3).Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            _billSummary.NetworkAccessCharge = Decimal.Parse(value, nfi);
        }

        private void GetDate(string rawText, int pos)
        {
            if (!_is2009Format)
            {
                var substring = rawText.Substring(pos);
                pos = substring.IndexOf("DIRECTDEPOSITDETAILBANK");
                substring = substring.Substring(0, pos);
                var lines = substring.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                var dateline = lines[4];
                pos = dateline.IndexOf(" ");
                _billSummary.Month = GetMonth(dateline.Substring(0, pos));
                _billSummary.Year = Int32.Parse(dateline.Substring(pos + 1, 4));
            }
            else
            {
                var substring = rawText.Substring(pos);
                pos = substring.IndexOf("STILFONTEIN");
                substring = substring.Substring(pos);
                var lines = substring.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                var dateline = lines[5];
                pos = dateline.IndexOf(" ");
                _billSummary.Month = GetMonth(dateline.Substring(0, pos));
                _billSummary.Year = Int32.Parse(dateline.Substring(pos + 1, 4));
            }
        }

        private int GetMonth(string month)
        {
            int value = 1;

            switch (month)
            {
                case "JANUARY":
                    value = 1;
                    break;
                case "FEBRUARY":
                    value = 2;
                    break;
                case "MARCH":
                    value = 3;
                    break;
                case "APRIL":
                    value = 4;
                    break;
                case "MAY":
                    value = 5;
                    break;
                case "JUNE":
                    value = 6;
                    break;
                case "JULY":
                    value = 7;
                    break;
                case "AUGUST":
                    value = 8;
                    break;
                case "SEPTEMBER":
                    value = 9;
                    break;
                case "OCTOBER":
                    value = 10;
                    break;
                case "NOVEMBER":
                    value = 11;
                    break;
                case "DECEMBER":
                    value = 12;
                    break;
            }

            return value;
        }

        private void AddData(string xlsFile)
        {
            _excel = GetExcelHelper(xlsFile);
            SetWorksheet(_excel, 1);

            _excel.SetWorkingColumn(_billSummary.Month, _billSummary.Year);
            AddSummary(_excel);
            AddPremiseData(_excel);
            string newExcelFile = GetSaveFile();

            if (newExcelFile.Equals(xlsFile))
            {
                _excel.Save();
            }
            else
            {
                _excel.SaveAs(newExcelFile);
            }
            _excel.Close();
        }

        private void AddPremiseData(ExcelHelper excel)
        {
            foreach (var premise in _premises.Values)
            {
                AddPremiseDetail(premise, excel);
            }
        }

        private void AddSummary(ExcelHelper excel)
        {
            int startRow = 1;

            UpdateProgress(string.Format("Adding summary data to Excel"));

            excel.SetDecimal(_billSummary.AdminCharge, startRow + 1, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.TransmissionNetworkCharge, startRow + 2, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.NetworkAccessCharge, startRow + 3, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.NetworkChargeDemand, startRow + 4, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.ExcessNetworkAccessCharge, startRow + 10, excel.WorkingColumn);
            excel.SetDouble(_billSummary.EnergyUsageOffPeak, startRow + 17, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.EnergyChargeOffPeak, startRow + 18, excel.WorkingColumn);
            excel.SetDouble(_billSummary.EnergyUsageStandard, startRow + 13, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.EnergyChargeStandard, startRow + 14, excel.WorkingColumn);
            excel.SetDouble(_billSummary.EnergyUsagePeak, startRow + 15, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.EnergyChargePeak, startRow + 16, excel.WorkingColumn);
            excel.SetDouble(_billSummary.ReactiveEnergyUsage, startRow + 20, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.ReactiveEnergy, startRow + 21, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.ElectrificationAndRuralCharge, startRow + 23, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.EnvironmentalLevy, startRow + 25, excel.WorkingColumn);
            excel.SetDecimal(_billSummary.ServiceCharge, startRow + 26, excel.WorkingColumn);
        }

        private void AddPremiseDetail(Premise premise, ExcelHelper excel)
        {
            UpdateProgress(string.Format("Adding premise {0} data to Excel", premise.PremiseId));
            int startRow = excel.RowOf(premise.PremiseId, 2);

            if (startRow != -1)
            {
                excel.SetDouble(premise.NotifiedMaxDemand, startRow + 1, excel.WorkingColumn);
                excel.SetDouble(premise.UtilisedCapacity, startRow + 2, excel.WorkingColumn);

                excel.SetDouble(premise.EnergyConsumptionOffPeak, startRow + 3, excel.WorkingColumn);
                excel.SetDouble(premise.EnergyConsumptionStandard, startRow + 4, excel.WorkingColumn);
                excel.SetDouble(premise.EnergyConsumptionPeak, startRow + 5, excel.WorkingColumn);
                excel.SetDouble(premise.EnergyConsumptionTotal, startRow + 6, excel.WorkingColumn);

                excel.SetDouble(premise.DemandConsumptionOffPeak, startRow + 7, excel.WorkingColumn);
                excel.SetDouble(premise.DemandConsumptionStandard, startRow + 8, excel.WorkingColumn);
                excel.SetDouble(premise.DemandConsumptionPeak, startRow + 9, excel.WorkingColumn);
                excel.SetDouble(premise.DemandConsumptionReading, startRow + 10, excel.WorkingColumn);

                excel.SetDouble(premise.ReactiveEnergyOffPeak, startRow + 11, excel.WorkingColumn);
                excel.SetDouble(premise.ReactiveEnergyStandard, startRow + 12, excel.WorkingColumn);
                excel.SetDouble(premise.ReactiveEnergyPeak, startRow + 13, excel.WorkingColumn);
                excel.SetDouble(premise.ExcessReactiveEnergy, startRow + 14, excel.WorkingColumn);

                excel.SetDouble(premise.LoadFactor, startRow + 15, excel.WorkingColumn);
                excel.SetInteger(premise.AdminPeriod, startRow + 16, excel.WorkingColumn);

                excel.SetDecimal(premise.AdminChargeRate, startRow + 18, excel.WorkingColumn);
                excel.SetDecimal(premise.TXNetworkAccessChargeRate, startRow + 20, excel.WorkingColumn);
                excel.SetDecimal(premise.NetworkAccessChargeRate, startRow + 22, excel.WorkingColumn);
                excel.SetDecimal(premise.NetworkDemandChargeRate, startRow + 24, excel.WorkingColumn);
                excel.SetInteger(premise.ExcessNACEvents, startRow + 26, excel.WorkingColumn);
                excel.SetDouble(premise.ExcessNACExceeded, startRow + 27, excel.WorkingColumn);
                excel.SetDecimal(premise.ExcessNACCharge, startRow + 28, excel.WorkingColumn);
                //excel.SetDouble(premise.ExcessNACUsage, startRow + 29, excel.WorkingColumn);

                if (premise.IsHighSeason)
                {
                    excel.SetDecimal(premise.SeasonOffPeakEnergyChargeRate, startRow + 32, excel.WorkingColumn);
                    excel.SetDecimal(premise.SeasonPeakEnergyChargeRate, startRow + 36, excel.WorkingColumn);
                    excel.SetDecimal(premise.SeasonStandardEnergyChargeRate, startRow + 40, excel.WorkingColumn);

                    excel.SetDecimal(premise.SeasonReactiveEnergyChargeRate, startRow + 43, excel.WorkingColumn);
                    excel.SetDouble(premise.SeasonReactiveEnergykvarh, startRow + 42, excel.WorkingColumn);
                }
                else
                {
                    excel.SetDecimal(premise.SeasonOffPeakEnergyChargeRate, startRow + 30, excel.WorkingColumn);
                    excel.SetDecimal(premise.SeasonPeakEnergyChargeRate, startRow + 34, excel.WorkingColumn);
                    excel.SetDecimal(premise.SeasonStandardEnergyChargeRate, startRow + 38, excel.WorkingColumn);
                }

                excel.SetDecimal(premise.ElectrificationAndRuralSubsidyRate, startRow + 45, excel.WorkingColumn);
                excel.SetDecimal(premise.RetailEnvironmentLevyChargeRate, startRow + 47, excel.WorkingColumn);
                excel.SetDecimal(premise.ServiceCharge, startRow + 49, excel.WorkingColumn);
                excel.SetDecimal(premise.TotalCharges, startRow + 51, excel.WorkingColumn);
            }
            else
            {
                UpdateProgress(string.Format("The premise {0} does not exist in the Excel spreadsheet", premise.PremiseId));
            }
        }

        private void SetWorksheet(ExcelHelper excel, int index)
        {
            excel.SetWorksheet(index);
        }

        private ExcelHelper GetExcelHelper(string xlsFile)
        {
            ExcelHelper excel = new ExcelHelper();
            excel.Initialize(false);
            excel.Load(xlsFile);
            return excel;
        }

        private void LoadAllPremises(string rawText)
        {
            var premiseBlocks = new List<StringBuilder>();
            var lines = rawText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            bool isInPremiseBlock = false;
            var premiseBlock = new StringBuilder();
            var lineCount = 0;

            foreach (var line in lines)
            {
                if (line.StartsWith("YOURACOUNTNOBILINGDATE"))
                {
                    isInPremiseBlock = true;
                    premiseBlock = new StringBuilder();
                    premiseBlock.AppendLine(line);
                }
                else if (line.StartsWith("YOUR ACCOUNT NO"))
                {
                    if (lines[lineCount + 1].StartsWith("BILLING DATE"))
                    {
                        isInPremiseBlock = true;
                        premiseBlock = new StringBuilder();
                        premiseBlock.AppendLine(line);
                    }
                }
                else if ((line.Contains("CENTRAL REGION")) && (isInPremiseBlock))
                {
                    if (!_is2009Format)
                    {
                        isInPremiseBlock = false;
                        premiseBlock.AppendLine(line);
                        var block = premiseBlock.ToString();

                        if ((block.IndexOf("REBILLED ADJUSTMENTS") == -1) &&
                            (block.IndexOf("CORRECTIONS") == -1))
                        {
                            premiseBlocks.Add(premiseBlock);
                        }
                    }
                    else
                    {
                        premiseBlock.AppendLine(line);
                    }

                }
                else if ((line.Contains("TOTAL CHARGES")) && (isInPremiseBlock))
                {
                    if (_is2009Format)
                    {
                        isInPremiseBlock = false;
                        premiseBlock.AppendLine(line);
                        premiseBlocks.Add(premiseBlock);
                    }
                    else
                    {
                        premiseBlock.AppendLine(line);
                    }

                }
                else if (isInPremiseBlock)
                {
                    premiseBlock.AppendLine(line);
                }

                lineCount++;
            }

            if (isInPremiseBlock)
            {
                var block = premiseBlock.ToString();
                if ((block.IndexOf("REBILLED ADJUSTMENTS") == -1) &&
                    (block.IndexOf("CORRECTIONS") == -1))
                {
                    premiseBlock.AppendLine("CENTRAL REGION");
                    premiseBlocks.Add(premiseBlock);
                }
            }

            UpdateProgress(string.Format("Total premise blocks loaded: {0}", premiseBlocks.Count));
            int counter = 1;

            foreach (var premiseBlockString in premiseBlocks)
            {
                UpdateProgress(string.Format("Loading Premise {0} details", counter));
                LoadPremise(premiseBlockString);
                counter++;
            }

        }

        private void LoadPremise(StringBuilder premiseBlock)
        {
            var block = premiseBlock.ToString().ToUpper();
            var premise = new Premise();
            premise.PremiseId = GetPremiseId(block);
            premise.AdminChargeRate = GetPremiseAdminChargeRate(block);
            premise.AdminPeriod = GetPremiseAdminChargePeriod(block);
            premise.TXNetworkAccessChargeRate = GetPremiseTXNetworkCharge(block);
            premise.TXkVA = GetPremiseTXkVA(block);
            premise.NetworkAccessChargeRate = GetPremiseNetworkAccessCharge(block);
            premise.NetworkAccesskVA = GetPremiseNetworkAccesskVA(block);
            premise.NetworkDemandChargeRate = GetPremiseNetworkDemandCharge(block);
            premise.NetworkDemandkVA = GetPremiseNetworkDemandkVA(block);
            premise.SeasonOffPeakEnergyChargeRate = GetPremiseSeasonOffPeakCharge(block);
            premise.SeasonOffPeakEnergykWh = GetPremiseSeasonOffPeakkWh(block);
            premise.SeasonPeakEnergyChargeRate = GetPremiseSeasonPeakCharge(block);
            premise.SeasonPeakEnergykWh = GetPremiseSeasonPeakkWh(block);
            premise.SeasonStandardEnergyChargeRate = GetPremiseSeasonStandardCharge(block);
            premise.SeasonStandardEnergykWh = GetPremiseSeasonStandardkWh(block);
            premise.SeasonReactiveEnergyChargeRate = GetPremiseSeasonReactiveCharge(block);
            premise.SeasonReactiveEnergykvarh = GetPremiseSeasonReactivekvarh(block);
            premise.ElectrificationAndRuralSubsidyRate = GetElectrificationAndRuralSubsidyRate(block);
            premise.ElectrificationAndRuralSubsidykWh = GetElectrificationAndRuralSubsidykWh(block);
            premise.RetailEnvironmentLevyChargeRate = GetRetailEnvironmentLevyChargeRate(block);
            premise.RetailEnviromentLevykWh = GetRetailEnviromentLevykWh(block);
            premise.IsHighSeason = GetPremiseIsHighSeason(block);

            premise.EnergyConsumptionOffPeak = GetPremiseEnergyConsumptionOffPeak(block);
            premise.EnergyConsumptionStandard = GetPremiseEnergyConsumptionStandard(block);
            premise.EnergyConsumptionPeak = GetPremiseEnergyConsumptionPeak(block);
            premise.EnergyConsumptionTotal = GetPremiseEnergyConsumptionTotal(block);

            premise.DemandConsumptionOffPeak = GetPremiseDemandConsumptionOffPeak(block);
            premise.DemandConsumptionStandard = GetPremiseDemandConsumptionStandard(block);
            premise.DemandConsumptionPeak = GetPremiseDemandConsumptionPeak(block);
            premise.DemandConsumptionReading = GetPremiseDemandConsumptionReading(block);

            premise.ReactiveEnergyOffPeak = GetPremiseReactiveEnergyOffPeak(block);
            premise.ReactiveEnergyStandard = GetPremiseReactiveEnergyStandard(block);
            premise.ReactiveEnergyPeak = GetPremiseReactiveEnergyPeak(block);
            premise.ExcessReactiveEnergy = GetPremiseExcessReactiveEnergy(block);
            premise.LoadFactor = GetPremiseLoadFactor(block);

            premise.NotifiedMaxDemand = GetPremiseNotifiedMaxDemand(block);
            premise.UtilisedCapacity = GetPremiseUtilisedCapacity(block);
            premise.TotalCharges = GetPremiseTotalCharges(block);

            premise.ExcessNACCharge = GetExcessNACCharge(block);
            premise.ExcessNACUsage = GetExcessNACUsage(block);
            premise.ExcessNACEvents = GetExcessNACEvents(block);
            premise.ExcessNACExceeded = GetExcessNACExceeded(block);

            premise.ServiceCharge = GetPremiseServiceCharge(block);

            _premises[premise.PremiseId] = premise;
        }

        private decimal GetPremiseServiceCharge(string block)
        {
            const string serviceCharge = "SERVICE CHARGE";
            var pos = block.IndexOf(serviceCharge);

            if (pos > -1)
            {
                pos += serviceCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("R") + "R".Length;
                var pos2 = substring.IndexOf(".") + 3;
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();
                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Decimal.Parse(rate, nfi);
            }
            else
            {
                return 0.0m;
            }
        }

        private double GetExcessNACExceeded(string block)
        {
            const string exceededBy = "EXCEEDED BY";
            var pos = block.IndexOf(exceededBy);

            if (pos > -1)
            {
                pos += exceededBy.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("KVA");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0;
            }
        }

        private int GetExcessNACEvents(string block)
        {
            const string numberOfEvents = "NUMBER OF EVENTS:";
            var pos = block.IndexOf(numberOfEvents);

            if (pos > -1)
            {
                pos += numberOfEvents.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("R");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Int32.Parse(usage, nfi);
            }
            else
            {
                return 0;
            }
        }

        private double GetExcessNACUsage(string block)
        {
            const string excessNacCharge = "EXCESS NAC CHARGE";
            var pos = block.IndexOf(excessNacCharge);

            if (pos > -1)
            {
                pos += excessNacCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("KVA @");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0;
            }
        }

        private decimal GetExcessNACCharge(string block)
        {
            const string reactiveEnergyCharge = "EXCESS NAC CHARGE";
            var pos = block.IndexOf(reactiveEnergyCharge);

            if (pos > -1)
            {
                pos += reactiveEnergyCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("@ R") + "@ R".Length;
                var pos2 = substring.IndexOf(":");
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();
                return Decimal.Parse(rate);
            }
            else
            {
                return 0.0m;
            }
        }

        private double GetPremiseUtilisedCapacity(string block)
        {
            if (!_is2009Format)
            {
                const string notifiedmaxdemandutilisedcapacity = "NOTIFIEDMAXDEMANDUTILISEDCAPACITY";
                var pos = block.IndexOf(notifiedmaxdemandutilisedcapacity);
                var substring = block.Substring(pos);
                pos = substring.IndexOf("CONSUMPTION DETAILS");
                var maxDemandBlock = substring.Substring(0, pos);
                var maxDemandBlockLines = maxDemandBlock.Split(new string[] { Environment.NewLine },
                                                               StringSplitOptions.None);
                var maxDemandLine = maxDemandBlockLines[4];
                var maxDemand = maxDemandLine.Substring(maxDemandLine.IndexOf(".") + 3);

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(maxDemand, nfi);
            }
            else
            {
                const string notifiedmaxdemandutilisedcapacity = "UTILISED CAPACITY";
                var pos = block.IndexOf(notifiedmaxdemandutilisedcapacity);
                var substring = block.Substring(pos + notifiedmaxdemandutilisedcapacity.Length);
                var maxDemand = substring.Substring(0, substring.IndexOf(".") + 3);

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(maxDemand, nfi);
            }
        }

        private double GetPremiseNotifiedMaxDemand(string block)
        {
            if (!_is2009Format)
            {
                const string notifiedmaxdemandutilisedcapacity = "NOTIFIEDMAXDEMANDUTILISEDCAPACITY";
                var pos = block.IndexOf(notifiedmaxdemandutilisedcapacity);
                var substring = block.Substring(pos);
                pos = substring.IndexOf("CONSUMPTION DETAILS");
                var maxDemandBlock = substring.Substring(0, pos);
                var maxDemandBlockLines = maxDemandBlock.Split(new string[] { Environment.NewLine },
                                                               StringSplitOptions.None);
                var maxDemandLine = maxDemandBlockLines[4];
                var maxDemand = maxDemandLine.Substring(0, maxDemandLine.IndexOf(".") + 3);

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(maxDemand, nfi);
            }
            else
            {
                const string notifiedmaxdemandutilisedcapacity = "NOTIFIED MAX DEMAND";
                var pos = block.IndexOf(notifiedmaxdemandutilisedcapacity);
                var substring = block.Substring(pos + notifiedmaxdemandutilisedcapacity.Length);
                var maxDemand = substring.Substring(0, substring.IndexOf(".") + 3);

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(maxDemand, nfi);
            }
        }

        private double GetPremiseLoadFactor(string block)
        {
            const string demandReadingKwKva = "LOAD FACTOR";
            var pos = block.IndexOf(demandReadingKwKva) + demandReadingKwKva.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf(".");
            var usage = substring.Substring(0, pos + 3);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseExcessReactiveEnergy(string block)
        {
            const string excessReactiveEnergy = "EXCESS REACTIVE ENERGY";

            var pos = block.IndexOf(excessReactiveEnergy);

            if (pos > -1)
            {
                pos += excessReactiveEnergy.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("LOAD FACTOR");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0;
            }
        }

        private double GetPremiseReactiveEnergyPeak(string block)
        {
            const string demandConsumptionPeak = "REACTIVE ENERGY - PEAK";
            var pos = block.IndexOf(demandConsumptionPeak);

            if (pos > -1)
            {
                pos += demandConsumptionPeak.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("EXCESS REACTIVE ENERGY");

                if (pos < 0)
                {
                    pos = substring.IndexOf("LOAD FACTOR");
                }

                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private double GetPremiseReactiveEnergyStandard(string block)
        {
            const string reactiveEnergyStd = "REACTIVE ENERGY - STD";
            var pos = block.IndexOf(reactiveEnergyStd);

            if (pos > -1)
            {
                pos += reactiveEnergyStd.Length;

                var substring = block.Substring(pos);
                pos = substring.IndexOf(".");
                var usage = substring.Substring(0, pos + 3);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private double GetPremiseReactiveEnergyOffPeak(string block)
        {
            const string reactiveEnergyOffPeak = "REACTIVE ENERGY - OFF PEAK";
            var pos = block.IndexOf(reactiveEnergyOffPeak);

            if (pos > -1)
            {
                pos += reactiveEnergyOffPeak.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf(".");
                var usage = substring.Substring(0, pos + 3);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private double GetPremiseDemandConsumptionReading(string block)
        {
            const string demandReadingKwKva = "DEMAND READING - KW/KVA";
            var pos = block.IndexOf(demandReadingKwKva) + demandReadingKwKva.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf(".");
            var usage = substring.Substring(0, pos + 3);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseDemandConsumptionPeak(string block)
        {
            const string demandConsumptionPeak = "DEMAND CONSUMPTION - PEAK";
            var pos = block.IndexOf(demandConsumptionPeak) + demandConsumptionPeak.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("DEMAND READING - KW/KVA");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseDemandConsumptionStandard(string block)
        {
            const string demandConsumptionStd = "DEMAND CONSUMPTION - STD";
            var pos = block.IndexOf(demandConsumptionStd) + demandConsumptionStd.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf(".");
            var usage = substring.Substring(0, pos + 3);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseDemandConsumptionOffPeak(string block)
        {
            const string demandConsumptionOffPeak = "DEMAND CONSUMPTION - OFF PEAK";
            var pos = block.IndexOf(demandConsumptionOffPeak) + demandConsumptionOffPeak.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf(".");
            var usage = substring.Substring(0, pos + 3);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseEnergyConsumptionTotal(string block)
        {
            const string energyConsumptionPeak = "ENERGY CONSUMPTION ALL KWH";
            var pos = block.IndexOf(energyConsumptionPeak) + energyConsumptionPeak.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("DEMAND CONSUMPTION");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseEnergyConsumptionPeak(string block)
        {
            const string energyConsumptionPeak = "ENERGY CONSUMPTION PEAK KWH";
            var pos = block.IndexOf(energyConsumptionPeak) + energyConsumptionPeak.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("ENERGY CONSUMPTION ALL");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseEnergyConsumptionStandard(string block)
        {
            const string energyConsumptionStd = "ENERGY CONSUMPTION STD KWH";
            var pos = block.IndexOf(energyConsumptionStd) + energyConsumptionStd.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("ENERGY CONSUMPTION PEAK");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetPremiseEnergyConsumptionOffPeak(string block)
        {
            const string energyConsumptionOffPeak = "ENERGY CONSUMPTION OFF PEAK KWH";
            var pos = block.IndexOf(energyConsumptionOffPeak) + energyConsumptionOffPeak.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("ENERGY CONSUMPTION STD");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private double GetRetailEnviromentLevykWh(string block)
        {
            const string retailEnvironmentalLevyCharge = "RETAIL ENVIRONMENTAL LEVY CHARGE";
            var pos = block.IndexOf(retailEnvironmentalLevyCharge);

            if (pos > -1)
            {
                pos += retailEnvironmentalLevyCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("KWH @");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private decimal GetRetailEnvironmentLevyChargeRate(string block)
        {
            const string retailEnvironmentalLevyCharge = "RETAIL ENVIRONMENTAL LEVY CHARGE";
            var pos = block.IndexOf(retailEnvironmentalLevyCharge);

            if (pos > -1)
            {
                pos += retailEnvironmentalLevyCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("@ R") + "@ R".Length;
                var pos2 = substring.IndexOf("/KWH");
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();
                return Decimal.Parse(rate);
            }
            else
            {
                return 0.0m;
            }
        }

        private decimal GetPremiseTotalCharges(string block)
        {
            const string totalCharges = "TOTAL CHARGES";
            var pos = block.IndexOf(totalCharges);

            if (pos > -1)
            {
                pos += totalCharges.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("R") + "R".Length;
                var pos2 = substring.IndexOf(".") + 3;
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Decimal.Parse(rate, nfi);
            }
            else
            {
                return 0.0m;
            }
        }

        private double GetElectrificationAndRuralSubsidykWh(string block)
        {
            const string electrificationAndRuralSubsidy = "ELECTRIFICATION AND RURAL SUBSIDY";
            var pos = block.IndexOf(electrificationAndRuralSubsidy) + electrificationAndRuralSubsidy.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KWH @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetElectrificationAndRuralSubsidyRate(string block)
        {
            const string electrificationAndRuralSubsidy = "ELECTRIFICATION AND RURAL SUBSIDY";
            var pos = block.IndexOf(electrificationAndRuralSubsidy) + electrificationAndRuralSubsidy.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf("/KWH");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private bool GetPremiseIsHighSeason(string block)
        {
            const string highSeason = "HIGH SEASON";
            return block.IndexOf(highSeason) > -1;
        }

        private double GetPremiseSeasonStandardkWh(string block)
        {
            const string seasonStandardEnergyCharge = "SEASON STANDARD ENERGY CHARGE";
            var pos = block.IndexOf(seasonStandardEnergyCharge) + seasonStandardEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KWH @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetPremiseSeasonStandardCharge(string block)
        {
            const string seasonStandardEnergyCharge = "SEASON STANDARD ENERGY CHARGE";
            var pos = block.IndexOf(seasonStandardEnergyCharge) + seasonStandardEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf("/KWH");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private double GetPremiseSeasonReactivekvarh(string block)
        {
            const string seasonReactiveEnergyCharge = "SEASON REACTIVE ENERGY CHARGE";
            var pos = block.IndexOf(seasonReactiveEnergyCharge);

            if (pos > -1)
            {
                pos += seasonReactiveEnergyCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("KVARH @");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private decimal GetPremiseSeasonReactiveCharge(string block)
        {
            const string seasonReactiveEnergyCharge = "SEASON REACTIVE ENERGY CHARGE";
            var pos = block.IndexOf(seasonReactiveEnergyCharge);

            if (pos > -1)
            {
                pos += seasonReactiveEnergyCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("@ R") + "@ R".Length;
                var pos2 = substring.IndexOf("/KVARH");
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();
                return Decimal.Parse(rate);
            }
            else
            {
                return 0.0m;
            }
        }

        private double GetPremiseSeasonPeakkWh(string block)
        {
            const string seasonPeakEnergyCharge = "SEASON PEAK ENERGY CHARGE";
            var pos = block.IndexOf(seasonPeakEnergyCharge) + seasonPeakEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KWH @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetPremiseSeasonPeakCharge(string block)
        {
            const string seasonPeakEnergyCharge = "SEASON PEAK ENERGY CHARGE";
            var pos = block.IndexOf(seasonPeakEnergyCharge) + seasonPeakEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf("/KWH");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private double GetPremiseSeasonOffPeakkWh(string block)
        {
            const string seasonOffPeakEnergyCharge = "SEASON OFF PEAK ENERGY CHARGE";
            var pos = block.IndexOf(seasonOffPeakEnergyCharge) + seasonOffPeakEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KWH @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetPremiseSeasonOffPeakCharge(string block)
        {
            const string seasonOffPeakEnergyCharge = "SEASON OFF PEAK ENERGY CHARGE";
            var pos = block.IndexOf(seasonOffPeakEnergyCharge) + seasonOffPeakEnergyCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf("/KWH");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private double GetPremiseNetworkDemandkVA(string block)
        {
            const string networkDemandCharge = "NETWORK DEMAND CHARGE";
            var pos = block.IndexOf(networkDemandCharge) + networkDemandCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KVA @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetPremiseNetworkDemandCharge(string block)
        {
            const string networkDemandCharge = "NETWORK DEMAND CHARGE";
            var pos = block.IndexOf(networkDemandCharge) + networkDemandCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf(":");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private double GetPremiseNetworkAccesskVA(string block)
        {
            const string txNetworkAccessCharge = "TX NETWORK ACCESS CHARGE";
            const string networkAccessCharge = "NETWORK ACCESS CHARGE";
            var pos = block.IndexOf(txNetworkAccessCharge) + txNetworkAccessCharge.Length;
            block = block.Substring(pos);
            pos = block.IndexOf(networkAccessCharge) + networkAccessCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("KVA @");
            var usage = substring.Substring(0, pos);
            usage = usage.Trim();

            var nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            nfi.NumberGroupSeparator = ",";
            return Double.Parse(usage, nfi);
        }

        private decimal GetPremiseNetworkAccessCharge(string block)
        {
            const string txNetworkAccessCharge = "TX NETWORK ACCESS CHARGE";
            const string networkAccessCharge = "NETWORK ACCESS CHARGE";
            var pos = block.IndexOf(txNetworkAccessCharge) + txNetworkAccessCharge.Length;
            block = block.Substring(pos);
            pos = block.IndexOf(networkAccessCharge) + networkAccessCharge.Length;
            var substring = block.Substring(pos);
            pos = substring.IndexOf("@ R") + "@ R".Length;
            var pos2 = substring.IndexOf(":");
            var rate = substring.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private double GetPremiseTXkVA(string block)
        {
            const string txNetworkAccessCharge = "TX NETWORK ACCESS CHARGE";
            var pos = block.IndexOf(txNetworkAccessCharge);

            if (pos > -1)
            {
                pos += txNetworkAccessCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("KVA @");
                var usage = substring.Substring(0, pos);
                usage = usage.Trim();

                var nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";
                nfi.NumberGroupSeparator = ",";
                return Double.Parse(usage, nfi);
            }
            else
            {
                return 0.0;
            }
        }

        private decimal GetPremiseTXNetworkCharge(string block)
        {
            const string txNetworkAccessCharge = "TX NETWORK ACCESS CHARGE";
            var pos = block.IndexOf(txNetworkAccessCharge);

            if (pos > -1)
            {
                pos += txNetworkAccessCharge.Length;
                var substring = block.Substring(pos);
                pos = substring.IndexOf("@ R") + "@ R".Length;
                var pos2 = substring.IndexOf(":");
                var rate = substring.Substring(pos, pos2 - pos);
                rate = rate.Trim();
                return Decimal.Parse(rate);
            }
            else
            {
                return 0.0m;
            }
        }

        private int GetPremiseAdminChargePeriod(string block)
        {
            const string perDayFor = "PER DAY FOR";
            var pos = block.IndexOf(perDayFor) + perDayFor.Length;
            var pos2 = block.IndexOf("DAYS");
            var days = block.Substring(pos, pos2 - pos);
            days = days.Trim();
            return int.Parse(days);
        }

        private decimal GetPremiseAdminChargeRate(string block)
        {
            const string adminString = "ADMINISTRATION CHARGE @ R";
            var pos = block.IndexOf(adminString) + adminString.Length;
            var pos2 = block.IndexOf("PER DAY");
            var rate = block.Substring(pos, pos2 - pos);
            rate = rate.Trim();
            return Decimal.Parse(rate);
        }

        private string GetPremiseId(string premiseBlock)
        {
            if (!_is2009Format)
            {
                string premisePrefix = "PREMISEIDNUMBER";
                int pos = premiseBlock.IndexOf(premisePrefix);
                var line = premiseBlock.Substring(pos + premisePrefix.Length);
                pos = line.IndexOf("TARIFF");
                line = line.Substring(0, pos);
                return line.Trim();
            }
            else
            {
                string premisePrefix = "PREMISE ID NUMBER";
                int pos = premiseBlock.IndexOf(premisePrefix);
                var line = premiseBlock.Substring(pos + premisePrefix.Length);
                pos = line.IndexOf(":");
                line = line.Substring(pos + 1);
                pos = line.IndexOf("MEGA");
                line = line.Substring(0, pos);
                return line.Trim();
            }
        }

        private void SaveRawText(string text, string file)
        {
            var writer = new StreamWriter(file);
            writer.Write(text);
            writer.Close();
        }

        private void UpdateProgress(string message)
        {
            if (ProgressOutput != null)
            {
                ProgressOutput.BeginInvoke((ProgressDelegate)DelegateUpdateProgress, new object[] { message });
            }
        }

        private string GetSaveFile()
        {
            var saveFile = (string)MainForm.Invoke((GetSaveFileDelegate)DelegateGetSaveFile);

            if (File.Exists(saveFile))
            {
                BackupFile(saveFile);
            }

            return saveFile;
        }

        private void BackupFile(string saveFile)
        {
            string backupFilename = Path.GetFileNameWithoutExtension(saveFile) + "." +
                                    DateTime.Today.ToString("yyyyMMdd") + Path.GetExtension(saveFile);
            if (!File.Exists(backupFilename))
            {
                File.Copy(saveFile, backupFilename);
            }
            else
            {
                File.Delete(backupFilename);
                File.Copy(saveFile, backupFilename);
                File.Delete(saveFile);
            }
        }

        private void DelegateUpdateProgress(string message)
        {
            ProgressOutput.AppendText(message + Environment.NewLine);
        }

        private string DelegateGetSaveFile()
        {
            return MainForm.GetSaveFile();
        }
        #endregion
    }
}
