using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FunWithWord
{
    public struct EstimateEquipment
    {
        double equipmentCost;
        double transportCost;
        double depotCost;

        public double EquipmentCost
        {
            get { return equipmentCost; }
            set { equipmentCost = value; }
        }

        public double TransportCost
        {
            get { return transportCost; }
            set { transportCost = value; }
        }

        public double DepotCost
        {
            get { return depotCost; }
            set { depotCost = value; }
        }

        public double TotalEquipmentCost
        {
            get { return equipmentCost + transportCost + depotCost; }
        }
    }

    class Estimate
    {
        List<EstimateString> estimateSet;
        EstimateString resumeString;
        EstimateEquipment equip;
        string name;
        double totalEstimateCost;
        double overheads;
        double estimateProfit;

        public Estimate(string name)
        {
            estimateSet = new List<EstimateString>();
            resumeString = new EstimateString(0);
            equip = new EstimateEquipment();
            this.name = name;
        }

        public EstimateString this[int index]
        {
            get { return estimateSet[index]; }
        }

        public void Add(EstimateString newEsStr)
        {
            estimateSet.Add(newEsStr);
        }

        public string Name
        {
            get { return name; }
        }

        public double TotalEstimateCost
        {
            get { return totalEstimateCost; }
            set { totalEstimateCost = value; }
        }

        public double Overheads
        {
            get { return overheads; }
            set { overheads = value; }
        }

        public double EstimateProfit 
        {
            get { return estimateProfit; }
            set { estimateProfit = value; }
        }

        public int StringCount
        {
            get { return estimateSet.Count; }
        }

        public void Remove(int number)
        {
            EstimateString deletingES = new EstimateString(number);
            foreach (EstimateString es in estimateSet)
	        {
                if (es.Number == number) deletingES = es;
	        }
            estimateSet.Remove(deletingES);
        }

        public void AddEquipment(double cost, double transport, double depot)
        {
            equip.EquipmentCost = cost;
            equip.TransportCost = transport;
            equip.DepotCost = depot;
        }

        public EstimateEquipment Equipment
        {
            get { return equip; }
            set { equip = value; }
        }

        public EstimateString ResumeString
        {
            get { return resumeString; }
        }

        public void AddResumeString(EstimateString resume)
        {
            resumeString = resume;
        }

        public List<EstimateString> EstimateMaterials()
        {
            List<EstimateString> esl = new List<EstimateString>();
            foreach (EstimateString es in estimateSet)
            {
                if ((es.CurrentFOT == 0) && (es.CurrentMachine == 0) && (es.CurrentCost > 0))
                    esl.Add(es);
            }
            return esl;
        }

        public List<int> CheckMissingStrings()
        {
            List<int> elementnumbers = (List<int>)estimateSet.Select(x => x.Number);
            int maxnum = estimateSet.Max(x => x.Number);
            List<int> outputlist = new List<int>();
            for (int i = 1; i <= maxnum; i++)
            {
                outputlist.Add(i);
            }
            foreach (int el in elementnumbers)
            {
                outputlist.Remove(el);
            }
            return outputlist;
        }

        public double SumStringsCost()
        {
            double sum = 0.0;
            foreach (EstimateString es in estimateSet)
            {
                sum += es.CurrentCost;
            }
            return sum;
        }

        public double CheckCostEquality()
        {
            return totalEstimateCost - SumStringsCost() - equip.DepotCost - equip.TransportCost - overheads - estimateProfit;
        }
    }
}
