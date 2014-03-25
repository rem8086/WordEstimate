﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FunWithWord
{
    public struct EstimateEquipment         //class for whole estimate (list of EstimateString and additional data)
    {
        public double EquipmentCost { get; set; }
        public double TransportCost { get; set; }
        public double DepotCost { get; set; }

        public double TotalEquipmentCost
        {
            get { return EquipmentCost + TransportCost + DepotCost; }
        }
    }

    class Estimate
    {
        List<EstimateString> estimateSet;
        EstimateString resumeString;
        EstimateEquipment equip;
        string name;
        public double TotalEstimateCost { get; set; }
        public double Overheads { get; set; }
        public double EstimateProfit { get; set; }

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

        public void AddResumeString(EstimateString resume)
        {
            resumeString = resume;
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

        public List<EstimateString> EstimateMaterials()
        {
            List<EstimateString> eslist = new List<EstimateString>();
            foreach (EstimateString es in estimateSet)
            {
                if ((es.CurrentFOT == 0) && (es.CurrentMachine == 0) && (es.CurrentCost > 0))
                    eslist.Add(es);
            }
            return eslist;
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
            return TotalEstimateCost - SumStringsCost() - equip.DepotCost - equip.TransportCost - Overheads - EstimateProfit;
        }
    }
}
