using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FunWithWord
{
    struct StringCostElements       //struct and class, which present string of estimate, just data
    {
        public double WorkerPay { get; set; }
        public double MachineWorkersPay { get; set; }
        public double MachineCost { get; set; }
        public double MaterialCost { get; set; }
        public double Cost { get; set; }
        public double Overheads { get; set; }
        public double Profit { get; set; }
        public double TotalCost { get; set; }

        public double FOT
        {
            get { return WorkerPay + MachineWorkersPay; }
        }
        public double CalcCost
        {
            get { return WorkerPay + MachineCost + MaterialCost; }
        }
    }

    class EstimateString : IComparable
    {
        int number;
        string caption;
        string name;
        double volume;
        StringCostElements currentCost;
        //StringCostElements basicCost;

        public EstimateString(int num)
        {
            number = num;
        }

        public int Number
        {
            get { return number; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Caption
        {
            get { return caption; }
            set { caption = value; }
        }

        public double Volume
        {
            get { return volume; }
            set { volume = value; }
        }

        public double CurrentFOT
        {
            get { return currentCost.FOT; }
        }

        public double CurrentMaterials
        {
            get { return currentCost.MaterialCost; }
            set { currentCost.MaterialCost = value; }
        }

        public double CurrentMachine
        {
            get { return currentCost.MachineCost; }
            set { currentCost.MachineCost = value; }
        }

        public double CurrentWorkers
        {
            get { return currentCost.WorkerPay; }
            set { currentCost.WorkerPay = value; }
        }

        public double CurrentMachineWorkers
        {
            get { return currentCost.MachineWorkersPay; }
            set { currentCost.MachineWorkersPay = value; }
        }

        public double CurrentCost
        {
            get { return currentCost.Cost; }
            set { currentCost.Cost = value; }
        }

        public double CurrentOverheads 
        {
            get { return currentCost.Overheads; }
            set { currentCost.Overheads = value; }
        }

        public double CurrentProfit
        {
            get { return currentCost.Profit; }
            set { currentCost.Profit = value; }
        }

        public double CurrentTotalCost
        {
            get { return currentCost.TotalCost; }
            set { currentCost.TotalCost = value; }
        }

        public int CompareTo(object otherstring)
        {
            EstimateString es = otherstring as EstimateString;
            return this.number.CompareTo(es.Number);
        }

        public override string ToString()
        {
            return String.Format("Element # {0}. {1}: Salary - {2}, Machine - {3}|{4}, Materials - {5}. Sum: {6}",
                number, caption, currentCost.WorkerPay, currentCost.MachineCost, 
                currentCost.MachineWorkersPay, currentCost.MaterialCost, currentCost.Cost);
        }

        public string[] ToStringArray()
        {
            return new string[]{
                number.ToString(), Name, Caption, Volume.ToString(), CurrentWorkers.ToString(), CurrentMachine.ToString(), 
                CurrentMachineWorkers.ToString(), CurrentMaterials.ToString(), CurrentCost.ToString(), CurrentOverheads.ToString(),
                CurrentProfit.ToString(), CurrentTotalCost.ToString()
            };
        }
    }
}
