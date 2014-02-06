using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FunWithWord
{
    struct StringCostElements
    {
        double workersPay;
        double machineWorkersPay;
        double machineCost;
        double materialCost;
        double cost;

        public double WorkerPay
        {
            get { return workersPay; }
            set { workersPay = value; }
        }
        public double MachineWorkersPay
        {
            get { return machineWorkersPay; }
            set { machineWorkersPay = value; }
        }
        public double MachineCost
        {
            get { return machineCost;}
            set { machineCost = value; }
        }
        public double MaterialCost
        {
            get { return materialCost; }
            set { materialCost = value; }
        }
        public double FOT
        {
            get { return workersPay + machineWorkersPay; }
        }
        public double CalcCost
        {
            get { return workersPay + machineCost + materialCost; }
        }
        public double Cost
        {
            get { return cost; }
            set { cost = value; }
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

        public int CompareTo(object otherstring)
        {
            EstimateString es = otherstring as EstimateString;
            return this.number.CompareTo(es.Number);
        }

        public double CurrentCost
        {
            get { return currentCost.Cost; }
            set { currentCost.Cost = value; }
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
                number.ToString(), name, caption, volume.ToString(), currentCost.WorkerPay.ToString(), currentCost.MachineCost.ToString(), 
                currentCost.MachineWorkersPay.ToString(), currentCost.MaterialCost.ToString(), currentCost.Cost.ToString()
            };
        }
    }
}
