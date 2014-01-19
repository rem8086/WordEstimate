using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FunWithWord
{
    class Estimate
    {
        List<EstimateString> estimateSet;
        string name;

        public Estimate(string name)
        {
            estimateSet = new List<EstimateString>();
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
    }
}
