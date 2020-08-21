using System;
using System.Collections.Generic;
using System.Text;

namespace NaicsStandard
{
    class NaicsStd
    {
        public string code { get; }
        public string revenue { get; }
        public string employees { get; }
        public string other { get; }
        public NaicsStd(string code, string revenue, string employees, string other)
        {
            this.code = code;
            this.revenue = revenue;
            this.employees = employees;
            this.other = other;
        }
    }
}
