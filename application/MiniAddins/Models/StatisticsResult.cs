using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MiniAddins.Models
{
    public class StatisticsResult
    {
        public string ElementType { set; get; }
        public string Author { set; get; }
        public int Count { set; get; }

    }

    public class EaElement
    {
        public string ElementGUID { set; get; }
        public string ElementType { set; get; }
        public string Author { set; get; }
        public string Name { set; get; }

    }

    public class CheckItem
    {
        public string Type { set; get; }
        public string Name { set; get; }
        public bool IsChecked { set; get; }
    }
}
