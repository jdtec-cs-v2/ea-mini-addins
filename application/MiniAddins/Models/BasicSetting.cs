using System;

namespace MiniAddins.Models
{
    public class BasicSetting 
    {
        public String RootPath { get; set; }

        public BasicSetting Clone()
        {
            return new BasicSetting() {  RootPath = RootPath };
        }
        
    }
}
