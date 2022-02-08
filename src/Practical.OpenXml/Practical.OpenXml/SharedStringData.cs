using System.Collections.Generic;

namespace Practical.OpenXml
{
    public class SharedStringData : Dictionary<string, int>
    {
        public int MaxIndex { get; set; } = 0;
    }
}
