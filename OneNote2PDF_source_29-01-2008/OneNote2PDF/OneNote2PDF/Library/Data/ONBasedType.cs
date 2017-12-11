using System;

namespace OneNote2PDF.Library.Data
{
    public class ONBasedType
    {
        public string Name { get; set; }
        public string ID { get; set; }
        public DateTime LastModifiedTime { get; set; }
        public ONBasedType Parent { get; set; }
    }
}
