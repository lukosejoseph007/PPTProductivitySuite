using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTProductivitySuite
{
    public class SlideLibraryItem
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string[] Tags { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime LastModified { get; set; }
        public string ThumbnailPath { get; set; }
        public string SlideFilePath { get; set; }
    }
}