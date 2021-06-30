using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordManipulationApp.Models
{
    public class ParagraphModel
    {
        public string CombinedParagraph { get; set; }

        public List<Component> ParagraphList { get; set; }
    }

    public class Component
    {
        public string ParagraphText { get; set; }
        public ComponentEnum SelectedComponent { get; set; }
    }

    public enum ComponentEnum
    {
        Header = 1,
        Name = 2,
        Signature = 3,
        Picture = 4,
        Chart = 5
    }
}