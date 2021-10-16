﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamList
    {
        [XmlAttribute("kind")]
        public DoxParamListKind Kind { get; set; }

        [XmlElement("parameteritem")]
        public List<DocParamListItem> ParameterItems { get; } = new();
    }
}
