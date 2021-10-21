﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Dox2Word.Model
{
    [Flags]
    public enum TextRunFormat
    {
        None,
        Bold = 1,
        Italic = 2,
        Monospace = 4,
    }

    public class TextRun
    {
        public TextRunFormat Format { get; }

        public string Text { get; }

        public TextRun(string text, TextRunFormat format)
        {
            this.Text = text;
            this.Format = format;
        }
    }
}