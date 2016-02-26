using System.Collections;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelDna.IntelliSense
{
    class FormattedText : IEnumerable<TextLine>
    {
        readonly List<TextLine> _lines;

        public FormattedText()
        {
            _lines = new List<TextLine>();
        }

        public void Add(TextLine line) { _lines.Add(line); }

        public void Add(IEnumerable<TextLine> lines) { _lines.AddRange(lines); }

        public IEnumerator<TextLine> GetEnumerator()
        {
            return _lines.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    class TextLine : IEnumerable<TextRun>
    {
        readonly List<TextRun> _runs;

        public TextLine()
        {
            _runs = new List<TextRun>();
        }

        public void Add(TextRun run) { _runs.Add(run); }


        public IEnumerator<TextRun> GetEnumerator()
        {
            return _runs.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    class TextRun
    {
        public string Text { get; set; }
        public FontStyle Style { get; set; }    
        // CONSIDER: Maybe allow links?
    }

}
