using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ExcelDna.IntelliSense
{
    // TODO: Needs cleaning up for efficiency
    // An alternative representation is as a string with extra attributes attached
    // each indicating a range in the string to modify.
    // (like NSAttributedString)
    class FormattedText : IEnumerable<TextLine>
    {
        readonly List<TextLine> _lines;

        public FormattedText()
        {
            _lines = new List<TextLine>();
        }

        public void Add(TextLine line) { _lines.Add(line); }

        public void Add(IEnumerable<TextLine> lines) { if (lines != null) _lines.AddRange(lines); }

        public IEnumerator<TextLine> GetEnumerator()
        {
            return _lines.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public override string ToString()
        {
            return string.Join("\r\n", _lines.Select(l => l.ToString()));
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

        public override string ToString()
        {
            return string.Concat(_runs.Select(r => r.Text));
        }
    }

    class TextRun
    {
        public string Text { get; set; }
        public FontStyle Style { get; set; }    
        public string LinkAddress { get; set; }
        public bool IsLink { get { return !string.IsNullOrEmpty(LinkAddress); } }
    }

}
