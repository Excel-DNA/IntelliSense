using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // CONSIDER: Support regular .NET XML documentation, either as files or packed?
    class XmlIntelliSenseProvider : IIntelliSenseProvider
    {
        public class XmlRegistrationInfo
        {
            string _fileName;              // Might be .xml file or Workbook path. Use only if _xmlIntelliSense is null.
            string _xmlIntelliSense;       // Might be null
            XmlIntelliSense _intelliSense; // Might be null - lazy parsed

            public XmlRegistrationInfo(string fileName, string xmlIntelliSense)
            {
                _fileName = fileName;
                _xmlIntelliSense = xmlIntelliSense;
            }

            // Called in a macro context
            public void Refresh()
            {
                if (_intelliSense != null)
                    return; // Already done
                
                try
                {
                    // Parse first
                    var xml = _xmlIntelliSense;
                    if (xml == null)
                    {
                        xml = File.ReadAllText(_fileName);
                    }
                    _intelliSense = XmlIntelliSense.Parse(xml);
                    if (_intelliSense?.XmlFunctionInfo?.FunctionsList != null)
                    {
                        // Fix up SourcePath (is this used?)
                        foreach (var func in _intelliSense.XmlFunctionInfo.FunctionsList)
                        {
                            func.SourcePath = _fileName;
                        }
                    }
                }
                catch // (Exception ex)
                {
                    // TODO: Log errors
                    _intelliSense = XmlIntelliSense.Empty;
                }
            }

            // Not in macro context - don't call Excel, could be any thread.
            public IEnumerable<FunctionInfo> GetFunctionInfos()
            {
                if (_intelliSense == null || _intelliSense.XmlFunctionInfo == null)
                    return Enumerable.Empty<FunctionInfo>();

                return _intelliSense.XmlFunctionInfo.FunctionsList;
            }


        }

        ExcelSynchronizationContext _syncContextExcel;
        Dictionary<string, XmlRegistrationInfo> _xmlRegistrationInfos;
        bool _isDirty;
        public event EventHandler Invalidate;

        public XmlIntelliSenseProvider()
        {
            _xmlRegistrationInfos = new Dictionary<string, XmlRegistrationInfo>();
            _syncContextExcel = new ExcelSynchronizationContext();
        }

        // May be called on the main Excel thread or on another thread (e.g. our automation thread)
        // Pass in the xmlFunctionInfo if available (from inside document), else file will be read
        // We make the parsing lazy...
        public void RegisterXmlFunctionInfo(string fileName, string xmlIntelliSense = null)
        {
            if (!File.Exists(fileName) && xmlIntelliSense == null)
                return;
            
            var regInfo = new XmlRegistrationInfo(fileName, xmlIntelliSense);
            lock (_xmlRegistrationInfos)
            {
                _xmlRegistrationInfos.Add(fileName, regInfo);

                if (!_isDirty)
                {
                    _isDirty = true;
                    _syncContextExcel.Post(OnInvalidate, null);
                }
            }
        }

        // Safe to call even if it wasn't registered
        public void UnregisterXmlFunctionInfo(string fileName)
        {
            if (_xmlRegistrationInfos.Remove(fileName))
                _isDirty = true;
            // Not Invalidating - we're not really worried about keeping the extra information around a bit longer
        }

        public void Initialize()
        {
            // CONSIDER: We might want to instantiate the XmlSerializer here ...?
            //           (Or not, since we are never likely to need it)
            // XmlRegistrationInfo.XmlIntelliSense.Initialize();
        }

        // Runs on the main thread
        public void Refresh()
        {
            Logger.Provider.Info("XmlIntelliSenseProvider.Refresh");
            lock (_xmlRegistrationInfos)
            {
                foreach (var regInfo in _xmlRegistrationInfos.Values)
                {
                    regInfo.Refresh();
                }
            }
        }

        // May be called from any thread
        public IList<FunctionInfo> GetFunctionInfos()
        {
            IList<FunctionInfo> functionInfos;
            lock (_xmlRegistrationInfos)
            {
                functionInfos = _xmlRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
            }
            Logger.Provider.Verbose("XmlIntelliSenseProvider.GetFunctionInfos Begin");
            foreach (var info in functionInfos)
            {
                Logger.Provider.Verbose($"\t{info.Name}({info.ArgumentList.Count}) - {info.Description} ");
            }
            Logger.Provider.Verbose("XmlIntelliSenseProvider.GetFunctionInfos End");
            return functionInfos;
        }
        
        void OnInvalidate(object _unused_)
        {
            Invalidate?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
        }
    }

    #region Serialized Xml structure
    [Serializable]
    [XmlType(AnonymousType = true)]
    [XmlRoot("IntelliSense", Namespace = XmlIntelliSense.Namespace, IsNullable = false)]
    public class XmlIntelliSense
    {
        [XmlElement("FunctionInfo")]
        public XmlFunctionInfo XmlFunctionInfo;

        // returns XmlIntelliSense.Empty on failure
        public static XmlIntelliSense Parse(string xmlFunctionInfo)
        {
            Initialize();
            try
            {
                using (var stringReader = new StringReader(xmlFunctionInfo))
                {
                    return (XmlIntelliSense)_serializer.Deserialize(stringReader);
                }
            }
            catch // (Exception ex)
            {
                // TODO: Log errors
                return Empty;
            }
        }

        public static void Initialize()
        {
            if (_serializer == null)
                _serializer = new XmlSerializer(typeof(XmlIntelliSense));
        }
        static XmlSerializer _serializer;
        public static XmlIntelliSense Empty { get; } = new XmlIntelliSense { XmlFunctionInfo = new XmlFunctionInfo { FunctionsList = new List<FunctionInfo>() } };
        public const string Namespace = "http://schemas.excel-dna.net/intellisense/1.0";
    }

    [Serializable]
    public class XmlFunctionInfo
    {
        [XmlElement("Function", typeof(FunctionInfo))]
        public List<FunctionInfo> FunctionsList;
    }

    #endregion
}
