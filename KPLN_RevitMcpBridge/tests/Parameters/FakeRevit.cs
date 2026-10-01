using System;
using System.Collections.Generic;
using System.Linq;
namespace Autodesk.Revit.DB
{
    public enum StorageType { None, String, Integer, Double, ElementId }
    public enum TransactionStatus { Uninitialized, Started, Committed, RolledBack }
    public enum FailureProcessingResult { Continue, ProceedWithRollBack }
    public class ElementId { public long Value; public ElementId(long value) { Value = value; } }
    public class DataType { public string TypeId => "test"; }
    public class Definition { public string Name = "Test"; public DataType GetDataType() => new DataType(); }
    public class Parameter
    {
        public ElementId Id; public Definition Definition = new Definition(); public StorageType StorageType;
        public object Value; public bool IsReadOnly; public bool IsShared; public Guid GUID; public bool FailSet;
        public bool HasValue => Value != null;
        public string AsString() => (string)Value; public int AsInteger() => (int)Value; public double AsDouble() => (double)Value;
        public ElementId AsElementId() => (ElementId)Value; public string AsValueString() => Convert.ToString(Value);
        private bool Put(object value) { if (FailSet) throw new Exception("Set failed"); Value = value; return true; }
        public bool Set(string v) => Put(v); public bool Set(int v) => Put(v); public bool Set(double v) => Put(v); public bool Set(ElementId v) => Put(v);
    }
    public class Element { public ElementId Id; public string UniqueId; public List<Parameter> Parameters = new List<Parameter>(); }
    public class Document
    {
        public bool IsReadOnly, IsLinked, IsModifiable, IsFamilyDocument, FailCommit;
        public Action OnCommit;
        public List<Element> Elements = new List<Element>();
        public Element GetElement(string id) => Elements.FirstOrDefault(e => e.UniqueId == id);
        public Element GetElement(ElementId id) => Elements.FirstOrDefault(e => e.Id.Value == id.Value);
        public void Regenerate() { }
        public Dictionary<Parameter, object> Snapshot() => Elements.SelectMany(e => e.Parameters).ToDictionary(p => p, p => p.Value);
        public void Restore(Dictionary<Parameter, object> values) { foreach (var item in values) item.Key.Value = item.Value; }
    }
    public class FailureMessage { public string GetDescriptionText() => "Synthetic warning"; }
    public class FailuresAccessor { public IList<FailureMessage> GetFailureMessages() => new List<FailureMessage> { new FailureMessage() }; }
    public interface IFailuresPreprocessor { FailureProcessingResult PreprocessFailures(FailuresAccessor accessor); }
    public class FailureHandlingOptions
    {
        public IFailuresPreprocessor Processor;
        public FailureHandlingOptions SetFailuresPreprocessor(IFailuresPreprocessor p) { Processor = p; return this; }
        public FailureHandlingOptions SetClearAfterRollback(bool value) => this;
        public FailureHandlingOptions SetForcedModalHandling(bool value) => this;
    }
    public class Transaction : IDisposable
    {
        protected Document Doc; protected TransactionStatus State; protected Dictionary<Parameter, object> Saved;
        private FailureHandlingOptions _options = new FailureHandlingOptions();
        public Transaction(Document doc, string name) { Doc = doc; }
        public TransactionStatus Start() { Saved = Doc.Snapshot(); return State = TransactionStatus.Started; }
        public FailureHandlingOptions GetFailureHandlingOptions() => _options;
        public void SetFailureHandlingOptions(FailureHandlingOptions value) { _options = value; }
        public TransactionStatus GetStatus() => State;
        public TransactionStatus RollBack() { Doc.Restore(Saved); return State = TransactionStatus.RolledBack; }
        public TransactionStatus Commit()
        {
            if (Doc.FailCommit && _options.Processor.PreprocessFailures(new FailuresAccessor()) == FailureProcessingResult.ProceedWithRollBack) return RollBack();
            Doc.OnCommit?.Invoke(); return State = TransactionStatus.Committed;
        }
        public void Dispose() { if (State == TransactionStatus.Started) RollBack(); }
    }
    public class TransactionGroup : Transaction
    {
        public TransactionGroup(Document doc, string name) : base(doc, name) { }
        public TransactionStatus Assimilate() => State = TransactionStatus.Committed;
    }
}
namespace KPLN_RevitMcpBridge.Services
{
    internal static class RevitService
    {
        public static long IdValue(Autodesk.Revit.DB.ElementId id) => id.Value;
        public static Autodesk.Revit.DB.ElementId Id(long value) => new Autodesk.Revit.DB.ElementId(value);
    }
}
