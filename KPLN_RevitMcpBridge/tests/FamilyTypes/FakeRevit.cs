using System;
using System.Collections.Generic;
using System.Linq;

namespace Autodesk.Revit.DB
{
    public enum StorageType { None, String, Integer, Double, ElementId }
    public enum TransactionStatus { Uninitialized, Started, Committed, RolledBack }
    public enum FailureProcessingResult { Continue, ProceedWithRollBack }
    public class ElementId { public long Value; public ElementId(long value) { Value = value; } }
    public class Definition { public string Name; public Definition(string name) { Name = name; } }
    public class Element { public ElementId Id; }
    public class FamilyParameter
    {
        public ElementId Id; public Definition Definition; public StorageType StorageType;
        public bool IsReadOnly; public string Formula;
    }
    public class FamilyType
    {
        public string Name; internal Dictionary<FamilyParameter, object> Values = new Dictionary<FamilyParameter, object>();
        public bool HasValue(FamilyParameter p) => Values.ContainsKey(p) && Values[p] != null;
        public string AsString(FamilyParameter p) => (string)Values[p];
        public int AsInteger(FamilyParameter p) => (int)Values[p];
        public double AsDouble(FamilyParameter p) => (double)Values[p];
        public ElementId AsElementId(FamilyParameter p) => (ElementId)Values[p];
    }
    public class FamilyManager
    {
        public List<FamilyParameter> ParameterList = new List<FamilyParameter>();
        public List<FamilyType> TypeList = new List<FamilyType>();
        public IEnumerable<FamilyParameter> Parameters => ParameterList;
        public IEnumerable<FamilyType> Types => TypeList;
        public FamilyType CurrentType { get; set; }
        public FamilyType NewType(string name)
        {
            if (TypeList.Any(t => t.Name == name)) throw new Exception("duplicate");
            var type = new FamilyType { Name = name, Values = CurrentType.Values.ToDictionary(p => p.Key, p => Clone(p.Value)) };
            TypeList.Add(type); CurrentType = type; return type;
        }
        private static object Clone(object value) => value is ElementId ? new ElementId(((ElementId)value).Value) : value;
        public void Set(FamilyParameter p, string value) { CurrentType.Values[p] = value; }
        public void Set(FamilyParameter p, int value) { CurrentType.Values[p] = value; }
        public void Set(FamilyParameter p, double value) { CurrentType.Values[p] = value; }
        public void Set(FamilyParameter p, ElementId value) { CurrentType.Values[p] = value; }
    }
    public class Document
    {
        public bool IsFamilyDocument = true, IsReadOnly, IsLinked, IsModifiable, FailCommit;
        public FamilyManager FamilyManager = new FamilyManager();
        public List<Element> Elements = new List<Element>();
        public Element GetElement(ElementId id) => Elements.FirstOrDefault(e => e.Id.Value == id.Value);
        public void Regenerate() { }
        internal Snapshot Save() => new Snapshot
        {
            Types = FamilyManager.TypeList.Select(t => new FamilyType { Name = t.Name, Values = t.Values.ToDictionary(p => p.Key, p => p.Value is ElementId ? (object)new ElementId(((ElementId)p.Value).Value) : p.Value) }).ToList(),
            Current = FamilyManager.CurrentType?.Name
        };
        internal void Restore(Snapshot saved)
        {
            FamilyManager.TypeList = saved.Types;
            FamilyManager.CurrentType = FamilyManager.TypeList.FirstOrDefault(t => t.Name == saved.Current);
        }
    }
    public class Snapshot { public List<FamilyType> Types; public string Current; }
    public class FailureMessage { public string GetDescriptionText() => "Synthetic warning"; }
    public class FailuresAccessor { public IList<FailureMessage> GetFailureMessages() => new List<FailureMessage> { new FailureMessage() }; }
    public interface IFailuresPreprocessor { FailureProcessingResult PreprocessFailures(FailuresAccessor accessor); }
    public class FailureHandlingOptions
    {
        public IFailuresPreprocessor Processor;
        public FailureHandlingOptions SetFailuresPreprocessor(IFailuresPreprocessor value) { Processor = value; return this; }
        public FailureHandlingOptions SetClearAfterRollback(bool value) => this;
        public FailureHandlingOptions SetForcedModalHandling(bool value) => this;
    }
    public class Transaction : IDisposable
    {
        protected Document Doc; protected TransactionStatus State; protected Snapshot Saved;
        private FailureHandlingOptions _options = new FailureHandlingOptions();
        public Transaction(Document doc, string name) { Doc = doc; }
        public TransactionStatus Start() { Saved = Doc.Save(); return State = TransactionStatus.Started; }
        public FailureHandlingOptions GetFailureHandlingOptions() => _options;
        public void SetFailureHandlingOptions(FailureHandlingOptions value) { _options = value; }
        public TransactionStatus GetStatus() => State;
        public TransactionStatus RollBack() { Doc.Restore(Saved); return State = TransactionStatus.RolledBack; }
        public TransactionStatus Commit()
        {
            if (Doc.FailCommit && _options.Processor.PreprocessFailures(new FailuresAccessor()) == FailureProcessingResult.ProceedWithRollBack) return RollBack();
            return State = TransactionStatus.Committed;
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
