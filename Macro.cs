using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Stages;

using static NewMacroNamespace.Guids;

namespace NewMacroNamespace
{
    public class NewMacro : MacroProvider
    {
        public NewMacro(MacroContext context)
            : base(context)
        {
            #region DEBUG
#if DEBUG

            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();

#endif
            #endregion
        }

        public override void Run() => new Macro(this, Macro, false);

        void Macro()
        {
            var dialog = GetDialog();
            if (dialog.Show() && GetReceiverObject(dialog) is object) Macro(dialog);
            else Break();
        }

        UserDialogObjectAccessor GetDialog()
        {
            var dialog = Context.GetUserDialog("Заимствовать На", UserDialogObjectAccessor.CreateInstance);

            //dialog.Caption = " ";
            dialog["Родители"] = false;
            dialog["Cостав"] = false;
            dialog["Техпроцессы"] = false;
            dialog["Маршруты"] = false;
            dialog["Файлы"] = false;
            dialog["Объект-приемник"] = $"{Context.ReferenceObject.GetNumber()}";
            dialog.RemoveLink("Объект ЭСИ приемник");

            return dialog;
        }

        void Macro(UserDialogObjectAccessor dialog)
        {
            var refs = Context.GetReference("Электронная структура изделий");
            var settings = new ConfigurationSettings(Context.Connection.ConfigurationSettings)
            {
                Date = "Сегодня",
                ApplyDate = true,

                DesignContext = null,
                ApplyDesignContext = true
            };

            using (refs.ChangeAndHoldConfigurationSettings(settings))
            {
                var receiver = GetReceiverObject(dialog).GetCorrectedObj(refs);
                var source = Context.ReferenceObject.GetCorrectedObj(refs);

                Macro(receiver, source, dialog);
            }
        }

        void Macro(ReferenceObject receiver, ReferenceObject source, UserDialogObjectAccessor dialog)
        {
            var node = new StageNode(Context);

            CopyParents(receiver, source, node, dialog);
            CopyChilds(receiver, source, node, dialog);
            CopyTechs(receiver, source, node, dialog);
            CopyFiles(receiver, source, node, dialog);

            receiver.EditOff("Заимствовать На");
            node.ChangeOff($"Заимствовать На   {CurrentUser}   of");
        }

        void CopyParents(ReferenceObject receiver, ReferenceObject source, StageNode node, UserDialogObjectAccessor dialog)
        {
            if (!dialog["Родители"])
                return;

            foreach (var group in GetAddedParents(receiver, source).GroupBy(e => e.ParentObjectId))
            {
                node.ChangeOn(group.First().ParentObject, $"Заимствовать На   {CurrentUser}   on");

                foreach (var e in group)
                    CopyLink(receiver, e, LinkCopy.Parent);

                node.ChangeOff($"Заимствовать На   {CurrentUser}   of");
            }
        }

        void CopyChilds(ReferenceObject receiver, ReferenceObject source, StageNode node, UserDialogObjectAccessor dialog)
        {
            if (!dialog["Cостав"])
                return;

            var added = GetAddedChilds(receiver, source);

            if (!added.Any())
                return;

            node.ChangeOn(receiver, $"Заимствовать На   {CurrentUser}   on");

            foreach (var e in added)
                CopyLink(receiver, e, LinkCopy.Child);
        }

        void CopyTechs(ReferenceObject receiver, ReferenceObject source, StageNode node, UserDialogObjectAccessor dialog)
        {
            var types = new List<string>();

            if (dialog["Техпроцессы"]) 
                types.Add("Технологический процесс");

            if (dialog["Маршруты"]) 
                types.Add("Маршрут");

            if (types.Count == 0) 
                return;

            var techs = source.Techs_loadfromserver(types.ToArray());

            node.ChangeOn(receiver, $"Заимствовать На   {CurrentUser}   on");
            receiver.EditOn();

            foreach (var e in techs)
                receiver.AddLinkedObject(СправочникЭлектроннаяСтруктура.Связь.ТехнологическиеОбъекты, e);
        }

        void CopyFiles(ReferenceObject receiver, ReferenceObject source, StageNode node, UserDialogObjectAccessor dialog)
        {
            if (!dialog["Файлы"])
                return;

            var files = source.Files_loadfromserver();

            node.ChangeOn(receiver, $"Заимствовать На   {CurrentUser}   on");

            var linkedObj = ((NomenclatureObject)receiver).LinkedObject;

            linkedObj.EditOn();

            foreach (var e in files)
                linkedObj.AddLinkedObject(СправочникЭлектроннаяСтруктура.Связь.Файлы, e);

            linkedObj.EditOff("Заимствовать На");
        }

        List<NomenclatureHierarchyLink> GetAddedParents(ReferenceObject receiver, ReferenceObject source)
        {
            var receiverParents = GetParents(receiver);
            var sourceParents = GetParents(source);

            return sourceParents
                .Where(e => !receiverParents.Contains(e, new CopyParentsComparer()))
                .ToList();
        }

        List<NomenclatureHierarchyLink> GetParents(ReferenceObject obj)
        {
            var parents = obj.Parents_loadfromserver().ToList();

            if (parents.ContainsCombinedStructs())
                Error("Родительские подключения объекта-приемника содержат совмещенные структуры");

            return parents;
        }

        List<NomenclatureHierarchyLink> GetAddedChilds(ReferenceObject receiver, ReferenceObject source)
        {
            var receiverChilds = GetChilds(receiver);
            var sourceChilds = GetChilds(source);

            return sourceChilds
                .Where(e => !receiverChilds.Contains(e, new CopyChildsComparer()))
                .ToList();
        }

        List<NomenclatureHierarchyLink> GetChilds(ReferenceObject obj)
        {
            var childs = obj.Childs_loadfromserver().ToList();

            if (childs.ContainsCombinedStructs())
                Error("Состав объекта-источника содержит совмещенные структуры");

            return childs;
        }

        ReferenceObject GetReceiverObject(UserDialogObjectAccessor dialog)
        {
            return ((ReferenceObject)dialog).GetObject(ДиалогЗаимствоватьНа.СвязьНаОбъектПриемник);
        }

        NomenclatureHierarchyLink CopyLink(ReferenceObject receiver, NomenclatureHierarchyLink parent, LinkCopy copyLink)
        {
            var link = CreateLink(receiver, parent, copyLink);
            link.UpdateFromLink(parent, true, true);
            link.UsingInStructuresLink.RemoveAll();
            link.UsingInStructuresLink.AddLinkedObject(parent.GetStructs().FirstOrDefault());
            link.EndChanges();

            return link;
        }

        NomenclatureHierarchyLink CreateLink(ReferenceObject receiver, NomenclatureHierarchyLink parent, LinkCopy linkcopy)
        {
            if (linkcopy == LinkCopy.Parent)
                return (NomenclatureHierarchyLink)receiver.CreateParentLink(parent.ParentObject);

            if (linkcopy == LinkCopy.Child)
                return (NomenclatureHierarchyLink)receiver.CreateChildLink(parent.ChildObject);

            throw new NotImplementedException();
        }
    }

    static partial class Extensions
    {
        public static Reference GetReference(this MacroContext context, string name)
        {
            return context.Connection.ReferenceCatalog.Find(name).CreateReference();
        }

        public static string GetVar(this ReferenceObject obj, Guid guid)
        {
            return obj[guid].GetString();
        }

        public static string GetNumber(this ReferenceObject obj)
        {
            return obj.GetVar(СправочникЭлектроннаяСтруктура.Параметр.Обозначение);
        }

        public static List<ReferenceObject> GetStructs(this ComplexHierarchyLink link)
        {
            return link.GetObjects(СправочникЭлектроннаяСтруктура.Связь.ПодключенныеСтруктуры);
        }

        public static int GetFirstStructId(this NomenclatureHierarchyLink link)
        {
            return link.GetStructs().FirstOrDefault().SystemFields.Id;
        }

        public static bool ContainsCombinedStructs(this List<NomenclatureHierarchyLink> links)
        {
            return links.Any(e => e.GetStructs().Count > 1);
        }

        public static List<ReferenceObject> Techs_loadfromserver(this ReferenceObject obj, params string[] types)
        {
            return obj.GetObjects(СправочникЭлектроннаяСтруктура.Связь.ТехнологическиеОбъекты)
                .Where(e => types.Contains(e.Class.Name))
                .ToList();
        }

        public static ReferenceObject GetCorrectedObj(this ReferenceObject obj, Reference refs)
        {
            return refs.Find(obj.Id);
        }

        public static IEnumerable<NomenclatureHierarchyLink> Parents_loadfromserver(this ReferenceObject obj)
        {
            return obj.GetParentHierarchy()
                .Select(x => obj.GetParentLinks(x)
                    .Where(e => e != null)
                    .Select(e => (NomenclatureHierarchyLink)e))
                .SelectMany(e => e);
        }

        static IEnumerable<ReferenceObject> GetParentHierarchy(this ReferenceObject obj)
        {
            return ((NomenclatureObject)obj).GetParentHierarchy();
        }

        static IEnumerable<ReferenceObject> GetParentHierarchy(this NomenclatureObject obj)
        {
            return obj.Reference.GetEntrances(obj, true, true).GetParentHierarchy();
        }

        static IEnumerable<ReferenceObject> GetParentHierarchy(this EntrancesTree tree)
        {
            return tree.Objects
                .Select(e => e.Object);
        }

        public static IEnumerable<NomenclatureHierarchyLink> Childs_loadfromserver(this ReferenceObject obj)
        {
            var loader = obj.Reference.CreateLoader(null, obj);
            loader.Load();
            return loader.GetLoadedHierarchyLinks().Select(e => (NomenclatureHierarchyLink)e);
        }

        public static List<ReferenceObject> Files_loadfromserver(this ReferenceObject obj)
        {
            return ((NomenclatureObject)obj).LinkedObject.GetObjects(СправочникЭлектроннаяСтруктура.Связь.Файлы);
        }

        public static void EditOn(this ReferenceObject obj)
        {
            if (obj.CanCheckOut)
                obj.CheckOut();

            if (!obj.Changing)
                obj.BeginChanges(false);
        }

        public static void EditOff(this ReferenceObject obj, string comment)
        {
            if (obj.Changing)
                obj.EndChanges();

            if (obj.CanCheckIn)
                Desktop.CheckIn(obj, comment, false);
        }
    }

    class StageNode
    {
        readonly MacroContext context;

        Stage stage;
        readonly List<string> skipped = new List<string>() { "Разработка", "Испр. Документовед", "Исправление" };

        ReferenceObject obj;
        Stage oldStage;

        public StageNode(MacroContext context)
        {
            this.context = context;
        }

        Stage Stage
        {
            get
            {
                if (!(stage is object))
                    stage = Stage.Find(context.Connection, "Испр. Документовед");
                return stage;
            }
        }

        public void ChangeOn(ReferenceObject obj, string comment)
        {
            oldStage = obj.SystemFields.Stage.Stage;

            if (skipped.Contains(oldStage.Name))
                return;

            this.obj = obj;
            Stage.Change(new List<ReferenceObject> { obj }, comment);
        }

        public void ChangeOff(string comment)
        {
            if (obj is object)
            {
                oldStage.Change(new List<ReferenceObject> { obj }, comment);
                obj = null;
            }
        }
    }

    class CopyParentsComparer : IEqualityComparer<NomenclatureHierarchyLink>
    {
        public bool Equals(NomenclatureHierarchyLink x, NomenclatureHierarchyLink y)
        {
            return x.ParentObjectId == y.ParentObjectId && x.GetFirstStructId() == y.GetFirstStructId();
        }

        public int GetHashCode(NomenclatureHierarchyLink obj)
        {
            return $"{obj.ParentObjectId} - {obj.GetFirstStructId()}".GetHashCode();
        }
    }

    class CopyChildsComparer : IEqualityComparer<NomenclatureHierarchyLink>
    {
        public bool Equals(NomenclatureHierarchyLink x, NomenclatureHierarchyLink y)
        {
            return x.ChildObjectId == y.ChildObjectId && x.GetFirstStructId() == y.GetFirstStructId();
        }

        public int GetHashCode(NomenclatureHierarchyLink obj)
        {
            return $"{obj.ChildObjectId} - {obj.GetFirstStructId()}".GetHashCode();
        }
    }

    enum LinkCopy
    {
        Parent,
        Child
    }

    class Macro
    {
        readonly MacroProvider _macro;

        public Macro(MacroProvider macro, Action action, bool showLog)
        {
            _macro = macro;
            Run(action, showLog);
        }

        void Run(Action action, bool showLog)
        {
            try
            {
                Action(action);
            }
            catch (MacroException)
            {
                throw;
            }
            catch (Exception)
            {
                Error();
                throw;
            }
            finally
            {
                if (showLog) 
                    Logger.Save(true);
            }
        }

        void Action(Action action)
        {
            var time = Stopwatch.StartNew();
            action();
            time.Stop();
            SuccessMessage(time);
        }

        void SuccessMessage(Stopwatch time)
        {
            //var message = $"Макрос успешно завершен\r\n" +
            //    $"Потрачено: {time.Elapsed:hh\\:mm\\:ss}";

            var message = $"Макрос успешно завершен\r\n";

            Message(message);
            Logger.AddToHead(message);
        }

        void Message(string message)
        {
            _macro.Message("Сообщение", message);
        }

        void Error()
        {
            var message = "При выполнении макроса возникли ошибки";

            Message(message);
            Logger.AddToHead(message);
        }
    }

    static class Logger
    {
        static readonly List<string> Head;
        static readonly List<string> Body;

        static Logger()
        {
            Head = new List<string>();
            Body = new List<string>();
        }

        static List<string> FirstSpace
        {
            get
            {
                if (Head.Any())
                    return new List<string>() { Environment.NewLine };
                return new List<string>();
            }
        }

        static List<string> SecondSpace
        {
            get
            {
                if (Body.Any())
                    return new List<string>() { Environment.NewLine };
                return new List<string>();
            }
        }

        public static string DefaultPath
        {
            get
            {
                return $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\{"! logs.log"}";
            }
        }

        static IEnumerable<string> ResultLog
        {
            get
            {
                return FirstSpace
                    .Concat(Head)
                    .Concat(SecondSpace)
                    .Concat(Body);
            }
        }

        public static void AddToHead(string line)
        {
            Head.Add(line);
        }

        public static void Add(string line)
        {
            Body.Add(line);
        }

        public static void Save(bool showLog)
        {
            WriteTo(DefaultPath);

            Head.Clear();
            Body.Clear();

            if (showLog) OpenLog();
        }

        static void WriteTo(string path)
        {
            if (ResultLog.FirstOrDefault() == null) return;
            if (File.Exists(path))
                File.Delete(path);
            File.AppendAllLines(path, ResultLog);
        }

        static void OpenLog()
        {
            if (File.Exists(DefaultPath))
                Process.Start("notepad.exe", DefaultPath);
        }
    }

    static class Guids
    {
        public static class СправочникЭлектроннаяСтруктура
        {
            public static class Параметр
            {
                public static Guid Обозначение { get; } = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            }

            public static class Связь
            {
                public static Guid ПодключенныеСтруктуры { get; } = new Guid("77726357-b0eb-4cea-afa5-182e21eb6373");
                public static Guid ТехнологическиеОбъекты { get; } = new Guid("e1e8fa07-6598-444d-8f57-3cfd1a3f4360");
                public static Guid Файлы { get; } = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
            }
        }

        public static class СправочникТипыСтруктур
        {
            public static class Параметр
            {
                public static Guid КороткоеНаименование { get; } = new Guid("ca6bf066-1494-415f-8eee-5ed8a02b6b6c");
            }
        }

        public static class ДиалогЗаимствоватьНа
        {
            public static Guid СвязьНаОбъектПриемник { get; } = new Guid("7e5ac86e-c8ec-4721-8c28-c419d768e637");
        }
    }
}