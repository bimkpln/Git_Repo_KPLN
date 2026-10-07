using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        internal sealed class ReviewObject
        {
            public string Key {get;set;}
            public string Title {get;set;}
            public string Source {get;set;}
            public string SourceKey {get;set;}
            public string Element {get;set;}
            public string Level {get;set;}
            public string Apartment {get;set;}
            public bool FloorContour {get;set;}
            public List<TEP.ReviewInput> Areas {get;set;}=new List<TEP.ReviewInput>();
            public TEP.Detail Detail {get;set;}
            public string Description {get{return string.Join(" | ",new[]{Level,Source,string.IsNullOrEmpty(Apartment)?null:"Квартира "+Apartment,FloorContour?null:"ID "+Element}.Where(s=>!string.IsNullOrEmpty(s)));}}
            public string AreaCount {get{return Areas.Count+" областей";}}
        }
        private List<ReviewObject> reviewObjects=new List<ReviewObject>();
        internal static List<ReviewObject> UniqueReviewObjects(IEnumerable<TEP.Detail> details,IEnumerable<TEP.ReviewInput> areas)
        {
            var objects=new Dictionary<string,ReviewObject>();
            Func<string,string,string,string> key=(sourceKey,source,id)=>{string scope=sourceKey??source??"";return scope.Length+":"+scope+"/"+id;};
            foreach(var detail in details??Enumerable.Empty<TEP.Detail>())
            {
                if(string.IsNullOrEmpty(detail.Element))continue;
                string id=key(detail.SourceKey,detail.Source,detail.Element);
                if(!objects.ContainsKey(id))objects[id]=new ReviewObject{Key=id,Title=detail.ObjectName??detail.PurposeName,Source=detail.Source,SourceKey=detail.SourceKey,Element=detail.Element,Level=detail.Level,Apartment=detail.Apartment,Detail=detail};
            }
            foreach(var area in (areas??Enumerable.Empty<TEP.ReviewInput>()).Where(a=>a!=null).GroupBy(a=>a.Key).Select(g=>g.First()))
            {
                bool floor=(area.Kind??"").StartsWith("Контур этажа");
                string id=floor?"floor:"+(area.Key??"").Substring(0,Math.Max(0,(area.Key??"").LastIndexOf('/'))):key(area.SourceKey,area.Source,area.Element);
                ReviewObject item;
                if(!objects.TryGetValue(id,out item))objects[id]=item=new ReviewObject{Key=id,Title=floor?"Контур этажа":area.ObjectName??area.Kind,Source=floor?"Общий расчётный контур":area.Source,SourceKey=area.SourceKey,Element=floor?null:area.Element,Level=area.Level,Apartment=floor?null:area.Apartment,FloorContour=floor};
                item.Areas.Add(area);
            }
            return objects.Values.OrderByDescending(o=>o.FloorContour).ThenByDescending(o=>o.Areas.Count>0).ThenBy(o=>o.Level).ThenBy(o=>o.Apartment).ThenBy(o=>o.Title).ThenBy(o=>o.Key).ToList();
        }
        private void RefreshObjectReview(TEP.Run report)
        {
            reviewObjects=report==null?new List<ReviewObject>():UniqueReviewObjects(report.Details,Engine.ReviewInputs.Count>0?Engine.ReviewInputs:report.ReviewAreas);
            FilterReviewObjects();
        }
        private void ReviewSearch_Changed(object sender,TextChangedEventArgs args){FilterReviewObjects();}
        private void FilterReviewObjects()
        {
            if(ObjectsList==null||reviewObjects==null)return;
            string search=(ObjectSearch?.Text??"").Trim();var selected=ObjectsList.SelectedItem as ReviewObject;
            var visible=reviewObjects.Where(o=>search.Length==0||(o.Title+" "+o.Description).IndexOf(search,StringComparison.CurrentCultureIgnoreCase)>=0).ToList();
            ObjectsList.ItemsSource=visible;ObjectsList.SelectedItem=visible.FirstOrDefault(o=>o.Key==selected?.Key)??visible.FirstOrDefault();
            if(ObjectCount!=null)ObjectCount.Text="Объектов: "+visible.Count+" / "+reviewObjects.Count;
            ReviewObject_Changed(null,null);
        }
        private void ReviewObject_Changed(object sender,SelectionChangedEventArgs args)
        {
            if(ReviewGrid==null||SelectedObjectTitle==null)return;
            var item=ObjectsList.SelectedItem as ReviewObject;
            SelectedObjectTitle.Text=item?.Title??"Выберите объект";
            SelectedObjectDescription.Text=item?.Description??"";
            ReviewGrid.ItemsSource=item?.Areas;ReviewGrid.SelectedIndex=item?.Areas.Count>0?0:-1;
            EmptyAreas.Visibility=item!=null&&item.Areas.Count>0?Visibility.Collapsed:Visibility.Visible;
            EmptyAreas.Text=item==null?"Выберите объект слева.":"У объекта нет расчётных 2D-областей. Для строительного объёма используется 3D-геометрия; ошибки построения указаны в проверках.";
            NavigateObjectButton.IsEnabled=item!=null&&!item.FloorContour&&!string.IsNullOrEmpty(item.Element);
        }
        private void NavigateObject_Click(object sender,RoutedEventArgs args)
        {Try(()=>{var item=ObjectsList.SelectedItem as ReviewObject;if(item==null||item.FloorContour)return;Engine.RequestAuditNavigation(new TEP.InputAuditRow{SourceKey=item.SourceKey,Source=item.Source,ElementId=item.Element});Close();});}
        private void ReviewAction_Click(object sender,RoutedEventArgs args)
        {
            var button=sender as Button;var input=button?.DataContext as TEP.ReviewInput;if(input==null)return;
            ReviewGrid.SelectedItem=input;
            if((button.Tag as string)=="open")OpenReview_Click(sender,args);else BindReview_Click(sender,args);
        }
    }
}
