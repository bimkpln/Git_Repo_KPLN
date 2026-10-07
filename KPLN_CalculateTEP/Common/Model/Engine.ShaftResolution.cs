using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            // Candidates that fail physical validation remain visible as draft areas,
            // but are never used as floor coverage, exclusion masks or vertical baselines.
            private readonly List<Record> unconfirmedAutomaticShafts=new List<Record>();

            private void ReadShaftReviewCandidate(Record record)
            {
                if(!Config.UseReviewRegions)return;
                string key=Config.Method+"/"+(Config.Phase??"")+"/shaft/"+record.Key;
                try
                {
                    var binding=reviewBindings.SingleOrDefault(b=>b.Key==key);
                    if(binding==null)return;
                    if(binding.Placement!=ReviewPlacement())
                        throw new InvalidOperationException("Изменилось размещение источников. Перепривяжите область шахты.");
                    var region=doc.GetElement(binding.RegionId) as FilledRegion;
                    if(region==null)throw new InvalidOperationException("Сохранённая область шахты удалена. Перепривяжите область.");
                    var view=doc.GetElement(region.OwnerViewId) as ViewPlan;
                    if(view?.GenLevel==null||Math.Abs(view.GenLevel.ProjectElevation-record.Z)>FloorGeometryTolerance||Math.Abs(binding.Elevation-record.Z)>FloorGeometryTolerance)
                        throw new InvalidOperationException("Сохранённая область шахты находится на другом расчётном уровне. Перепривяжите область.");
                    var actual=RegionFromLoops(region.GetBoundaries(),true);
                    if(actual.IsEmpty)throw new InvalidOperationException("У сохранённой области шахты нет замкнутой границы.");
                    record.AutomaticShaftRegion=actual;
                    record.Manual=binding.InitialBoundary==null||OverlayBoundaryChanged(PlanarRegion.FromRings(binding.InitialBoundary),actual);
                }
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex){record.AutomaticShaftError=ex.Message;}
            }

            private List<Record> ConfirmAutomaticShaftCandidates(List<Record> candidates)
            {
                var confirmed=new List<Record>();
                foreach(var record in candidates)
                {
                    Progress("Проверка шахты по геометрии: "+record.Source.Name+" / "+record.Level.Name+" / ID "+IDHelper.ElIdValue(record.Element.Id));
                    if(record.AutomaticShaftError==null)
                        try
                        {
                            var evidence=ReadShaftPhysicalEvidence(record);
                            var matches=evidence.Voids.Where(v=>ShaftMatchesPhysicalVoid(record.AutomaticShaftRegion,v.Region)).ToList();
                            if(matches.Count>0&&ShaftInteriorIsClear(record.AutomaticShaftRegion,evidence.WallMaterial,evidence.WallsComplete))
                                record.AutomaticShaftProof=string.Join("; ",matches.Select(v=>v.Description).Distinct(StringComparer.Ordinal));
                            else
                            {
                                string reason=matches.Count==0?"Замкнутая граница не совпадает целиком с подтверждённым горизонтальным проёмом или замкнутой пустотой в сечении стен / перекрытий.":
                                    "Граница совпадает с проёмом, но свободное пространство внутри него не подтверждено сечениями ограждений.";
                                if(evidence.WallMaterial!=null)
                                {
                                    double overlap=RegionIntersection(record.AutomaticShaftRegion,evidence.WallMaterial).Area*.09290304;
                                    if(overlap>.001)reason+=" Внутри границы находится материал стен или других ограждений: "+overlap.ToString("0.###",CultureInfo.CurrentCulture)+" м².";
                                }
                                if(evidence.Errors.Count>0)reason+=" Не удалось проверить часть геометрии: "+string.Join("; ",evidence.Errors.Distinct().Take(4));
                                record.AutomaticShaftError=reason;
                            }
                        }
                        catch(OperationCanceledException){throw;}
                        catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                        catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                        catch(Exception ex){record.AutomaticShaftError="Не удалось проверить физическую геометрию: "+ex.Message;}
                    if(record.AutomaticShaftError!=null)
                    {
                        string location="Этаж «"+record.Level.Name+"», отметка "+(record.Z*.3048).ToString("+0.000;-0.000;0.000",CultureInfo.CurrentCulture)+" м. ";
                        record.AutomaticShaftError=location+record.AutomaticShaftError+
                            " Источники границы: "+string.Join("; ",record.AutomaticShaftOrigins??new List<string>())+
                            ". Контур оставлен для проверки и не используется в расчёте шахт. Название группы или схемы само по себе не подтверждает шахту.";
                        unconfirmedAutomaticShafts.Add(record);
                        Notice("AUTO_SHAFT_UNCONFIRMED","Ошибка",record.AutomaticShaftError,record);
                        continue;
                    }
                    var duplicate=confirmed.FirstOrDefault(other=>other.Source.Mode==record.Source.Mode&&other.Building==record.Building&&
                        Math.Abs(other.Z-record.Z)<=FloorGeometryTolerance&&ShaftMatchesPhysicalVoid(other.AutomaticShaftRegion,record.AutomaticShaftRegion));
                    if(duplicate!=null)
                    {
                        duplicate.AutomaticShaftOrigins.AddRange(record.AutomaticShaftOrigins);
                        continue;
                    }
                    confirmed.Add(record);
                    Notice("AUTO_SHAFT_CONFIRMED","Информация","Граница шахты подтверждена физической геометрией: "+record.AutomaticShaftProof,record);
                }
                return confirmed;
            }

            private static void ValidateConfirmedShaftRegion(Record record,PlanarRegion actual)
            {
                if(record.AutomaticShaftError!=null)
                    throw new UnconfirmedEnvelopeException(record.AutomaticShaftError,actual??record.AutomaticShaftRegion,"Неподтверждённая граница группы «шахты»");
                if(string.IsNullOrWhiteSpace(record.AutomaticShaftProof))
                    throw new UnconfirmedEnvelopeException("Граница шахты ещё не подтверждена физической геометрией в текущем расчёте.",actual??record.AutomaticShaftRegion,"Непроверенная область шахты");
                if(!ShaftMatchesPhysicalVoid(actual,record.AutomaticShaftRegion))
                    throw new UnconfirmedEnvelopeException("Граница сохранённой области шахты отличается от проверенной физической геометрии. Пересчитайте ТЭП для повторной проверки.",actual,"Изменённая область шахты");
            }
        }
    }
}
