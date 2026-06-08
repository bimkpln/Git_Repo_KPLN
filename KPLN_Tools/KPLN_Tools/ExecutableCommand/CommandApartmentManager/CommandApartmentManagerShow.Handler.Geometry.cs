using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private static bool TryDeleteSource2DApartmentInstance(Document doc, ElementId apartmentId, List<string> debugMessages)
        {
            if (doc == null || apartmentId == null || apartmentId == ElementId.InvalidElementId)
                return false;

            try
            {
                Element element = doc.GetElement(apartmentId);
                if (element == null)
                    return false;

                doc.Delete(apartmentId);
                return true;
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось удалить 2D-квартиру ID = " + IDHelper.ElIdValue(apartmentId) + ": " + ex.Message);

                return false;
            }
        }

        private static bool CanFlipHand(FamilyInstance fi)
        {
            if (fi == null)
                return false;

            try
            {
                return fi.CanFlipHand;
            }
            catch
            {
                return false;
            }
        }

        private static bool CanFlipFacing(FamilyInstance fi)
        {
            if (fi == null)
                return false;

            try
            {
                return fi.CanFlipFacing;
            }
            catch
            {
                return false;
            }
        }

        private static XYZ GetFamilyInstanceHandDirection2D(FamilyInstance fi, Transform fallbackTransform)
        {
            if (fi != null)
            {
                try
                {
                    XYZ hand = fi.HandOrientation;
                    if (hand != null)
                    {
                        XYZ normalized = Normalize2D(hand);
                        if (normalized != null)
                            return normalized;
                    }
                }
                catch
                {
                }
            }

            if (fallbackTransform == null && fi != null)
            {
                try
                {
                    fallbackTransform = fi.GetTransform();
                }
                catch
                {
                }
            }

            if (fallbackTransform != null)
            {
                XYZ basisX = fallbackTransform.BasisX;
                if (basisX != null)
                    return Normalize2D(basisX);
            }

            return null;
        }

        private static XYZ GetFamilyInstanceFacingDirection2D(FamilyInstance fi, Transform fallbackTransform)
        {
            if (fi != null)
            {
                try
                {
                    XYZ facing = fi.FacingOrientation;
                    if (facing != null)
                    {
                        XYZ normalized = Normalize2D(facing);
                        if (normalized != null)
                            return normalized;
                    }
                }
                catch
                {
                }
            }

            if (fallbackTransform == null && fi != null)
            {
                try
                {
                    fallbackTransform = fi.GetTransform();
                }
                catch
                {
                }
            }

            if (fallbackTransform != null)
            {
                XYZ basisY = fallbackTransform.BasisY;
                if (basisY != null)
                    return Normalize2D(basisY);
            }

            return null;
        }

        private static string FormatPointMm(XYZ point)
        {
            if (point == null)
                return "<нет>";

            return "(" +
                   Math.Round(IDHelper.ConvertInternalToMm(point.X)).ToString("0") + "; " +
                   Math.Round(IDHelper.ConvertInternalToMm(point.Y)).ToString("0") + "; " +
                   Math.Round(IDHelper.ConvertInternalToMm(point.Z)).ToString("0") + ") мм";
        }

        private static string FormatLengthMm(double valueInternal)
        {
            return Math.Round(IDHelper.ConvertInternalToMm(valueInternal)).ToString("0") + " мм";
        }

        private static string FormatDouble(double value)
        {
            if (double.IsNaN(value))
                return "<nan>";

            if (double.IsNegativeInfinity(value) || double.IsPositiveInfinity(value))
                return "<нет>";

            return Math.Round(value, 3).ToString("0.###");
        }

        private static void AddApartmentDiagnostics(ApartmentProcessState state, List<string> debugMessages, IEnumerable<string> messages)
        {
            if (messages == null)
                return;

            foreach (string message in messages)
                AddApartmentDiagnostic(state, debugMessages, message);
        }

        private static string FormatVector2D(XYZ vector)
        {
            if (vector == null)
                return "<нет>";

            return "(" +
                   Math.Round(vector.X, 3).ToString("0.###") + "; " +
                   Math.Round(vector.Y, 3).ToString("0.###") + ")";
        }

        private static string FormatElementIdForDiagnostic(ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId)
                return "<из семейства>";

            return IDHelper.ElIdValue(id).ToString();
        }

        private static XYZ GetApartmentInteriorReferencePoint(FamilyInstance apartmentFi, Document doc)
        {
            List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartmentFi);
            if (rooms != null && rooms.Count > 0)
            {
                double x = 0;
                double y = 0;
                double z = 0;
                int count = 0;

                foreach (FamilyInstance roomFi in rooms)
                {
                    XYZ center = GetRoomCenterPoint(roomFi);
                    if (center == null)
                        continue;

                    x += center.X;
                    y += center.Y;
                    z += center.Z;
                    count++;
                }

                if (count > 0)
                    return new XYZ(x / count, y / count, z / count);
            }

            Transform tr = apartmentFi != null ? apartmentFi.GetTransform() : null;
            return tr != null ? tr.Origin : null;
        }

        private static XYZ Normalize2D(XYZ v)
        {
            if (v == null)
                return null;

            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-12)
                return null;

            return new XYZ(v.X / len, v.Y / len, 0);
        }

        private static XYZ CanonicalizeDirection(XYZ dir)
        {
            if (dir == null)
                return null;

            if (dir.X < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            if (Math.Abs(dir.X) < 1e-9 && dir.Y < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            return new XYZ(dir.X, dir.Y, 0);
        }

        private static double Dot2D(XYZ a, XYZ b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static bool PointOnSegment2D(XYZ p, XYZ a, XYZ b, double tol)
        {
            double ab = Distance2D(a, b);
            double ap = Distance2D(a, p);
            double pb = Distance2D(p, b);
            return Math.Abs((ap + pb) - ab) <= tol;
        }

        private static bool AreCollinear2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            if (Math.Abs(Cross2D(ad, bd)) > tol)
                return false;

            if (Math.Abs(Cross2D(b0 - a0, ad)) > tol)
                return false;

            return true;
        }

        private static bool TryIntersectInfiniteLines2D(XYZ p1, XYZ d1, XYZ p2, XYZ d2, out XYZ intersection)
        {
            const double tol = 1e-12;
            intersection = null;

            double cross = Cross2D(d1, d2);
            if (Math.Abs(cross) < tol)
                return false;

            XYZ delta = p2 - p1;
            double t = Cross2D(delta, d2) / cross;

            intersection = new XYZ(
                p1.X + d1.X * t,
                p1.Y + d1.Y * t,
                0);

            return true;
        }

        private static bool TryIntersectSegments2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, out XYZ intersection, double tol)
        {
            intersection = null;

            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            XYZ inter;
            if (!TryIntersectInfiniteLines2D(a0, ad, b0, bd, out inter))
                return false;

            if (!PointOnSegment2D(inter, a0, a1, tol))
                return false;

            if (!PointOnSegment2D(inter, b0, b1, tol))
                return false;

            intersection = inter;
            return true;
        }

    }
}
