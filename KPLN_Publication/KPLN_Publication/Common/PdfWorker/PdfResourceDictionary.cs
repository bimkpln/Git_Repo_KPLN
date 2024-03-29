﻿#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;


namespace KPLN_Publication.PdfWorker
{
    public class PdfResourceDictionary : PdfDictionary
    {

        #region Private variables

        private List<PdfDictionary> _resourceStack = new List<PdfDictionary>();

        #endregion

        #region Push

        public void Push(PdfDictionary resources)
        {
            _resourceStack.Add(resources);
        }

        #endregion

        #region Pop

        public void Pop()
        {
            _resourceStack.RemoveAt(_resourceStack.Count - 1);
        }

        #endregion

        #region GetDirectObject

        public override PdfObject GetDirectObject(PdfName key)
        {
            for (int index = _resourceStack.Count - 1; index >= 0; index--)
            {
                PdfDictionary subResource = _resourceStack[index];

                if (subResource != null)
                {
                    PdfObject obj = subResource.GetDirectObject(key);
                    if (obj != null)
                    {
                        return obj;
                    }
                }
            }
            return base.GetDirectObject(key);
        }

        #endregion

    }
}
