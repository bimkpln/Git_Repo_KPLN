#region License
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

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace KPLN_Publication.PdfWorker
{
    public class PdfContentModifier
    {

        #region Private variables

        private Dictionary<string, PdfContentOperatorHandler> _operators = new Dictionary<string, PdfContentOperatorHandler>();
        private Stack<PdfContentStreamBuilder> _contentStreamBuilderStack = new Stack<PdfContentStreamBuilder>();
        private PdfResourceDictionary _resourceDictionaryStack = new PdfResourceDictionary();
        private Func<string, List<PdfObject>, bool> _operatorFilter;

        #endregion

        #region Modify

        public void Modify(PdfReader reader, int pageNumber)
        {
            Modify(reader, pageNumber, null);
        }

        public void Modify(PdfReader reader, int pageNumber, Func<string, List<PdfObject>, bool> operatorFilter)
        {
            PdfDictionary pageDictionary = reader.GetPageN(pageNumber);
            PdfDictionary resourcesDictionary = pageDictionary.GetAsDict(PdfName.RESOURCES);

            byte[] contentBytes = reader.GetPageContent(pageNumber);

            // Ограничение действует и во вложенных Form XObject, обрабатываемых тем же модификатором.
            Func<string, List<PdfObject>, bool> previousFilter = _operatorFilter;
            _operatorFilter = operatorFilter;
            try
            {
                contentBytes = this.Modify(contentBytes, resourcesDictionary);
                reader.SetPageContent(pageNumber, contentBytes);
            }
            finally
            {
                _operatorFilter = previousFilter;
            }
        }

        #endregion

        #region Modify

        public byte[] Modify(byte[] contentBytes, PdfDictionary resourcesDictionary)
        {
            _contentStreamBuilderStack.Push(new PdfContentStreamBuilder());
            _resourceDictionaryStack.Push(resourcesDictionary);
            PRTokeniser tokeniser = new PRTokeniser(new RandomAccessFileOrArray(contentBytes));
            PdfContentParser ps = new PdfContentParser(tokeniser);

            List<PdfObject> operands = new List<PdfObject>();

            try
            {
                while (true)
                {
                    long operationStart = tokeniser.FilePointer;
                    if (ps.Parse(operands).Count == 0)
                        break;

                    PdfLiteral oper = (PdfLiteral)operands[operands.Count - 1];

                    // После BI идут бинарные данные изображения, а не обычные PDF-операторы.
                    // Штатный парсер определяет конец изображения; исходные байты сохраняем без изменений.
                    if (oper.ToString() == "BI")
                    {
                        InlineImageUtils.ParseInlineImage(ps, _resourceDictionaryStack.GetAsDict(PdfName.COLORSPACE));
                        int imageLength = checked((int)(tokeniser.FilePointer - operationStart));
                        byte[] imageContent = new byte[imageLength];
                        Buffer.BlockCopy(contentBytes, checked((int)operationStart), imageContent, 0, imageLength);
                        _contentStreamBuilderStack.Peek().Push(new PdfLiteral(imageContent));
                        continue;
                    }

                    //System.Diagnostics.Debug.WriteLine("[Debug] Opr: " + oper.ToString());

                    PdfContentOperatorHandler operHandler = null;

                    if (_operators.TryGetValue(oper.ToString(), out operHandler))
                    {
                        try
                        {
                            if (_operatorFilter == null || _operatorFilter(oper.ToString(), operands))
                                operands = operHandler(oper, operands);
                        }
                        catch (System.Runtime.InteropServices.ExternalException)
                        {
                            // Игнорирую элементы на листе, которые невозможно перекрасить (например печать организации)
                        }
                        catch (Exception)
                        {
                            // Игнорирую неподдерживаемые или нестандартные элементы (например облака точек в связях),
                            // чтобы преобразование в монохром продолжалось для остального содержимого листа.
                        }
                    }

                    _contentStreamBuilderStack.Peek().Push(operands);
                }

                return _contentStreamBuilderStack.Peek().GetBytes();
            }
            finally
            {
                // Ошибка во вложенном Form XObject не должна оставлять его состояние в стеке страницы.
                _resourceDictionaryStack.Pop();
                _contentStreamBuilderStack.Pop();
                tokeniser.Close();
            }
        }

        #endregion

        #region RegisterOperator

        public void RegisterOperator(string oper, PdfContentOperatorHandler callback)
        {
            _operators.Add(oper, callback);
        }

        #endregion

        #region ResourceDictionary

        public PdfResourceDictionary ResourceDictionary
        {
            get { return _resourceDictionaryStack; }
        }

        #endregion

    }
}
