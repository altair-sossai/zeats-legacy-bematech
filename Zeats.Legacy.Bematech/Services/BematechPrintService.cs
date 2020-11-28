using System;
using Zeats.Bematech.Impressoras;
using Zeats.Legacy.PlainTextTable.Print.Enums;
using Zeats.Legacy.PlainTextTable.Print.Print;

namespace Zeats.Legacy.Bematech.Services
{
    public class BematechPrintService : IPrintService
    {
        private static readonly object Lock = new object();

        public void Print(PrintCollection printCollection)
        {
            lock (Lock)
            {
                var retorno = MP2032.IniciaPorta(printCollection.Options.PortName);
                if (retorno == 0)
                    throw new Exception("Não foi possível estabelecer comunicação com a impressora");

                foreach (var printItem in printCollection)
                {
                    retorno = printItem.FontType == FontType.Text ? PrintText(printItem) : PrintBarCode(printItem);

                    if (retorno == 0)
                        throw new Exception("Ocorreu um erro ao enviar um comando de impressão para a impressora");
                }

                retorno = MP2032.FechaPorta();
                if (retorno == 0)
                    throw new Exception("Não foi possível estabelecer comunicação com a impressora");
            }
        }

        public void Cut(string portName, CutType cutType = CutType.Full)
        {
            lock (Lock)
            {
                MP2032.IniciaPorta(portName);
                MP2032.AcionaGuilhotina(cutType == CutType.Partial ? 0 : 1);
                MP2032.FechaPorta();
            }
        }

        private static int PrintBarCode(PrintItem printItem)
        {
            MP2032.ConfiguraCodigoBarras(100, 2, 2, 1, 20);
            var retorno = MP2032.ImprimeCodigoBarrasEAN13(printItem.Content);
            return retorno;
        }

        private static int PrintText(PrintItem printItem)
        {
            var tipoLetra = printItem.FontSize == FontSize.Small ? 1
                : printItem.FontSize == FontSize.Normal ? 2
                : printItem.FontSize == FontSize.Large ? 3
                : 2;

            var italico = printItem.Italic ? 1 : 0;
            var sublinhado = printItem.Underline ? 1 : 0;
            var expandido = printItem.FontSize == FontSize.Large ? 1 : 0;
            var enfatizado = printItem.Bold ? 1 : 0;

            var retorno = MP2032.FormataTX(printItem.Content, tipoLetra, italico, sublinhado, expandido, enfatizado);
            return retorno;
        }
    }
}