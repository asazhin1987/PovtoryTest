using System;

namespace Povtory.Models
{
    class OtstuplenueObychnoe : Otstuplenie
    {
        public int Shtuk { get; set; }

        public OtstuplenueObychnoe(int pSnumber, DateTime dateOfInspection, byte distanciaPuti,
                               byte okolotok, string uchastok, string wayNumber, int kmCoord,
                               int pkCoord, int metrCoord, string stepen, string neispravnost,
                               double velichinaNeispravnosti, int dlinaNeispravnosti, int shtuk,
                               string povtory, string vidProverki)
                               : base(pSnumber, dateOfInspection, distanciaPuti, okolotok, uchastok, wayNumber,
                                     kmCoord, pkCoord, metrCoord, stepen, neispravnost, velichinaNeispravnosti,
                                     dlinaNeispravnosti, povtory, vidProverki)
        {
            Shtuk = shtuk;
        }
    }
}
