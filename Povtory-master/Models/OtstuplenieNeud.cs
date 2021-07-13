using System;

namespace Povtory.Models
{
    class OtstuplenieNeud : Otstuplenie
    {
        //Ограничение скорости
        public string SpeedReduction { get; set; }

        //Время ограничения
        public string TimeOfRestriction { get; set; }

		public OtstuplenieNeud()
		{

		}

		public OtstuplenieNeud(int pSnumber, DateTime dateOfInspection, byte distanciaPuti, byte okolotok, string uchastok,
                               string wayNumber, int kmCoord, int pkCoord, int metrCoord, string stepen,
                               string neispravnost, double velichinaNeispravnosti, int dlinaNeispravnosti, string povtory, string vidProverki,
                               string speedReduction, string timeOfRestriction)
                               : base(pSnumber, dateOfInspection, distanciaPuti, okolotok, uchastok, wayNumber,
                                     kmCoord, pkCoord, metrCoord, stepen, neispravnost, velichinaNeispravnosti,
                                     dlinaNeispravnosti, povtory, vidProverki)
        {
            SpeedReduction = speedReduction;
            TimeOfRestriction = timeOfRestriction;
        }
    }
}
