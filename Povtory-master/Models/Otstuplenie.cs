using System;

namespace Povtory.Models
{
    abstract class Otstuplenie
    {
        //1
		/// <summary>
		/// 
		/// </summary>
        public int PSnumber { get; set; }
        //2
		/// <summary>
		/// 
		/// </summary>
        public DateTime DateOfInspection { get; set; }
        //3
		/// <summary>
		/// 
		/// </summary>
        public byte DistanciaPuti { get; set; }
        //4
        public byte Okolotok { get; set; }
        //5
        public string Uchastok { get; set; }
        //6
        public string WayNumber { get; set; }
        //7
        public int KmCoord { get; set; }
        //8
        public int PkCoord { get; set; }
        //9
        public int MetrCoord { get; set; }
        //10
        public string Stepen { get; set; }
        //11
        public string Neispravnost { get; set; }
        //12
        public double VelichinaNeispravnosti { get; set; }
        //13
        public int DlinaNeispravnosti { get; set; }

        //14
        public string Povtory { get; set; }
        //15
        public string VidProverki { get; set; }

		public Otstuplenie()
		{
		}

		public Otstuplenie(int pSnumber, DateTime dateOfInspection, byte distanciaPuti, byte okolotok, string uchastok,
                           string wayNumber, int kmCoord, int pkCoord, int metrCoord, string stepen, string neispravnost,
                           double velichinaNeispravnosti, int dlinaNeispravnosti, string povtory, string vidProverki)
        {
            PSnumber = pSnumber;
            DateOfInspection = dateOfInspection;
            DistanciaPuti = distanciaPuti;
            Okolotok = okolotok;
            Uchastok = uchastok;
            WayNumber = wayNumber;
            KmCoord = kmCoord;
            PkCoord = pkCoord;
            MetrCoord = metrCoord;
            Stepen = stepen;
            Neispravnost = neispravnost;
            VelichinaNeispravnosti = velichinaNeispravnosti;
            DlinaNeispravnosti = dlinaNeispravnosti;
            Povtory = povtory;
            VidProverki = vidProverki;
        }
    }
}
