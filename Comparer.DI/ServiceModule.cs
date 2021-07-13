using Comparer.ServiceOX;
using ExcelComparer.ISvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Comparer.DI
{
	public class ServiceModule
	{
		//Lazy<ServiceModule> lazy 
		public static ServiceModule Instance()
		{
			return new ServiceModule();
		}

		IComparerSvc comparerSvc;
		public IComparerSvc ComparerSvc
		{
			get
			{
				if (comparerSvc == null)
					comparerSvc = new OpenXmlService();
				return comparerSvc;
			}
		}

	}
}
