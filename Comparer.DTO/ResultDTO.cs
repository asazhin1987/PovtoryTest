using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Comparer.DTO
{
	public class ResultDTO
	{

	}

	public class OrderDTO
	{
		public string userName;
	}



	/**/
	//service
	public class sss
	{
		Context db;
		OrderDTO aaaa()
		{
			var dbEml = db.GetEmployee(1);
			return new OrderDTO
			{
				userName = $"{dbEml.Name} {dbEml.Sername}"
			};
		}
	}

	/**/

	public class Context
	{
		public class Order
		{
			public int EmployeeId;
		}

		public class Employeer
		{
			public string Name;

			public string Sername;
		}


		public Employeer GetEmployee(int Id)
		{
			return new Employeer()
			{
				Name = "NNN", Sername = "SSSSS"
			};
		}
	}

	
}
