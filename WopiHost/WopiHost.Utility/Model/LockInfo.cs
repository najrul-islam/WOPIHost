using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.Model
{
	public class LockInfo
	{
		public string Lock { get; set; }

		public DateTime DateCreated { get; set; }

		public bool Expired { get { return DateCreated.AddHours(24) < DateTime.UtcNow; } }

		//public bool Expired { get { return DateCreated.AddSeconds(5) < DateTime.UtcNow; } }

	}
}
