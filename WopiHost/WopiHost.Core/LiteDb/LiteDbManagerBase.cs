using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Abstractions;

namespace WopiHost.Core.LiteDb
{
    /*public class LiteDbManagerBase<T> 
    {

        LiteDB.LiteDatabase _context;
        LiteDB.LiteCollection<T> _index { get; set; }
        private readonly IOptionsSnapshot<WopiHostOptions> _wopiHostOptions;

        public LiteDbManagerBase(IOptionsSnapshot<WopiHostOptions> wopiHostOptions)
        {
            
            this._wopiHostOptions = wopiHostOptions;
            _context = new LiteDB.LiteDatabase(this._wopiHostOptions.Value.WebRootPath + "\\LiteDb\\request.db");
            _index = _context.GetCollection<T>(typeof(T).ToString());
        }
        public async Task<bool> InsertRequest(T Item)
        {
            _index.Insert(Item);
            return await Task.FromResult(true);
        }
        public async Task<T> Get(string id,string propname)
        {
            T item = Activator.CreateInstance<T>();
            item.GetType().GetProperty(propname).SetValue(item, id);
            //var fitem = _index.FindOne(x=>x.GetType().GetProperty(propname).GetValue(id, null));
            return await Task.FromResult(item);
        }
        public async Task<bool> Update(T item)
        {
            var output = false;
            _index.Update(item);
            output = true;
            return await Task.FromResult(output);
        }
    }*/
}
