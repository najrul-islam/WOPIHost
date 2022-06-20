using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Data;
using WopiHost.Data.Models;

namespace WopiHost.Service.UserService
{
    public interface IUserService
    {
        //User
        Task<User> GetUser(int userId);
        Task<User> GetUser(string email);
        Task<User> AddUser(User user);
        Task<IQueryable<User>> GetLoginInfo(string email, string password);

        //Document
        Task<Data.Models.Document> GetDocument(int documentId);
        Task<IEnumerable<Data.Models.Document>> GetDocumentList();
        Task<Data.Models.Document> AddDocument(Data.Models.Document document);
    }
    public class UserService : IUserService
    {
        private readonly WopiHostContext _wopiHostContext;
        public UserService(WopiHostContext wopiHostContext)
        {
            _wopiHostContext = wopiHostContext;
        }
        #region User

        public async Task<User> GetUser(int userId)
        {
            User user;
            user = _wopiHostContext.User.FirstOrDefault(u => u.UserId == userId);
            return await Task.FromResult(user);
        }
        public async Task<User> GetUser(string email)
        {
            User user;
            user = _wopiHostContext.User.FirstOrDefault(u => u.Email == email);
            return await Task.FromResult(user);
        }
        public async Task<User> AddUser(User user)
        {
            if (!string.IsNullOrEmpty(user.Email))
            {
                User existingUser = await GetUser(user.Email);
                if (existingUser == null)
                {
                    user.UserName = user.Email;
                    user.CreatedDate = DateTime.UtcNow;
                    _wopiHostContext.User.Add(user);
                    await _wopiHostContext.SaveChangesAsync();
                }
            }
            return await Task.FromResult(user);
        }
        public async Task<IQueryable<User>> GetLoginInfo(string email, string password)
        {
            var user = _wopiHostContext.User.Where(u => u.Email == email && u.Password == password).AsNoTrackingWithIdentityResolution();
            return await Task.FromResult(user);
        }
        #endregion

        #region Document

        public async Task<Data.Models.Document> GetDocument(int documentId)
        {
            var document = _wopiHostContext.Document.FirstOrDefault(x => x.DocumentId == documentId);
            return await Task.FromResult(document);
        }
        public async Task<IEnumerable<Data.Models.Document>> GetDocumentList()
        {
            var document = _wopiHostContext.Document.AsQueryable().AsNoTracking();
            return await Task.FromResult(document);
        }
        public async Task<Data.Models.Document> AddDocument(Data.Models.Document document)
        {
            if (document != null && !string.IsNullOrEmpty(document?.DocumentName))
            {
                _wopiHostContext.Document.Add(document);
                await _wopiHostContext.SaveChangesAsync();
            }
            return await Task.FromResult(document);
        }
        #endregion

    }
}
