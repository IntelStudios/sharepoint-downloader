using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

namespace SharepointDownloader.Sharepoint
{
    public class Sharepoint365Context
    {
        private readonly string _login;
        private readonly string _password;
        private readonly string _siteUrl;
        private readonly string _rootFolder;
        private readonly int _connectionTimeOut;

        public Sharepoint365Context(Sharepoint365Options options)
        {
            _siteUrl = options.Url;
            _login = options.Login;
            _password = options.Password;
            _rootFolder = options.RootFolder;
            _connectionTimeOut = options.ConnectionTimeout;
        }

        public async Task<List<FileModel>> GetFilesRecursivelyWithoutData()
        {
            using (ClientContext ctx = new ClientContext(_siteUrl))
            {
                //ctx.RequestTimeout = -1;
                ctx.Credentials = new SharePointOnlineCredentials(_login, _password);

                Web web = ctx.Web;
                List list = web.Lists.GetByTitle("Documents");
                ctx.Load(list);
                await ctx.ExecuteQueryAsync();

                Folder folder = web.GetFolderByServerRelativeUrl(_rootFolder);
                ctx.Load(folder);
                await ctx.ExecuteQueryAsync();

                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = @"<View Scope='Recursive'>
                                     <Query>
                                     </Query>
                                 </View>",
                    FolderServerRelativeUrl = folder.ServerRelativeUrl
                };

                ListItemCollection items = list.GetItems(camlQuery);
                ctx.Load(items);
                await ctx.ExecuteQueryAsync();

                List<FileModel> files = new List<FileModel>();

                foreach (var item in items)
                {
                    string path = item.FieldValues["FileRef"].ToString();
                    DateTime created = DateTime.Parse(item.FieldValues["Created"].ToString());
                    DateTime modified = DateTime.Parse(item.FieldValues["Modified"].ToString());
                    string fileName = item.FieldValues["FileLeafRef"].ToString();

                    FileModel fileModel = new FileModel
                    {
                        DateCreated = created,
                        DateModified = modified,
                        FileName = fileName,
                        FileStorageIdentifier = path
                    };

                    files.Add(fileModel);
                }

                return files;
            }
        }

        public async Task LoadFilesData(List<FileModel> filesModels)
        {
            using (ClientContext ctx = new ClientContext(_siteUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(_login, _password);

                Web web = ctx.Web;

                foreach (var fileModel in filesModels)
                {
                    var fileData = web.GetFileByServerRelativeUrl(fileModel.FileStorageIdentifier);
                    if (fileData != null)
                    {
                        ClientResult<Stream> data = fileData.OpenBinaryStream();
                        ctx.Load(fileData);
                        await ctx.ExecuteQueryAsync();

                        using (MemoryStream mStream = new MemoryStream())
                        {
                            if (data != null)
                            {
                                data.Value.CopyTo(mStream);
                                fileModel.FileData = mStream.ToArray();
                            }
                        }
                    }
                }
            }
        }

        public async Task<byte[]> LoadFile(string path)
        {
            using (ClientContext ctx = new ClientContext(_siteUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(_login, _password);

                Web web = ctx.Web;
                var file = web.GetFileByServerRelativeUrl(path);

                if (file != null)
                {
                    ClientResult<Stream> data = file.OpenBinaryStream();
                    ctx.Load(file);
                    await ctx.ExecuteQueryAsync();

                    using (MemoryStream mStream = new MemoryStream())
                    {
                        if (data != null)
                        {
                            data.Value.CopyTo(mStream);
                            return mStream.ToArray();
                        }
                    }
                }

                return null;
            }
        }
    }
}