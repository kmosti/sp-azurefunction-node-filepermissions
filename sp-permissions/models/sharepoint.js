var pnp = require('sp-pnp-js');
var NodeFetchClient = require('node-pnp-js').default;
var config = require('../config');
var Q = require('q');
var JsomNode = require('sp-jsom-node').JsomNode;

function spWeb(){

}

module.exports = spWeb;

spWeb.prototype = {
    init: function ( functioncontext ) {
        pnp.setup({
            sp: {
                fetchClientFactory: () => {
                    let credentials =  {
                        username: config.spUserName,
                        password:  config.spPassword
                    };
                    return new NodeFetchClient(credentials);
                }
            }
        });
    },
    getContext: function ( functioncontext ) {
        let deferred = Q.defer();

        let jsomNodeOptions = {
            siteUrl: config.spSite,
            authOptions: {
                username: config.spUserName,
                password: config.spPassword
            },
            config: {
                encryptPassword: false
            },
            envCode: "spo"
        };

        (new JsomNode(jsomNodeOptions)).init();

        let ctx = SP.ClientContext.get_current();

        deferred.resolve(ctx);

        return deferred.promise;
    },
    getListData: function( functioncontext ) {
        var deferred = Q.defer();
        let Web = new pnp.Web(config.spSite);

        Web.getFolderByServerRelativeUrl(config.folderRelativeUrl)
        .files
        .expand('Files/ListItemAllFields')
        .select('Name,MajorVersion,MinorVersion,ServerRelativeUrl')
        .get()
        .then( data => {
            var promises = [];
            let items = [];
            for (f of data) {
                let spObject = {
                    name: f.Name,
                    MajorVersion: f.MajorVersion,
                    MinorVersion: f.MinorVersion,
                    ServerRelativeUrl: f.ServerRelativeUrl
                };

                var itemPromise = Web.getFileByServerRelativeUrl(f.ServerRelativeUrl).getItem().then( item => {
                    spObject.ID = item.ID;
                    spObject.Sendername = item.Sendername;
                    spObject.file = item;
                    items.push(spObject);
                }).catch( e => {
                    itemPromise.reject(e);
                });
                promises.push(itemPromise);
            }
            Q.all(promises).then( p => {
                deferred.resolve(items);
            }).catch( e => {
                deferred.reject(e);
            });
        }).catch( e => {
            deferred.reject(e);
        });
        
        return deferred.promise;
    },
    shareFile: function( item, context, functioncontext ) {
        var deferred = Q.defer();
        
        const web = context.get_web();
        const oList = web.get_lists().getByTitle(config.spList);
        const oUser = web.ensureUser("i:0#.f|membership|" + item.Sendername);
        const oAdmin = web.ensureUser("i:0#.f|membership|" + config.listAdmin);
        const oFile = oList.getItemById(item.ID);
        const oRoles = SP.RoleDefinitionBindingCollection.newObject(context);

        oFile.breakRoleInheritance(false, true);
        oRoles.add(web.get_roleDefinitions().getByType(SP.RoleType.administrator));
        oFile.get_roleAssignments().add(oUser, oRoles);
        oFile.get_roleAssignments().add(oAdmin, oRoles);

        context.load(web); 
        context.load(oUser);
        context.load(oAdmin);
        context.load(oList);
        context.load(oFile);
        context.executeQueryAsync(() => {
            deferred.resolve();
        }, (sender, args) => {
            deferred.reject( args.get_message() );
        });
        return deferred.promise;
    },
    checkinFile: function(itemID) {
        var deferred = Q.defer();
        return deferred.promise;
        //https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Files#check-in-check-out-and-approve--deny
    }
}