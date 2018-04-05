module.exports = function(context) {
    var spWeb = require('./models/sharepoint');
    
    context.log("function triggered");
    var getWeb = new spWeb();
    getWeb.init(context);
    getWeb.getContext(context).then( ctx => {
        getWeb.getListData(context).then( data => {
            const index = data.length;
            let counter = 1;
            for (let item of data) {
                if (counter == data.length) {
                    getWeb.shareFile(item, ctx).then( result => {
                        context.log("LAST ITEM: successfully shared " + item.name + " with " + item.Sendername);
                        context.bindings.response = { status: 201, body: "Completed successfully" };
                        context.done( null ,{ status: 201, body: "successfully completed" } );
                    }).catch(e => {
                        context.done( e ,{ status: 500, body: e } );
                    });
                } else {
                    getWeb.shareFile(item, ctx).then( result => {
                        context.log("successfully shared " + item.name + " with " + item.Sendername);
                    }).catch(e => {
                        context.done( e ,{ status: 500, body: e } );
                    });
                }
                counter++
            }
        }).catch( e => {
            context.done( e ,{ status: 500, body: e } );
        });;
    }).catch( e => {
        context.done( e ,{ status: 500, body: e } );
    });
};
