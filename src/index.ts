export interface IJsomContext {
    url: string;
    clientContext: SP.ClientContext;
    web: SP.Web;
    lists: SP.ListCollection;
    propBag: SP.FieldStringValues;
}

const __getClientContext = (url: string) => {
    return new Promise<SP.ClientContext>((resolve, reject) => {
        SP.SOD.executeFunc("sp.js", "SP.ClientContext", () => {
            const clientContext = new SP.ClientContext(url);
            resolve(clientContext);
        }), reject(`Unable to resolve clientContext`);
    })
}

export class JsomContext implements IJsomContext {
    public url: string;
    public clientContext: SP.ClientContext;
    public web: SP.Web;
    public site: SP.Site;
    public rootWeb: SP.Web;
    public lists: SP.ListCollection;
    public propBag: SP.FieldStringValues;

    constructor(url?: string) {
        this.url = url || _spPageContextInfo.webAbsoluteUrl;
    }

    public async load(): Promise<JsomContext> {
        try {
            this.clientContext = await __getClientContext(this.url);
            this.web = this.clientContext.get_web();
            this.site = this.clientContext.get_site();
            this.rootWeb = this.clientContext.get_site().get_rootWeb();
            this.lists = this.web.get_lists();
            this.propBag = this.web.get_allProperties();
            return this;
        } catch (err) {
            throw `Failed to load context for url ${this.url}`;
        }
    }
}

export const CreateJsomContext = async (url: string): Promise<JsomContext> => await new JsomContext(url).load()

export const ExecuteJsomQuery = async (ctx: JsomContext, clientObjectsToLoad: SP.ClientObject[] = []) => {
    return new Promise<{ sender, args }>((resolve, reject) => {
        clientObjectsToLoad.forEach(clientObj => ctx.clientContext.load(clientObj));
        ctx.clientContext.executeQueryAsync((sender, args) => {
            resolve({ sender, args });
        }, (sender, args) => {
            reject({ sender, args });
        })
    });
}
