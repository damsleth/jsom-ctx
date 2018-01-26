# jsom-ctx
Simplifies JSOM usage, and allows for async-await.

# Examples
Get web title for current web

```typescript
(async function () {
    try {
        let ctx = new JsomContext(_spPageContextInfo.webAbsoluteUrl)
        ctx = await ctx.load()
        await ExecuteJsomQuery(ctx, [ctx.web, ctx.lists])
        console.log(ctx.web.get_title())
    } catch (err) {
        console.log(err)
    }
})();
````

Get PropertyBag Values  
JsomContext url defaults to ````_spPageContextInfo.webAbsoluteUrl```` (current web url) unless otherwise specified

```typescript
let GetPropertyBagValues = async () => {
        let ctx = new JsomContext()
        ctx = await ctx.load()
        await ExecuteJsomQuery(ctx, [ctx.propBag])
        return ctx.propBag.get_fieldValues()
}

await GetPropertyBagValues();
````
