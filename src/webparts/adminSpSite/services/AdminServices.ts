// import {  Caching } from "@pnp/queryable";
import { SPFI, spfi} from '@pnp/sp';
import { CopyFrom } from "@pnp/core";
import { getSP, getSPAdmin } from "../../PnPJsConfig";
// import { ISearchBuilder, SearchQueryBuilder, SearchResults } from "@pnp/sp/search";
// import Constants from "../../Constants";
import "@pnp/sp/sites";
import Constants from '../../Constants';
// import {  Caching } from "@pnp/queryable";

// import { CopyFrom } from "@pnp/core";
import { InjectHeaders } from "@pnp/queryable";
// import {  SearchResults } from "@pnp/sp/search";

import { parseString } from "xml2js";
import "@pnp/sp/items/get-all";

export default class AdminServices {
  
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public static async UpdateSiteProperties(nomSite:string, updatedProperties: Record<string, any>): Promise<void> {
    const spAdmin: SPFI = getSPAdmin();

    const id = await this.getId(nomSite)
    console.log("mon id ##"+id+"##");
    
  
    await spfi(Constants.TENANT_ADMIN_URL).using(CopyFrom(spAdmin.site)).using(InjectHeaders({ 
      "Accept": "application/json;odata=verbose", 
      "Content-Type": "application/json;odata=verbose",
      "X-HTTP-Method": "MERGE"
    })).admin.tenant.call<void>("Sites('"+id+"')", {
    // })).admin.tenant.call<void>("Sites('06fb41ec-2922-4be5-94f4-6cb32f843636')", {
      "__metadata": {
        "type": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties"
      },
      ...updatedProperties
    });
  }

  public static async getId(siteName: string): Promise<String> {
    const _sp: SPFI = getSP();
    
    const urlName = siteName.replace(/\s/g, '');
  
    var id;
    try{
      await fetch(Constants.TENANT_URL+"/sites/"+urlName+"/_api/site/id")
        .then(res => res.text())
        .then(str => parseString(str, function(err,res){
          if(err)console.log("erreur fetch url -> parsage de l'XML");
          else id = res['d:Id']['_']; 
        }))
    }catch (e){
      const allItem = await _sp.web.lists.getByTitle("IDSite").items.getAll();
      allItem.forEach(item => {if(item.Title == siteName){ id = item.IDsite}})
    }
    finally {
      return id
    }
  }

  public static async ajoutItemList(params:any): Promise<void>{
    const _sp: SPFI = getSP();

    const res = await _sp.web.lists.getByTitle("IDSite").items()
    
    const res2 = res.filter(e => {return e.Title==params.Title})

    if (res2.length==0){
      await _sp.web.lists.getByTitle("IDSite").items.add(
        params,
      );
    }else {
      console.log("update with ");
      console.log(params);
      await _sp.web.lists.getByTitle("IDSite").items.getById(res2[0].Id).update(params)
      
    }
  }
}


