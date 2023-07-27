import * as React from 'react';
import { getSPAdmin, getSP} from '../../PnPJsConfig';
import {
    // IPowerAppsEnvironment,
    // ISPOWebTemplatesInfo,
    // ITenantSitePropertiesInfo,
    PersonalSiteFilter
  } from '@pnp/sp-admin';
import { SPFI } from '@pnp/sp';
import {Table, Loader, CheckmarkCircleIcon, BanIcon, Flex, Button, MoreIcon} from "@fluentui/react-northstar";
import SiteLock from './actions/SiteLock';
import "@pnp/sp/search";
import "@pnp/sp/fields";

import "@pnp/sp/batching";
import SiteLockNoAccess from './actions/SiteLockNoAccess';
import SearchBar from './actions/SearchBar';
import MoreAction from './actions/MoreAction';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface ISitesProps {
  spHttpClient: any;
}
interface IsitesState {
    sites: any[],
    isLoading: boolean,
    isPanelMoreActionOpen: boolean,
    selectedSite: any, 
    homeSite: string,
}

class SPTemplates {
    static readonly POINTPUBLISHING_HUB = "pointpublishinghub#0";
    static readonly SEARCH_CENTER = "srchcen#0";
    static readonly PERSONNAL_SITE = "spsmsitehost#0";
    static readonly APP_CATALOG = "appcatalog#0";
    static readonly COMMUNICATION_SITE = "sitepagepublishing#0";
    static readonly PRIVATE_CHANNEL_SITE = "teamchannel#0";
    static readonly TEAM_SITE_CLASSIC = "sts#0";
    static readonly TEAM_SITE_NO_GROUP = "sts#3";
    static readonly TEAM_SITE = "group#0";
    static readonly REDIRECT_SITE = "redirectsite#0";
  }

const header = {
    key: 'header',
    items: [
        {
            content: 'nom',
            key: 'nom'
        },
        {
            content: 'Url',
            key: 'url'
        },
        {
            content: 'Template',
            key: 'template'
        },
        {
            content: 'Dans un group ?',
            key: 'group'
        },
        {
            content: "Catalogue d'Application",
            key: 'appCat'
        },
        {
            content: 'Actions',
            key: 'action'
        }
    ]
    
}

export default class Sites extends React.Component<ISitesProps, IsitesState> {
    private _spAdmin: SPFI;
    private _sp: SPFI;

    public constructor(props: ISitesProps){
        super(props);

        this.state = {
            sites: [],
            isLoading: true,
            isPanelMoreActionOpen: false,
            selectedSite: null,
            homeSite: "",
        }

        this._spAdmin = getSPAdmin();
        this._sp = getSP();

        this.getSites = this.getSites.bind(this);
        this.getSiteFromList = this.getSiteFromList.bind(this)
    }

    public componentDidMount(): void {
        this.getSites();   
    }

    public render(): React.ReactElement<{}> {
        return (
            <div>   
                {this.state.isLoading ?
                <Loader label={"Chargement"} labelPosition={"end"}></Loader> :
                <div><p>{this.state.homeSite}</p>
                  <Table header={header} rows={this.state.sites}/></div>
                }

              <MoreAction
                spHttpCtx={this.props.spHttpClient}
                isPanelOpen={this.state.isPanelMoreActionOpen}
                isHomeSite={this.state.homeSite === this.state.selectedSite?.Url || this.state.homeSite + "/" === this.state.selectedSite?.Url}
                isCommSite={this.state.selectedSite?.Template.toLowerCase() === SPTemplates.COMMUNICATION_SITE}
                site={this.state.selectedSite}
                closePanel={() => this.setState({ isPanelMoreActionOpen: false, selectedSite: null })}
                refreshHomeSite={this._updateHomeSite} />                  
            </div>
            
        )
    }

    public async getSites(){

      await this._getHomeSite();

      const sites = await this._spAdmin.admin.tenant.getSitePropertiesFromSharePointByFilters({ IncludePersonalSite: PersonalSiteFilter.UseServerDefault, StartIndex: null, IncludeDetail: true, GroupIdDefined: 0 });
   
      const siteListe = await this.getSiteFromList()

      

      siteListe.forEach(siteL=> {
        sites.forEach(site => {
          if(site.Title==siteL.Title){
            if(siteL.DisableFlows!= null){site.DisableFlows = siteL.DisableFlows;}
            if(siteL.CommentsOnSitePagesDisabled!= null){site.CommentsOnSitePagesDisabled = siteL.CommentsOnSitePagesDisabled;}
            
            
          } 
        })
      })

      console.log(sites);
      const siteRow = sites.map((site, index:number) => {

      const hasGroup: boolean = site.GroupId !== "00000000-0000-0000-0000-000000000000";
      return {
        key:index+1,
        items:[
          {
              content: site.Title
          },
          {
              content: site.Url,
              truncateContent: true
          },
          {
              content: this._getTemplateDisplayName(site.Template)
          },
          {
              content: hasGroup? <CheckmarkCircleIcon/> : <BanIcon/>,
              styles: { left: "20px" }
          },
          {
              content: ""
          },
          {
            content: (
            <Flex gap='gap.small' vAlign='center'>
              <SiteLock etatActuel={site.LockState} siteName={site.Title} disabled={this._isSiteSpecial(site.Template)}/>
              <SiteLockNoAccess site={site} disabled={this._isSiteSpecial(site.Template)}></SiteLockNoAccess>
              <SearchBar disabled={this._isSiteSpecial(site.Template)} siteUrl={site.Url} />
              <Button
                  icon={<MoreIcon />}
                  iconOnly
                  onClick={()=> {
                    this.setState({selectedSite: site, isPanelMoreActionOpen: true});
                    console.log(site);
                  }}/>
            </Flex>),
            styles: { position: "relative", right: "5%" }
              
          }
        ]
      }
      })
      this.setState({
      sites: siteRow,
      isLoading: false
      });

    }

    private async getSiteFromList() : Promise<any[]> {
      const res = await this._sp.web.lists.ensure("IDSite");

      const [batchedWeb, execute] = this._sp.web.batched();
      batchedWeb.lists.getByTitle("IDSite").fields.addText("IDsite", { MaxLength: 255});
      batchedWeb.lists.getByTitle("IDSite").fields.addText("GroupID", { MaxLength: 255});
      batchedWeb.lists.getByTitle("IDSite").fields.addText("Url", { MaxLength: 255});
      batchedWeb.lists.getByTitle("IDSite").fields.addText("Template", { MaxLength: 255});
      batchedWeb.lists.getByTitle("IDSite").fields.addText("LockState", { MaxLength: 255});
      batchedWeb.lists.getByTitle("IDSite").fields.addBoolean("CommentsOnSitePagesDisabled");
      batchedWeb.lists.getByTitle("IDSite").fields.addNumber("DisableFlows");

      if(res.created){
        await execute();
        return [];
      }else{
        let returnarray:any[] = []
        const items: any[] = await this._sp.web.lists.getByTitle("IDSite").items();
        items.forEach(item => {
          returnarray.push({
            Title: item.Title,
            Url: item.Url,
            Template: item.Template,
            LockState: item.LockState,
            GroupId: item.GroupID,
            DisableFlows: item.DisableFlows,
            CommentsOnSitePagesDisabled: item.CommentsOnSitePagesDisabled
          })
        })
        return returnarray; 
      } 
    }

    private _getHomeSite = async (): Promise<void> => {
      const getURL: string = (await getSP().site()).Url + "/_api/SPHSite/Details/";
  
      const currentHomeSite: any = await this.props.spHttpClient.get(
        getURL,
        SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse): Promise<SPHttpClientResponse> => {
          return response.json();
        })
        .catch((err: any) => { throw err });
  
      console.log(currentHomeSite);
  
      this.setState({
        homeSite: currentHomeSite.Url, //+ "/"
      });
    }
    private _updateHomeSite = (siteUrl: string): void => {
      this.setState({
        homeSite: siteUrl,
      });
    }

    private _getTemplateDisplayName = (templateId: string): string => {
        let templateName: string;
        switch (templateId.toLowerCase()) {
          case SPTemplates.TEAM_SITE_NO_GROUP:
          case SPTemplates.TEAM_SITE:
          case SPTemplates.TEAM_SITE_CLASSIC: {
            templateName = "Team Site" + (templateId.toLowerCase() === SPTemplates.TEAM_SITE_CLASSIC ? " (classic experience)" : "");
            break;
          }
    
          case SPTemplates.POINTPUBLISHING_HUB: {
            templateName = "PointPublishing Hub";
            break;
          }
    
          case SPTemplates.COMMUNICATION_SITE: {
            templateName = "Communication Site";
            break;
          }
    
          case SPTemplates.PRIVATE_CHANNEL_SITE: {
            templateName = "Private Channel Site";
            break;
          }
    
          case SPTemplates.SEARCH_CENTER: {
            templateName = "Enterprise Search Center";
            break;
          }
    
          case SPTemplates.PERSONNAL_SITE: {
            templateName = "Personnal Site";
            break;
          }
    
          case SPTemplates.APP_CATALOG: {
            templateName = "Tenant App Catalog";
            break;
          }
    
          default: {
            templateName = templateId;
            break;
          }
        }
    
        return templateName;
      }

      private _isSiteSpecial = (templateId: string): boolean => {
        return [
          SPTemplates.SEARCH_CENTER,
          SPTemplates.PERSONNAL_SITE,
          SPTemplates.POINTPUBLISHING_HUB,
          SPTemplates.APP_CATALOG,
          SPTemplates.REDIRECT_SITE
        ].some(val => val === templateId.toLowerCase());
      }

      
}