import { Table } from "@fluentui/react-northstar";
import { SPFI } from "@pnp/sp";
import * as React from "react";
import { getSP } from "../../PnPJsConfig";
import { IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp/hubsites";


interface IHubSitesProps {

}
interface IHubSitesState {
    sites: any[];
}
const header = {
    key: 'header',
    items: [
        {
            content: 'nom',
            key: 'nom'
        }
    ]
}
export default class HubSites extends React.Component<IHubSitesProps, IHubSitesState> {
    private _sp: SPFI;

    public constructor(props: IHubSitesProps){
        super(props)
        this.state = {
            sites: null
        }
        this._sp = getSP()
    }

    public render(): React.ReactElement<{}> {
        return (
            <div>
                <Table header={header} rows={this.state.sites}/>
            </div>
        )
    }

    public componentDidMount(): void {
        this.getHubSites();   
    }

    public async getHubSites(): Promise<void> {
        const hubsites: IHubSiteInfo[] = await this._sp.hubSites();
        console.log(hubsites);
        
        const siteRow = hubsites.map((site, index:number)=> {
            return {
                key:index+1,
                items:[
                  {
                      content: site.Title
                  }]
            }
        })
        this.setState({sites: siteRow})

    }
}